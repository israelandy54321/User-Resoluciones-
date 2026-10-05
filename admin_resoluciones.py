import sys
import os
import tempfile
import base64
import requests
import hashlib
import json
import re
import subprocess
import shutil
import zipfile
import pymupdf

from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path

# Firma digital PDF
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509.oid import NameOID
import qrcode
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

try:
    from pyhanko import stamp
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign import fields, signers
except ImportError as e:
    raise SystemExit(
        "Falta pyHanko para la firma digital.\n\n"
        "Ejecuta:\n"
        "python -m pip install pyHanko pywin32 reportlab "
        "'qrcode[pil]' cryptography pypdf\n\n"
        f"Detalle: {e}"
    )

from PySide6.QtCore import (
    Qt,
    QTimer,
    QDate,
    QObject,
    QSize,
    QRunnable,
    QThreadPool,
    Signal
)

from PySide6.QtGui import (
    QPixmap,
    QImage,
    QFont
)

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QDateEdit,
    QListWidget,
    QCheckBox,
    QListWidgetItem,
    QFileDialog,
    QMessageBox,
    QScrollArea,
    QFrame,
    QSizePolicy,
    QSplitter,
    QProgressBar,
    QDialog,
    QInputDialog
)


# =========================================================
# CONFIGURACIÓN
# =========================================================

GOOGLE_SHEETS_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbzjphcwy8Dpyc_OoSG3geYjpyrGtKDJY-FpaPyf7lVJm7rZ7XDAy2zzfcc_oup70G7s"
    "/exec"
)


# =========================================================
# ESTADOS DE RESOLUCIONES
# =========================================================

ESTADO_PENDIENTE = "PENDIENTE"
ESTADO_ENVIADO = "ENVIADO"
ESTADO_RECHAZADO = "RECHAZADO"
ESTADOS_VALIDOS = (
    ESTADO_PENDIENTE,
    ESTADO_ENVIADO,
    ESTADO_RECHAZADO
)


# =========================================================
# CORREO Y PLANTILLA EXCEL
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    RECURSOS_DIR = sys._MEIPASS
else:
    RECURSOS_DIR = BASE_DIR


# =========================================================
# LOGO DEL SISTEMA
# =========================================================
# Busca automáticamente el logo dentro de la carpeta de
# recursos del programa. Se prueban nombres habituales y,
# si no se encuentra, se revisan imágenes cuyo nombre
# contenga "logo", "isra" o "solutions".
def buscar_logo_sistema():
    nombres_preferidos = [
        "isra.png",
    ]

    for nombre in nombres_preferidos:
        ruta = os.path.join(RECURSOS_DIR, nombre)
        if os.path.isfile(ruta):
            return ruta

    try:
        extensiones = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".ico")

        for nombre in os.listdir(RECURSOS_DIR):
            nombre_minuscula = nombre.lower()

            if not nombre_minuscula.endswith(extensiones):
                continue

            if any(
                palabra in nombre_minuscula
                for palabra in ("logo", "isra", "solutions")
            ):
                ruta = os.path.join(RECURSOS_DIR, nombre)

                if os.path.isfile(ruta):
                    return ruta

    except Exception:
        pass

    return None


LOGO_SISTEMA = buscar_logo_sistema()


PLANTILLA_EXCEL = os.path.join(
    RECURSOS_DIR,
    "plantilla_AMT.xlsx"
)

CORREO_DESTINO_PRUEBAS = "aiy.delacruz@yavirac.edu.ec"

CORREO_REMITENTE_OUTLOOK = "gestion.movilidad@outlook.com"

ASUNTO_CORREO = "Documentación AMT - {nombre}"

CUERPO_CORREO = """Estimado/a {nombre}:

Por medio del presente se remite el formato de verificación documental correspondiente,
la cual deberá ser completada con la información solicitada.

Una vez completada, por favor remitirla por el medio indicado.

Saludos cordiales,
Administrador
Sistema de Resoluciones AMT
"""


# =========================================================
# MICROSOFT GRAPH
# =========================================================

MICROSOFT_CLIENT_ID = "b07fecb9-0133-4f67-88ca-908968a8154d"

GRAPH_AUTHORITY = "https://login.microsoftonline.com/consumers"

GRAPH_SCOPES = [
    "Mail.Send",
    "User.Read",
]

GRAPH_SEND_URL = (
    "https://graph.microsoft.com/v1.0/me/sendMail"
)

GRAPH_ME_URL = (
    "https://graph.microsoft.com/v1.0/me"
)

GRAPH_CACHE_DIR = os.path.join(
    os.environ.get(
        "LOCALAPPDATA",
        BASE_DIR
    ),
    "SistemaResolucionesAMT"
)

GRAPH_CACHE_FILE = os.path.join(
    GRAPH_CACHE_DIR,
    "msal_token_cache.bin"
)


# =========================================================
# CERTIFICADO DE FIRMA GUARDADO
# =========================================================
# Solo se guarda la RUTA del certificado.
# La contraseña nunca se guarda.
FIRMA_CERT_CONFIG_FILE = os.path.join(
    GRAPH_CACHE_DIR,
    "certificado_firma.json"
)


def cargar_ruta_certificado_guardada():
    try:
        if not os.path.isfile(FIRMA_CERT_CONFIG_FILE):
            return ""

        with open(
            FIRMA_CERT_CONFIG_FILE,
            "r",
            encoding="utf-8"
        ) as archivo:
            datos = json.load(archivo)

        ruta = str(
            datos.get("p12_path", "")
        ).strip()

        if ruta and os.path.isfile(ruta):
            return ruta

    except Exception:
        pass

    return ""


def guardar_ruta_certificado(p12_path):
    os.makedirs(
        GRAPH_CACHE_DIR,
        exist_ok=True
    )

    with open(
        FIRMA_CERT_CONFIG_FILE,
        "w",
        encoding="utf-8"
    ) as archivo:
        json.dump(
            {
                "p12_path": str(p12_path)
            },
            archivo,
            ensure_ascii=False,
            indent=2
        )


# =========================================================
# OBTENER TOKEN MICROSOFT
# =========================================================

def obtener_token_microsoft_graph():

    if (
        not MICROSOFT_CLIENT_ID
        or MICROSOFT_CLIENT_ID == "PON_AQUI_TU_CLIENT_ID"
    ):

        raise RuntimeError(
            "Falta configurar MICROSOFT_CLIENT_ID.\n\n"
            "Debes crear una aplicación en Microsoft Entra/Azure, "
            "agregar los permisos Mail.Send y User.Read, habilitar "
            "cuentas personales de Microsoft y colocar aquí el Client ID."
        )

    try:

        import msal

    except ImportError:

        raise RuntimeError(
            "Falta instalar MSAL para Microsoft Graph.\n\n"
            "Ejecuta en la terminal:\n\n"
            "pip install msal"
        )

    os.makedirs(
        GRAPH_CACHE_DIR,
        exist_ok=True
    )

    cache = msal.SerializableTokenCache()

    if os.path.isfile(GRAPH_CACHE_FILE):

        try:

            with open(
                GRAPH_CACHE_FILE,
                "r",
                encoding="utf-8"
            ) as archivo_cache:

                cache.deserialize(
                    archivo_cache.read()
                )

        except Exception:

            pass

    app = msal.PublicClientApplication(
        MICROSOFT_CLIENT_ID,
        authority=GRAPH_AUTHORITY,
        token_cache=cache
    )

    cuentas = app.get_accounts(
        username=CORREO_REMITENTE_OUTLOOK
    )

    resultado = None

    if cuentas:

        resultado = app.acquire_token_silent(
            GRAPH_SCOPES,
            account=cuentas[0]
        )

    if (
        not resultado
        or "access_token" not in resultado
    ):

        resultado = app.acquire_token_interactive(
            scopes=GRAPH_SCOPES,
            prompt="select_account"
        )

    if (
        not resultado
        or "access_token" not in resultado
    ):

        detalle = (
            resultado.get(
                "error_description",
                "No se recibió un token."
            )
            if resultado
            else
            "No se recibió respuesta de Microsoft."
        )

        raise RuntimeError(
            "No se pudo autenticar con Microsoft Graph.\n\n"
            + str(detalle)
        )

    if cache.has_state_changed:

        try:

            with open(
                GRAPH_CACHE_FILE,
                "w",
                encoding="utf-8"
            ) as archivo_cache:

                archivo_cache.write(
                    cache.serialize()
                )

        except Exception:

            pass

    return resultado["access_token"]


# =========================================================
# COMPROBAR CUENTA GRAPH
# =========================================================

def obtener_cuenta_graph(token):

    respuesta = requests.get(
        GRAPH_ME_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        },
        timeout=30
    )

    if not respuesta.ok:

        raise RuntimeError(
            "Microsoft Graph no pudo identificar la cuenta.\n\n"
            f"HTTP {respuesta.status_code}: {respuesta.text}"
        )

    datos = respuesta.json()

    correo = (
        datos.get("mail")
        or datos.get("userPrincipalName")
        or ""
    ).strip().lower()

    esperado = (
        CORREO_REMITENTE_OUTLOOK
        .strip()
        .lower()
    )

    if correo and correo != esperado:

        raise RuntimeError(
            "La cuenta de Microsoft iniciada no coincide con "
            "la cuenta configurada para enviar.\n\n"
            f"Configurada: {CORREO_REMITENTE_OUTLOOK}\n"
            f"Iniciada: {correo}\n\n"
            "Cierra la sesión o selecciona la cuenta correcta."
        )

    return correo


# =========================================================
# ENVIAR CORREO
# =========================================================

def enviar_correo_graph(
    token,
    nombre,
    archivo_adjunto,
    pdf_bytes=None,
    pdf_nombre=None,
    destinatario=None
):

    if not os.path.isfile(archivo_adjunto):

        raise FileNotFoundError(
            f"No se encontró el archivo PDF adjunto: "
            f"{archivo_adjunto}"
        )

    with open(
        archivo_adjunto,
        "rb"
    ) as archivo:

        contenido_pdf_personalizado = (
            base64.b64encode(
                archivo.read()
            ).decode("ascii")
        )

    nombre_archivo_pdf_personalizado = (
        os.path.basename(
            archivo_adjunto
        )
    )

    attachments = [
        {
            "@odata.type":
                "#microsoft.graph.fileAttachment",

            "name":
                nombre_archivo_pdf_personalizado,

            "contentType":
                "application/pdf",

            "contentBytes":
                contenido_pdf_personalizado,
        }
    ]

    if pdf_bytes:

        if not pdf_bytes.startswith(b"%PDF"):

            raise ValueError(
                "Los datos de la resolución "
                "no corresponden a un PDF válido."
            )

        attachments.append(
            {
                "@odata.type":
                    "#microsoft.graph.fileAttachment",

                "name":
                    pdf_nombre or "Resolucion.pdf",

                "contentType":
                    "application/pdf",

                "contentBytes":
                    base64.b64encode(
                        pdf_bytes
                    ).decode("ascii"),
            }
        )

    payload = {

        "message": {

            "subject":
                ASUNTO_CORREO.format(
                    nombre=nombre
                ),

            "body": {

                "contentType":
                    "Text",

                "content":
                    CUERPO_CORREO.format(
                        nombre=nombre
                    )
            },

            "toRecipients": [

                {
                    "emailAddress": {

                        "address":
                            (
                                destinatario
                                or CORREO_DESTINO_PRUEBAS
                            )
                    }
                }
            ],

            "attachments":
                attachments,
        },

        "saveToSentItems":
            True,
    }

    respuesta = requests.post(
        GRAPH_SEND_URL,

        headers={
            "Authorization":
                f"Bearer {token}",

            "Content-Type":
                "application/json",

            "Accept":
                "application/json",
        },

        json=payload,

        timeout=60
    )

    if respuesta.status_code != 202:

        try:
            detalle = respuesta.json()

        except Exception:
            detalle = respuesta.text

        raise RuntimeError(
            "Microsoft Graph rechazó el envío.\n\n"
            f"HTTP {respuesta.status_code}\n"
            f"{detalle}"
        )



# =========================================================
# ENVIAR NOTIFICACIÓN DE RECHAZO
# =========================================================

def enviar_correo_rechazo_graph(token, nombre, destinatario, motivo, placa=""):

    destinatario = str(destinatario or "").strip().lower()
    motivo = str(motivo or "").strip()
    nombre = str(nombre or "Usuario").strip() or "Usuario"
    placa = str(placa or "").strip().upper()

    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", destinatario):
        raise ValueError("El registro no tiene un correo electrónico válido.")
    if not motivo:
        raise ValueError("No se puede enviar una notificación de rechazo sin motivo.")

    payload = {
        "message": {
            "subject": "Notificación de rechazo - Resolución AMT",
            "body": {
                "contentType": "Text",
                "content": (
                    f"Estimado/a {nombre}:\n\n"
                    "Le informamos que su trámite/resolución ha sido rechazado.\n\n"
                    f"Placa: {placa or 'No registrada'}\n\n"
                    "Motivo del rechazo:\n"
                    f"{motivo}\n\n"
                    "Por favor, revise la observación indicada y realice las acciones correspondientes.\n\n"
                    "Saludos cordiales,\nAdministrador\nSistema de Resoluciones AMT"
                )
            },
            "toRecipients": [{"emailAddress": {"address": destinatario}}]
        },
        "saveToSentItems": True
    }

    respuesta = requests.post(
        GRAPH_SEND_URL,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json", "Accept": "application/json"},
        json=payload,
        timeout=60
    )

    if respuesta.status_code != 202:
        try:
            detalle = respuesta.json()
        except Exception:
            detalle = respuesta.text
        raise RuntimeError(
            "Microsoft Graph rechazó la notificación de rechazo.\n\n"
            f"HTTP {respuesta.status_code}\n{detalle}"
        )


# =========================================================
# RESTAURAR IMÁGENES DE LA PLANTILLA
# =========================================================

def restaurar_imagenes_originales_excel(
    archivo_salida
):

    if not os.path.isfile(
        PLANTILLA_EXCEL
    ):

        return

    archivos_graficos = {

        "xl/drawings/drawing1.xml",

        "xl/drawings/_rels/"
        "drawing1.xml.rels",
    }

    try:

        with zipfile.ZipFile(
            PLANTILLA_EXCEL,
            "r"
        ) as original:

            nombres_originales = (
                original.namelist()
            )

            for nombre in nombres_originales:

                if nombre.startswith(
                    "xl/media/"
                ):

                    archivos_graficos.add(
                        nombre
                    )

            temporal = (
                archivo_salida
                + ".graficos.tmp"
            )

            with zipfile.ZipFile(
                archivo_salida,
                "r"
            ) as entrada, zipfile.ZipFile(
                temporal,
                "w",
                zipfile.ZIP_DEFLATED
            ) as salida:

                for info in entrada.infolist():

                    if info.filename in archivos_graficos:
                        continue

                    salida.writestr(
                        info,
                        entrada.read(
                            info.filename
                        )
                    )

                for nombre in archivos_graficos:

                    if (
                        nombre
                        in nombres_originales
                    ):

                        salida.writestr(
                            nombre,
                            original.read(
                                nombre
                            )
                        )

            os.replace(
                temporal,
                archivo_salida
            )

    except Exception as error:

        print(
            "Advertencia: no se pudo restaurar "
            "la imagen original:",
            error
        )


# =========================================================
# CONVERTIR EXCEL A PDF
# =========================================================

def convertir_excel_a_pdf(
    archivo_excel
):

    carpeta = os.path.dirname(
        archivo_excel
    )

    base = os.path.splitext(
        os.path.basename(
            archivo_excel
        )
    )[0]

    archivo_pdf = os.path.join(
        carpeta,
        base + ".pdf"
    )

    # -----------------------------------------------------
    # MICROSOFT EXCEL
    # -----------------------------------------------------

    try:

        import win32com.client

        excel = (
            win32com.client.DispatchEx(
                "Excel.Application"
            )
        )

        excel.Visible = False
        excel.DisplayAlerts = False

        libro = None

        try:

            libro = excel.Workbooks.Open(
                os.path.abspath(
                    archivo_excel
                ),
                ReadOnly=True
            )

            libro.ExportAsFixedFormat(
                0,
                os.path.abspath(
                    archivo_pdf
                )
            )

        finally:

            if libro is not None:

                try:
                    libro.Close(False)
                except Exception:
                    pass

            try:
                excel.Quit()
            except Exception:
                pass

        if os.path.isfile(
            archivo_pdf
        ):

            return archivo_pdf

    except Exception as error_excel:

        print(
            "Excel COM no disponible; "
            "se intentará LibreOffice:",
            error_excel
        )

    # -----------------------------------------------------
    # LIBREOFFICE
    # -----------------------------------------------------

    ejecutable = (
        shutil.which("libreoffice")
        or
        shutil.which("soffice")
    )

    posibles = [

        r"C:\Program Files\LibreOffice"
        r"\program\soffice.exe",

        r"C:\Program Files (x86)"
        r"\LibreOffice\program\soffice.exe",
    ]

    if not ejecutable:

        for ruta_posible in posibles:

            if os.path.isfile(
                ruta_posible
            ):

                ejecutable = (
                    ruta_posible
                )

                break

    if not ejecutable:

        raise RuntimeError(
            "No se pudo convertir el Excel a PDF.\n\n"
            "Se necesita Microsoft Excel o LibreOffice instalado."
        )

    resultado = subprocess.run(

        [
            ejecutable,

            "--headless",

            "--convert-to",
            "pdf",

            "--outdir",
            carpeta,

            archivo_excel
        ],

        capture_output=True,

        text=True,

        timeout=120
    )

    if resultado.returncode != 0:

        raise RuntimeError(
            "LibreOffice no pudo convertir "
            "el Excel a PDF.\n\n"
            +
            (
                resultado.stderr
                or
                resultado.stdout
            )
        )

    if not os.path.isfile(
        archivo_pdf
    ):

        raise RuntimeError(
            "LibreOffice terminó, pero "
            "no se encontró el PDF generado."
        )

    return archivo_pdf


# =========================================================
# CONFIGURACIÓN DE TIEMPOS
# =========================================================

INTERVALO_ACTUALIZACION = 20000

TIMEOUT_LISTA = 30

TIMEOUT_PDF = 120


# =========================================================
# VISOR PDF
# =========================================================

class PDFViewer(QWidget):

    def __init__(self):

        super().__init__()

        self.doc = None

        self.page_num = 0

        self.zoom = 1.0

        self.init_ui()

    def init_ui(self):

        layout = QVBoxLayout(
            self
        )

        layout.setContentsMargins(
            0,
            0,
            0,
            0
        )

        layout.setSpacing(8)

        controles = QHBoxLayout()

        controles.setSpacing(6)

        self.btn_anterior = QPushButton(
            "◀"
        )

        self.btn_siguiente = QPushButton(
            "▶"
        )

        self.lbl_pagina = QLabel(
            "Página 0 de 0"
        )

        self.lbl_pagina.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.btn_zoom_menos = QPushButton(
            "-"
        )

        self.btn_zoom_mas = QPushButton(
            "+"
        )

        self.btn_anterior.clicked.connect(
            self.pagina_anterior
        )

        self.btn_siguiente.clicked.connect(
            self.pagina_siguiente
        )

        self.btn_zoom_menos.clicked.connect(
            self.zoom_menos
        )

        self.btn_zoom_mas.clicked.connect(
            self.zoom_mas
        )

        for boton in (

            self.btn_anterior,

            self.btn_siguiente,

            self.btn_zoom_menos,

            self.btn_zoom_mas

        ):

            boton.setFixedHeight(
                36
            )

            boton.setMinimumWidth(
                36
            )

        controles.addWidget(
            self.btn_anterior
        )

        controles.addWidget(
            self.lbl_pagina
        )

        controles.addWidget(
            self.btn_siguiente
        )

        controles.addStretch()

        controles.addWidget(
            QLabel("Zoom:")
        )

        controles.addWidget(
            self.btn_zoom_menos
        )

        controles.addWidget(
            self.btn_zoom_mas
        )

        layout.addLayout(
            controles
        )

        self.scroll = QScrollArea()

        self.scroll.setWidgetResizable(
            True
        )

        self.scroll.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.lbl_pdf = QLabel()

        self.lbl_pdf.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.lbl_pdf.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        self.scroll.setWidget(
            self.lbl_pdf
        )

        layout.addWidget(
            self.scroll,
            1
        )

        # -------------------------------------------------
        # CARGADOR EXCLUSIVO DEL VISOR PDF
        # -------------------------------------------------
        # Se muestra únicamente dentro del área donde aparece
        # el PDF. No bloquea ni cubre el resto de la aplicación.
        self.pdf_loading = QFrame(self)
        self.pdf_loading.setObjectName(
            "pdfLoading"
        )

        self.pdf_loading.setStyleSheet("""
            #pdfLoading {
                background: rgba(255, 255, 255, 235);
                border: 1px solid #dce5ef;
                border-radius: 10px;
            }

            #pdfLoadingTitulo {
                color: #17324d;
                font-size: 11pt;
                font-weight: 700;
                background: transparent;
            }

            #pdfLoadingTexto {
                color: #718096;
                font-size: 9pt;
                background: transparent;
            }

            QProgressBar {
                background: #e9eef3;
                border: none;
                border-radius: 5px;
                height: 8px;
            }

            QProgressBar::chunk {
                background: #3978a8;
                border-radius: 5px;
            }
        """)

        loading_layout = QVBoxLayout(
            self.pdf_loading
        )

        loading_layout.setContentsMargins(
            24,
            18,
            24,
            18
        )

        loading_layout.setSpacing(
            7
        )

        loading_layout.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.pdf_loading_titulo = QLabel(
            "⏳ Cargando PDF..."
        )

        self.pdf_loading_titulo.setObjectName(
            "pdfLoadingTitulo"
        )

        self.pdf_loading_titulo.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        loading_layout.addWidget(
            self.pdf_loading_titulo
        )

        self.pdf_loading_texto = QLabel(
            "Preparando la resolución..."
        )

        self.pdf_loading_texto.setObjectName(
            "pdfLoadingTexto"
        )

        self.pdf_loading_texto.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        loading_layout.addWidget(
            self.pdf_loading_texto
        )

        self.pdf_loading_barra = QProgressBar()

        # Indeterminado: muestra que el sistema está trabajando
        # sin inventar un porcentaje de carga.
        self.pdf_loading_barra.setRange(
            0,
            0
        )

        self.pdf_loading_barra.setTextVisible(
            False
        )

        loading_layout.addWidget(
            self.pdf_loading_barra
        )

        self.pdf_loading.hide()

        self.actualizar_controles()

    def resizeEvent(self, event):
        super().resizeEvent(event)

        if hasattr(self, "pdf_loading") and hasattr(self, "scroll"):
            self.pdf_loading.setGeometry(
                self.scroll.geometry()
            )

    def mostrar_cargando(self):
        if not hasattr(self, "pdf_loading"):
            return

        self.pdf_loading.setGeometry(
            self.scroll.geometry()
        )

        self.pdf_loading.raise_()
        self.pdf_loading.show()

        # Permite que el usuario vea el cargador antes de
        # comenzar cualquier operación que pueda tardar.
        QApplication.processEvents()

    def ocultar_cargando(self):
        if hasattr(self, "pdf_loading"):
            self.pdf_loading.hide()

    def abrir_pdf(
        self,
        ruta
    ):

        try:

            self.cerrar()

            self.doc = pymupdf.open(
                ruta
            )

            self.page_num = 0

            self.zoom = 1.0

            self.mostrar_pagina()

        except Exception as e:

            QMessageBox.critical(
                self,
                "Error PDF",
                f"No se pudo abrir el PDF:\n\n{e}"
            )

    def cerrar(self):

        if self.doc is not None:

            try:
                self.doc.close()
            except Exception:
                pass

            self.doc = None

        self.lbl_pdf.clear()

        self.lbl_pagina.setText(
            "Página 0 de 0"
        )

        self.actualizar_controles()

    def limpiar(self):

        self.cerrar()

    def mostrar_pagina(self):

        if self.doc is None:
            return

        try:

            page = self.doc[
                self.page_num
            ]

            matriz = pymupdf.Matrix(
                self.zoom,
                self.zoom
            )

            pix = page.get_pixmap(
                matrix=matriz,
                alpha=False
            )

            imagen = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format.Format_RGB888
            )

            pixmap = QPixmap.fromImage(
                imagen
            )

            self.lbl_pdf.setPixmap(
                pixmap
            )

            self.lbl_pagina.setText(
                f"Página {self.page_num + 1} "
                f"de {len(self.doc)}"
            )

            self.actualizar_controles()

        except Exception as e:

            QMessageBox.critical(
                self,
                "Error",
                f"No se pudo mostrar la página:\n\n{e}"
            )

    def pagina_anterior(self):

        if self.doc is None:
            return

        if self.page_num > 0:

            self.page_num -= 1

            self.mostrar_pagina()

    def pagina_siguiente(self):

        if self.doc is None:
            return

        if (
            self.page_num
            <
            len(self.doc) - 1
        ):

            self.page_num += 1

            self.mostrar_pagina()

    def zoom_menos(self):

        if self.doc is None:
            return

        self.zoom -= 0.1

        if self.zoom < 0.3:
            self.zoom = 0.3

        self.mostrar_pagina()

    def zoom_mas(self):

        if self.doc is None:
            return

        self.zoom += 0.1

        if self.zoom > 4.0:
            self.zoom = 4.0

        self.mostrar_pagina()

    def actualizar_controles(self):

        tiene_pdf = (
            self.doc is not None
        )

        self.btn_anterior.setEnabled(
            tiene_pdf
            and
            self.page_num > 0
        )

        self.btn_siguiente.setEnabled(
            tiene_pdf
            and
            self.page_num
            <
            len(self.doc) - 1
            if tiene_pdf
            else False
        )

        self.btn_zoom_menos.setEnabled(
            tiene_pdf
        )

        self.btn_zoom_mas.setEnabled(
            tiene_pdf
        )


# =========================================================
# CAMBIAR ESTADO REMOTO
# =========================================================

def cambiar_estado_remoto(
    pdf_id,
    estado,
    motivo_rechazo=""
):

    pdf_id = str(pdf_id or "").strip()
    estado = str(estado or "").strip().upper()
    motivo_rechazo = str(motivo_rechazo or "").strip()

    if not pdf_id:
        raise ValueError("Falta el PDF ID para cambiar el estado.")

    if estado not in ESTADOS_VALIDOS:
        raise ValueError("Estado inválido: " + estado)

    if estado == ESTADO_RECHAZADO and not motivo_rechazo:
        raise ValueError("El rechazo requiere un motivo.")

    respuesta = requests.post(
        GOOGLE_SHEETS_URL,
        json={
            "accion": "cambiar_estado",
            "pdf_id": pdf_id,
            "estado": estado,
            "motivo_rechazo": motivo_rechazo
        },
        timeout=30
    )

    if respuesta.status_code != 200:
        raise RuntimeError(
            "Google Apps Script respondió con HTTP "
            f"{respuesta.status_code}"
        )

    resultado = respuesta.json()

    if resultado.get("estado") != "ok":
        raise RuntimeError(
            resultado.get(
                "mensaje",
                "No se pudo cambiar el estado."
            )
        )

    return resultado


# =========================================================
# WORKER DE PETICIONES
# =========================================================

class WorkerSignals(QObject):

    finished = Signal(object)

    error = Signal(str)


class RequestWorker(QRunnable):

    def __init__(
        self,
        url,
        params,
        timeout
    ):

        super().__init__()

        self.url = url

        self.params = params

        self.timeout = timeout

        self.signals = WorkerSignals()

    def run(self):

        try:

            respuesta = requests.get(

                self.url,

                params=self.params,

                timeout=self.timeout
            )

            if respuesta.status_code != 200:

                raise Exception(
                    "Google Apps Script respondió "
                    f"con HTTP {respuesta.status_code}"
                )

            resultado = respuesta.json()

            self.signals.finished.emit(
                resultado
            )

        except Exception as e:

            self.signals.error.emit(
                str(e)
            )



# =========================================================
# FIRMA DIGITAL PDF
# =========================================================
#
# Se utiliza exactamente la misma apariencia y coordenadas
# del firmador independiente que ya funciona correctamente.
#
# La plantilla ya está cargada en el sistema; aquí solamente
# se firma el PDF personalizado que se genera al enviar.
# =========================================================

SIGNATURE_BOX = (197.5, 53, 414.5, 108)
STAMP_W = SIGNATURE_BOX[2] - SIGNATURE_BOX[0]
STAMP_H = SIGNATURE_BOX[3] - SIGNATURE_BOX[1]
SIGNATURE_FIELD_NAME = "FirmaElectronica"


def sha256_archivo_firma(path):
    h = hashlib.sha256()

    with open(path, "rb") as f:
        for bloque in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):
            h.update(bloque)

    return h.hexdigest()


def cargar_datos_certificado_firma(
    p12_path,
    password
):
    """
    Lee el certificado .P12/.PFX para obtener el nombre y serial.
    La clave privada se utiliza para firmar, pero no se guarda.
    """

    with open(p12_path, "rb") as f:
        data = f.read()

    private_key, certificate, additional_certs = (
        pkcs12.load_key_and_certificates(
            data,
            password.encode("utf-8")
            if password
            else None
        )
    )

    if private_key is None or certificate is None:
        raise ValueError(
            "El archivo .P12/.PFX no contiene "
            "una clave privada y certificado utilizables."
        )

    nombre = None

    try:
        cn = certificate.subject.get_attributes_for_oid(
            NameOID.COMMON_NAME
        )

        if cn:
            nombre = cn[0].value

    except Exception:
        pass

    if not nombre:

        try:
            nombres = (
                certificate.subject.get_attributes_for_oid(
                    NameOID.GIVEN_NAME
                )
            )

            apellidos = (
                certificate.subject.get_attributes_for_oid(
                    NameOID.SURNAME
                )
            )

            partes = []

            if nombres:
                partes.append(
                    str(nombres[0].value)
                )

            if apellidos:
                partes.append(
                    str(apellidos[0].value)
                )

            nombre = " ".join(partes).strip()

        except Exception:
            pass

    if not nombre:
        nombre = certificate.subject.rfc4514_string()

    return {
        "nombre": str(nombre).upper(),
        "serial": format(
            certificate.serial_number,
            "X"
        ),
        "not_before": certificate.not_valid_before_utc,
        "not_after": certificate.not_valid_after_utc,
    }


def crear_estampa_firma_correo(
    salida_pdf,
    nombre,
    numero_serie,
    hash_pdf,
    fecha_hora
):
    """
    Crea exactamente la apariencia visual de firma
    que ya funciona en el firmador independiente.
    """

    qr_texto = (
        "FIRMA ELECTRONICA\n"
        f"Nombre: {nombre}\n"
        f"Serial: {numero_serie}\n"
        f"Fecha: {fecha_hora}\n"
        f"SHA256: {hash_pdf}"
    )

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=1,
    )

    qr.add_data(qr_texto)
    qr.make(fit=True)

    qr_img = qr.make_image(
        fill_color="black",
        back_color="white"
    )

    qr_temp = Path(
        salida_pdf
    ).with_name(
        "qr_firma_temp.png"
    )

    qr_img.save(qr_temp)

    c = canvas.Canvas(
        str(salida_pdf),
        pagesize=(
            STAMP_W,
            STAMP_H
        )
    )

    margen = 4
    qr_size = STAMP_H - (margen * 2)

    c.drawImage(
        str(qr_temp),
        margen,
        margen,
        width=qr_size,
        height=qr_size,
        preserveAspectRatio=True,
        mask="auto",
    )

    x_text = (
        margen
        + qr_size
        + 5
    )

    c.setFillColorRGB(
        0,
        0,
        0
    )

    c.setFont(
        "Courier",
        5.5
    )

    nombre_limpio = (
        " ".join(
            nombre.split()
        )
    )

    palabras = (
        nombre_limpio.split()
    )

    if (
        len(nombre_limpio) > 27
        and len(palabras) >= 3
    ):
        mitad = (
            len(palabras) // 2
        )

        linea1 = " ".join(
            palabras[:mitad]
        )

        linea2 = " ".join(
            palabras[mitad:]
        )

    else:
        linea1 = nombre_limpio
        linea2 = ""

    centro_qr_y = (
        margen
        + (qr_size / 2)
    )

    c.drawString(
        x_text,
        centro_qr_y + 7,
        "Firmado electrónicamente por:"
    )

    c.setFont(
        "Courier-Bold",
        7.8
    )

    c.drawString(
        x_text,
        centro_qr_y - 1,
        linea1
    )

    if linea2:
        c.drawString(
            x_text,
            centro_qr_y - 9,
            linea2
        )

    c.save()

    try:
        qr_temp.unlink()

    except Exception:
        pass


def normalizar_pdf_para_firma_correo(
    pdf_entrada,
    pdf_salida
):
    """
    Reescribe el PDF antes de firmarlo para evitar
    problemas de hybrid cross-reference.
    """

    try:
        reader = PdfReader(
            str(pdf_entrada),
            strict=False
        )

        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        try:

            if reader.metadata:

                metadata = {
                    str(k): str(v)
                    for k, v
                    in reader.metadata.items()
                    if k and v is not None
                }

                if metadata:
                    writer.add_metadata(
                        metadata
                    )

        except Exception:
            pass

        with open(
            pdf_salida,
            "wb"
        ) as f:

            writer.write(f)

        if (
            not Path(pdf_salida).exists()
            or Path(pdf_salida).stat().st_size == 0
        ):
            raise RuntimeError(
                "No se pudo crear el PDF "
                "normalizado para la firma."
            )

        return Path(
            pdf_salida
        )

    except Exception as e:

        raise RuntimeError(
            "No se pudo normalizar el PDF "
            "antes de firmarlo.\n\n"
            f"Detalle: {e}"
        )


def firmar_pdf_correo(
    pdf_entrada,
    pdf_salida,
    p12_path,
    password,
    datos_certificado
):
    """
    Firma criptográficamente el PDF personalizado
    con el certificado .P12/.PFX.
    """

    signer = (
        signers.SimpleSigner.load_pkcs12(
            pfx_file=str(
                p12_path
            ),
            passphrase=(
                password.encode("utf-8")
                if password
                else None
            ),
        )
    )

    pdf_salida = Path(
        pdf_salida
    )

    apariencia_pdf = (
        pdf_salida.with_name(
            pdf_salida.stem
            + "_apariencia.pdf"
        )
    )

    pdf_normalizado = (
        pdf_salida.with_name(
            pdf_salida.stem
            + "_BASE.pdf"
        )
    )

    try:

        hash_pdf = (
            sha256_archivo_firma(
                pdf_entrada
            )
        )

        fecha_hora = datetime.now().strftime(
            "%d/%m/%Y %H:%M:%S"
        )

        crear_estampa_firma_correo(
            apariencia_pdf,
            datos_certificado["nombre"],
            datos_certificado["serial"],
            hash_pdf,
            fecha_hora
        )

        normalizar_pdf_para_firma_correo(
            pdf_entrada,
            pdf_normalizado
        )

        with open(
            pdf_normalizado,
            "rb"
        ) as entrada, open(
            pdf_salida,
            "wb"
        ) as salida:

            writer = (
                IncrementalPdfFileWriter(
                    entrada,
                    strict=True
                )
            )

            fields.append_signature_field(
                writer,
                sig_field_spec=fields.SigFieldSpec(
                    SIGNATURE_FIELD_NAME,
                    box=SIGNATURE_BOX,
                    on_page=0,
                ),
            )

            metadata = (
                signers.PdfSignatureMetadata(
                    field_name=SIGNATURE_FIELD_NAME,
                    md_algorithm="sha256",
                    reason=(
                        "Firma electrónica "
                        "del documento"
                    ),
                )
            )

            pdf_signer = (
                signers.PdfSigner(
                    metadata,
                    signer=signer,
                    stamp_style=(
                        stamp.StaticStampStyle.from_pdf_file(
                            str(
                                apariencia_pdf
                            ),
                            border_width=0,
                        )
                    ),
                )
            )

            pdf_signer.sign_pdf(
                writer,
                output=salida
            )

        if (
            not pdf_salida.exists()
            or pdf_salida.stat().st_size == 0
        ):
            raise RuntimeError(
                "No se pudo generar el "
                "PDF firmado."
            )

        return str(
            pdf_salida
        )

    finally:

        for temporal in (
            apariencia_pdf,
            pdf_normalizado
        ):

            try:

                if temporal.exists():
                    temporal.unlink()

            except Exception:
                pass


# =========================================================
# DIÁLOGO PARA CERTIFICADO DE FIRMA
# =========================================================

class FirmaCorreoDialog(QDialog):

    def __init__(
        self,
        parent=None,
        p12_path=""
    ):
        super().__init__(parent)

        self.setWindowTitle(
            "Firmar documentos"
        )

        self.setMinimumWidth(
            620
        )

        self.p12_path = str(
            p12_path or ""
        )

        self.password = ""
        self.datos_certificado = None

        layout = QVBoxLayout(
            self
        )

        titulo = QLabel(
            "FIRMA ELECTRÓNICA"
        )

        titulo.setFont(
            QFont(
                "Arial",
                14,
                QFont.Weight.Bold
            )
        )

        titulo.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        layout.addWidget(
            titulo
        )

        if self.p12_path:
            texto_certificado = (
                "Introduce la contraseña de tu certificado .P12/.PFX.\\n"
                "El documento se firmará automáticamente "
                "antes de enviarse por correo."
            )
        else:
            texto_certificado = (
                "Selecciona tu certificado .P12/.PFX y luego "
                "introduce su contraseña.\\n"
                "El documento se firmará automáticamente "
                "antes de enviarse por correo."
            )

        texto = QLabel(
            texto_certificado
        )

        texto.setWordWrap(
            True
        )

        layout.addWidget(
            texto
        )

        fila_cert = QHBoxLayout()

        fila_cert.addWidget(
            QLabel(
                "Certificado:"
            )
        )

        self.txt_certificado = QLineEdit()

        self.txt_certificado.setReadOnly(
            True
        )

        self.txt_certificado.setText(
            self.p12_path
        )

        fila_cert.addWidget(
            self.txt_certificado,
            1
        )

        self.btn_buscar_certificado = QPushButton(
            "Cambiar..."
        )

        self.btn_buscar_certificado.clicked.connect(
            self.seleccionar_certificado
        )

        fila_cert.addWidget(
            self.btn_buscar_certificado
        )

        # El selector SOLO aparece al momento de enviar
        # cuando todavía no existe un certificado cargado.
        # Si ya existe una ruta guardada, se muestra únicamente
        # la contraseña.
        self.btn_buscar_certificado.setVisible(
            not bool(self.p12_path)
        )

        layout.addLayout(
            fila_cert
        )

        fila_pass = QHBoxLayout()

        fila_pass.addWidget(
            QLabel(
                "Contraseña:"
            )
        )

        self.txt_password = QLineEdit()

        self.txt_password.setEchoMode(
            QLineEdit.EchoMode.Password
        )

        fila_pass.addWidget(
            self.txt_password,
            1
        )

        layout.addLayout(
            fila_pass
        )

        self.lbl_firmante = QLabel(
            "Firmante: --"
        )

        self.lbl_firmante.setWordWrap(
            True
        )

        layout.addWidget(
            self.lbl_firmante
        )

        botones = QHBoxLayout()

        botones.addStretch()

        btn_cancelar = QPushButton(
            "Cancelar"
        )

        btn_cancelar.clicked.connect(
            self.reject
        )

        btn_firmar = QPushButton(
            "FIRMAR Y ENVIAR"
        )

        btn_firmar.setDefault(
            True
        )

        btn_firmar.clicked.connect(
            self.aceptar
        )

        botones.addWidget(
            btn_cancelar
        )

        botones.addWidget(
            btn_firmar
        )

        layout.addLayout(
            botones
        )

    def seleccionar_certificado(
        self
    ):

        ruta, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar certificado digital",
            "",
            (
                "Certificado PKCS#12 "
                "(*.p12 *.pfx);;"
                "Todos los archivos (*.*)"
            )
        )

        if not ruta:
            return

        self.p12_path = ruta

        self.txt_certificado.setText(
            ruta
        )

        self.btn_buscar_certificado.setVisible(
            False
        )

        self._validar_certificado(
            mostrar_error=False
        )

    def _validar_certificado(
        self,
        mostrar_error=True
    ):

        if not self.p12_path:

            if mostrar_error:

                QMessageBox.warning(
                    self,
                    "Certificado",
                    "Selecciona el certificado .P12/.PFX."
                )

            return False

        if not os.path.isfile(
            self.p12_path
        ):

            if mostrar_error:

                QMessageBox.warning(
                    self,
                    "Certificado",
                    "El certificado guardado ya no existe. "
                    "Selecciona un .P12/.PFX nuevo."
                )

            return False

        try:

            self.datos_certificado = (
                cargar_datos_certificado_firma(
                    self.p12_path,
                    self.txt_password.text()
                )
            )

            self.lbl_firmante.setText(
                "Firmante: "
                + self.datos_certificado[
                    "nombre"
                ]
            )

            return True

        except Exception as e:

            self.datos_certificado = None

            self.lbl_firmante.setText(
                "Firmante: --"
            )

            if mostrar_error:

                QMessageBox.critical(
                    self,
                    "Certificado no válido",
                    "No se pudo abrir el certificado.\n\n"
                    + str(e)
                )

            return False

    def aceptar(self):

        if not self.p12_path:

            QMessageBox.warning(
                self,
                "Falta certificado",
                "No hay un certificado .P12/.PFX configurado."
            )

            return

        if not self.txt_password.text():

            QMessageBox.warning(
                self,
                "Falta contraseña",
                "Introduce la contraseña del certificado."
            )

            return

        if not self._validar_certificado():

            return

        self.password = (
            self.txt_password.text()
        )

        self.accept()




# =========================================================
# SEÑALES DEL WORKER DE CORREOS
# =========================================================

class EmailWorkerSignals(QObject):

    progress = Signal(
        int,
        int,
        str
    )

    finished = Signal(
        int,
        object
    )

    error_fatal = Signal(
        str
    )


# =========================================================
# WORKER PARA ENVÍO DE CORREOS
# =========================================================

class EmailWorker(QRunnable):

    def __init__(
        self,
        registros,
        cache_pdfs,
        p12_path,
        password,
        datos_certificado
    ):

        super().__init__()

        self.registros = registros

        self.cache_pdfs = cache_pdfs

        # Datos del certificado para firmar cada PDF
        # personalizado antes de enviarlo.
        self.p12_path = p12_path
        self.password = password
        self.datos_certificado = datos_certificado

        self.signals = EmailWorkerSignals()

        self.setAutoDelete(
            True
        )

    def crear_excel_personalizado(
        self,
        registro
    ):

        fecha = str(
            registro.get(
                "fecha",
                ""
            )
        ).strip()

        if " " in fecha:

            fecha = fecha.split(
                " "
            )[0]

        elif "T" in fecha:

            fecha = fecha.split(
                "T"
            )[0]

        try:

            from datetime import datetime as _datetime

            for _formato in (
                "%Y-%m-%d",
                "%d/%m/%Y"
            ):

                try:

                    fecha = (
                        _datetime.strptime(
                            fecha,
                            _formato
                        )
                        .strftime(
                            "%d/%m/%Y"
                        )
                    )

                    break

                except ValueError:

                    pass

        except Exception:

            pass

        placa = str(
            registro.get(
                "placa",
                ""
            )
        ).strip()

        chasis = str(
            registro.get(
                "chasis",
                ""
            )
        ).strip()

        socio = str(
            registro.get(
                "socio",
                ""
            )
        ).strip()

        libro = load_workbook(
            PLANTILLA_EXCEL
        )

        hoja = libro[
            "Verificacion Documental"
        ]

        # -------------------------------------------------
        # ASEGURAR FIRMA
        # -------------------------------------------------

        if not getattr(
            hoja,
            "_images",
            []
        ):

            try:

                with zipfile.ZipFile(
                    PLANTILLA_EXCEL,
                    "r"
                ) as _zip:

                    _media = [

                        n

                        for n
                        in _zip.namelist()

                        if n.startswith(
                            "xl/media/"
                        )

                        and n.lower().endswith(
                            (
                                ".png",
                                ".jpg",
                                ".jpeg"
                            )
                        )
                    ]

                    if _media:

                        _tmp_firma = os.path.join(

                            tempfile.gettempdir(),

                            "firma_plantilla_amt.png"
                        )

                        with open(
                            _tmp_firma,
                            "wb"
                        ) as _f:

                            _f.write(
                                _zip.read(
                                    _media[0]
                                )
                            )

                        from openpyxl.drawing.image import (
                            Image as XLImage
                        )

                        _firma = XLImage(
                            _tmp_firma
                        )

                        _firma.anchor = "A45"

                        hoja.add_image(
                            _firma
                        )

            except Exception as _error_firma:

                print(
                    "Advertencia: no se pudo "
                    "recuperar la imagen:",
                    _error_firma
                )

        hoja["F3"] = fecha

        hoja["C8"] = placa

        hoja["C9"] = chasis

        hoja["C10"] = socio

        placa_archivo = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            placa
        ) or "resolucion"

        nombre_archivo = (
            f"Verificacion_{placa_archivo}.xlsx"
        )

        ruta = os.path.join(
            tempfile.gettempdir(),
            nombre_archivo
        )

        libro.save(
            ruta
        )

        libro.close()

        restaurar_imagenes_originales_excel(
            ruta
        )

        return ruta

    def obtener_pdf(
        self,
        pdf_id
    ):

        if pdf_id in self.cache_pdfs:

            return self.cache_pdfs[
                pdf_id
            ]

        respuesta_pdf = requests.get(

            GOOGLE_SHEETS_URL,

            params={
                "accion": "pdf",
                "id": pdf_id
            },

            timeout=TIMEOUT_PDF
        )

        if respuesta_pdf.status_code != 200:

            raise Exception(
                "No se pudo obtener el PDF "
                f"(HTTP {respuesta_pdf.status_code})."
            )

        resultado_pdf = (
            respuesta_pdf.json()
        )

        if resultado_pdf.get(
            "estado"
        ) != "ok":

            raise Exception(
                resultado_pdf.get(
                    "mensaje",
                    "No se pudo obtener el PDF."
                )
            )

        pdf_base64 = (
            resultado_pdf.get(
                "data"
            )
        )

        if not pdf_base64:

            raise Exception(
                "El servidor no devolvió el PDF."
            )

        pdf_bytes = base64.b64decode(
            pdf_base64
        )

        if not pdf_bytes.startswith(
            b"%PDF"
        ):

            raise Exception(
                "El archivo recibido "
                "no es un PDF válido."
            )

        self.cache_pdfs[
            pdf_id
        ] = pdf_bytes

        return pdf_bytes

    def run(self):

        enviados = 0

        errores = []

        try:

            # -------------------------------------------------
            # AUTENTICACIÓN
            # -------------------------------------------------

            self.signals.progress.emit(
                0,
                len(self.registros),
                "Autenticando con Microsoft..."
            )

            token = (
                obtener_token_microsoft_graph()
            )

            obtener_cuenta_graph(
                token
            )

            total = len(
                self.registros
            )

            # -------------------------------------------------
            # PROCESAR CADA PERSONA
            # -------------------------------------------------

            for indice, registro in enumerate(
                self.registros,
                start=1
            ):

                nombre = str(
                    registro.get(
                        "socio",
                        ""
                    )
                ).strip() or "Usuario"

                self.signals.progress.emit(
                    indice - 1,
                    total,
                    f"Procesando: {nombre}"
                )

                excel_personalizado = None

                pdf_personalizado = None

                try:

                    # -----------------------------------------
                    # PDF ORIGINAL
                    # -----------------------------------------

                    pdf_id = str(
                        registro.get(
                            "pdf_id",
                            ""
                        )
                    ).strip()

                    if not pdf_id:

                        raise Exception(
                            "Esta resolución no tiene "
                            "pdf_id asociado."
                        )

                    pdf_bytes = (
                        self.obtener_pdf(
                            pdf_id
                        )
                    )

                    # -----------------------------------------
                    # PLACA
                    # -----------------------------------------

                    placa = str(
                        registro.get(
                            "placa",
                            ""
                        )
                    ).strip() or "resolucion"

                    # -----------------------------------------
                    # CREAR EXCEL
                    # -----------------------------------------

                    self.signals.progress.emit(
                        indice - 1,
                        total,
                        f"Generando documento de {nombre}..."
                    )

                    excel_personalizado = (
                        self.crear_excel_personalizado(
                            registro
                        )
                    )

                    # -----------------------------------------
                    # CONVERTIR A PDF
                    # -----------------------------------------

                    self.signals.progress.emit(
                        indice - 1,
                        total,
                        f"Preparando PDF de {nombre}..."
                    )

                    pdf_personalizado = (
                        convertir_excel_a_pdf(
                            excel_personalizado
                        )
                    )

                    # -----------------------------------------
                    # FIRMAR PDF PERSONALIZADO
                    # -----------------------------------------
                    self.signals.progress.emit(
                        indice - 1,
                        total,
                        f"Firmando documento de {nombre}..."
                    )

                    pdf_firmado = os.path.join(
                        tempfile.gettempdir(),
                        f"Verificacion_{placa}_FIRMADO.pdf"
                    )

                    try:
                        if os.path.isfile(pdf_firmado):
                            os.remove(pdf_firmado)
                    except Exception:
                        pass

                    firmar_pdf_correo(
                        pdf_personalizado,
                        pdf_firmado,
                        self.p12_path,
                        self.password,
                        self.datos_certificado
                    )

                    # Desde aquí se adjunta la versión firmada.
                    pdf_personalizado = pdf_firmado

                    # -----------------------------------------
                    # ENVIAR
                    # -----------------------------------------

                    self.signals.progress.emit(
                        indice - 1,
                        total,
                        f"Enviando correo de {nombre}..."
                    )

                    enviar_correo_graph(

                        token=token,

                        nombre=nombre,

                        archivo_adjunto=
                            pdf_personalizado,

                        pdf_bytes=
                            pdf_bytes,

                        pdf_nombre=
                            f"Resolucion_{placa}.pdf"
                    )

                    # Solo después de que Microsoft Graph acepta el correo,
                    # se marca la resolución como ENVIADO.
                    try:
                        cambiar_estado_remoto(
                            pdf_id,
                            ESTADO_ENVIADO,
                            ""
                        )
                    except Exception as error_estado:
                        errores.append(
                            f"{nombre}: correo enviado, pero no se pudo "
                            f"marcar como ENVIADO: {error_estado}"
                        )

                    enviados += 1

                except Exception as e:

                    errores.append(
                        f"{nombre}: {e}"
                    )

                finally:

                    if excel_personalizado:

                        try:

                            os.remove(
                                excel_personalizado
                            )

                        except Exception:

                            pass

                    if pdf_personalizado:

                        try:

                            os.remove(
                                pdf_personalizado
                            )

                        except Exception:

                            pass

                # -----------------------------------------
                # ACTUALIZAR PROGRESO
                # -----------------------------------------

                self.signals.progress.emit(
                    indice,
                    total,
                    f"Completado: {nombre}"
                )

            # -------------------------------------------------
            # FINAL
            # -------------------------------------------------

            self.signals.finished.emit(
                enviados,
                errores
            )

        except Exception as e:

            self.signals.error_fatal.emit(
                str(e)
            )


# =========================================================
# VENTANA DE CARGA
# =========================================================

class LoadingOverlay(QFrame):

    def __init__(
        self,
        parent=None
    ):

        super().__init__(
            parent
        )

        self.setObjectName(
            "loadingOverlay"
        )

        self.setAttribute(
            Qt.WidgetAttribute.WA_StyledBackground,
            True
        )

        self.setStyleSheet("""
            #loadingOverlay {
                background: rgba(23, 50, 77, 235);
                border-radius: 14px;
            }

            #loadingTitulo {
                color: white;
                font-size: 18pt;
                font-weight: 700;
                background: transparent;
            }

            #loadingTexto {
                color: #e5edf5;
                font-size: 10pt;
                background: transparent;
            }

            #loadingContador {
                color: white;
                font-size: 11pt;
                font-weight: 700;
                background: transparent;
            }

            QProgressBar {
                background: rgba(255,255,255,35);
                border: 1px solid rgba(255,255,255,80);
                border-radius: 7px;
                height: 14px;
                text-align: center;
            }

            QProgressBar::chunk {
                background: #ffffff;
                border-radius: 6px;
            }
        """)

        layout = QVBoxLayout(
            self
        )

        layout.setContentsMargins(
            35,
            30,
            35,
            30
        )

        layout.setSpacing(
            12
        )

        layout.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.titulo = QLabel(
            "📧 Enviando correos"
        )

        self.titulo.setObjectName(
            "loadingTitulo"
        )

        self.titulo.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        layout.addWidget(
            self.titulo
        )

        self.texto = QLabel(
            "Procesando..."
        )

        self.texto.setObjectName(
            "loadingTexto"
        )

        self.texto.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.texto.setWordWrap(
            True
        )

        layout.addWidget(
            self.texto
        )

        self.contador = QLabel(
            "Preparando..."
        )

        self.contador.setObjectName(
            "loadingContador"
        )

        self.contador.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        layout.addWidget(
            self.contador
        )

        self.progress = QProgressBar()

        self.progress.setRange(
            0,
            100
        )

        self.progress.setValue(
            0
        )

        self.progress.setTextVisible(
            False
        )

        layout.addWidget(
            self.progress
        )

        self.animacion_timer = QTimer(
            self
        )

        self.animacion_timer.timeout.connect(
            self.animar_titulo
        )

        self.puntos = 0

        self.setFixedWidth(
            480
        )

        self.hide()

    def animar_titulo(self):

        self.puntos += 1

        if self.puntos > 3:
            self.puntos = 0

        texto = (
            "📧 Enviando correos"
            +
            "." * self.puntos
        )

        self.titulo.setText(
            texto
        )

    def mostrar(
        self,
        total
    ):

        self.puntos = 0

        self.progress.setValue(
            0
        )

        self.contador.setText(
            f"Preparando envío de {total} correo(s)..."
        )

        self.texto.setText(
            "No cierres la aplicación mientras "
            "el proceso esté en ejecución."
        )

        self.animacion_timer.start(
            400
        )

        self.raise_()

        self.show()

    def actualizar(
        self,
        procesados,
        total,
        mensaje
    ):

        if total > 0:

            porcentaje = int(
                (
                    procesados
                    /
                    total
                )
                * 100
            )

        else:

            porcentaje = 0

        self.progress.setValue(
            porcentaje
        )

        self.contador.setText(
            f"{procesados} de {total}"
        )

        self.texto.setText(
            mensaje
        )

    def ocultar(self):

        self.animacion_timer.stop()

        self.hide()


# =========================================================
# FILA MODERNA DE RESOLUCIÓN
# =========================================================

class ResolutionRow(QWidget):

    # Señal independiente del checkbox.
    # Sirve para que hacer clic sobre los datos abra la resolución
    # y cargue el PDF, sin confundirlo con la selección para correo.
    clicked = Signal(object)

    # Señal exclusiva del botón VER DETALLE.
    detalle_clicked = Signal(object)

    def __init__(self, registro, item, parent=None):

        super().__init__(parent)

        self.registro = registro
        self.item = item

        self.setObjectName("resolutionRow")
        self.setMinimumHeight(76)

        # El propio contenedor recibe el clic.
        self.setCursor(
            Qt.CursorShape.PointingHandCursor
        )

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(10)

        # CHECKBOX
        self.checkbox = QCheckBox()
        self.checkbox.setObjectName("resolutionCheck")
        self.checkbox.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        self.checkbox.setToolTip(
            "Seleccionar resolución para enviar"
        )

        self.checkbox.stateChanged.connect(
            self._sincronizar_item
        )

        layout.addWidget(
            self.checkbox,
            0,
            Qt.AlignmentFlag.AlignVCenter
        )

        # CONTENIDO
        contenido = QVBoxLayout()
        contenido.setContentsMargins(0, 0, 0, 0)
        contenido.setSpacing(4)

        fecha = str(
            registro.get("fecha", "")
        ).strip()

        socio = str(
            registro.get("socio", "")
        ).strip() or "Sin socio"

        placa = str(
            registro.get("placa", "")
        ).strip() or "Sin placa"

        sri = str(
            registro.get("valores_sri", "")
        ).strip() or "NO"

        estado = str(
            registro.get(
                "estado",
                ESTADO_PENDIENTE
            )
        ).strip().upper() or ESTADO_PENDIENTE

        self.lbl_principal = QLabel(
            f"{fecha}    •    {socio}"
        )

        self.lbl_principal.setObjectName(
            "resolutionMain"
        )

        self.lbl_secundario = QLabel(
            f"Placa: {placa}    |    SRI: {sri}    |    Estado: {estado}"
        )

        self.lbl_secundario.setObjectName(
            "resolutionSecondary"
        )

        # Los textos no deben comerse el clic de la fila.
        self.lbl_principal.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True
        )

        self.lbl_secundario.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True
        )

        contenido.addWidget(
            self.lbl_principal
        )

        contenido.addWidget(
            self.lbl_secundario
        )

        layout.addLayout(
            contenido,
            1
        )

        # -------------------------------------------------
        # BOTÓN VER DETALLE
        # -------------------------------------------------

        self.btn_detalle = QPushButton(
            "VER DETALLE"
        )

        self.btn_detalle.setObjectName(
            "btnVerDetalle"
        )

        self.btn_detalle.setCursor(
            Qt.CursorShape.PointingHandCursor
        )

        self.btn_detalle.setMinimumSize(
            108,
            34
        )

        self.btn_detalle.setMaximumWidth(
            118
        )

        self.btn_detalle.setToolTip(
            "Ver todos los datos de esta resolución"
        )

        self.btn_detalle.clicked.connect(
            self._abrir_detalle
        )

        layout.addWidget(
            self.btn_detalle,
            0,
            Qt.AlignmentFlag.AlignVCenter
        )

    def _abrir_detalle(self):

        self.detalle_clicked.emit(
            self.registro
        )

    def mousePressEvent(self, event):

        if event.button() == Qt.MouseButton.LeftButton:

            # Clic sobre los datos = visualizar resolución/PDF.
            self.clicked.emit(
                self.item
            )

        super().mousePressEvent(event)

    def _sincronizar_item(self, estado):

        if not self.item:
            return

        if estado == int(
            Qt.CheckState.Checked.value
        ):

            self.item.setCheckState(
                Qt.CheckState.Checked
            )

        else:

            self.item.setCheckState(
                Qt.CheckState.Unchecked
            )

    def actualizar_checkbox(self):

        if not self.item:
            return

        estado = (
            self.item.checkState()
            == Qt.CheckState.Checked
        )

        self.checkbox.blockSignals(True)

        self.checkbox.setChecked(
            estado
        )

        self.checkbox.blockSignals(False)


# =========================================================
# VENTANA DE DETALLE DE RESOLUCIÓN
# =========================================================
class ImagenEvidenciaDialog(QDialog):

    def __init__(self, registro, parent=None):

        super().__init__(parent)

        self.registro = registro or {}
        self.pixmap_original = QPixmap()

        self.setWindowTitle(
            "Evidencia SRI"
        )

        self.resize(
            760,
            620
        )

        self.setMinimumSize(
            500,
            400
        )

        layout = QVBoxLayout(
            self
        )

        layout.setContentsMargins(
            12,
            12,
            12,
            12
        )

        layout.setSpacing(
            8
        )

        titulo = QLabel(
            "IMAGEN DE EVIDENCIA SRI"
        )

        titulo.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        titulo.setFont(
            QFont(
                "Arial",
                12,
                QFont.Weight.Bold
            )
        )

        layout.addWidget(
            titulo
        )

        self.lbl_estado = QLabel(
            "Cargando imagen..."
        )

        self.lbl_estado.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        layout.addWidget(
            self.lbl_estado
        )

        self.scroll = QScrollArea()

        self.scroll.setWidgetResizable(
            True
        )

        self.scroll.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.scroll.setFrameShape(
            QFrame.Shape.StyledPanel
        )

        self.lbl_imagen = QLabel()

        self.lbl_imagen.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.lbl_imagen.setText(
            "Cargando..."
        )

        self.lbl_imagen.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        self.scroll.setWidget(
            self.lbl_imagen
        )

        layout.addWidget(
            self.scroll,
            1
        )

        btn = QPushButton(
            "CERRAR"
        )

        btn.setMinimumHeight(
            36
        )

        btn.clicked.connect(
            self.accept
        )

        layout.addWidget(
            btn
        )

        self._cargar_imagen()

    def _cargar_imagen(self):

        # Primero se intenta usar una imagen Base64 que ya venga
        # directamente en el registro.
        evidencia = str(
            self.registro.get(
                "evidencia_sri",
                ""
            )
        ).strip()

        try:

            if evidencia.startswith(
                "data:image/"
            ):

                datos = evidencia.split(
                    ",",
                    1
                )[1]

                pix = QPixmap()

                if pix.loadFromData(
                    base64.b64decode(
                        datos
                    )
                ):

                    self._mostrar_pixmap(
                        pix
                    )

                    return

        except Exception:
            pass

        # También se acepta Base64 puro si el registro lo contiene.
        try:

            if len(evidencia) > 100:

                datos = base64.b64decode(
                    evidencia,
                    validate=True
                )

                pix = QPixmap()

                if pix.loadFromData(
                    datos
                ):

                    self._mostrar_pixmap(
                        pix
                    )

                    return

        except Exception:
            pass

        # La imagen persistida por el Apps Script se solicita por
        # el ID del PDF. No se usa Drive para visualizarla.
        pdf_id = str(
            self.registro.get(
                "pdf_id",
                ""
            )
        ).strip()

        if not pdf_id:

            self._error(
                "El registro no tiene ID de PDF para localizar la evidencia SRI."
            )

            return

        self.lbl_estado.setText(
            "Consultando evidencia SRI en Base64..."
        )

        worker = RequestWorker(
            GOOGLE_SHEETS_URL,
            {
                "accion": "evidencia_base64",
                "pdf_id": pdf_id
            },
            TIMEOUT_PDF
        )

        self.worker = worker

        worker.signals.finished.connect(
            self._procesar_respuesta
        )

        worker.signals.error.connect(
            self._error
        )

        QThreadPool.globalInstance().start(
            worker
        )

    def _procesar_respuesta(
        self,
        resultado
    ):

        try:

            if resultado.get(
                "estado"
            ) != "ok":

                raise Exception(
                    resultado.get(
                        "mensaje",
                        "No se pudo obtener la evidencia SRI."
                    )
                )

            data = resultado.get(
                "data",
                ""
            )

            if data:

                pix = QPixmap()

                if pix.loadFromData(
                    base64.b64decode(
                        data
                    )
                ):

                    self._mostrar_pixmap(
                        pix
                    )

                    return

            url = resultado.get(
                "url",
                ""
            )

            if url:

                self.lbl_estado.setText(
                    "Evidencia SRI disponible"
                )

                self.lbl_imagen.setText(
                    url
                )

                return

            raise Exception(
                "El servidor no devolvió una imagen válida."
            )

        except Exception as error:

            self._error(
                str(error)
            )

    def _mostrar_pixmap(
        self,
        pix
    ):

        if pix.isNull():

            self._error(
                "La imagen recibida no es válida."
            )

            return

        self.pixmap_original = pix

        self.lbl_estado.setText(
            "Evidencia SRI cargada"
        )

        self.lbl_imagen.setText(
            ""
        )

        self._ajustar_imagen()

    def _ajustar_imagen(self):

        if self.pixmap_original.isNull():
            return

        area = self.scroll.viewport().size()

        margen = 20

        ancho = max(
            100,
            area.width() - margen
        )

        alto = max(
            100,
            area.height() - margen
        )

        pix = self.pixmap_original.scaled(
            ancho,
            alto,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )

        self.lbl_imagen.setPixmap(
            pix
        )

    def resizeEvent(
        self,
        event
    ):

        super().resizeEvent(
            event
        )

        self._ajustar_imagen()

    def _error(
        self,
        error
    ):

        self.lbl_estado.setText(
            "No se pudo cargar la evidencia SRI."
        )

        self.lbl_imagen.setText(
            str(error)
        )



class DetalleResolucionDialog(QDialog):

    def __init__(self, registro, parent=None):

        super().__init__(parent)

        self.registro = registro or {}

        self.setWindowTitle(
            "Detalle de la resolución"
        )

        self.setModal(False)

        self.setMinimumSize(
            390,
            520
        )

        self.resize(
            430,
            600
        )

        self.setWindowFlags(
            Qt.WindowType.Window
            |
            Qt.WindowType.WindowCloseButtonHint
            |
            Qt.WindowType.WindowMinimizeButtonHint
        )

        self._crear_interfaz()
        self.actualizar_datos(self.registro)
        self._colocar_a_la_izquierda()

    def _crear_interfaz(self):

        principal = QVBoxLayout(self)

        principal.setContentsMargins(
            14,
            14,
            14,
            14
        )

        principal.setSpacing(10)

        titulo = QLabel(
            "DETALLE DE LA RESOLUCIÓN"
        )

        titulo.setObjectName(
            "tituloDetalleVentana"
        )

        titulo.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        principal.addWidget(
            titulo
        )

        subtitulo = QLabel(
            "Información completa del registro seleccionado"
        )

        subtitulo.setObjectName(
            "subtituloDetalleVentana"
        )

        subtitulo.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        subtitulo.setWordWrap(True)

        principal.addWidget(
            subtitulo
        )

        scroll = QScrollArea()

        scroll.setWidgetResizable(True)
        scroll.setFrameShape(
            QFrame.Shape.NoFrame
        )

        contenido = QFrame()

        contenido.setObjectName(
            "contenidoDetalleVentana"
        )

        layout = QVBoxLayout(contenido)

        layout.setContentsMargins(
            8,
            8,
            8,
            8
        )

        layout.setSpacing(8)

        self.lbl_fecha = self._crear_campo(
            "FECHA"
        )

        self.lbl_socio = self._crear_campo(
            "SOCIO"
        )

        self.lbl_placa = self._crear_campo(
            "PLACA"
        )

        self.lbl_chasis = self._crear_campo(
            "CHASIS"
        )

        self.lbl_sri = self._crear_campo(
            "VALORES SRI"
        )

        self.lbl_evidencia = self._crear_campo_evidencia_sri()

        self.lbl_motivo_rechazo = self._crear_campo(
            "MOTIVO DEL RECHAZO"
        )

        self.lbl_pdf_id = self._crear_campo(
            "ID PDF"
        )

        self.lbl_pdf_url = self._crear_campo(
            "URL PDF"
        )

        for bloque in [
            self.lbl_fecha,
            self.lbl_socio,
            self.lbl_placa,
            self.lbl_chasis,
            self.lbl_sri,
            self.lbl_evidencia,
            self.lbl_motivo_rechazo,
            self.lbl_pdf_id,
            self.lbl_pdf_url
        ]:

            layout.addWidget(
                bloque
            )

        layout.addStretch()

        scroll.setWidget(
            contenido
        )

        principal.addWidget(
            scroll,
            1
        )

    def _crear_campo(self, titulo):

        frame = QFrame()

        frame.setObjectName(
            "campoDetalleVentana"
        )

        layout = QVBoxLayout(frame)

        layout.setContentsMargins(
            10,
            8,
            10,
            8
        )

        layout.setSpacing(3)

        lbl_titulo = QLabel(
            titulo
        )

        lbl_titulo.setObjectName(
            "tituloCampoDetalle"
        )

        lbl_valor = QLabel(
            "-"
        )

        lbl_valor.setObjectName(
            "valorCampoDetalle"
        )

        lbl_valor.setWordWrap(True)

        lbl_valor.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )

        layout.addWidget(
            lbl_titulo
        )

        layout.addWidget(
            lbl_valor
        )

        frame.lbl_valor = lbl_valor

        return frame

    def _crear_campo_evidencia_sri(self):

        frame = QFrame()

        frame.setObjectName(
            "campoDetalleVentana"
        )

        layout = QVBoxLayout(frame)

        layout.setContentsMargins(
            10,
            8,
            10,
            8
        )

        layout.setSpacing(5)

        lbl_titulo = QLabel(
            "EVIDENCIA SRI"
        )

        lbl_titulo.setObjectName(
            "tituloCampoDetalle"
        )

        layout.addWidget(
            lbl_titulo
        )

        self.btn_ver_evidencia = QPushButton(
            "VER IMAGEN DEL SRI"
        )

        self.btn_ver_evidencia.setObjectName(
            "btnVerEvidencia"
        )

        self.btn_ver_evidencia.setCursor(
            Qt.CursorShape.PointingHandCursor
        )

        self.btn_ver_evidencia.setMinimumHeight(
            36
        )

        self.btn_ver_evidencia.clicked.connect(
            self.ver_evidencia_sri
        )

        layout.addWidget(
            self.btn_ver_evidencia
        )

        lbl_valor = QLabel(
            "-"
        )

        lbl_valor.setObjectName(
            "valorCampoDetalle"
        )

        lbl_valor.setWordWrap(
            True
        )

        lbl_valor.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )

        layout.addWidget(
            lbl_valor
        )

        frame.lbl_valor = lbl_valor

        return frame


    def ver_evidencia_sri(self):

        valor = str(
            self.registro.get(
                "evidencia_sri",
                ""
            )
        ).strip()

        if not valor:

            QMessageBox.information(
                self,
                "Evidencia SRI",
                "Este registro no tiene una imagen de evidencia SRI."
            )

            return

        dialogo = ImagenEvidenciaDialog(
            self.registro,
            self
        )

        dialogo.exec()


    def _colocar_a_la_izquierda(self):

        ventana_principal = self.parentWidget()

        if not ventana_principal:
            return

        try:

            geo = ventana_principal.frameGeometry()

            ancho = self.width()

            separacion = 10

            x = (
                geo.left()
                -
                ancho
                -
                separacion
            )

            y = geo.top()

            pantalla = (
                QApplication.screenAt(
                    geo.center()
                )
                or QApplication.primaryScreen()
            )

            if pantalla:

                disponible = (
                    pantalla.availableGeometry()
                )

                if x < disponible.left():

                    x = disponible.left()

                if y < disponible.top():

                    y = disponible.top()

                if (
                    y + self.height()
                    >
                    disponible.bottom()
                ):

                    y = max(
                        disponible.top(),
                        disponible.bottom()
                        -
                        self.height()
                    )

            self.move(
                x,
                y
            )

        except Exception:
            pass

    def actualizar_datos(self, registro):

        self.registro = registro or {}

        self.lbl_fecha.lbl_valor.setText(
            str(
                self.registro.get(
                    "fecha",
                    ""
                )
            )
        )

        self.lbl_socio.lbl_valor.setText(
            str(
                self.registro.get(
                    "socio",
                    ""
                )
            )
        )

        self.lbl_placa.lbl_valor.setText(
            str(
                self.registro.get(
                    "placa",
                    ""
                )
            )
        )

        self.lbl_chasis.lbl_valor.setText(
            str(
                self.registro.get(
                    "chasis",
                    ""
                )
            )
        )

        self.lbl_sri.lbl_valor.setText(
            str(
                self.registro.get(
                    "valores_sri",
                    ""
                )
            )
        )

        evidencia_actual = str(
            self.registro.get(
                "evidencia_sri",
                ""
            )
        ).strip()

        self.lbl_evidencia.lbl_valor.setText(
            evidencia_actual
        )

        self.btn_ver_evidencia.setEnabled(
            bool(evidencia_actual)
        )

        estado_actual = str(
            self.registro.get(
                "estado",
                ESTADO_PENDIENTE
            )
        ).strip().upper()

        motivo_rechazo_actual = str(
            self.registro.get(
                "motivo_rechazo",
                ""
            )
        ).strip()

        self.lbl_motivo_rechazo.lbl_valor.setText(
            motivo_rechazo_actual
            if motivo_rechazo_actual
            else "-"
        )

        # El motivo de rechazo solamente se muestra cuando
        # el registro esta realmente en estado RECHAZADO.
        self.lbl_motivo_rechazo.setVisible(
            estado_actual == ESTADO_RECHAZADO
        )

        self.lbl_pdf_id.lbl_valor.setText(
            str(
                self.registro.get(
                    "pdf_id",
                    ""
                )
            )
        )

        self.lbl_pdf_url.lbl_valor.setText(
            str(
                self.registro.get(
                    "pdf_url",
                    ""
                )
            )
        )

        self._colocar_a_la_izquierda()


# =========================================================
# VENTANA ADMINISTRADOR
# =========================================================

class AdminWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle(
            "ADMIN - Resoluciones AMT"
        )

        self.resize(
            1450,
            850
        )

        # -------------------------------------------------
        # DATOS
        # -------------------------------------------------

        self.registros = []

        self.registro_actual = None

        # Ventana independiente de detalle. Se abre a la izquierda
        # de la ventana principal y no ocupa espacio del visor PDF.
        self.ventana_detalle = None

        self.fecha_desde = None

        self.fecha_hasta = None

        # Filtro actual de estado.
        self.filtro_estado_actual = "TODOS"

        self.archivo_temporal = None

        # -------------------------------------------------
        # CONTROL CONSULTAS
        # -------------------------------------------------

        self.actualizando = False

        self.cerrando = False

        self.thread_pool = (
            QThreadPool.globalInstance()
        )

        # -------------------------------------------------
        # CONTROL ENVÍO
        # -------------------------------------------------

        self.enviando_correos = False

        self.worker_correos = None

        # -------------------------------------------------
        # CACHÉ PDF
        # -------------------------------------------------

        self.cache_pdfs = {}

        self.pdf_id_actual = None

        # -------------------------------------------------
        # FIRMA DATOS
        # -------------------------------------------------

        self.firma_registros = ""
        # -------------------------------------------------
        # CERTIFICADO DE FIRMA
        # -------------------------------------------------
        # Se conserva la ruta entre ejecuciones.
        # La contraseña nunca se guarda.
        self.p12_path = (
            cargar_ruta_certificado_guardada()
        )


        # -------------------------------------------------
        # INTERFAZ
        # -------------------------------------------------

        self.init_ui()

        # -------------------------------------------------
        # CARGADOR
        # -------------------------------------------------

        self.loading_overlay = (
            LoadingOverlay(
                self.centralWidget()
            )
        )

        self.loading_overlay.setGeometry(
            self.centralWidget().rect()
        )

        self.centralWidget().installEventFilter(
            self
        )
        # -------------------------------------------------
        # CERTIFICADO DE FIRMA
        # -------------------------------------------------
        # NO se solicita el certificado al iniciar.
        # Si existe una ruta guardada, se conserva silenciosamente.
        # Si no existe, se solicitará únicamente al momento de
        # enviar/firma los documentos.
        # -------------------------------------------------


        # -------------------------------------------------
        # PRIMERA CARGA
        # -------------------------------------------------

        self.actualizar_lista(
            mostrar_errores=True
        )

        # -------------------------------------------------
        # ACTUALIZACIÓN AUTOMÁTICA
        # -------------------------------------------------

        self.timer_actualizacion = QTimer(
            self
        )

        self.timer_actualizacion.setInterval(
            INTERVALO_ACTUALIZACION
        )

        self.timer_actualizacion.timeout.connect(
            self.actualizar_automaticamente
        )

        self.timer_actualizacion.start()

    # =====================================================
    # REDIMENSIONAR OVERLAY
    # =====================================================

    def resizeEvent(
        self,
        event
    ):

        super().resizeEvent(
            event
        )

        if hasattr(
            self,
            "loading_overlay"
        ):

            self.loading_overlay.setGeometry(
                self.centralWidget().rect()
            )

        if (
            hasattr(self, "ventana_detalle")
            and
            self.ventana_detalle
            and
            self.ventana_detalle.isVisible()
        ):

            self.ventana_detalle._colocar_a_la_izquierda()

    # =====================================================
    # INTERFAZ
    # =====================================================

    def init_ui(self):

        central = QWidget()

        self.setCentralWidget(
            central
        )

        principal = QVBoxLayout(
            central
        )

        principal.setContentsMargins(
            10,
            8,
            10,
            10
        )

        principal.setSpacing(
            8
        )

        # -------------------------------------------------
        # CABECERA
        # -------------------------------------------------

        cabecera = QFrame()

        cabecera.setObjectName(
            "cabecera"
        )

        cabecera_layout = QHBoxLayout(
            cabecera
        )

        cabecera_layout.setContentsMargins(
            22,
            14,
            22,
            14
        )

        # -------------------------------------------------
        # LOGO + NOMBRE DEL SISTEMA
        # -------------------------------------------------

        logoSistema = QLabel()

        logoSistema.setObjectName(
            "logoSistema"
        )

        logoSistema.setFixedSize(
            62,
            62
        )

        logoSistema.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        logoSistema.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents,
            True
        )

        if LOGO_SISTEMA and os.path.isfile(LOGO_SISTEMA):

            try:

                pixmap_logo = QPixmap(
                    LOGO_SISTEMA
                )

                if not pixmap_logo.isNull():

                    pixmap_logo = pixmap_logo.scaled(
                        56,
                        56,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation
                    )

                    logoSistema.setPixmap(
                        pixmap_logo
                    )

            except Exception as error_logo:

                print(
                    "Advertencia: no se pudo cargar el logo:",
                    error_logo
                )

        cabecera_layout.addWidget(
            logoSistema,
            0,
            Qt.AlignmentFlag.AlignVCenter
        )

        bloque_titulo = QVBoxLayout()

        bloque_titulo.setSpacing(
            2
        )

        titulo = QLabel(
            "Solutions &sra"
        )

        titulo.setObjectName(
            "tituloPrincipal"
        )

        titulo.setFont(
            QFont(
                "Arial",
                20,
                QFont.Weight.Bold
            )
        )

        subtitulo = QLabel(
            "Panel de administración y consulta documental"
        )

        subtitulo.setObjectName(
            "subtituloPrincipal"
        )

        bloque_titulo.addWidget(
            titulo
        )

        bloque_titulo.addWidget(
            subtitulo
        )

        cabecera_layout.addLayout(
            bloque_titulo
        )

        cabecera_layout.addStretch()

        rol = QLabel(
            "ADMINISTRADOR"
        )

        rol.setObjectName(
            "badgeRol"
        )

        rol.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        cabecera_layout.addWidget(
            rol
        )

        principal.addWidget(
            cabecera
        )

        # -------------------------------------------------
        # BARRA ESTADO
        # -------------------------------------------------

        estado_frame = QFrame()

        estado_frame.setObjectName(
            "barraEstado"
        )

        estado_layout = QHBoxLayout(
            estado_frame
        )

        estado_layout.setContentsMargins(
            14,
            7,
            14,
            7
        )

        self.lbl_estado = QLabel(
            "🟡 Conectando..."
        )

        self.lbl_estado.setObjectName(
            "estadoConexion"
        )

        self.lbl_estado.setFont(
            QFont(
                "Arial",
                9,
                QFont.Weight.Bold
            )
        )

        self.lbl_ultima_actualizacion = QLabel(
            "Última actualización: --"
        )

        self.lbl_ultima_actualizacion.setAlignment(
            Qt.AlignmentFlag.AlignRight
        )

        estado_layout.addWidget(
            self.lbl_estado
        )

        estado_layout.addStretch()

        estado_layout.addWidget(
            self.lbl_ultima_actualizacion
        )

        principal.addWidget(
            estado_frame
        )

        # -------------------------------------------------
        # SPLITTER
        # -------------------------------------------------

        splitter = QSplitter(
            Qt.Orientation.Horizontal
        )

        splitter.setChildrenCollapsible(
            False
        )

        splitter.setHandleWidth(
            3
        )

        principal.addWidget(
            splitter,
            1
        )

        # =================================================
        # PANEL IZQUIERDO
        # =================================================

        panel_izquierdo = QFrame()

        panel_izquierdo.setObjectName(
            "panelIzquierdo"
        )

        layout_izq = QVBoxLayout(
            panel_izquierdo
        )

        layout_izq.setContentsMargins(
            8,
            8,
            8,
            8
        )

        layout_izq.setSpacing(
            7
        )

        label_lista = QLabel(
            "LISTA DE RESOLUCIONES"
        )

        label_lista.setFont(
            QFont(
                "Arial",
                11,
                QFont.Weight.Bold
            )
        )

        layout_izq.addWidget(
            label_lista
        )

        self.txt_buscar = QLineEdit()

        self.txt_buscar.setPlaceholderText(
            "🔎 Buscar..."
        )

        self.txt_buscar.setMinimumHeight(
            38
        )

        self.txt_buscar.textChanged.connect(
            self.filtrar_lista
        )

        layout_izq.addWidget(
            self.txt_buscar
        )

        # -------------------------------------------------
        # FILTRO FECHA
        # -------------------------------------------------

        filtro_fecha = QFrame()

        filtro_fecha.setObjectName(
            "filtroFecha"
        )

        filtro_layout = QVBoxLayout(
            filtro_fecha
        )

        filtro_layout.setContentsMargins(
            10,
            8,
            10,
            8
        )

        filtro_layout.setSpacing(
            6
        )

        lbl_filtro = QLabel(
            "FILTRAR POR FECHA"
        )

        lbl_filtro.setObjectName(
            "tituloFiltro"
        )

        filtro_layout.addWidget(
            lbl_filtro
        )

        fechas_layout = QHBoxLayout()

        fechas_layout.setSpacing(
            6
        )

        lbl_desde = QLabel(
            "Desde"
        )

        lbl_desde.setObjectName(
            "labelFecha"
        )

        lbl_hasta = QLabel(
            "Hasta"
        )

        lbl_hasta.setObjectName(
            "labelFecha"
        )

        self.fecha_desde_edit = QDateEdit()

        self.fecha_desde_edit.setCalendarPopup(
            True
        )

        self.fecha_desde_edit.setDisplayFormat(
            "dd/MM/yyyy"
        )

        self.fecha_desde_edit.setDate(
            QDate(
                2026,
                1,
                1
            )
        )

        self.fecha_desde_edit.setMinimumWidth(
            105
        )

        self.fecha_hasta_edit = QDateEdit()

        self.fecha_hasta_edit.setCalendarPopup(
            True
        )

        self.fecha_hasta_edit.setDisplayFormat(
            "dd/MM/yyyy"
        )

        self.fecha_hasta_edit.setDate(
            QDate.currentDate()
        )

        self.fecha_hasta_edit.setMinimumWidth(
            105
        )

        self.btn_aplicar_fecha = QPushButton(
            "Aplicar"
        )

        self.btn_aplicar_fecha.setObjectName(
            "btnFiltro"
        )

        self.btn_aplicar_fecha.clicked.connect(
            self.aplicar_filtro_fecha
        )

        self.btn_limpiar_fecha = QPushButton(
            "Limpiar"
        )

        self.btn_limpiar_fecha.setObjectName(
            "btnLimpiarFiltro"
        )

        self.btn_limpiar_fecha.clicked.connect(
            self.limpiar_filtro_fecha
        )

        fechas_layout.addWidget(
            lbl_desde
        )

        fechas_layout.addWidget(
            self.fecha_desde_edit
        )

        fechas_layout.addWidget(
            lbl_hasta
        )

        fechas_layout.addWidget(
            self.fecha_hasta_edit
        )

        fechas_layout.addWidget(
            self.btn_aplicar_fecha
        )

        fechas_layout.addWidget(
            self.btn_limpiar_fecha
        )

        filtro_layout.addLayout(
            fechas_layout
        )

        layout_izq.addWidget(
            filtro_fecha
        )

        # -------------------------------------------------
        # FILTRO POR ESTADO
        # -------------------------------------------------

        estado_layout = QHBoxLayout()
        estado_layout.setSpacing(
            5
        )

        self.btn_estado_todos = QPushButton(
            "TODOS"
        )

        self.btn_estado_pendientes = QPushButton(
            "PENDIENTES"
        )

        self.btn_estado_enviados = QPushButton(
            "ENVIADOS"
        )

        self.btn_estado_rechazados = QPushButton(
            "RECHAZADOS"
        )

        self.btn_estado_todos.clicked.connect(
            lambda: self.seleccionar_filtro_estado("TODOS")
        )

        self.btn_estado_pendientes.clicked.connect(
            lambda: self.seleccionar_filtro_estado(ESTADO_PENDIENTE)
        )

        self.btn_estado_enviados.clicked.connect(
            lambda: self.seleccionar_filtro_estado(ESTADO_ENVIADO)
        )

        self.btn_estado_rechazados.clicked.connect(
            lambda: self.seleccionar_filtro_estado(ESTADO_RECHAZADO)
        )

        for boton in (
            self.btn_estado_todos,
            self.btn_estado_pendientes,
            self.btn_estado_enviados,
            self.btn_estado_rechazados
        ):
            boton.setMinimumHeight(
                34
            )
            estado_layout.addWidget(
                boton
            )

        layout_izq.addLayout(
            estado_layout
        )

        self.lbl_contador = QLabel(
            "Mostrando 0 de 0 resoluciones"
        )

        self.lbl_contador.setObjectName(
            "contadorLista"
        )

        layout_izq.addWidget(
            self.lbl_contador
        )

        # -------------------------------------------------
        # LISTA
        # -------------------------------------------------

        self.lista = QListWidget()

        self.lista.setSpacing(
            1
        )

        self.lista.itemClicked.connect(
            self.seleccionar_registro
        )

        layout_izq.addWidget(
            self.lista,
            1
        )

        # -------------------------------------------------
        # ACTUALIZAR
        # -------------------------------------------------

        self.btn_actualizar = QPushButton(
            "↻  ACTUALIZAR DATOS"
        )

        self.btn_actualizar.setMinimumHeight(
            38
        )

        self.btn_actualizar.clicked.connect(
            self.actualizar_lista_manual
        )

        layout_izq.addWidget(
            self.btn_actualizar
        )

        # -------------------------------------------------
        # CORREOS
        # -------------------------------------------------

        correo_layout = QHBoxLayout()

        correo_layout.setSpacing(
            6
        )

        self.btn_seleccionar_todos = QPushButton(
            "☑  SELECCIONAR TODOS"
        )

        self.btn_seleccionar_todos.setMinimumHeight(
            40
        )

        self.btn_seleccionar_todos.clicked.connect(
            self.seleccionar_todos_visibles
        )

        self.btn_enviar_correos = QPushButton(
            "📧  ENVIAR CORREOS"
        )

        self.btn_enviar_correos.setObjectName(
            "btnEnviarCorreo"
        )

        self.btn_enviar_correos.setMinimumHeight(
            40
        )

        self.btn_enviar_correos.clicked.connect(
            self.enviar_correos_seleccionados
        )

        correo_layout.addWidget(
            self.btn_seleccionar_todos
        )

        correo_layout.addWidget(
            self.btn_enviar_correos
        )

        self.btn_rechazar = QPushButton(
            "✖  RECHAZAR"
        )

        self.btn_rechazar.setObjectName(
            "btnRechazar"
        )

        self.btn_rechazar.setMinimumHeight(
            40
        )

        self.btn_rechazar.clicked.connect(
            self.rechazar_registro_actual
        )

        correo_layout.addWidget(
            self.btn_rechazar
        )

        layout_izq.addLayout(
            correo_layout
        )

        # =================================================
        # PANEL DERECHO
        # =================================================

        panel_derecho = QFrame()

        panel_derecho.setObjectName(
            "panelDerecho"
        )

        layout_der = QVBoxLayout(
            panel_derecho
        )

        layout_der.setContentsMargins(
            8,
            8,
            8,
            8
        )

        layout_der.setSpacing(
            7
        )

        titulo_pdf = QLabel(
            "RESOLUCIÓN SELECCIONADA"
        )

        titulo_pdf.setFont(
            QFont(
                "Arial",
                11,
                QFont.Weight.Bold
            )
        )

        titulo_pdf.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        titulo_pdf.setMinimumHeight(
            30
        )

        layout_der.addWidget(
            titulo_pdf
        )

        self.pdf_viewer = PDFViewer()

        layout_der.addWidget(
            self.pdf_viewer,
            1
        )

        self.btn_descargar = QPushButton(
            "📥 DESCARGAR PDF"
        )

        self.btn_descargar.setMinimumHeight(
            40
        )

        self.btn_descargar.setEnabled(
            False
        )

        self.btn_descargar.clicked.connect(
            self.descargar_pdf
        )

        layout_der.addWidget(
            self.btn_descargar
        )

        splitter.addWidget(
            panel_izquierdo
        )

        splitter.addWidget(
            panel_derecho
        )

        splitter.setSizes([
            390,
            1040
        ])

        # =================================================
        # ESTILO
        # =================================================

        self.setStyleSheet("""

            QMainWindow {
                background: #f4f7fb;
            }

            QWidget {
                background: #f4f7fb;
                color: #17212b;
                font-family: "Segoe UI", Arial;
                font-size: 10pt;
            }

            QLabel {
                color: #17212b;
                background: transparent;
            }

            #cabecera {
                background: #ffffff;
                border: 1px solid #dce5ef;
                border-radius: 12px;
            }

            #logoSistema {
                background: transparent;
                border: none;
            }

            #tituloPrincipal {
                color: #17324d;
                font-size: 20pt;
                font-weight: 700;
            }

            #subtituloPrincipal {
                color: #718096;
                font-size: 9.5pt;
            }

            #badgeRol {
                background: #17324d;
                color: #ffffff;
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: 700;
            }

            #barraEstado {
                background: #ffffff;
                border: 1px solid #dce5ef;
                border-radius: 9px;
            }

            #estadoConexion {
                color: #1f7a4d;
            }

            #panelIzquierdo,
            #panelDerecho {
                background: #ffffff;
                border: 1px solid #dce5ef;
                border-radius: 12px;
            }

            #panelIzquierdo {
                padding: 2px;
            }

            #datosFrame {
                background: #f8fafc;
                border: 1px solid #d9e4ee;
                border-radius: 11px;
            }

            #tituloDatos {
                color: #17324d;
                font-size: 10pt;
                font-weight: 800;
                padding: 2px 2px 5px 2px;
            }

            #datoResolucion {
                background: #ffffff;
                border: 1px solid #e3eaf1;
                border-radius: 7px;
                padding: 7px 9px;
                color: #33485b;
                font-size: 9pt;
            }

            #datoResolucion:hover {
                border: 1px solid #b9ccdd;
                background: #fbfdff;
            }

            QLineEdit {
                padding: 9px 11px;
                border: 1px solid #ccd8e5;
                border-radius: 8px;
                background: #ffffff;
                color: #17212b;
                selection-background-color: #d9e9f7;
                selection-color: #17212b;
            }

            QLineEdit:focus {
                border: 1px solid #4f8bbd;
            }

            #filtroFecha {
                background: #f8fafc;
                border: 1px solid #e1e8f0;
                border-radius: 9px;
            }

            #tituloFiltro {
                color: #526579;
                font-size: 8.5pt;
                font-weight: 700;
            }

            #contadorLista {
                color: #718096;
                font-size: 8.5pt;
                padding-left: 3px;
            }

            QDateEdit {
                padding: 7px 9px;
                border: 1px solid #ccd8e5;
                border-radius: 8px;
                background: #ffffff;
                color: #263746;
                font-size: 9pt;
                font-weight: 600;
            }

            QDateEdit:hover {
                border: 1px solid #8eabc2;
                background: #fbfdff;
            }

            QDateEdit:focus {
                border: 2px solid #3978a8;
                background: #ffffff;
            }

            QDateEdit::drop-down {
                width: 28px;
                border: none;
                border-left: 1px solid #e1e8f0;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }

            #labelFecha {
                color: #718096;
                font-size: 8pt;
                font-weight: 600;
            }

            QListWidget {
                background: #f8fafc;
                border: 1px solid #dce5ef;
                border-radius: 10px;
                outline: none;
                padding: 5px;
            }

            QListWidget::item {
                margin: 2px 1px;
                padding: 0px;
                border: 1px solid #e3eaf1;
                border-radius: 9px;
                background: #ffffff;
            }

            QListWidget::item:hover {
                background: #f7fbff;
                border: 1px solid #bfd2e2;
            }

            QListWidget::item:selected {
                background: #eaf3fb;
                color: #17212b;
                border: 1px solid #3978a8;
                border-left: 4px solid #3978a8;
            }

            #resolutionRow {
                background: transparent;
            }

            #resolutionMain {
                background: transparent;
                color: #18354f;
                font-size: 9.3pt;
                font-weight: 700;
            }

            #resolutionSecondary {
                background: transparent;
                color: #64778a;
                font-size: 8.7pt;
                font-weight: 600;
            }

            QCheckBox#resolutionCheck {
                background: transparent;
                spacing: 0px;
            }

            QCheckBox#resolutionCheck::indicator {
                width: 20px;
                height: 20px;
                border: 2px solid #9db0c2;
                border-radius: 5px;
                background: #ffffff;
            }

            QCheckBox#resolutionCheck::indicator:hover {
                border: 2px solid #3978a8;
                background: #f1f7fc;
            }

            QCheckBox#resolutionCheck::indicator:checked {
                border: 2px solid #17324d;
                background: #17324d;
            }

            QCheckBox#resolutionCheck::indicator:checked:hover {
                border: 2px solid #234d70;
                background: #234d70;
            }

            #btnVerDetalle {
                padding: 6px 10px;
                border-radius: 7px;
                border: 1px solid #3978a8;
                background: #eef6fc;
                color: #245b84;
                font-size: 8.5pt;
                font-weight: 700;
            }

            #btnVerDetalle:hover {
                background: #dceefb;
                border: 1px solid #245b84;
            }

            #btnVerDetalle:pressed {
                background: #cfe4f4;
            }

            #tituloDetalleVentana {
                color: #17324d;
                font-size: 13pt;
                font-weight: 800;
            }

            #subtituloDetalleVentana {
                color: #718096;
                font-size: 9pt;
            }

            #contenidoDetalleVentana {
                background: #f8fafc;
            }

            #campoDetalleVentana {
                background: #ffffff;
                border: 1px solid #dce5ef;
                border-radius: 9px;
            }

            #tituloCampoDetalle {
                color: #6b7c8e;
                font-size: 8pt;
                font-weight: 800;
            }

            #valorCampoDetalle {
                color: #18354f;
                font-size: 10pt;
                font-weight: 600;
            }

            QPushButton {
                padding: 8px 12px;
                border-radius: 8px;
                border: 1px solid #ccd8e5;
                background: #ffffff;
                color: #263746;
                font-weight: 600;
            }

            QPushButton:hover {
                background: #f1f6fa;
                border: 1px solid #8eabc2;
            }

            QPushButton:pressed {
                background: #e5eef6;
            }

            QPushButton:disabled {
                background: #f1f3f5;
                color: #9aa4ae;
                border: 1px solid #e0e4e8;
            }

            #btnRechazar {
                background: #fff4f4;
                color: #9f2d2d;
                border: 1px solid #e0aaaa;
                font-weight: 700;
            }

            #btnRechazar:hover {
                background: #ffe6e6;
                border: 1px solid #c97979;
            }

            #btnRechazar:pressed {
                background: #ffd8d8;
            }

            #btnEnviarCorreo {
                background: #17324d;
                color: #ffffff;
                border: 1px solid #17324d;
                font-weight: 700;
            }

            #btnEnviarCorreo:hover {
                background: #234d70;
                border: 1px solid #234d70;
            }

            #btnEnviarCorreo:pressed {
                background: #10263a;
            }

            #btnFiltro {
                background: #17324d;
                color: #ffffff;
                border: 1px solid #17324d;
            }

            #btnFiltro:hover {
                background: #234d70;
                border: 1px solid #234d70;
            }

            QPushButton[activo="true"] {
                background: #17324d;
                color: #ffffff;
                border: 1px solid #17324d;
                font-weight: 800;
            }

            QPushButton[activo="true"]:hover {
                background: #234d70;
                border: 1px solid #234d70;
            }

            #btnLimpiarFiltro {
                color: #526579;
                background: #ffffff;
                border: 1px solid #ccd8e5;
            }

            #btnLimpiarFiltro:hover {
                background: #f1f6fa;
                border: 1px solid #8eabc2;
            }

            QScrollArea {
                background: #f7f9fb;
                border: 1px solid #dce5ef;
                border-radius: 9px;
            }

            QScrollBar:vertical {
                background: #eef2f6;
                width: 12px;
                border: none;
            }

            QScrollBar::handle:vertical {
                background: #b9c8d6;
                min-height: 30px;
                border-radius: 6px;
            }

            QScrollBar::handle:vertical:hover {
                background: #8fa8bc;
            }

            QScrollBar:horizontal {
                background: #eef2f6;
                height: 12px;
                border: none;
            }

            QScrollBar::handle:horizontal {
                background: #b9c8d6;
                min-width: 30px;
                border-radius: 6px;
            }

            QSplitter::handle {
                background: #dce5ef;
            }

        """)

    # =====================================================
    # ACTUALIZACIÓN MANUAL
    # =====================================================

    def actualizar_lista_manual(self):

        self.actualizar_lista(
            mostrar_errores=True
        )

    # =====================================================
    # ACTUALIZACIÓN AUTOMÁTICA
    # =====================================================

    def actualizar_automaticamente(self):

        if self.enviando_correos:
            return

        self.actualizar_lista(
            mostrar_errores=False
        )

    # =====================================================
    # ACTUALIZAR LISTA
    # =====================================================

    def actualizar_lista(
        self,
        mostrar_errores=False
    ):

        if self.cerrando:
            return

        if self.actualizando:
            return

        self.actualizando = True

        self.btn_actualizar.setEnabled(
            False
        )

        self.lbl_estado.setText(
            "🟡 Actualizando..."
        )

        worker = RequestWorker(

            GOOGLE_SHEETS_URL,

            {
                "accion": "listar"
            },

            TIMEOUT_LISTA
        )

        self.worker_actual = worker

        worker.signals.finished.connect(

            lambda resultado:

            self.procesar_respuesta_lista(
                resultado,
                mostrar_errores
            )
        )

        worker.signals.error.connect(

            lambda error:

            self.error_actualizacion_lista(
                error,
                mostrar_errores
            )
        )

        self.thread_pool.start(
            worker
        )

    # =====================================================
    # PROCESAR LISTA
    # =====================================================

    def procesar_respuesta_lista(
        self,
        resultado,
        mostrar_errores
    ):

        if self.cerrando:
            return

        try:

            if resultado.get(
                "estado"
            ) != "ok":

                raise Exception(
                    resultado.get(
                        "mensaje",
                        "Error desconocido."
                    )
                )

            nuevos_registros = (
                resultado.get(
                    "registros",
                    []
                )
            )

            nueva_firma = (
                self.crear_firma_registros(
                    nuevos_registros
                )
            )

            if (
                self.firma_registros
                and
                nueva_firma
                ==
                self.firma_registros
            ):

                self.registros = (
                    nuevos_registros
                )

                self.actualizar_contadores_estado()
                self.actualizar_botones_estado()

                self.lbl_estado.setText(
                    "🟢 Conectado · Sin cambios"
                )

                self.lbl_ultima_actualizacion.setText(
                    "Última consulta: "
                    +
                    self.hora_actual()
                )

                return

            registros_anteriores = (
                self.registros
            )

            self.firma_registros = (
                nueva_firma
            )

            self.registros = (
                nuevos_registros
            )

            self.actualizar_contadores_estado()
            self.actualizar_botones_estado()

            cantidad_nuevos = (
                self.contar_registros_nuevos(
                    registros_anteriores,
                    nuevos_registros
                )
            )

            id_registro_actual = None

            if self.registro_actual:

                id_registro_actual = (
                    self.obtener_identificador_registro(
                        self.registro_actual
                    )
                )

            self.mostrar_lista(
                id_registro_actual
            )

            if cantidad_nuevos > 0:

                texto = (
                    f"🟢 Conectado · "
                    f"{cantidad_nuevos} registro(s) nuevo(s)"
                )

            else:

                texto = (
                    "🟢 Conectado · "
                    "Datos actualizados"
                )

            self.lbl_estado.setText(
                texto
            )

            self.lbl_ultima_actualizacion.setText(
                "Última actualización: "
                +
                self.hora_actual()
            )

        except Exception as e:

            self.error_actualizacion_lista(
                str(e),
                mostrar_errores
            )

        finally:

            self.actualizando = False

            self.btn_actualizar.setEnabled(
                True
            )

            self.worker_actual = None

    # =====================================================
    # ERROR LISTA
    # =====================================================

    def error_actualizacion_lista(
        self,
        error,
        mostrar_errores
    ):

        if self.cerrando:
            return

        self.actualizando = False

        self.btn_actualizar.setEnabled(
            True
        )

        self.lbl_estado.setText(
            "🔴 Sin conexión"
        )

        if mostrar_errores:

            QMessageBox.critical(
                self,
                "Error",
                "No se pudo cargar la lista:\n\n"
                +
                str(error)
            )

        else:

            print(
                "Error actualización automática:",
                error
            )

        self.worker_actual = None

    # =====================================================
    # FIRMA
    # =====================================================

    def crear_firma_registros(
        self,
        registros
    ):

        partes = []

        for registro in registros:

            partes.append(
                "|".join([

                    str(
                        registro.get(
                            "fila",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "fecha",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "socio",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "placa",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "chasis",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "valores_sri",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "evidencia_sri",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "pdf_id",
                            ""
                        )
                    ),

                    str(
                        registro.get(
                            "estado",
                            ESTADO_PENDIENTE
                        )
                    ),

                    str(
                        registro.get(
                            "motivo_rechazo",
                            ""
                        )
                    )
                ])
            )

        contenido = "\n".join(
            partes
        )

        return hashlib.sha256(
            contenido.encode(
                "utf-8"
            )
        ).hexdigest()

    # =====================================================
    # CONTAR NUEVOS
    # =====================================================

    def contar_registros_nuevos(
        self,
        registros_anteriores,
        registros_nuevos
    ):

        anteriores = set()

        for registro in registros_anteriores:

            anteriores.add(
                self.obtener_identificador_registro(
                    registro
                )
            )

        nuevos = 0

        for registro in registros_nuevos:

            identificador = (
                self.obtener_identificador_registro(
                    registro
                )
            )

            if identificador not in anteriores:

                nuevos += 1

        return nuevos

    # =====================================================
    # HORA
    # =====================================================

    def hora_actual(self):

        return datetime.now().strftime(
            "%H:%M:%S"
        )

    # =====================================================
    # IDENTIFICADOR
    # =====================================================

    def obtener_identificador_registro(
        self,
        registro
    ):

        pdf_id = str(
            registro.get(
                "pdf_id",
                ""
            )
        ).strip()

        if pdf_id:

            return (
                "PDF:"
                +
                pdf_id
            )

        fila = str(
            registro.get(
                "fila",
                ""
            )
        ).strip()

        if fila:

            return (
                "FILA:"
                +
                fila
            )

        return "|".join([

            str(
                registro.get(
                    "fecha",
                    ""
                )
            ),

            str(
                registro.get(
                    "socio",
                    ""
                )
            ),

            str(
                registro.get(
                    "placa",
                    ""
                )
            ),

            str(
                registro.get(
                    "chasis",
                    ""
                )
            )
        ])

    # =====================================================
    # MOSTRAR LISTA
    # =====================================================

    def mostrar_lista(
        self,
        id_registro_actual=None
    ):

        self.lista.blockSignals(
            True
        )

        self.lista.clear()

        indice_a_seleccionar = -1

        for indice, registro in enumerate(
            self.registros
        ):

            item = QListWidgetItem()

            # El estado se mantiene en el item para que el
            # resto del sistema siga funcionando igual.
            item.setFlags(
                item.flags()
                |
                Qt.ItemFlag.ItemIsUserCheckable
            )

            item.setCheckState(
                Qt.CheckState.Unchecked
            )

            item.setData(
                Qt.ItemDataRole.UserRole,
                registro
            )

            fila = ResolutionRow(
                registro,
                item
            )

            # IMPORTANTE:
            # Como la fila es un QWidget personalizado, QListWidget.itemClicked
            # no siempre recibe el clic. Conectamos explícitamente el clic
            # de la fila para visualizar el PDF y los datos.
            fila.clicked.connect(
                self.seleccionar_registro
            )

            fila.detalle_clicked.connect(
                self.abrir_detalle_registro
            )

            item.setSizeHint(
                QSize(
                    0,
                    82
                )
            )

            self.lista.addItem(
                item
            )

            self.lista.setItemWidget(
                item,
                fila
            )

            if id_registro_actual:

                id_nuevo = (
                    self.obtener_identificador_registro(
                        registro
                    )
                )

                if (
                    id_nuevo
                    ==
                    id_registro_actual
                ):

                    indice_a_seleccionar = (
                        indice
                    )

        if indice_a_seleccionar >= 0:

            self.lista.setCurrentRow(
                indice_a_seleccionar
            )

            item = self.lista.item(
                indice_a_seleccionar
            )

            if item:

                registro = item.data(
                    Qt.ItemDataRole.UserRole
                )

                self.registro_actual = (
                    registro
                )


        else:

            if id_registro_actual:

                self.registro_actual = None

                self.pdf_viewer.limpiar()

                self.pdf_id_actual = None

                self.btn_descargar.setEnabled(
                    False
                )

        self.lista.blockSignals(
            False
        )

        self.filtrar_lista()

    # =====================================================
    # FECHA
    # =====================================================

    def _fecha_registro(
        self,
        registro
    ):

        valor = str(
            registro.get(
                "fecha",
                ""
            )
        ).strip()

        formatos = (

            "%d/%m/%Y %H:%M",

            "%d/%m/%Y %H:%M:%S",

            "%d/%m/%Y",

            "%Y-%m-%d %H:%M",

            "%Y-%m-%d %H:%M:%S",

            "%Y-%m-%d",
        )

        for formato in formatos:

            try:

                return datetime.strptime(
                    valor,
                    formato
                )

            except ValueError:

                continue

        for separador in (
            " ",
            "T"
        ):

            if separador in valor:

                parte = valor.split(
                    separador
                )[0]

                for formato in (
                    "%d/%m/%Y",
                    "%Y-%m-%d"
                ):

                    try:

                        return datetime.strptime(
                            parte,
                            formato
                        )

                    except ValueError:

                        pass

        return None

    # =====================================================
    # APLICAR FILTRO
    # =====================================================

    def aplicar_filtro_fecha(
        self
    ):

        desde = (
            self.fecha_desde_edit.date()
        )

        hasta = (
            self.fecha_hasta_edit.date()
        )

        if desde > hasta:

            QMessageBox.warning(
                self,
                "Filtro de fechas",
                "La fecha 'Desde' no puede "
                "ser posterior a la fecha 'Hasta'."
            )

            return

        self.fecha_desde = desde

        self.fecha_hasta = hasta

        self.filtrar_lista()

    # =====================================================
    # LIMPIAR FILTRO
    # =====================================================

    def limpiar_filtro_fecha(
        self
    ):

        self.fecha_desde = None

        self.fecha_hasta = None

        self.filtrar_lista()

    # =====================================================
    # FILTRO POR ESTADO
    # =====================================================

    def seleccionar_filtro_estado(
        self,
        estado
    ):

        self.filtro_estado_actual = (
            str(estado or "TODOS").strip().upper()
        )

        self.actualizar_botones_estado()

        self.filtrar_lista()


    def actualizar_botones_estado(
        self
    ):

        botones = {
            "TODOS": self.btn_estado_todos,
            ESTADO_PENDIENTE: self.btn_estado_pendientes,
            ESTADO_ENVIADO: self.btn_estado_enviados,
            ESTADO_RECHAZADO: self.btn_estado_rechazados
        }

        for estado, boton in botones.items():
            boton.setProperty(
                "activo",
                estado == self.filtro_estado_actual
            )
            boton.style().unpolish(boton)
            boton.style().polish(boton)


    def actualizar_contadores_estado(
        self
    ):

        conteos = {
            ESTADO_PENDIENTE: 0,
            ESTADO_ENVIADO: 0,
            ESTADO_RECHAZADO: 0
        }

        for registro in self.registros:
            estado = str(
                registro.get(
                    "estado",
                    ESTADO_PENDIENTE
                )
            ).strip().upper()

            if estado not in conteos:
                estado = ESTADO_PENDIENTE

            conteos[estado] += 1

        self.btn_estado_pendientes.setText(
            f"PENDIENTES ({conteos[ESTADO_PENDIENTE]})"
        )

        self.btn_estado_enviados.setText(
            f"ENVIADOS ({conteos[ESTADO_ENVIADO]})"
        )

        self.btn_estado_rechazados.setText(
            f"RECHAZADOS ({conteos[ESTADO_RECHAZADO]})"
        )

        self.btn_estado_todos.setText(
            f"TODOS ({len(self.registros)})"
        )


    # =====================================================
    # FILTRAR
    # =====================================================

    def filtrar_lista(
        self
    ):

        texto = (
            self.txt_buscar
            .text()
            .strip()
            .lower()
        )

        visibles = 0

        for i in range(
            self.lista.count()
        ):

            item = self.lista.item(
                i
            )

            registro = item.data(
                Qt.ItemDataRole.UserRole
            )

            contenido = " ".join([

                str(
                    registro.get(
                        "fecha",
                        ""
                    )
                ),

                str(
                    registro.get(
                        "socio",
                        ""
                    )
                ),

                str(
                    registro.get(
                        "placa",
                        ""
                    )
                ),

                str(
                    registro.get(
                        "chasis",
                        ""
                    )
                ),

                str(
                    registro.get(
                        "valores_sri",
                        ""
                    )
                ),

                str(
                    registro.get(
                        "evidencia_sri",
                        ""
                    )
                )

            ]).lower()

            coincide_texto = (
                not texto
                or
                texto in contenido
            )

            coincide_fecha = True

            if (
                self.fecha_desde
                or
                self.fecha_hasta
            ):

                fecha_registro = (
                    self._fecha_registro(
                        registro
                    )
                )

                if fecha_registro is None:

                    coincide_fecha = False

                else:

                    fecha_solo = (
                        fecha_registro.date()
                    )

                    if self.fecha_desde:

                        coincide_fecha = (
                            coincide_fecha
                            and
                            fecha_solo
                            >=
                            self.fecha_desde.toPython()
                        )

                    if self.fecha_hasta:

                        coincide_fecha = (
                            coincide_fecha
                            and
                            fecha_solo
                            <=
                            self.fecha_hasta.toPython()
                        )

            estado_registro = str(
                registro.get(
                    "estado",
                    ESTADO_PENDIENTE
                )
            ).strip().upper()

            coincide_estado = (
                self.filtro_estado_actual == "TODOS"
                or
                estado_registro == self.filtro_estado_actual
            )

            visible = (
                coincide_texto
                and
                coincide_fecha
                and
                coincide_estado
            )

            item.setHidden(
                not visible
            )

            if visible:

                visibles += 1

        total = self.lista.count()

        self.lbl_contador.setText(
            f"Mostrando {visibles} "
            f"de {total} resoluciones"
        )

    # =====================================================
    # SELECCIONAR TODOS
    # =====================================================

    def seleccionar_todos_visibles(
        self
    ):

        visibles = 0

        for i in range(
            self.lista.count()
        ):

            item = self.lista.item(
                i
            )

            if item.isHidden():
                continue

            item.setCheckState(
                Qt.CheckState.Checked
            )

            fila = self.lista.itemWidget(
                item
            )

            if fila and hasattr(
                fila,
                "checkbox"
            ):

                fila.actualizar_checkbox()

            visibles += 1

        if visibles == 0:

            QMessageBox.information(
                self,
                "Selección",
                "No hay registros visibles "
                "para seleccionar."
            )

    # =====================================================
    # ENVÍO DE CORREOS
    # =====================================================

    def configurar_certificado_inicio(
        self
    ):
        """
        Método conservado por compatibilidad.
        El certificado NO se solicita al iniciar.
        La selección se realiza desde el diálogo de firma
        únicamente cuando se intenta enviar y no hay una ruta
        de certificado disponible.
        """
        if self.p12_path and os.path.isfile(
            self.p12_path
        ):
            return

        ruta, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar certificado digital",
            "",
            (
                "Certificado PKCS#12 "
                "(*.p12 *.pfx);;"
                "Todos los archivos (*.*)"
            )
        )

        if not ruta:
            QMessageBox.warning(
                self,
                "Certificado requerido",
                "No seleccionaste un certificado .P12/.PFX.\n\n"
                "Podrás configurarlo nuevamente al reiniciar "
                "la aplicación."
            )
            return

        self.p12_path = ruta

        try:
            guardar_ruta_certificado(
                ruta
            )
        except Exception as e:
            QMessageBox.warning(
                self,
                "Aviso",
                "El certificado se cargó correctamente, pero "
                "no se pudo guardar su ruta para el próximo inicio.\n\n"
                + str(e)
            )

    def rechazar_registro_actual(
        self
    ):

        registro = self.registro_actual

        if not registro:
            QMessageBox.information(
                self,
                "Rechazar resolución",
                "Selecciona primero una resolución."
            )
            return

        pdf_id = str(
            registro.get(
                "pdf_id",
                ""
            )
        ).strip()

        if not pdf_id:
            QMessageBox.warning(
                self,
                "Rechazar resolución",
                "La resolución seleccionada no tiene PDF ID."
            )
            return

        motivo, aceptado = QInputDialog.getMultiLineText(
            self,
            "Rechazar resolución",
            "Indica el motivo del rechazo:",
            str(
                registro.get(
                    "motivo_rechazo",
                    ""
                )
            ).strip()
        )

        if not aceptado:
            return

        motivo = str(motivo or "").strip()

        if not motivo:
            QMessageBox.warning(
                self,
                "Motivo requerido",
                "Debes indicar el motivo del rechazo."
            )
            return

        try:
            cambiar_estado_remoto(
                pdf_id,
                ESTADO_RECHAZADO,
                motivo
            )

            correo_usuario = str(registro.get("correo_electronico", "")).strip().lower()

            if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", correo_usuario):
                raise ValueError(
                    "El registro fue rechazado, pero no tiene un correo electrónico válido para notificar al usuario."
                )

            token = obtener_token_microsoft_graph()
            obtener_cuenta_graph(token)

            enviar_correo_rechazo_graph(
                token=token,
                nombre=str(registro.get("socio", "Usuario")).strip() or "Usuario",
                destinatario=correo_usuario,
                motivo=motivo,
                placa=str(registro.get("placa", "")).strip()
            )

            for item_index in range(
                self.lista.count()
            ):
                item = self.lista.item(item_index)
                registro_item = item.data(
                    Qt.ItemDataRole.UserRole
                )
                if (
                    registro_item
                    and
                    str(registro_item.get("pdf_id", "")).strip()
                    == pdf_id
                ):
                    registro_item["estado"] = ESTADO_RECHAZADO
                    registro_item["motivo_rechazo"] = motivo
                    break

            registro["estado"] = ESTADO_RECHAZADO
            registro["motivo_rechazo"] = motivo
            self.firma_registros = ""
            self.actualizar_lista(
                mostrar_errores=True
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error al rechazar",
                str(e)
            )


    def enviar_correos_seleccionados(
        self
    ):

        if self.enviando_correos:

            return

        seleccionados = []

        for i in range(
            self.lista.count()
        ):

            item = self.lista.item(
                i
            )

            if item.isHidden():
                continue

            if (
                item.checkState()
                !=
                Qt.CheckState.Checked
            ):

                continue

            registro = item.data(
                Qt.ItemDataRole.UserRole
            )

            if registro:

                estado_registro = str(
                    registro.get(
                        "estado",
                        ESTADO_PENDIENTE
                    )
                ).strip().upper()

                # PENDIENTE y RECHAZADO pueden entrar nuevamente
                # al proceso de envío. ENVIADO queda bloqueado para
                # evitar que un correo ya enviado se duplique.
                if estado_registro == ESTADO_ENVIADO:
                    continue

                if estado_registro not in (
                    ESTADO_PENDIENTE,
                    ESTADO_RECHAZADO
                ):
                    continue

                seleccionados.append(
                    registro
                )

        if not seleccionados:

            QMessageBox.information(
                self,
                "Enviar correos",
                "Selecciona al menos "
                "un registro."
            )

            return

        if not os.path.isfile(
            PLANTILLA_EXCEL
        ):

            QMessageBox.critical(
                self,
                "Plantilla Excel no encontrada",
                "No se encontró la plantilla Excel "
                "incluida en el sistema.\n\n"
                f"Ruta buscada:\n{PLANTILLA_EXCEL}\n\n"
                "Si estás usando el .EXE, vuelve a "
                "generarlo incluyendo plantilla_AMT.xlsx."
            )

            return

        confirmacion = QMessageBox.question(

            self,

            "Confirmar envío",

            "Se enviarán "
            f"{len(seleccionados)} correo(s) "
            "de prueba a:\n\n"

            f"{CORREO_DESTINO_PRUEBAS}\n\n"

            "Cada correo llevará el formato "
            "personalizado en PDF, firmado "
            "electrónicamente con el certificado "
            "seleccionado.\n\n"

            "Durante el proceso aparecerá una "
            "pantalla de carga para indicar que "
            "el sistema sigue trabajando.\n\n"

            "¿Deseas continuar?",

            QMessageBox.StandardButton.Yes
            |
            QMessageBox.StandardButton.No
        )

        if (
            confirmacion
            !=
            QMessageBox.StandardButton.Yes
        ):

            return

        # -------------------------------------------------
        # CERTIFICADO PARA LA FIRMA
        # -------------------------------------------------

        # Si ya existe un certificado cargado/guardado,
        # el diálogo solicitará únicamente la contraseña.
        # Si no existe, el diálogo permitirá seleccionar el
        # .P12/.PFX en este momento, justo al enviar.
        dialogo_firma = FirmaCorreoDialog(
            self,
            p12_path=self.p12_path
        )

        if (
            dialogo_firma.exec()
            != QDialog.DialogCode.Accepted
        ):
            return

        p12_path = dialogo_firma.p12_path
        password_firma = dialogo_firma.password
        datos_certificado = (
            dialogo_firma.datos_certificado
        )


        if p12_path != self.p12_path:
            self.p12_path = p12_path
            try:
                guardar_ruta_certificado(
                    p12_path
                )
            except Exception:
                pass

        # -------------------------------------------------
        # ACTIVAR ESTADO DE ENVÍO
        # -------------------------------------------------

        self.enviando_correos = True

        # -------------------------------------------------
        # DESACTIVAR CONTROLES
        # -------------------------------------------------

        self.btn_enviar_correos.setEnabled(
            False
        )

        self.btn_seleccionar_todos.setEnabled(
            False
        )

        self.btn_actualizar.setEnabled(
            False
        )

        self.txt_buscar.setEnabled(
            False
        )

        self.btn_aplicar_fecha.setEnabled(
            False
        )

        self.btn_limpiar_fecha.setEnabled(
            False
        )

        # -------------------------------------------------
        # DETENER ACTUALIZACIÓN AUTOMÁTICA
        # -------------------------------------------------

        self.timer_actualizacion.stop()

        self.lbl_estado.setText(
            "🟡 Enviando correos..."
        )

        # -------------------------------------------------
        # MOSTRAR LOADING
        # -------------------------------------------------

        self.loading_overlay.setGeometry(
            self.centralWidget().rect()
        )

        self.loading_overlay.mostrar(
            len(seleccionados)
        )

        # -------------------------------------------------
        # CREAR WORKER
        # -------------------------------------------------

        worker = EmailWorker(
            seleccionados,
            self.cache_pdfs,
            p12_path,
            password_firma,
            datos_certificado
        )

        self.worker_correos = worker

        worker.signals.progress.connect(
            self.actualizar_progreso_correos
        )

        worker.signals.finished.connect(
            self.finalizar_envio_correos
        )

        worker.signals.error_fatal.connect(
            self.error_fatal_envio_correos
        )

        self.thread_pool.start(
            worker
        )

    # =====================================================
    # ACTUALIZAR LOADING
    # =====================================================

    def actualizar_progreso_correos(
        self,
        procesados,
        total,
        mensaje
    ):

        if self.cerrando:
            return

        self.loading_overlay.actualizar(
            procesados,
            total,
            mensaje
        )

        self.lbl_estado.setText(
            "🟡 "
            +
            mensaje
        )

    # =====================================================
    # FINALIZAR ENVÍO
    # =====================================================

    def finalizar_envio_correos(
        self,
        enviados,
        errores
    ):

        if self.cerrando:
            return

        self.loading_overlay.progress.setValue(
            100
        )

        self.loading_overlay.contador.setText(
            "Proceso terminado"
        )

        self.loading_overlay.texto.setText(
            "Los correos han terminado de procesarse."
        )

        # -------------------------------------------------
        # PEQUEÑA PAUSA VISUAL
        # -------------------------------------------------

        QTimer.singleShot(
            700,
            lambda:
            self._mostrar_resultado_envio(
                enviados,
                errores
            )
        )

    # =====================================================
    # MOSTRAR RESULTADO
    # =====================================================

    def _mostrar_resultado_envio(
        self,
        enviados,
        errores
    ):

        if self.cerrando:
            return

        self.loading_overlay.ocultar()

        self.enviando_correos = False

        self.worker_correos = None

        # -------------------------------------------------
        # REACTIVAR CONTROLES
        # -------------------------------------------------

        self.btn_enviar_correos.setEnabled(
            True
        )

        self.btn_seleccionar_todos.setEnabled(
            True
        )

        self.btn_actualizar.setEnabled(
            True
        )

        self.txt_buscar.setEnabled(
            True
        )

        self.btn_aplicar_fecha.setEnabled(
            True
        )

        self.btn_limpiar_fecha.setEnabled(
            True
        )

        # -------------------------------------------------
        # REINICIAR ACTUALIZACIÓN
        # -------------------------------------------------

        self.timer_actualizacion.start()

        # -------------------------------------------------
        # RESULTADO
        # -------------------------------------------------

        if errores:

            detalle = "\n".join(
                errores[:5]
            )

            if len(errores) > 5:

                detalle += (
                    f"\n... y "
                    f"{len(errores) - 5} "
                    "error(es) más."
                )

            QMessageBox.warning(

                self,

                "Resultado del envío",

                f"Enviados correctamente: "
                f"{enviados}\n"

                f"Con error: "
                f"{len(errores)}\n\n"

                f"{detalle}"
            )

            self.lbl_estado.setText(

                f"🟡 Correos enviados: "
                f"{enviados} · "
                f"Errores: "
                f"{len(errores)}"
            )

        else:

            QMessageBox.information(

                self,

                "Resultado del envío",

                f"Se enviaron "
                f"{enviados} correo(s) "
                "correctamente a\n"

                f"{CORREO_DESTINO_PRUEBAS}."
            )

            self.lbl_estado.setText(

                f"🟢 {enviados} "
                "correo(s) enviado(s)"
            )

        # -------------------------------------------------
        # DESMARCAR SELECCIONADOS
        # -------------------------------------------------

        for i in range(
            self.lista.count()
        ):

            item = self.lista.item(
                i
            )

            item.setCheckState(
                Qt.CheckState.Unchecked
            )

        # Refrescar inmediatamente los estados desde Apps Script.
        # Así los enviados salen de PENDIENTES y aparecen en ENVIADOS
        # sin esperar al siguiente ciclo automático.
        self.firma_registros = ""
        self.actualizar_lista(
            mostrar_errores=False
        )

    # =====================================================
    # ERROR FATAL
    # =====================================================

    def error_fatal_envio_correos(
        self,
        error
    ):

        if self.cerrando:
            return

        self.loading_overlay.ocultar()

        self.enviando_correos = False

        self.worker_correos = None

        self.btn_enviar_correos.setEnabled(
            True
        )

        self.btn_seleccionar_todos.setEnabled(
            True
        )

        self.btn_actualizar.setEnabled(
            True
        )

        self.txt_buscar.setEnabled(
            True
        )

        self.btn_aplicar_fecha.setEnabled(
            True
        )

        self.btn_limpiar_fecha.setEnabled(
            True
        )

        self.timer_actualizacion.start()

        self.lbl_estado.setText(
            "🔴 Error durante el envío"
        )

        QMessageBox.critical(

            self,

            "Nuevo Outlook / Microsoft Graph",

            "No se pudo completar "
            "el proceso de envío.\n\n"
            +
            str(error)
        )

    # =====================================================
    # SELECCIONAR REGISTRO
    # =====================================================

    def seleccionar_registro(
        self,
        item
    ):

        registro = item.data(
            Qt.ItemDataRole.UserRole
        )

        if not registro:
            return

        self.registro_actual = (
            registro
        )

        pdf_url = str(
            registro.get(
                "pdf_url",
                ""
            )
        ).strip()

        pdf_id = str(
            registro.get(
                "pdf_id",
                ""
            )
        ).strip()

        if pdf_id:

            self.cargar_pdf_desde_drive(
                pdf_url,
                pdf_id
            )

        else:

            self.pdf_viewer.ocultar_cargando()
            self.pdf_viewer.limpiar()

            self.pdf_id_actual = None

            self.btn_descargar.setEnabled(
                False
            )

            QMessageBox.warning(
                self,
                "PDF",
                "Este registro no tiene "
                "ID de PDF guardado."
            )

    # =====================================================
    # ABRIR DETALLE
    # =====================================================

    def abrir_detalle_registro(
        self,
        registro
    ):

        if not registro:
            return

        self.registro_actual = registro

        if (
            hasattr(self, "ventana_detalle")
            and
            self.ventana_detalle
            and
            self.ventana_detalle.isVisible()
        ):

            self.ventana_detalle.actualizar_datos(
                registro
            )

            self.ventana_detalle.raise_()
            self.ventana_detalle.activateWindow()

            return

        self.ventana_detalle = DetalleResolucionDialog(
            registro,
            self
        )

        self.ventana_detalle.finished.connect(
            self._detalle_cerrado
        )

        self.ventana_detalle.show()

        self.ventana_detalle.raise_()
        self.ventana_detalle.activateWindow()

    def _detalle_cerrado(
        self
    ):

        self.ventana_detalle = None

    # =====================================================
    # CARGAR PDF
    # =====================================================

    def cargar_pdf_desde_drive(
        self,
        url,
        pdf_id=None
    ):

        if not pdf_id:

            QMessageBox.warning(
                self,
                "PDF",
                "No se encontró el ID "
                "del archivo PDF."
            )

            return

        # Mostrar inmediatamente el cargador SOLO en el visor PDF.
        # Así el usuario ve que el sistema está procesando la
        # resolución mientras se obtiene y prepara el archivo.
        self.pdf_viewer.mostrar_cargando()

        if pdf_id in self.cache_pdfs:

            self.mostrar_pdf_bytes(
                self.cache_pdfs[pdf_id],
                pdf_id
            )

            return

        if (
            self.pdf_id_actual
            ==
            pdf_id
            and
            self.pdf_viewer.doc
            is not None
        ):

            return

        self.btn_descargar.setEnabled(
            False
        )

        self.lbl_estado.setText(
            "🟡 Cargando PDF..."
        )

        worker = RequestWorker(

            GOOGLE_SHEETS_URL,

            {
                "accion": "pdf",
                "id": pdf_id
            },

            TIMEOUT_PDF
        )

        self.worker_pdf = worker

        worker.signals.finished.connect(

            lambda resultado:

            self.procesar_pdf(
                resultado,
                pdf_id
            )
        )

        worker.signals.error.connect(

            lambda error:

            self.error_pdf(
                error
            )
        )

        self.thread_pool.start(
            worker
        )

    # =====================================================
    # PROCESAR PDF
    # =====================================================

    def procesar_pdf(
        self,
        resultado,
        pdf_id
    ):

        if self.cerrando:
            return

        try:

            if resultado.get(
                "estado"
            ) != "ok":

                raise Exception(
                    resultado.get(
                        "mensaje",
                        "No se pudo obtener el PDF."
                    )
                )

            pdf_base64 = (
                resultado.get(
                    "data"
                )
            )

            if not pdf_base64:

                raise Exception(
                    "El servidor no devolvió "
                    "los datos del PDF."
                )

            pdf_bytes = (
                base64.b64decode(
                    pdf_base64
                )
            )

            if not pdf_bytes.startswith(
                b"%PDF"
            ):

                raise Exception(
                    "Los datos recibidos no "
                    "corresponden a un PDF válido."
                )

            self.cache_pdfs[
                pdf_id
            ] = pdf_bytes

            self.mostrar_pdf_bytes(
                pdf_bytes,
                pdf_id
            )

            self.lbl_estado.setText(
                "🟢 PDF cargado"
            )

        except Exception as e:

            self.error_pdf(
                str(e)
            )

        finally:

            self.worker_pdf = None

    # =====================================================
    # MOSTRAR PDF
    # =====================================================

    def mostrar_pdf_bytes(
        self,
        pdf_bytes,
        pdf_id
    ):

        try:

            # El cargador está únicamente dentro del visor PDF.
            # Se mantiene visible hasta que el PDF termina de
            # abrirse y mostrarse.
            self.pdf_viewer.mostrar_cargando()

            self.pdf_viewer.cerrar()

            if self.archivo_temporal:

                try:

                    os.remove(
                        self.archivo_temporal
                    )

                except Exception:

                    pass

                self.archivo_temporal = None

            archivo = (
                tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".pdf"
                )
            )

            archivo.write(
                pdf_bytes
            )

            archivo.close()

            self.archivo_temporal = (
                archivo.name
            )

            self.pdf_viewer.abrir_pdf(
                self.archivo_temporal
            )

            self.pdf_id_actual = (
                pdf_id
            )

            self.btn_descargar.setEnabled(
                True
            )

            # El PDF ya está visible: quitar el cargador.
            self.pdf_viewer.ocultar_cargando()

        except Exception as e:

            self.pdf_viewer.ocultar_cargando()

            self.error_pdf(
                str(e)
            )

    # =====================================================
    # ERROR PDF
    # =====================================================

    def error_pdf(
        self,
        error
    ):

        # Si ocurrió un error, retirar el cargador del visor.
        self.pdf_viewer.ocultar_cargando()

        self.btn_descargar.setEnabled(
            False
        )

        self.lbl_estado.setText(
            "🔴 Error al cargar PDF"
        )

        if not self.cerrando:

            QMessageBox.critical(
                self,
                "Error al cargar PDF",
                str(error)
            )

        self.worker_pdf = None

    # =====================================================
    # DESCARGAR PDF
    # =====================================================

    def descargar_pdf(
        self
    ):

        if not self.registro_actual:

            QMessageBox.warning(
                self,
                "Descargar PDF",
                "Primero seleccione un registro."
            )

            return

        pdf_id = str(
            self.registro_actual.get(
                "pdf_id",
                ""
            )
        ).strip()

        if not pdf_id:

            QMessageBox.warning(
                self,
                "Descargar PDF",
                "No se encontró el ID "
                "del PDF."
            )

            return

        placa = self.registro_actual.get(
            "placa",
            "resolucion"
        )

        nombre = (
            f"Resolucion_{placa}.pdf"
        )

        ruta, _ = (
            QFileDialog.getSaveFileName(

                self,

                "Guardar PDF",

                nombre,

                "Archivos PDF (*.pdf)"
            )
        )

        if not ruta:
            return

        if pdf_id in self.cache_pdfs:

            try:

                with open(
                    ruta,
                    "wb"
                ) as archivo:

                    archivo.write(
                        self.cache_pdfs[
                            pdf_id
                        ]
                    )

                QMessageBox.information(
                    self,
                    "PDF descargado",
                    "El PDF se guardó correctamente."
                )

                return

            except Exception as e:

                QMessageBox.critical(
                    self,
                    "Error",
                    str(e)
                )

                return

        self.lbl_estado.setText(
            "🟡 Descargando PDF..."
        )

        self.btn_descargar.setEnabled(
            False
        )

        worker = RequestWorker(

            GOOGLE_SHEETS_URL,

            {
                "accion": "pdf",
                "id": pdf_id
            },

            TIMEOUT_PDF
        )

        self.worker_descarga = worker

        worker.signals.finished.connect(

            lambda resultado:

            self.procesar_descarga(
                resultado,
                pdf_id,
                ruta
            )
        )

        worker.signals.error.connect(

            lambda error:

            self.error_descarga(
                error
            )
        )

        self.thread_pool.start(
            worker
        )

    # =====================================================
    # PROCESAR DESCARGA
    # =====================================================

    def procesar_descarga(
        self,
        resultado,
        pdf_id,
        ruta
    ):

        try:

            if resultado.get(
                "estado"
            ) != "ok":

                raise Exception(
                    resultado.get(
                        "mensaje",
                        "No se pudo descargar el PDF."
                    )
                )

            pdf_bytes = (
                base64.b64decode(
                    resultado["data"]
                )
            )

            if not pdf_bytes.startswith(
                b"%PDF"
            ):

                raise Exception(
                    "El archivo recibido "
                    "no es un PDF válido."
                )

            self.cache_pdfs[
                pdf_id
            ] = pdf_bytes

            with open(
                ruta,
                "wb"
            ) as archivo:

                archivo.write(
                    pdf_bytes
                )

            QMessageBox.information(
                self,
                "PDF descargado",
                "El PDF se guardó correctamente."
            )

            self.lbl_estado.setText(
                "🟢 Conectado"
            )

        except Exception as e:

            self.error_descarga(
                str(e)
            )

        finally:

            self.btn_descargar.setEnabled(
                True
            )

            self.worker_descarga = None

    # =====================================================
    # ERROR DESCARGA
    # =====================================================

    def error_descarga(
        self,
        error
    ):

        self.btn_descargar.setEnabled(
            True
        )

        self.lbl_estado.setText(
            "🔴 Error de conexión"
        )

        QMessageBox.critical(
            self,
            "Error al descargar PDF",
            str(error)
        )

        self.worker_descarga = None

    # =====================================================
    # CERRAR
    # =====================================================

    def closeEvent(
        self,
        event
    ):

        if self.enviando_correos:

            respuesta = QMessageBox.question(

                self,

                "Envío en progreso",

                "Hay correos que todavía se están "
                "enviando.\n\n"
                "Si cierras la aplicación, el proceso "
                "se interrumpirá.\n\n"
                "¿Seguro que deseas cerrar?",

                QMessageBox.StandardButton.Yes
                |
                QMessageBox.StandardButton.No
            )

            if (
                respuesta
                !=
                QMessageBox.StandardButton.Yes
            ):

                event.ignore()

                return

        self.cerrando = True

        if hasattr(
            self,
            "timer_actualizacion"
        ):

            self.timer_actualizacion.stop()

        try:

            self.pdf_viewer.cerrar()

        except Exception:

            pass

        if self.archivo_temporal:

            try:

                os.remove(
                    self.archivo_temporal
                )

            except Exception:

                pass

            self.archivo_temporal = None

        self.cache_pdfs.clear()

        event.accept()


# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":

    app = QApplication(
        sys.argv
    )

    ventana = AdminWindow()

    ventana.show()

    sys.exit(
        app.exec()
    )
