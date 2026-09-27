# -*- coding: utf-8 -*-
"""
Firmador Excel -> PDF con firma digital visible

Flujo:
1. Selecciona el Excel.
2. Selecciona el certificado .P12/.PFX.
3. Introduce la contraseña del certificado.
4. El programa elimina la palabra "FIRMA" de la zona A45:C48.
5. Convierte el Excel a PDF usando Microsoft Excel (LibreOffice queda como respaldo).
6. Crea automáticamente una estampa visible con:
      - QR único para ese documento
      - "Firmado electrónicamente por:"
      - nombre obtenido del certificado
      - fecha/hora
7. Firma criptográficamente el PDF con el certificado .P12.
8. Guarda: NOMBRE_ORIGINAL_FIRMADO.pdf

IMPORTANTE:
- El .P12 y su contraseña permanecen en el PC.
- El QR generado aquí contiene datos únicos del documento y de la firma.
  No es, por sí solo, un servicio público de validación. La validez
  criptográfica está en la firma PDF.
- La zona de firma está configurada para la plantilla proporcionada:
      A45:C48
  en la página 1.
"""

import hashlib
import os
import shutil
import subprocess
import tempfile
import tkinter as tk
import win32com.client
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# Dependencias externas:
#   pip install openpyxl reportlab qrcode[pil] cryptography pyHanko

try:
    from openpyxl import load_workbook
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.x509.oid import NameOID
    import qrcode
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
except ImportError as e:
    raise SystemExit(
        "Faltan dependencias. Ejecuta:\n\n"
        "python -m pip install openpyxl reportlab \"qrcode[pil]\" cryptography pyHanko pywin32\n\n"
        f"Detalle: {e}"
    )

try:
    from pyhanko import stamp
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign import fields, signers
except ImportError as e:
    raise SystemExit(
        "Falta pyHanko. Ejecuta:\n\n"
        "python -m pip install pyHanko pywin32\n\n"
        f"Detalle: {e}"
    )


# ============================================================
# CONFIGURACIÓN DE LA PLANTILLA
# ============================================================

# La plantilla entregada tiene la palabra FIRMA en A45 y el área
# de firma está combinada desde A45 hasta C48.
CELDA_FIRMA = "A45"

# Coordenadas PDF en puntos para la plantilla proporcionada.
# Página Letter: 612 x 792 puntos.
# Área correspondiente aproximadamente a A45:C48.
SIGNATURE_BOX = (197.5, 53, 414.5, 108)  # x1, y1, x2, y2

# Tamaño de la estampa interna.
STAMP_W = SIGNATURE_BOX[2] - SIGNATURE_BOX[0]
STAMP_H = SIGNATURE_BOX[3] - SIGNATURE_BOX[1]

# Nombre del campo de firma PDF.
SIGNATURE_FIELD_NAME = "FirmaElectronica"


# ============================================================
# UTILIDADES
# ============================================================

def buscar_libreoffice():
    """Busca LibreOffice como alternativa secundaria."""
    candidatos = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for ruta in candidatos:
        if ruta and Path(ruta).exists():
            return str(ruta)

    return None


def convertir_excel_con_microsoft_excel(excel_path, pdf_path):
    """
    Convierte el Excel a PDF usando Microsoft Excel mediante COM.
    Esto conserva mejor márgenes, celdas combinadas, fuentes y diseño.
    """
    excel = None
    libro = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        libro = excel.Workbooks.Open(
            str(Path(excel_path).resolve()),
            UpdateLinks=0,
            ReadOnly=False,
        )

        hoja = libro.Worksheets(1)

        # La plantilla tiene la zona de firma en A45:C48.
        # A45 puede ser una celda combinada; vaciarla elimina "FIRMA".
        hoja.Range("A45").Value = ""

        libro.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=str(Path(pdf_path).resolve()),
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )

        if not Path(pdf_path).exists():
            raise RuntimeError("Microsoft Excel no generó el PDF.")

        return Path(pdf_path)

    except Exception as e:
        raise RuntimeError(
            "Microsoft Excel no pudo abrir o exportar la plantilla.\n\n"
            f"Detalle: {e}"
        )

    finally:
        try:
            if libro is not None:
                libro.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def convertir_excel_con_libreoffice(excel_path, carpeta_salida, libreoffice):
    """Conversión alternativa mediante LibreOffice."""
    carpeta_salida = Path(carpeta_salida)
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    cmd = [
        libreoffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(carpeta_salida),
        str(excel_path),
    ]

    startupinfo = None
    creationflags = 0

    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW

    resultado = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        startupinfo=startupinfo,
        creationflags=creationflags,
        timeout=120,
    )

    pdf_path = carpeta_salida / (Path(excel_path).stem + ".pdf")

    if resultado.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            "LibreOffice no pudo convertir el Excel a PDF.\n\n"
            + resultado.stdout
            + "\n"
            + resultado.stderr
        )

    return pdf_path


def convertir_excel_a_pdf(excel_path, carpeta_salida):
    """
    Usa primero Microsoft Excel. LibreOffice queda como respaldo.
    """
    carpeta_salida = Path(carpeta_salida)
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    pdf_path = carpeta_salida / (Path(excel_path).stem + ".pdf")

    try:
        return convertir_excel_con_microsoft_excel(
            excel_path,
            pdf_path,
        )
    except Exception as excel_error:
        libreoffice = buscar_libreoffice()

        if libreoffice:
            try:
                return convertir_excel_con_libreoffice(
                    excel_path,
                    carpeta_salida,
                    libreoffice,
                )
            except Exception as lo_error:
                raise RuntimeError(
                    "No se pudo convertir el Excel a PDF.\n\n"
                    f"Microsoft Excel:\n{excel_error}\n\n"
                    f"LibreOffice:\n{lo_error}"
                )

        raise RuntimeError(
            "No se pudo convertir el Excel a PDF con Microsoft Excel.\n\n"
            "No necesitas LibreOffice para este programa.\n\n"
            "Verifica que Microsoft Excel esté instalado y que puedas "
            "abrir normalmente el archivo .xlsx.\n\n"
            f"Error real de Excel:\n{excel_error}\n\n"
            "Si acabas de instalar pywin32, ejecuta también:\n"
            "python -m pywin32_postinstall -install"
        )

def sha256_archivo(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloque in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloque)
    return h.hexdigest()


def cargar_datos_certificado(p12_path, password):
    """
    Lee el certificado para obtener nombre y número de serie.
    La clave privada no se guarda ni se exporta.
    """
    with open(p12_path, "rb") as f:
        data = f.read()

    private_key, certificate, additional_certs = pkcs12.load_key_and_certificates(
        data,
        password.encode("utf-8") if password else None
    )

    if private_key is None or certificate is None:
        raise ValueError(
            "El archivo .P12/.PFX no contiene una clave privada y certificado utilizables."
        )

    nombre = None
    try:
        cn = certificate.subject.get_attributes_for_oid(NameOID.COMMON_NAME)
        if cn:
            nombre = cn[0].value
    except Exception:
        pass

    if not nombre:
        try:
            nombres = certificate.subject.get_attributes_for_oid(NameOID.GIVEN_NAME)
            apellidos = certificate.subject.get_attributes_for_oid(NameOID.SURNAME)
            partes = []
            if nombres:
                partes.append(str(nombres[0].value))
            if apellidos:
                partes.append(str(apellidos[0].value))
            nombre = " ".join(partes).strip()
        except Exception:
            pass

    if not nombre:
        nombre = certificate.subject.rfc4514_string()

    return {
        "nombre": str(nombre).upper(),
        "serial": format(certificate.serial_number, "X"),
        "not_before": certificate.not_valid_before_utc,
        "not_after": certificate.not_valid_after_utc,
    }


def preparar_excel(excel_original, carpeta_temp):
    """
    Crea una copia temporal del Excel y elimina la palabra FIRMA.
    El archivo original del usuario nunca se modifica.
    """
    destino = Path(carpeta_temp) / "plantilla_para_firmar.xlsx"

    wb = load_workbook(excel_original)
    ws = wb[wb.sheetnames[0]]

    # Quitamos el texto "FIRMA" de la zona donde aparecerá
    # automáticamente la firma electrónica.
    ws[CELDA_FIRMA] = ""

    # Aseguramos el área de impresión de la plantilla.
    ws.print_area = "A1:F49"

    wb.save(destino)
    wb.close()

    return destino



def crear_estampa_firma(
    salida_pdf,
    nombre,
    numero_serie,
    hash_pdf,
    fecha_hora,
):
    """
    Crea un PDF pequeño que se utilizará como apariencia visible
    de la firma criptográfica.

    El QR se genera automáticamente y cambia para cada documento.
    """
    # Datos compactos y únicos para el QR.
    # El hash corresponde al PDF antes de insertar la firma.
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
    qr_img = qr.make_image(fill_color="black", back_color="white")

    qr_temp = Path(salida_pdf).with_name("qr_firma_temp.png")
    qr_img.save(qr_temp)

    c = canvas.Canvas(str(salida_pdf), pagesize=(STAMP_W, STAMP_H))

    # QR a la izquierda.
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

    # Una sola coordenada X para TODO el texto.
    # Así "Firmado electrónicamente por:" queda perfectamente alineado
    # con el inicio del nombre del firmante.
    x_text = margen + qr_size + 5

    # Texto de la firma.
    c.setFillColorRGB(0, 0, 0)

    # La etiqueta y el nombre usan exactamente el mismo punto inicial X.
    # Además, todo el bloque de texto se centra verticalmente respecto al QR.
    c.setFont("Courier", 5.5)

    # Nombre en dos líneas si es largo.
    nombre_limpio = " ".join(nombre.split())
    palabras = nombre_limpio.split()

    if len(nombre_limpio) > 27 and len(palabras) >= 3:
        mitad = len(palabras) // 2
        linea1 = " ".join(palabras[:mitad])
        linea2 = " ".join(palabras[mitad:])
    else:
        linea1 = nombre_limpio
        linea2 = ""

    # Centro vertical del QR.
    centro_qr_y = margen + (qr_size / 2)

    # En la referencia hay muy poco interlineado y las tres líneas
    # quedan centradas verticalmente con respecto al QR.
    c.drawString(x_text, centro_qr_y + 7, "Firmado electrónicamente por:")

    c.setFont("Courier-Bold", 7.8)
    c.drawString(x_text, centro_qr_y - 1, linea1)

    if linea2:
        c.drawString(x_text, centro_qr_y - 9, linea2)

    c.save()

    try:
        qr_temp.unlink()
    except Exception:
        pass



def normalizar_pdf_para_firma(pdf_entrada, pdf_salida):
    """Reescribe el PDF de Excel para eliminar hybrid cross-reference."""
    try:
        reader = PdfReader(str(pdf_entrada), strict=False)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        try:
            if reader.metadata:
                metadata = {
                    str(k): str(v)
                    for k, v in reader.metadata.items()
                    if k and v is not None
                }
                if metadata:
                    writer.add_metadata(metadata)
        except Exception:
            pass

        with open(pdf_salida, "wb") as f:
            writer.write(f)

        if not Path(pdf_salida).exists() or Path(pdf_salida).stat().st_size == 0:
            raise RuntimeError("No se pudo crear el PDF normalizado.")

        return Path(pdf_salida)

    except Exception as e:
        raise RuntimeError(
            "No se pudo normalizar el PDF generado por Excel.\n\n"
            f"Detalle: {e}"
        )


def firmar_pdf(pdf_entrada, pdf_salida, p12_path, password, apariencia_pdf):
    """
    Firma criptográficamente el PDF con el certificado PKCS#12.
    El PDF de Excel se normaliza antes para eliminar hybrid xrefs.
    """
    signer = signers.SimpleSigner.load_pkcs12(
        pfx_file=str(p12_path),
        passphrase=password.encode("utf-8") if password else None,
    )

    pdf_normalizado = Path(pdf_salida).with_name(
        Path(pdf_salida).stem + "_BASE.pdf"
    )

    normalizar_pdf_para_firma(pdf_entrada, pdf_normalizado)

    try:
        with open(pdf_normalizado, "rb") as entrada, open(pdf_salida, "wb") as salida:
            writer = IncrementalPdfFileWriter(
                entrada,
                strict=True,
            )

            fields.append_signature_field(
                writer,
                sig_field_spec=fields.SigFieldSpec(
                    SIGNATURE_FIELD_NAME,
                    box=SIGNATURE_BOX,
                    on_page=0,
                ),
            )

            metadata = signers.PdfSignatureMetadata(
                field_name=SIGNATURE_FIELD_NAME,
                md_algorithm="sha256",
                reason="Firma electrónica del documento",
            )

            pdf_signer = signers.PdfSigner(
                metadata,
                signer=signer,
                stamp_style=stamp.StaticStampStyle.from_pdf_file(
                    str(apariencia_pdf),
                    border_width=0,
                ),
            )

            pdf_signer.sign_pdf(
                writer,
                output=salida,
            )
    finally:
        try:
            if pdf_normalizado.exists():
                pdf_normalizado.unlink()
        except Exception:
            pass


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def procesar(excel_path, p12_path, password, carpeta_salida):
    excel_path = Path(excel_path)
    p12_path = Path(p12_path)
    carpeta_salida = Path(carpeta_salida)

    # Crear automáticamente la carpeta de salida si no existe.
    # Esto evita el error [Errno 2] No such file or directory al guardar
    # el PDF firmado.
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        raise FileNotFoundError("No se encontró el archivo Excel.")

    if excel_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("Selecciona un archivo .xlsx o .xlsm.")

    if not p12_path.exists():
        raise FileNotFoundError("No se encontró el certificado .P12/.PFX.")

    # Microsoft Excel se utiliza directamente para convertir la plantilla.
    # LibreOffice NO es obligatorio.
    datos_cert = cargar_datos_certificado(p12_path, password)

    with tempfile.TemporaryDirectory(prefix="firmador_excel_") as temp:
        temp = Path(temp)

        # 1. Copiar la plantilla. El original nunca se modifica.
        excel_preparado = temp / "plantilla_para_firmar.xlsx"
        shutil.copy2(excel_path, excel_preparado)

        # 2. Convertir con Microsoft Excel.
        # Durante la conversión se elimina automáticamente A45 ("FIRMA").
        pdf_sin_firma = convertir_excel_a_pdf(
            excel_preparado,
            temp,
        )

        # 3. Hash del PDF antes de insertar la firma.
        hash_pdf = sha256_archivo(pdf_sin_firma)

        # 4. Fecha/hora local.
        fecha_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # 5. Crear estampa visible automática.
        apariencia = temp / "apariencia_firma.pdf"
        crear_estampa_firma(
            apariencia,
            datos_cert["nombre"],
            datos_cert["serial"],
            hash_pdf,
            fecha_hora,
        )

        # 6. Nombre final.
        salida = carpeta_salida / (
            excel_path.stem + "_FIRMADO.pdf"
        )

        # La carpeta puede haber sido eliminada/cambiada mientras se
        # procesaba el documento; la creamos nuevamente por seguridad.
        salida.parent.mkdir(parents=True, exist_ok=True)

        # 7. Firma digital real.
        firmar_pdf(
            pdf_sin_firma,
            salida,
            p12_path,
            password,
            apariencia,
        )

    return salida, datos_cert


# ============================================================
# INTERFAZ GRÁFICA
# ============================================================

class Aplicacion(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Firmador Excel - Firma electrónica")
        self.geometry("720x390")
        self.resizable(False, False)

        self.excel_var = tk.StringVar()
        self.p12_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.salida_var = tk.StringVar(
            value=str(Path.home() / "Documentos" / "Documentos_Firmados")
        )
        self.estado_var = tk.StringVar(
            value="Seleccione el Excel y su certificado .P12."
        )

        self.crear_interfaz()

    def crear_interfaz(self):
        margen = 18

        titulo = ttk.Label(
            self,
            text="FIRMADOR ELECTRÓNICO DE DOCUMENTOS",
            font=("Segoe UI", 16, "bold"),
        )
        titulo.pack(pady=(18, 4))

        subtitulo = ttk.Label(
            self,
            text="Excel → PDF → firma digital visible automática",
            font=("Segoe UI", 10),
        )
        subtitulo.pack(pady=(0, 18))

        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=margen)

        # Excel
        ttk.Label(frame, text="Archivo Excel:").grid(
            row=0, column=0, sticky="w", pady=7
        )
        ttk.Entry(
            frame,
            textvariable=self.excel_var,
            width=70,
        ).grid(row=0, column=1, padx=8)
        ttk.Button(
            frame,
            text="Buscar...",
            command=self.seleccionar_excel,
        ).grid(row=0, column=2)

        # P12
        ttk.Label(frame, text="Certificado .P12/.PFX:").grid(
            row=1, column=0, sticky="w", pady=7
        )
        ttk.Entry(
            frame,
            textvariable=self.p12_var,
            width=70,
        ).grid(row=1, column=1, padx=8)
        ttk.Button(
            frame,
            text="Buscar...",
            command=self.seleccionar_p12,
        ).grid(row=1, column=2)

        # Password
        ttk.Label(frame, text="Contraseña del certificado:").grid(
            row=2, column=0, sticky="w", pady=7
        )
        ttk.Entry(
            frame,
            textvariable=self.password_var,
            width=35,
            show="*",
        ).grid(row=2, column=1, sticky="w", padx=8)

        # Output
        ttk.Label(frame, text="Carpeta de salida:").grid(
            row=3, column=0, sticky="w", pady=7
        )
        ttk.Entry(
            frame,
            textvariable=self.salida_var,
            width=70,
        ).grid(row=3, column=1, padx=8)
        ttk.Button(
            frame,
            text="Buscar...",
            command=self.seleccionar_salida,
        ).grid(row=3, column=2)

        ttk.Separator(self, orient="horizontal").pack(
            fill="x", padx=margen, pady=15
        )

        self.boton = ttk.Button(
            self,
            text="  FIRMAR DOCUMENTO  ",
            command=self.firmar,
        )
        self.boton.pack(ipady=8)

        ttk.Label(
            self,
            textvariable=self.estado_var,
            wraplength=670,
            justify="center",
        ).pack(pady=16)

        ttk.Label(
            self,
            text="La firma se colocará automáticamente en la zona que dice FIRMA.",
            font=("Segoe UI", 9, "italic"),
        ).pack()

    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar plantilla Excel",
            filetypes=[
                ("Excel", "*.xlsx *.xlsm"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if ruta:
            self.excel_var.set(ruta)

    def seleccionar_p12(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar certificado digital",
            filetypes=[
                ("Certificado PKCS#12", "*.p12 *.pfx"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if ruta:
            self.p12_var.set(ruta)

    def seleccionar_salida(self):
        ruta = filedialog.askdirectory(
            title="Seleccionar carpeta de salida"
        )
        if ruta:
            self.salida_var.set(ruta)

    def firmar(self):
        excel = self.excel_var.get().strip()
        p12 = self.p12_var.get().strip()
        password = self.password_var.get()
        salida = self.salida_var.get().strip()

        if not excel:
            messagebox.showwarning(
                "Falta Excel",
                "Seleccione el archivo Excel."
            )
            return

        if not p12:
            messagebox.showwarning(
                "Falta certificado",
                "Seleccione el certificado .P12/.PFX."
            )
            return

        if not salida:
            messagebox.showwarning(
                "Falta carpeta",
                "Seleccione la carpeta de salida."
            )
            return

        self.boton.config(state="disabled")
        self.estado_var.set("Firmando... espere, no cierre el programa.")
        self.update_idletasks()

        try:
            resultado, datos = procesar(
                excel,
                p12,
                password,
                salida,
            )

            self.estado_var.set(
                "Documento firmado correctamente."
            )

            messagebox.showinfo(
                "Firma completada",
                "PDF firmado correctamente.\n\n"
                f"Firmante: {datos['nombre']}\n"
                f"Archivo:\n{resultado}",
            )

            # Abrir carpeta de salida en Windows.
            if os.name == "nt":
                os.startfile(str(resultado.parent))

        except Exception as e:
            self.estado_var.set("Ocurrió un error durante la firma.")
            messagebox.showerror(
                "Error",
                f"No se pudo firmar el documento:\n\n{e}"
            )

        finally:
            self.boton.config(state="normal")


if __name__ == "__main__":
    app = Aplicacion()
    app.mainloop()


