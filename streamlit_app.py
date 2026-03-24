import io
import re
import zipfile
import unicodedata
import shutil
from pathlib import Path

import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

st.set_page_config(
    page_title="Extractor de Certificados Electorales",
    page_icon="🗳️",
    layout="wide",
)

# =========================================================
# UTILIDADES
# =========================================================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    valor = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("utf-8")
    valor = valor.upper()
    valor = re.sub(r"\s+", " ", valor)
    return valor.strip()


def solo_digitos(valor):
    return re.sub(r"\D", "", str(valor or ""))


def slug_nombre(valor):
    valor = normalizar_texto(valor)
    valor = re.sub(r"[^A-Z0-9]+", "_", valor)
    valor = re.sub(r"_+", "_", valor).strip("_")
    return valor[:120]


def tesseract_disponible():
    return shutil.which("tesseract") is not None


def mejorar_imagen(img):
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = ImageEnhance.Contrast(img).enhance(1.2)
    img = img.filter(ImageFilter.MedianFilter(size=3))
    return img


def extraer_texto(img):
    img = mejorar_imagen(img)

    try:
        texto = pytesseract.image_to_string(img, lang="spa")
        if texto and texto.strip():
            return texto
    except Exception:
        pass

    return pytesseract.image_to_string(img)


def buscar_patron(patron, texto, flags=0, grupo=1, default=""):
    m = re.search(patron, texto, flags)
    return m.group(grupo).strip() if m else default


def limpiar_nombre_extraido(valor):
    valor = normalizar_texto(valor)
    valor = re.sub(r"FIRMA.*$", "", valor)
    valor = re.sub(r"0\d{9,}.*$", "", valor)
    valor = re.sub(r"\b(REPUBLICA|COLOMBIA|ELECTORAL|CERTIFICADO|ELECCIONES)\b.*$", "", valor)
    valor = re.sub(r"\s+", " ", valor)
    return valor.strip(" -_:")


def extraer_campos(texto, nombre_archivo):
    t = normalizar_texto(texto)
    t = t.replace("MUNICIPIO/DISTRITO", "MUNICIPIO DISTRITO")

    departamento = buscar_patron(r"\b(ATLANTICO)\b", t)
    municipio = buscar_patron(r"\b(BARRANQUILLA)\b", t)
    zona = buscar_patron(r"\bZONA\s*(\d{1,3})\b", t)
    mesa = buscar_patron(r"\bMESA\s*(\d{1,3})\b", t)

    puesto = buscar_patron(
        r"(?:ATLANTICO\s+DEPARTAMENTO\s+)?(?:BARRANQUILLA\s+MUNICIPIO DISTRITO\s+)?(.+?)\s+PUESTO DE VOTACION",
        t,
        flags=re.DOTALL,
    )
    puesto = limpiar_nombre_extraido(puesto)

    cedula = buscar_patron(r"\bNUM\.?\s*([0-9 .]{6,20})", t)
    cedula = solo_digitos(cedula)

    if not cedula:
        candidatos = re.findall(r"\b\d{6,12}\b", t)
        candidatos = [c for c in candidatos if not c.startswith("031")]
        cedula = candidatos[0] if candidatos else ""

    nombre = buscar_patron(
        r"NOMBRES Y APELLIDOS\s+(.+?)(?:\b0\d{9,}|\bFIRMA\b|$)",
        t,
        flags=re.DOTALL,
    )
    nombre = limpiar_nombre_extraido(nombre)

    if not cedula:
        base_nombre = Path(nombre_archivo).stem
        m = re.search(r"(\d{6,12})", base_nombre)
        cedula = m.group(1) if m else ""

    return {
        "CEDULA": cedula,
        "NOMBRE": nombre,
        "DEPARTAMENTO": departamento,
        "MUNICIPIO": municipio,
        "PUESTO": puesto,
        "ZONA": zona,
        "MESA": mesa,
        "TEXTO_OCR": t,
    }


def generar_nombre_archivo(cedula, nombre, original):
    ext = Path(original).suffix.lower() or ".jpg"
    if cedula:
        return f"{cedula}_{slug_nombre(nombre) or 'SIN_NOMBRE'}{ext}"
    return f"SIN_CEDULA_{slug_nombre(nombre) or slug_nombre(Path(original).stem) or 'ARCHIVO'}{ext}"


def generar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="certificados")
    output.seek(0)
    return output.getvalue()


def generar_zip(archivos_dict):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_destino, contenido in archivos_dict.items():
            zf.writestr(nombre_destino, contenido)
    output.seek(0)
    return output.getvalue()


# =========================================================
# INTERFAZ
# =========================================================
st.title("🗳️ Extractor de certificados electorales")
st.caption("Sube certificados, extrae los datos y genera una tabla con el certificado adjunto por nombre de archivo.")

if not tesseract_disponible():
    st.error("Tesseract no está disponible. Revisa packages.txt.")
    st.stop()

certificados = st.file_uploader(
    "Sube certificados (JPG, JPEG, PNG)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if not certificados:
    st.info("Sube uno o varios certificados para comenzar.")
    st.stop()

if st.button("Procesar certificados", type="primary", use_container_width=True):
    resultados = []
    archivos_zip = {}

    barra = st.progress(0)
    estado = st.empty()

    total = len(certificados)

    for i, archivo in enumerate(certificados, start=1):
        estado.write(f"Procesando {i}/{total}: {archivo.name}")

        try:
            imagen = Image.open(archivo).convert("RGB")
            texto = extraer_texto(imagen)
            extraido = extraer_campos(texto, archivo.name)

            nombre_final = generar_nombre_archivo(
                extraido["CEDULA"],
                extraido["NOMBRE"],
                archivo.name,
            )

            archivos_zip[f"CERTIFICADOS_RENOMBRADOS/{nombre_final}"] = archivo.getvalue()

            resultados.append({
                "CEDULA": extraido["CEDULA"],
                "NOMBRE": extraido["NOMBRE"],
                "DEPARTAMENTO": extraido["DEPARTAMENTO"],
                "MUNICIPIO": extraido["MUNICIPIO"],
                "PUESTO": extraido["PUESTO"],
                "ZONA": extraido["ZONA"],
                "MESA": extraido["MESA"],
                "ARCHIVO_ORIGINAL": archivo.name,
                "ARCHIVO_CERTIFICADO": nombre_final,
                "LINK_CERTIFICADO": f"PENDIENTE_SUBIR_A_DRIVE/{nombre_final}",
            })

        except Exception as e:
            resultados.append({
                "CEDULA": "",
                "NOMBRE": "",
                "DEPARTAMENTO": "",
                "MUNICIPIO": "",
                "PUESTO": "",
                "ZONA": "",
                "MESA": "",
                "ARCHIVO_ORIGINAL": archivo.name,
                "ARCHIVO_CERTIFICADO": archivo.name,
                "LINK_CERTIFICADO": "",
                "ERROR": str(e),
            })

        barra.progress(i / total)

    df_resultados = pd.DataFrame(resultados)

    st.success("Procesamiento terminado.")
    st.dataframe(df_resultados, use_container_width=True, height=500)

    st.download_button(
        "Descargar Excel",
        data=generar_excel(df_resultados),
        file_name="certificados_extraidos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.download_button(
        "Descargar ZIP de certificados renombrados",
        data=generar_zip(archivos_zip),
        file_name="certificados_renombrados.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.markdown("### Apps Script opcional para pegar links en Google Sheets")
    st.code(
        """
function vincularCertificados() {
  const carpetaId = 'PEGA_AQUI_ID_CARPETA';
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = hoja.getDataRange().getValues();
  const encabezados = data[0];

  const colArchivo = encabezados.indexOf('ARCHIVO_CERTIFICADO');
  const colLink = encabezados.indexOf('LINK_CERTIFICADO');

  if (colArchivo === -1 || colLink === -1) {
    throw new Error('La hoja debe tener ARCHIVO_CERTIFICADO y LINK_CERTIFICADO');
  }

  const mapa = {};
  for (let i = 1; i < data.length; i++) {
    const archivo = String(data[i][colArchivo] || '').trim();
    if (archivo) mapa[archivo] = i + 1;
  }

  const archivos = DriveApp.getFolderById(carpetaId).getFiles();

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombre = archivo.getName();

    if (mapa[nombre]) {
      hoja.getRange(mapa[nombre], colLink + 1).setValue(archivo.getUrl());
    }
  }
}
        """,
        language="javascript",
    )

    st.info("Si luego subes el ZIP a Drive, puedes usar ese script para llenar automáticamente la columna LINK_CERTIFICADO en Google Sheets.")
