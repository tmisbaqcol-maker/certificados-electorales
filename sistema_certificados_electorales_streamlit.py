import io
import os
import re
import zipfile
import shutil
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageFilter, ImageEnhance

# OCR
import easyocr

st.set_page_config(page_title="Sistema de Certificados Electorales", page_icon="🗳️", layout="wide")

# ============================================================
# CONFIG
# ============================================================
COLUMNAS_SUGERIDAS = [
    "CEDULA",
    "NOMBRE",
    "DEPARTAMENTO",
    "MUNICIPIO",
    "PUESTO",
    "ZONA",
    "MESA",
    "ESTADO_CRUCE",
    "OBSERVACION",
    "ARCHIVO_CERTIFICADO",
    "LINK_CERTIFICADO",
]

ALIASES_CEDULA = ["cedula", "cédula", "cc", "documento", "num", "numero", "número", "identificacion", "identificación"]
ALIASES_NOMBRE = ["nombre", "nombres", "apellidos", "nombres y apellidos"]
ALIASES_PUESTO = ["puesto", "puesto de votacion", "puesto de votación"]
ALIASES_ZONA = ["zona"]
ALIASES_MESA = ["mesa"]
ALIASES_LINK = ["link", "link_certificado", "certificado", "url", "enlace"]

# ============================================================
# UTILS
# ============================================================
def normalizar_texto(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
    s = s.upper().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def slug_nombre(s: str) -> str:
    s = normalizar_texto(s)
    s = re.sub(r"[^A-Z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s[:120]


def solo_digitos(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))


def detectar_columna(cols, aliases):
    mapa = {c: normalizar_texto(c) for c in cols}
    for col, norm in mapa.items():
        for a in aliases:
            if normalizar_texto(a) == norm:
                return col
    for col, norm in mapa.items():
        for a in aliases:
            if normalizar_texto(a) in norm:
                return col
    return None


def preparar_base(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    col_cedula = detectar_columna(df.columns, ALIASES_CEDULA)
    col_nombre = detectar_columna(df.columns, ALIASES_NOMBRE)
    col_puesto = detectar_columna(df.columns, ALIASES_PUESTO)
    col_zona = detectar_columna(df.columns, ALIASES_ZONA)
    col_mesa = detectar_columna(df.columns, ALIASES_MESA)
    col_link = detectar_columna(df.columns, ALIASES_LINK)

    if col_cedula is None:
        st.error("La base debe tener una columna de cédula. Ejemplo: CEDULA o CC")
        st.stop()

    if col_nombre is None:
        df["NOMBRE"] = ""
        col_nombre = "NOMBRE"

    if col_puesto is None:
        df["PUESTO"] = ""
        col_puesto = "PUESTO"

    if col_zona is None:
        df["ZONA"] = ""
        col_zona = "ZONA"

    if col_mesa is None:
        df["MESA"] = ""
        col_mesa = "MESA"

    if col_link is None:
        df["LINK_CERTIFICADO"] = ""
        col_link = "LINK_CERTIFICADO"

    df["_CEDULA"] = df[col_cedula].astype(str).apply(solo_digitos)
    df["_NOMBRE"] = df[col_nombre].astype(str).apply(normalizar_texto)
    df["_PUESTO"] = df[col_puesto].astype(str).apply(normalizar_texto)
    df["_ZONA"] = df[col_zona].astype(str).apply(solo_digitos)
    df["_MESA"] = df[col_mesa].astype(str).apply(solo_digitos)

    df["ESTADO_CRUCE"] = df.get("ESTADO_CRUCE", "")
    df["OBSERVACION"] = df.get("OBSERVACION", "")
    df["ARCHIVO_CERTIFICADO"] = df.get("ARCHIVO_CERTIFICADO", "")
    df["LINK_CERTIFICADO"] = df[col_link].astype(str)

    df.attrs["col_cedula"] = col_cedula
    df.attrs["col_nombre"] = col_nombre
    df.attrs["col_puesto"] = col_puesto
    df.attrs["col_zona"] = col_zona
    df.attrs["col_mesa"] = col_mesa
    df.attrs["col_link"] = col_link
    return df


def leer_base(archivo):
    nombre = archivo.name.lower()
    if nombre.endswith(".csv"):
        try:
            return pd.read_csv(archivo)
        except Exception:
            archivo.seek(0)
            return pd.read_csv(archivo, encoding="latin-1")
    return pd.read_excel(archivo)


def mejorar_imagen(img: Image.Image) -> Image.Image:
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = img.filter(ImageFilter.MedianFilter(size=3))
    return img


def extraer_texto(reader, img: Image.Image):
    imagen = mejorar_imagen(img)
    resultado = reader.readtext(
        image=np.array(imagen),
        detail=0,
        paragraph=True,
        width_ths=0.7,
        link_threshold=0.4,
    )
    texto = "\n".join([str(x) for x in resultado])
    return texto


def buscar_patron(pattern, texto, flags=0, group=1, default=""):
    m = re.search(pattern, texto, flags)
    return m.group(group).strip() if m else default


def limpiar_nombre_extraido(valor: str) -> str:
    valor = normalizar_texto(valor)
    valor = re.sub(r"FIRMA.*$", "", valor)
    valor = re.sub(r"0\d{9,}.*$", "", valor)
    valor = re.sub(r"\b(REPUBLICA|COLOMBIA|ELECTORAL|CERTIFICADO|ELECCIONES)\b.*$", "", valor)
    valor = re.sub(r"\s+", " ", valor).strip(" -_:")
    return valor


def extraer_campos(texto: str, nombre_archivo: str):
    t = normalizar_texto(texto)
    t = t.replace("MUNICIPIO/DISTRITO", "MUNICIPIO DISTRITO")
    t = t.replace("PUESTO DE VOTACION", "PUESTO DE VOTACION")
    t = t.replace("NOMBRES Y APELLIDOS", "NOMBRES Y APELLIDOS")

    departamento = buscar_patron(r"\b(ATLANTICO)\b", t)
    municipio = buscar_patron(r"\b(BARRANQUILLA)\b", t)
    zona = buscar_patron(r"\bZONA\s*(\d{1,3})\b", t)
    mesa = buscar_patron(r"\bMESA\s*(\d{1,3})\b", t)

    puesto = buscar_patron(
        r"(?:DEPARTAMENTO\s+ATLANTICO\s+)?(?:BARRANQUILLA\s+)?(.+?)\s+PUESTO DE VOTACION",
        t,
        flags=re.DOTALL,
    )
    puesto = limpiar_nombre_extraido(puesto)

    cedula = buscar_patron(r"\bNUM\.?\s*([0-9 .]{6,20})", t)
    cedula = solo_digitos(cedula)
    if not cedula:
        candidatos = re.findall(r"\b\d{6,12}\b", t)
        candidatos = [c for c in candidatos if len(c) >= 6 and not c.startswith("031")]
        cedula = candidatos[0] if candidatos else ""

    nombre = buscar_patron(r"NOMBRES Y APELLIDOS\s+(.+?)(?:\b0\d{9,}|\bFIRMA\b|$)", t, flags=re.DOTALL)
    nombre = limpiar_nombre_extraido(nombre)

    if not cedula:
        base_nombre = Path(nombre_archivo).stem
        m = re.search(r"(\d{6,12})", base_nombre)
        cedula = m.group(1) if m else ""

    return {
        "CEDULA_EXTRAIDA": cedula,
        "NOMBRE_EXTRAIDO": nombre,
        "DEPARTAMENTO_EXTRAIDO": departamento,
        "MUNICIPIO_EXTRAIDO": municipio,
        "PUESTO_EXTRAIDO": puesto,
        "ZONA_EXTRAIDA": zona,
        "MESA_EXTRAIDA": mesa,
        "TEXTO_OCR": t,
    }


def score_match(base_row, extraido):
    score = 0
    razones = []

    if base_row["_CEDULA"] and extraido["CEDULA_EXTRAIDA"]:
        if base_row["_CEDULA"] == extraido["CEDULA_EXTRAIDA"]:
            score += 100
            razones.append("CEDULA")

    if base_row["_MESA"] and extraido["MESA_EXTRAIDA"]:
        if base_row["_MESA"] == extraido["MESA_EXTRAIDA"]:
            score += 15
            razones.append("MESA")

    if base_row["_ZONA"] and extraido["ZONA_EXTRAIDA"]:
        if base_row["_ZONA"] == extraido["ZONA_EXTRAIDA"]:
            score += 10
            razones.append("ZONA")

    if base_row["_PUESTO"] and extraido["PUESTO_EXTRAIDO"]:
        if base_row["_PUESTO"] in extraido["PUESTO_EXTRAIDO"] or extraido["PUESTO_EXTRAIDO"] in base_row["_PUESTO"]:
            score += 12
            razones.append("PUESTO")

    if base_row["_NOMBRE"] and extraido["NOMBRE_EXTRAIDO"]:
        base_tokens = set(base_row["_NOMBRE"].split())
        ext_tokens = set(extraido["NOMBRE_EXTRAIDO"].split())
        inter = len(base_tokens.intersection(ext_tokens))
        if inter >= 2:
            score += min(inter * 5, 20)
            razones.append("NOMBRE")

    return score, ", ".join(razones)


def cruzar_registro(df_base, extraido):
    candidatos = []

    if extraido["CEDULA_EXTRAIDA"]:
        sub = df_base[df_base["_CEDULA"] == extraido["CEDULA_EXTRAIDA"]].copy()
        if not sub.empty:
            for idx, row in sub.iterrows():
                score, razones = score_match(row, extraido)
                candidatos.append((idx, score, razones))

    if not candidatos and extraido["MESA_EXTRAIDA"]:
        sub = df_base[df_base["_MESA"] == extraido["MESA_EXTRAIDA"]].copy()
        if extraido["ZONA_EXTRAIDA"]:
            sub = sub[sub["_ZONA"] == extraido["ZONA_EXTRAIDA"]]
        for idx, row in sub.iterrows():
            score, razones = score_match(row, extraido)
            candidatos.append((idx, score, razones))

    if not candidatos and extraido["NOMBRE_EXTRAIDO"]:
        nombre_tokens = set(extraido["NOMBRE_EXTRAIDO"].split())
        for idx, row in df_base.iterrows():
            base_tokens = set(row["_NOMBRE"].split())
            if len(nombre_tokens.intersection(base_tokens)) >= 2:
                score, razones = score_match(row, extraido)
                candidatos.append((idx, score, razones))

    if not candidatos:
        return None, 0, "SIN COINCIDENCIA"

    candidatos = sorted(candidatos, key=lambda x: x[1], reverse=True)
    best = candidatos[0]
    return best[0], best[1], best[2]


def generar_nombre_archivo(cedula, nombre, original):
    ext = Path(original).suffix.lower() or ".jpg"
    if cedula:
        return f"{cedula}_{slug_nombre(nombre) or 'SIN_NOMBRE'}{ext}"
    return f"SIN_CEDULA_{slug_nombre(nombre) or slug_nombre(Path(original).stem) or 'ARCHIVO'}{ext}"


def dataframe_para_descarga(df):
    salida = df.copy()
    cols_aux = [c for c in salida.columns if c.startswith("_") or c == "TEXTO_OCR"]
    salida = salida.drop(columns=cols_aux, errors="ignore")
    return salida


def a_excel_bytes(df1, df2, df3):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe_para_descarga(df1).to_excel(writer, index=False, sheet_name="base_actualizada")
        dataframe_para_descarga(df2).to_excel(writer, index=False, sheet_name="certificados_procesados")
        dataframe_para_descarga(df3).to_excel(writer, index=False, sheet_name="no_registrados")
    output.seek(0)
    return output.getvalue()


def zip_certificados(archivos_dict):
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_destino, contenido in archivos_dict.items():
            zf.writestr(nombre_destino, contenido)
    mem.seek(0)
    return mem.getvalue()


# numpy import tardío para evitar error si no se usa OCR todavía
import numpy as np

# ============================================================
# SIDEBAR
# ============================================================
st.title("🗳️ Sistema completo de certificados electorales")
st.caption("Carga base + certificados, extrae datos, cruza registros, adjunta evidencia y genera salidas listas para Google Sheets.")

with st.sidebar:
    st.header("Configuración")
    idioma_ocr = st.selectbox("Idioma OCR", ["es", "es,en"], index=0)
    umbral_match = st.slider("Umbral de coincidencia aceptada", min_value=60, max_value=130, value=100, step=5)
    st.markdown("**Flujo recomendado**")
    st.markdown("1. Exporta tu Google Sheet a Excel o CSV.\n2. Carga la base aquí.\n3. Carga los certificados en JPG, PNG o PDF convertido a imagen.\n4. Descarga el Excel final y el ZIP renombrado.")

# ============================================================
# INPUTS
# ============================================================
col1, col2 = st.columns([1, 1])
with col1:
    archivo_base = st.file_uploader("Sube la base exportada de Google Sheets", type=["xlsx", "xls", "csv"])
with col2:
    certificados = st.file_uploader(
        "Sube certificados (JPG, JPEG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )

if archivo_base is None:
    st.info("Sube primero la base exportada desde Google Sheets.")
    st.stop()

# ============================================================
# BASE
# ============================================================
df_base_raw = leer_base(archivo_base)
df_base = preparar_base(df_base_raw)

with st.expander("Vista previa de la base"):
    st.dataframe(dataframe_para_descarga(df_base).head(20), use_container_width=True)

if not certificados:
    st.warning("Ahora sube los certificados para procesarlos.")
    st.stop()

# ============================================================
# OCR READER CACHE
# ============================================================
@st.cache_resource
def get_reader(langs):
    return easyocr.Reader(langs, gpu=False)

langs = ["es"] if idioma_ocr == "es" else ["es", "en"]
reader = get_reader(langs)

# ============================================================
# PROCESS
# ============================================================
if st.button("Procesar certificados", type="primary", use_container_width=True):
    resultados = []
    no_registrados = []
    archivos_zip = {}
    progreso = st.progress(0)
    estado = st.empty()

    total = len(certificados)

    for i, archivo in enumerate(certificados, start=1):
        estado.write(f"Procesando {i}/{total}: {archivo.name}")

        try:
            img = Image.open(archivo).convert("RGB")
            texto = extraer_texto(reader, img)
            extraido = extraer_campos(texto, archivo.name)
            idx_match, score, razones = cruzar_registro(df_base, extraido)

            nombre_final = generar_nombre_archivo(
                extraido["CEDULA_EXTRAIDA"],
                extraido["NOMBRE_EXTRAIDO"],
                archivo.name,
            )

            contenido = archivo.getvalue()
            archivos_zip[f"CERTIFICADOS_RENOMBRADOS/{nombre_final}"] = contenido

            if idx_match is not None and score >= umbral_match:
                df_base.at[idx_match, "ESTADO_CRUCE"] = "ENCONTRADO"
                df_base.at[idx_match, "OBSERVACION"] = f"MATCH {score} - {razones}"
                df_base.at[idx_match, "ARCHIVO_CERTIFICADO"] = nombre_final
                if not str(df_base.at[idx_match, "LINK_CERTIFICADO"]).strip():
                    df_base.at[idx_match, "LINK_CERTIFICADO"] = f"PENDIENTE_SUBIR_A_DRIVE/{nombre_final}"
                estado_cruce = "ENCONTRADO"
                observacion = f"MATCH {score} - {razones}"
                base_ref = idx_match
            else:
                estado_cruce = "NO_REGISTRADO"
                observacion = f"SIN MATCH SUFICIENTE ({score}) - {razones}"
                base_ref = None
                no_registrados.append({
                    "CEDULA": extraido["CEDULA_EXTRAIDA"],
                    "NOMBRE": extraido["NOMBRE_EXTRAIDO"],
                    "DEPARTAMENTO": extraido["DEPARTAMENTO_EXTRAIDO"],
                    "MUNICIPIO": extraido["MUNICIPIO_EXTRAIDO"],
                    "PUESTO": extraido["PUESTO_EXTRAIDO"],
                    "ZONA": extraido["ZONA_EXTRAIDA"],
                    "MESA": extraido["MESA_EXTRAIDA"],
                    "ESTADO_CRUCE": estado_cruce,
                    "OBSERVACION": observacion,
                    "ARCHIVO_CERTIFICADO": nombre_final,
                    "LINK_CERTIFICADO": f"PENDIENTE_SUBIR_A_DRIVE/{nombre_final}",
                })

            resultados.append({
                **extraido,
                "ARCHIVO_ORIGINAL": archivo.name,
                "ARCHIVO_CERTIFICADO": nombre_final,
                "ESTADO_CRUCE": estado_cruce,
                "OBSERVACION": observacion,
                "INDICE_BASE": base_ref,
                "SCORE_MATCH": score,
                "CRITERIOS_MATCH": razones,
            })

        except Exception as e:
            resultados.append({
                "CEDULA_EXTRAIDA": "",
                "NOMBRE_EXTRAIDO": "",
                "DEPARTAMENTO_EXTRAIDO": "",
                "MUNICIPIO_EXTRAIDO": "",
                "PUESTO_EXTRAIDO": "",
                "ZONA_EXTRAIDA": "",
                "MESA_EXTRAIDA": "",
                "TEXTO_OCR": "",
                "ARCHIVO_ORIGINAL": archivo.name,
                "ARCHIVO_CERTIFICADO": archivo.name,
                "ESTADO_CRUCE": "ERROR",
                "OBSERVACION": str(e),
                "INDICE_BASE": None,
                "SCORE_MATCH": 0,
                "CRITERIOS_MATCH": "",
            })

        progreso.progress(i / total)

    estado.success("Procesamiento terminado.")

    df_resultados = pd.DataFrame(resultados)
    df_no_reg = pd.DataFrame(no_registrados)

    st.subheader("Resumen operativo")
    a = int((df_resultados["ESTADO_CRUCE"] == "ENCONTRADO").sum())
    b = int((df_resultados["ESTADO_CRUCE"] == "NO_REGISTRADO").sum())
    c = int((df_resultados["ESTADO_CRUCE"] == "ERROR").sum())
    m1, m2, m3 = st.columns(3)
    m1.metric("Encontrados", a)
    m2.metric("No registrados", b)
    m3.metric("Errores", c)

    tab1, tab2, tab3 = st.tabs(["Base actualizada", "Certificados procesados", "No registrados"])

    with tab1:
        st.dataframe(dataframe_para_descarga(df_base), use_container_width=True, height=500)

    with tab2:
        st.dataframe(dataframe_para_descarga(df_resultados), use_container_width=True, height=500)

    with tab3:
        if df_no_reg.empty:
            st.success("No hay no registrados.")
        else:
            st.dataframe(dataframe_para_descarga(df_no_reg), use_container_width=True, height=500)

    excel_bytes = a_excel_bytes(df_base, df_resultados, df_no_reg)
    zip_bytes = zip_certificados(archivos_zip)

    st.download_button(
        "Descargar Excel final",
        data=excel_bytes,
        file_name="resultado_certificados_electorales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.download_button(
        "Descargar ZIP de certificados renombrados",
        data=zip_bytes,
        file_name="certificados_renombrados.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.markdown("### Apps Script para pegar links automáticamente en Google Sheets")
    st.code(
        """
function vincularCertificados() {
  const carpetaId = 'PEGA_AQUI_ID_CARPETA';
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = hoja.getDataRange().getValues();
  const encabezados = data[0];

  const colCedula = encabezados.indexOf('CEDULA');
  const colArchivo = encabezados.indexOf('ARCHIVO_CERTIFICADO');
  const colLink = encabezados.indexOf('LINK_CERTIFICADO');

  if (colCedula === -1 || colArchivo === -1 || colLink === -1) {
    throw new Error('La hoja debe tener CEDULA, ARCHIVO_CERTIFICADO y LINK_CERTIFICADO');
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

    st.info("Sube el ZIP a una carpeta de Drive, descomprímelo, pega el Apps Script en la hoja y ejecuta la función vincularCertificados().")
