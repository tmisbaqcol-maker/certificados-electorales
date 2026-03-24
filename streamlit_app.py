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
    page_title="Sistema de Certificados Electorales",
    page_icon="🗳️",
    layout="wide",
)

ALIASES_CEDULA = [
    "cedula", "cédula", "cc", "documento", "num", "numero",
    "número", "identificacion", "identificación"
]
ALIASES_NOMBRE = ["nombre", "nombres", "apellidos", "nombres y apellidos"]
ALIASES_PUESTO = ["puesto", "puesto de votacion", "puesto de votación"]
ALIASES_ZONA = ["zona"]
ALIASES_MESA = ["mesa"]
ALIASES_LINK = ["link", "link_certificado", "certificado", "url", "enlace"]


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


def detectar_columna(columnas, aliases):
    columnas_norm = {c: normalizar_texto(c) for c in columnas}

    for col, col_norm in columnas_norm.items():
        for alias in aliases:
            if col_norm == normalizar_texto(alias):
                return col

    for col, col_norm in columnas_norm.items():
        for alias in aliases:
            if normalizar_texto(alias) in col_norm:
                return col

    return None


def leer_base(archivo):
    nombre = archivo.name.lower()

    if nombre.endswith(".csv"):
        try:
            return pd.read_csv(archivo)
        except Exception:
            archivo.seek(0)
            return pd.read_csv(archivo, encoding="latin-1")

    return pd.read_excel(archivo)


def preparar_base(df):
    df = df.copy()

    col_cedula = detectar_columna(df.columns, ALIASES_CEDULA)
    col_nombre = detectar_columna(df.columns, ALIASES_NOMBRE)
    col_puesto = detectar_columna(df.columns, ALIASES_PUESTO)
    col_zona = detectar_columna(df.columns, ALIASES_ZONA)
    col_mesa = detectar_columna(df.columns, ALIASES_MESA)
    col_link = detectar_columna(df.columns, ALIASES_LINK)

    if col_cedula is None:
        st.error("La base debe tener una columna de cédula.")
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

    if "ESTADO_CRUCE" not in df.columns:
        df["ESTADO_CRUCE"] = ""

    if "OBSERVACION" not in df.columns:
        df["OBSERVACION"] = ""

    if "ARCHIVO_CERTIFICADO" not in df.columns:
        df["ARCHIVO_CERTIFICADO"] = ""

    df["_CEDULA"] = df[col_cedula].astype(str).apply(solo_digitos)
    df["_NOMBRE"] = df[col_nombre].astype(str).apply(normalizar_texto)
    df["_PUESTO"] = df[col_puesto].astype(str).apply(normalizar_texto)
    df["_ZONA"] = df[col_zona].astype(str).apply(solo_digitos)
    df["_MESA"] = df[col_mesa].astype(str).apply(solo_digitos)
    df["LINK_CERTIFICADO"] = df[col_link].astype(str)

    return df


def dataframe_limpio(df):
    df = df.copy()
    cols_aux = [c for c in df.columns if c.startswith("_") or c == "TEXTO_OCR"]
    return df.drop(columns=cols_aux, errors="ignore")


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

    if base_row["_CEDULA"] and extraido["CEDULA_EXTRAIDA"] and base_row["_CEDULA"] == extraido["CEDULA_EXTRAIDA"]:
        score += 100
        razones.append("CEDULA")

    if base_row["_MESA"] and extraido["MESA_EXTRAIDA"] and base_row["_MESA"] == extraido["MESA_EXTRAIDA"]:
        score += 15
        razones.append("MESA")

    if base_row["_ZONA"] and extraido["ZONA_EXTRAIDA"] and base_row["_ZONA"] == extraido["ZONA_EXTRAIDA"]:
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
        sub = df_base[df_base["_CEDULA"] == extraido["CEDULA_EXTRAIDA"]]
        for idx, row in sub.iterrows():
            score, razones = score_match(row, extraido)
            candidatos.append((idx, score, razones))

    if not candidatos and extraido["MESA_EXTRAIDA"]:
        sub = df_base[df_base["_MESA"] == extraido["MESA_EXTRAIDA"]]
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

    candidatos.sort(key=lambda x: x[1], reverse=True)
    return candidatos[0]


def generar_nombre_archivo(cedula, nombre, original):
    ext = Path(original).suffix.lower() or ".jpg"
    if cedula:
        return f"{cedula}_{slug_nombre(nombre) or 'SIN_NOMBRE'}{ext}"
    return f"SIN_CEDULA_{slug_nombre(nombre) or slug_nombre(Path(original).stem) or 'ARCHIVO'}{ext}"


def generar_excel(df_base, df_resultados, df_no_reg):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe_limpio(df_base).to_excel(writer, index=False, sheet_name="base_actualizada")
        dataframe_limpio(df_resultados).to_excel(writer, index=False, sheet_name="certificados_procesados")
        dataframe_limpio(df_no_reg).to_excel(writer, index=False, sheet_name="no_registrados")
    output.seek(0)
    return output.getvalue()


def generar_zip(archivos_dict):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_destino, contenido in archivos_dict.items():
            zf.writestr(nombre_destino, contenido)
    output.seek(0)
    return output.getvalue()


st.title("🗳️ Sistema de certificados electorales")
st.caption("Cruza certificados contra tu base y prepara archivos para Google Sheets.")

with st.sidebar:
    st.header("Configuración")
    umbral_match = st.slider("Umbral de coincidencia", 60, 130, 100, 5)

if not tesseract_disponible():
    st.error("Tesseract no está disponible en el entorno. Revisa packages.txt.")
    st.stop()

col1, col2 = st.columns(2)

with col1:
    archivo_base = st.file_uploader("Sube la base", type=["xlsx", "xls", "csv"])

with col2:
    certificados = st.file_uploader(
        "Sube certificados (JPG, JPEG, PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )

if archivo_base is None:
    st.info("Sube primero la base.")
    st.stop()

try:
    df_base = preparar_base(leer_base(archivo_base))
except Exception as e:
    st.error(f"Error leyendo la base: {e}")
    st.stop()

with st.expander("Vista previa de la base"):
    st.dataframe(dataframe_limpio(df_base).head(20), use_container_width=True)

if not certificados:
    st.warning("Ahora sube los certificados.")
    st.stop()

if st.button("Procesar certificados", type="primary", use_container_width=True):
    resultados = []
    no_registrados = []
    archivos_zip = {}

    barra = st.progress(0)
    total = len(certificados)

    for i, archivo in enumerate(certificados, start=1):
        try:
            imagen = Image.open(archivo).convert("RGB")
            texto = extraer_texto(imagen)
            extraido = extraer_campos(texto, archivo.name)
            idx_match, score, razones = cruzar_registro(df_base, extraido)

            nombre_final = generar_nombre_archivo(
                extraido["CEDULA_EXTRAIDA"],
                extraido["NOMBRE_EXTRAIDO"],
                archivo.name,
            )

            archivos_zip[f"CERTIFICADOS_RENOMBRADOS/{nombre_final}"] = archivo.getvalue()

            if idx_match is not None and score >= umbral_match:
                df_base.at[idx_match, "ESTADO_CRUCE"] = "ENCONTRADO"
                df_base.at[idx_match, "OBSERVACION"] = f"MATCH {score} - {razones}"
                df_base.at[idx_match, "ARCHIVO_CERTIFICADO"] = nombre_final

                if not str(df_base.at[idx_match, "LINK_CERTIFICADO"]).strip():
                    df_base.at[idx_match, "LINK_CERTIFICADO"] = f"PENDIENTE_SUBIR_A_DRIVE/{nombre_final}"

                estado_cruce = "ENCONTRADO"
                observacion = f"MATCH {score} - {razones}"
                indice_base = idx_match
            else:
                estado_cruce = "NO_REGISTRADO"
                observacion = f"SIN MATCH SUFICIENTE ({score}) - {razones}"
                indice_base = None

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
                "INDICE_BASE": indice_base,
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

        barra.progress(i / total)

    df_resultados = pd.DataFrame(resultados)
    df_no_reg = pd.DataFrame(no_registrados)

    c1, c2, c3 = st.columns(3)
    c1.metric("Encontrados", int((df_resultados["ESTADO_CRUCE"] == "ENCONTRADO").sum()))
    c2.metric("No registrados", int((df_resultados["ESTADO_CRUCE"] == "NO_REGISTRADO").sum()))
    c3.metric("Errores", int((df_resultados["ESTADO_CRUCE"] == "ERROR").sum()))

    tab1, tab2, tab3 = st.tabs(["Base actualizada", "Procesados", "No registrados"])

    with tab1:
        st.dataframe(dataframe_limpio(df_base), use_container_width=True, height=500)

    with tab2:
        st.dataframe(dataframe_limpio(df_resultados), use_container_width=True, height=500)

    with tab3:
        st.dataframe(dataframe_limpio(df_no_reg), use_container_width=True, height=500)

    st.download_button(
        "Descargar Excel final",
        data=generar_excel(df_base, df_resultados, df_no_reg),
        file_name="resultado_certificados_electorales.xlsx",
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
