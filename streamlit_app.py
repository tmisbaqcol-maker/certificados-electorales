import streamlit as st
import re
import os
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import pytesseract
import gdown
import shutil

st.set_page_config(page_title="Extractor Num desde Drive", layout="wide")

# =========================
# FUNCIONES
# =========================
def extraer_id_drive(url):
    match = re.search(r"folders/([a-zA-Z0-9_-]+)", url)
    return match.group(1) if match else None


def solo_digitos(texto):
    return re.sub(r"\D", "", texto or "")


def recortar_zona_num(img):
    w, h = img.size
    return img.crop((
        int(w * 0.03),
        int(h * 0.42),
        int(w * 0.62),
        int(h * 0.72)
    ))


def mejorar(img):
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = img.filter(ImageFilter.MedianFilter(size=3))
    return img


def extraer_num(img):
    zona = recortar_zona_num(img)
    zona = mejorar(zona)

    try:
        texto = pytesseract.image_to_string(zona, lang="spa")
    except:
        texto = pytesseract.image_to_string(zona)

    texto = texto.upper()

    match = re.search(r"NUM\.?\s*([0-9\s]{6,20})", texto)
    if match:
        return solo_digitos(match.group(1))

    nums = re.findall(r"\b\d{6,12}\b", texto)
    nums = [n for n in nums if not n.startswith("031")]

    return nums[0] if nums else ""


# =========================
# INTERFAZ
# =========================
st.title("🗳️ Extractor de certificados desde Drive")

url = st.text_input("Pega el link de la carpeta de Drive")

if st.button("Procesar"):

    folder_id = extraer_id_drive(url)

    if not folder_id:
        st.error("Link inválido")
        st.stop()

    carpeta_local = "certificados"

    # limpiar si existe
    if os.path.exists(carpeta_local):
        shutil.rmtree(carpeta_local)

    st.write("Descargando archivos...")

    try:
        gdown.download_folder(
            id=folder_id,
            output=carpeta_local,
            quiet=False,
            use_cookies=False
        )
    except Exception as e:
        st.error(f"Error descargando: {e}")
        st.stop()

    st.success("Descarga completa")

    resultados = []

    for archivo in os.listdir(carpeta_local):
        if archivo.lower().endswith((".jpg", ".jpeg", ".png")):

            path = os.path.join(carpeta_local, archivo)

            try:
                img = Image.open(path).convert("RGB")
                num = extraer_num(img)

                resultados.append({
                    "ARCHIVO": archivo,
                    "NUM": num
                })

            except Exception as e:
                resultados.append({
                    "ARCHIVO": archivo,
                    "NUM": "",
                    "ERROR": str(e)
                })

    df = pd.DataFrame(resultados)

    st.dataframe(df)

    st.download_button(
        "Descargar Excel",
        df.to_csv(index=False).encode(),
        "resultado.csv",
        "text/csv"
    )
