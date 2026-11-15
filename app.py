import datetime
import json
import os
import pathlib
import tempfile
import openpyxl
import streamlit as st
from dotenv import load_dotenv

from google_connector import (
    obtener_maestra,
    obtener_mapa_sheets,
    guardar_mapa_sheets,
    registrar_auditoria_sheets,
)
from llenadores import llenar_excel, llenar_word, llenar_pdf, llenar_csv
from mapeo_gemini import generar_mapa_con_gemini

load_dotenv()

# Cargar GEMINI_API_KEY desde st.secrets si estamos en la nube
if "GEMINI_API_KEY" in st.secrets:
    os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]

CRED_GOOGLE = "credenciales_google.json"
NOMBRE_SHEET = "Base_Maestra_Proveedores"

st.set_page_config(page_title="Portal de Formularios de Proveedores", layout="wide")
st.title("Portal de Diligenciamiento Automático de Formularios")

# Control de acceso opcional si se configura CLAVE_ACCESO en los secretos de la nube
clave_configurada = st.secrets.get("CLAVE_ACCESO", "") if hasattr(st, "secrets") else ""
if clave_configurada:
    clave_ingresada = st.sidebar.text_input("Clave de acceso corporativo", type="password")
    if clave_ingresada != clave_configurada:
        st.info("Ingresa la clave corporativa en el panel izquierdo para habilitar el sistema.")
        st.stop()

try:
    datos_maestra = obtener_maestra(CRED_GOOGLE, NOMBRE_SHEET)
except Exception as e:
    st.error(f"Error al leer la Base Maestra en Google Sheets: {e}")
    st.stop()

with st.expander("Consultar datos vigentes de la Base Maestra (Google Sheets)"):
    st.json(datos_maestra)

col1, col2 = st.columns(2)
with col1:
    usuario = st.text_input("Correo corporativo de quien diligencia")
with col2:
    proveedor = st.text_input("Nombre del Proveedor o Entidad")

archivo = st.file_uploader(
    "Sube el formulario del proveedor (.pdf, .xlsx, .docx, .csv)",
    type=["pdf", "xlsx", "docx", "csv"]
)

if archivo and usuario and proveedor:
    ext = pathlib.Path(archivo.name).suffix.lower()
    identificador_plantilla = f"{pathlib.Path(archivo.name).stem}_{ext.replace('.', '')}"

    with tempfile.TemporaryDirectory() as tmp_dir:
        ruta_temp = pathlib.Path(tmp_dir) / archivo.name
        ruta_temp.write_bytes(archivo.getvalue())

        if ext == ".xlsx":
            wb_info = openpyxl.load_workbook(ruta_temp, read_only=True)
            st.caption(f"Hojas detectadas en el archivo Excel: {', '.join(wb_info.sheetnames)}")
            wb_info.close()

        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            forzar_reescaneo = st.button("Re-escanear con IA")

        mapa_actual = None if forzar_reescaneo else obtener_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla)

        if mapa_actual is None:
            with st.spinner("Formulario nuevo detectado. Analizando estructura con Gemini..."):
                mapa_actual = generar_mapa_con_gemini(ruta_temp, list(datos_maestra.keys()))
                guardar_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla, mapa_actual)
            st.info("Mapeo inicial generado. Revisa las celdas o coordenadas antes de aprobar.")

        st.subheader("1. Revisión de campos detectados")
        campos_editados = st.data_editor(
            mapa_actual.get("campos", []),
            num_rows="dynamic",
            width="stretch"
        )

        if st.button("2. Aprobar Mapa y Generar Formulario", type="primary"):
            mapa_actual["campos"] = campos_editados
            guardar_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla, mapa_actual)

            marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta_salida = pathlib.Path(tmp_dir) / f"{proveedor}_{marca}_{archivo.name}"

            if ext == ".xlsx":
                llenar_excel(ruta_temp, ruta_salida, mapa_actual, datos_maestra)
            elif ext == ".docx":
                llenar_word(ruta_temp, ruta_salida, mapa_actual, datos_maestra)
            elif ext == ".pdf":
                llenar_pdf(ruta_temp, ruta_salida, mapa_actual, datos_maestra)
            elif ext == ".csv":
                llenar_csv(ruta_temp, ruta_salida, mapa_actual, datos_maestra)

            consecutivo = registrar_auditoria_sheets(
                CRED_GOOGLE, NOMBRE_SHEET, usuario, proveedor, ext, ruta_salida
            )
            nombre_final = f"Consecutivo_{consecutivo:04d}_{proveedor}_{archivo.name}"

            st.success(f"Formulario diligenciado correctamente. Consecutivo registrado: #{consecutivo}")
            st.download_button(
                label="3. Descargar Formulario Diligenciado",
                data=ruta_salida.read_bytes(),
                file_name=nombre_final
            )