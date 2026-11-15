import datetime
import importlib
import os
import pathlib
import tempfile
import openpyxl
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

import google_connector
import llenadores
import mapeo_gemini

# Forzar a Streamlit Cloud a recargar en memoria la última versión subida a GitHub
importlib.reload(google_connector)
importlib.reload(llenadores)
importlib.reload(mapeo_gemini)

from google_connector import (
    cargar_datos_completos_maestra,
    obtener_maestra,
    obtener_mapa_sheets,
    guardar_mapa_sheets,
    registrar_auditoria_sheets,
)
from llenadores import llenar_excel, llenar_word, llenar_pdf, llenar_csv
from mapeo_gemini import generar_mapa_con_gemini

CRED_GOOGLE = "credenciales_google.json"
NOMBRE_SHEET = "Base_Maestra_Proveedores"

st.set_page_config(page_title="Portal de Formularios de Proveedores", layout="wide")
st.title("Portal de Diligenciamiento Automático de Formularios")

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

# Contador de campos vacíos en la pestaña Maestra
claves_vacias_maestra = [k for k, v in datos_maestra.items() if str(v).strip() in ("", "VACIO", "None")]

col_estado1, col_estado2 = st.columns([3, 2])
with col_estado1:
    if claves_vacias_maestra:
        st.warning(
            f"Atención: Tu pestaña 'Maestra' en Google Sheets tiene {len(claves_vacias_maestra)} campos marcados como 'VACIO': "
            f"{', '.join(claves_vacias_maestra[:8])}."
        )
    else:
        st.success(f"Base Maestra activa y completa ({len(datos_maestra)} campos listos con información).")

with col_estado2:
    if st.button("Cargar todos los datos oficiales en la pestaña Maestra"):
        cargar_datos_completos_maestra(CRED_GOOGLE, NOMBRE_SHEET)
        st.success("Se cargaron los 29 campos completos en la pestaña Maestra de tu Google Sheets.")
        st.rerun()

with st.expander("Ver tabla completa de la pestaña Maestra (muestra VACIO donde falta información)"):
    df_maestra = pd.DataFrame(
        [{"Clave en Maestra": k, "Valor Actual": v, "Estado": "VACIO" if str(v) == "VACIO" else "LLENO"}
         for k, v in datos_maestra.items()]
    )
    st.dataframe(df_maestra, width="stretch")

col1, col2, col3 = st.columns(3)
with col1:
    usuario = st.text_input("Correo de quien solicita")
with col2:
    area_solicitante = st.text_input("Área que lo solicita (ej. Compras, Jurídica)")
with col3:
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

        forzar_reescaneo = st.button("Actualizar / Reconstruir Mapa del Formulario")

        mapa_actual = None if forzar_reescaneo else obtener_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla)

        if mapa_actual is None:
            with st.spinner("Construyendo mapa exacto de celdas/coordenadas..."):
                mapa_actual = generar_mapa_con_gemini(ruta_temp, list(datos_maestra.keys()))
                guardar_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla, mapa_actual, datos_maestra)

        # Agregar en vivo la columna de vista previa del valor y si está LLENO o VACIO
        campos_con_estado = []
        for c in mapa_actual.get("campos", []):
            item = dict(c)
            val = str(datos_maestra.get(item.get("clave", ""), "VACIO")).strip()
            item["valor_a_escribir"] = val if val != "" else "VACIO"
            item["estado"] = "VACIO" if item["valor_a_escribir"] == "VACIO" else "LLENO"
            campos_con_estado.append(item)

        st.subheader("1. Control de llenado: Dónde se escribe cada dato y cuáles están en VACIO")
        campos_editados = st.data_editor(
            campos_con_estado,
            num_rows="dynamic",
            width="stretch"
        )

        if st.button("2. Guardar en Google Sheets y Generar Archivo Diligenciado", type="primary"):
            mapa_actual["campos"] = campos_editados
            guardar_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla, mapa_actual, datos_maestra)

            # Solo enviar a los llenadores los datos que no estén marcados como VACIO
            datos_limpios = {k: v for k, v in datos_maestra.items() if str(v).strip() not in ("", "VACIO")}

            marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta_salida = pathlib.Path(tmp_dir) / f"{proveedor}_{marca}_{archivo.name}"

            if ext == ".xlsx":
                llenar_excel(ruta_temp, ruta_salida, mapa_actual, datos_limpios)
            elif ext == ".docx":
                llenar_word(ruta_temp, ruta_salida, mapa_actual, datos_limpios)
            elif ext == ".pdf":
                llenar_pdf(ruta_temp, ruta_salida, mapa_actual, datos_limpios)
            elif ext == ".csv":
                llenar_csv(ruta_temp, ruta_salida, mapa_actual, datos_limpios)

            faltantes = [c["clave"] for c in campos_editados if str(datos_maestra.get(c.get("clave"), "VACIO")) == "VACIO"]
            llenados = len(campos_editados) - len(faltantes)
            usuario_completo = f"{usuario} ({area_solicitante})" if area_solicitante else usuario

            consecutivo = registrar_auditoria_sheets(
                CRED_GOOGLE, NOMBRE_SHEET, usuario_completo, proveedor, ext, ruta_salida, llenados, faltantes
            )
            nombre_final = f"Consecutivo_{consecutivo:04d}_{proveedor}_{archivo.name}"

            st.success(
                f"Consecutivo #{consecutivo} registrado en Auditoría. "
                f"Campos llenados: {llenados} | Campos en VACIO: {len(faltantes)}."
            )
            st.download_button(
                label="3. Descargar Formulario Diligenciado",
                data=ruta_salida.read_bytes(),
                file_name=nombre_final
            )