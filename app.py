import datetime
import os
import pathlib
import tempfile
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from google_connector import (
    cargar_datos_completos_maestra,
    obtener_maestra,
    obtener_mapa_sheets,
    guardar_mapa_sheets,
    registrar_auditoria_sheets,
    sincronizar_respuestas_forms_a_maestra,
)
from llenadores import llenar_excel, llenar_word, llenar_pdf, llenar_csv
from mapeo_gemini import generar_mapa_con_gemini

load_dotenv()

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

claves_vacias_maestra = [k for k, v in datos_maestra.items() if str(v).strip() in ("", "VACIO", "None")]

col_estado1, col_btn_cargar, col_btn_sync = st.columns([2, 1, 1])
with col_estado1:
    if claves_vacias_maestra:
        st.warning(
            f"Tu pestaña 'Maestra' tiene {len(claves_vacias_maestra)} campos en 'VACIO': "
            f"{', '.join(claves_vacias_maestra[:6])}..."
        )
    else:
        st.success(f"Base Maestra activa en memoria ({len(datos_maestra)} campos listos).")

with col_btn_cargar:
    if st.button("Cargar los 29 datos oficiales en Maestra"):
        cargar_datos_completos_maestra(CRED_GOOGLE, NOMBRE_SHEET)
        st.success("Datos oficiales cargados.")
        st.rerun()

with col_btn_sync:
    if st.button("Sincronizar respuestas de Google Forms"):
        n_sync = sincronizar_respuestas_forms_a_maestra(CRED_GOOGLE, NOMBRE_SHEET)
        st.success(f"Sincronizados {n_sync} campos desde Google Forms.")
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

        campos_con_estado = []
        for c in mapa_actual.get("campos", []):
            item = dict(c)
            val = str(datos_maestra.get(item.get("clave", ""), "VACIO")).strip()
            item["valor_a_escribir"] = val if val != "" else "VACIO"
            item["estado"] = "VACIO" if item["valor_a_escribir"] == "VACIO" else "LLENO"
            campos_con_estado.append(item)

        st.subheader("1. Control de llenado: Edita todas las celdas que necesites sin recargar la página")

        # st.form evita que cada cambio de valor en la tabla dispare llamadas a Google Sheets
        with st.form(key=f"form_edicion_{identificador_plantilla}"):
            campos_editados = st.data_editor(
                campos_con_estado,
                num_rows="dynamic",
                width="stretch"
            )
            aprobar_y_generar = st.form_submit_button(
                "2. Guardar en Google Sheets y Generar Archivo Diligenciado",
                type="primary"
            )

        if aprobar_y_generar:
            mapa_actual["campos"] = campos_editados
            guardar_mapa_sheets(CRED_GOOGLE, NOMBRE_SHEET, identificador_plantilla, mapa_actual, datos_maestra)

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