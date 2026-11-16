import datetime
import json
import os
import pathlib
import openpyxl
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from llenadores import llenar_excel, llenar_word, llenar_pdf, llenar_csv

load_dotenv()

CRED_GOOGLE = "credenciales_google.json"
NOMBRE_SHEET = "Base_Maestra_Proveedores"

st.set_page_config(page_title="Portal de Proveedores", layout="wide")
st.title("Portal de Diligenciamiento Automático de Formularios")

fuente_datos = st.radio(
    "Selecciona de dónde leer la Base Maestra:",
    ["Modo Prueba Rápida (Datos en memoria)", "Google Sheets en vivo (usa credenciales_google.json)"],
    horizontal=True
)

if fuente_datos == "Google Sheets en vivo (usa credenciales_google.json)":
    try:
        from google_connector import obtener_maestra, registrar_auditoria_sheets
        datos_maestra = obtener_maestra(CRED_GOOGLE, NOMBRE_SHEET)
        st.success("Conectado exitosamente a Google Sheets.")
    except Exception as e:
        st.error(f"Revisa que la hoja '{NOMBRE_SHEET}' exista en Drive y esté compartida con el correo del bot: {e}")
        st.stop()
else:
    datos_maestra = {
        "ciudad_diligenciamiento": "BOGOTA D.C.",
        "razon_social": "COLOMBIA TELECOMUNICACIONES S.A. E.S.P.",
        "tipo_id": "NIT",
        "nit": "830122566-1",
        "direccion": "TRANSVERSAL 60 NO. 114A-55",
        "ciudad": "BOGOTA D.C.",
        "correo_principal": "notificaciones@empresa.com",
        "ciiu": "6120",
        "activos_totales": 4086449000000,
        "pasivos_totales": 5517599000000,
        "patrimonio": -1431150000000,
        "ingresos_totales": 733953000000,
        "rep_legal_primer_nombre": "CARLOS",
        "rep_legal_primer_apellido": "RODRIGUEZ",
        "rep_legal_id": "79123456"
    }

with st.expander("Ver datos actuales de la Base Maestra"):
    st.json(datos_maestra)

# Función que inspecciona tu Excel o CSV real para armar el mapa con las hojas y celdas verdaderas
def autodetectar_estructura_local(ruta_archivo: pathlib.Path, claves_maestra: list) -> dict:
    ext = ruta_archivo.suffix.lower()
    sinonimos = {
        "razon_social": ["razón social", "razon social", "nombre de la empresa", "nombre o razón"],
        "nit": ["nit", "número de identificación", "numero de identificacion", "identificación tributaria"],
        "ciiu": ["ciiu", "actividad económica", "actividad economica"],
        "activos_totales": ["total activos", "activos totales", "activos"],
        "pasivos_totales": ["total pasivos", "pasivos totales", "pasivos"],
        "patrimonio": ["patrimonio"],
        "ingresos_totales": ["ingresos", "ventas anuales"],
        "direccion": ["dirección", "direccion"],
        "ciudad": ["ciudad", "municipio"],
        "correo_principal": ["correo", "email", "e-mail"],
        "rep_legal_primer_nombre": ["representante legal", "primer nombre"],
        "rep_legal_id": ["documento representante", "cédula", "cedula"]
    }

    if ext == ".xlsx":
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        hoja_principal = wb.active.title
        campos_detectados = []
        claves_usadas = set()

        for nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
            for fila in ws.iter_rows(max_row=150, max_col=20):
                for celda in fila:
                    if celda.value is None:
                        continue
                    texto_celda = str(celda.value).strip().lower()
                    for clave, palabras in sinonimos.items():
                        if clave in claves_usadas:
                            continue
                        if any(p in texto_celda for p in palabras):
                            # Buscar la celda vacía a la derecha o abajo para escribir el dato
                            celda_derecha = ws.cell(row=celda.row, column=celda.column + 1)
                            celda_abajo = ws.cell(row=celda.row + 1, column=celda.column)
                            destino = celda_derecha.coordinate if celda_derecha.value is None else celda_abajo.coordinate
                            campos_detectados.append({
                                "clave": clave,
                                "etiqueta": str(celda.value)[:40],
                                "hoja": nombre_hoja,
                                "celda": destino
                            })
                            claves_usadas.add(clave)
                            break

        if not campos_detectados:
            campos_detectados = [
                {"clave": "razon_social", "etiqueta": "Razón Social", "hoja": hoja_principal, "celda": "C5"},
                {"clave": "nit", "etiqueta": "NIT", "hoja": hoja_principal, "celda": "C6"},
                {"clave": "ciiu", "etiqueta": "Código CIIU", "hoja": hoja_principal, "celda": "C7"},
                {"clave": "activos_totales", "etiqueta": "Activos", "hoja": hoja_principal, "celda": "C8"},
                {"clave": "pasivos_totales", "etiqueta": "Pasivos", "hoja": hoja_principal, "celda": "C9"},
                {"clave": "patrimonio", "etiqueta": "Patrimonio", "hoja": hoja_principal, "celda": "C10"},
                {"clave": "ingresos_totales", "etiqueta": "Ingresos", "hoja": hoja_principal, "celda": "C11"}
            ]
        return {"campos": campos_detectados}

    elif ext == ".csv":
        try:
            df = pd.read_csv(ruta_archivo, nrows=2, sep=None, engine="python")
            columnas = list(df.columns)
        except Exception:
            columnas = []
        campos_csv = []
        if len(columnas) > 1:
            for col in columnas:
                col_lower = col.strip().lower()
                clave_sugerida = "razon_social"
                for k in claves_maestra:
                    if k.lower() in col_lower or col_lower in k.lower():
                        clave_sugerida = k
                        break
                campos_csv.append({"clave": clave_sugerida, "etiqueta": col, "columna_csv": col})
        else:
            for k in claves_maestra:
                campos_csv.append({"clave": k, "etiqueta": k, "columna_csv": k})
        return {"campos": campos_csv}

    return {"campos": [{"clave": "razon_social", "etiqueta": "Razón Social", "ubicacion": "tabla", "tabla_idx": 0, "fila_idx": 0, "col_idx": 1}]}

col1, col2 = st.columns(2)
with col1:
    usuario = st.text_input("Tu correo corporativo", value="analista@empresa.com")
with col2:
    proveedor = st.text_input("Nombre del Proveedor", value="Proveedor_Prueba")

archivo = st.file_uploader(
    "Arrastra aquí el formulario (.pdf, .xlsx, .docx, .csv)",
    type=["pdf", "xlsx", "docx", "csv"]
)

MAPA_AVALFIDUCIARIA = {
    "campos": [
        {"clave": "ciudad_diligenciamiento", "etiqueta": "*Ciudad", "pagina": 0, "tipo": "texto", "bbox": [55, 160, 250, 178]},
        {"clave": "razon_social", "etiqueta": "*Nombre de la Empresa / Razón Social", "pagina": 0, "tipo": "texto", "bbox": [45, 480, 520, 498]},
        {"clave": "tipo_id", "etiqueta": "Checkbox NIT", "pagina": 0, "tipo": "checkbox", "opcion": "NIT", "bbox": [550, 486, 565, 502]},
        {"clave": "nit", "etiqueta": "*Número de Identificación", "pagina": 0, "tipo": "texto", "bbox": [695, 480, 940, 498]},
        {"clave": "direccion", "etiqueta": "*Dirección de la Empresa", "pagina": 0, "tipo": "texto", "bbox": [200, 718, 480, 736]},
        {"clave": "ciudad", "etiqueta": "*Ciudad/Dpto", "pagina": 0, "tipo": "texto", "bbox": [500, 718, 730, 736]},
        {"clave": "correo_principal", "etiqueta": "*Correo Electrónico Principal", "pagina": 0, "tipo": "texto", "bbox": [45, 792, 480, 810]},
        {"clave": "ciiu", "etiqueta": "*Código CIIU", "pagina": 0, "tipo": "texto", "bbox": [740, 792, 940, 810]},
        {"clave": "rep_legal_primer_nombre", "etiqueta": "*Primer Nombre Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [45, 690, 260, 708]},
        {"clave": "rep_legal_primer_apellido", "etiqueta": "*Primer Apellido Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [505, 690, 720, 708]},
        {"clave": "rep_legal_id", "etiqueta": "*Número de Identificación Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [365, 740, 650, 758]}
    ]
}

if archivo and usuario and proveedor:
    ext = pathlib.Path(archivo.name).suffix.lower()
    ruta_temp = pathlib.Path("plantillas") / archivo.name
    ruta_temp.parent.mkdir(exist_ok=True)
    ruta_temp.write_bytes(archivo.getvalue())

    # Mostrar las hojas reales cuando se sube un Excel para facilitar la verificación
    if ext == ".xlsx":
        wb_info = openpyxl.load_workbook(ruta_temp, read_only=True)
        st.info(f"Pestañas encontradas dentro de este Excel: {wb_info.sheetnames}")
        wb_info.close()

    ruta_mapa = pathlib.Path("mapas") / f"{ruta_temp.stem}_{ext.replace('.', '')}.json"
    ruta_mapa.parent.mkdir(exist_ok=True)

    if st.button("🔄 Volver a escanear archivo (Reiniciar mapa)"):
        if ruta_mapa.exists():
            ruta_mapa.unlink()

    if not ruta_mapa.exists():
        if "avalfiduciaria" in archivo.name.lower() and ext == ".pdf":
            ruta_mapa.write_text(json.dumps(MAPA_AVALFIDUCIARIA, ensure_ascii=False, indent=2), encoding="utf-8")
        elif os.environ.get("GEMINI_API_KEY"):
            from mapeo_gemini import generar_mapa_con_gemini
            with st.spinner("Analizando formulario con Gemini..."):
                mapa_propuesto = generar_mapa_con_gemini(ruta_temp, list(datos_maestra.keys()))
                ruta_mapa.write_text(json.dumps(mapa_propuesto, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            mapa_auto = autodetectar_estructura_local(ruta_temp, list(datos_maestra.keys()))
            ruta_mapa.write_text(json.dumps(mapa_auto, ensure_ascii=False, indent=2), encoding="utf-8")

    mapa_actual = json.loads(ruta_mapa.read_text(encoding="utf-8"))

    st.subheader("1. Verifica o edita el mapa de campos")
    campos_editados = st.data_editor(mapa_actual.get("campos", []), num_rows="dynamic", width="stretch")

    if st.button("2. Generar Formulario Diligenciado", type="primary"):
        mapa_actual["campos"] = campos_editados
        ruta_mapa.write_text(json.dumps(mapa_actual, ensure_ascii=False, indent=2), encoding="utf-8")

        carpeta_salidas = pathlib.Path("salidas")
        carpeta_salidas.mkdir(exist_ok=True)
        marca = datetime.datetime.now().strftime("%H%M%S")
        ruta_salida = carpeta_salidas / f"{proveedor}_{marca}_{archivo.name}"

        if ext == ".xlsx":
            llenar_excel(ruta_temp, ruta_salida, mapa_actual, datos_maestra)
        elif ext == ".docx":
            llenar_word(ruta_temp, ruta_salida, mapa_actual, datos_maestra)
        elif ext == ".pdf":
            llenar_pdf(ruta_temp, ruta_salida, mapa_actual, datos_maestra)
        elif ext == ".csv":
            llenar_csv(ruta_temp, ruta_salida, mapa_actual, datos_maestra)

        if fuente_datos == "Google Sheets en vivo (usa credenciales_google.json)":
            consecutivo = registrar_auditoria_sheets(CRED_GOOGLE, NOMBRE_SHEET, usuario, proveedor, ext, ruta_salida)
        else:
            consecutivo = 1

        nombre_descarga = f"Consecutivo_{consecutivo:04d}_{proveedor}_{archivo.name}"
        st.success(f"¡Listo! Consecutivo #{consecutivo} generado sin errores.")
        st.download_button(
            label="3. Descargar Archivo Diligenciado",
            data=ruta_salida.read_bytes(),
            file_name=nombre_descarga
        )