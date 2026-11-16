import datetime
import hashlib
import json
import pathlib
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

DATOS_COMPLETOS_EMPRESA = [
    ["clave", "valor", "area_responsable", "actualizado_por", "fecha_actualizacion"],
    ["tipo_solicitud", "Vinculación", "Comercial", "sistema@empresa.com", "2026-11-15"],
    ["tipo_contraparte", "Proveedor", "Comercial", "sistema@empresa.com", "2026-11-15"],
    ["tipo_persona", "Jurídica", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["tipo_id", "NIT", "Fiscal", "sistema@empresa.com", "2026-11-15"],
    ["nit", "830122566-1", "Fiscal", "sistema@empresa.com", "2026-11-15"],
    ["razon_social", "COLOMBIA TELECOMUNICACIONES S.A. E.S.P.", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["tipo_organizacion", "Sociedad Anónima", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["direccion", "TRANSVERSAL 60 NO. 114A-55", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["ciudad", "BOGOTA D.C.", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["ciudad_diligenciamiento", "BOGOTA D.C.", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["fecha_diligenciamiento_dd", "15", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["fecha_diligenciamiento_mm", "11", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["fecha_diligenciamiento_aaaa", "2026", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["oficina", "PRINCIPAL BOGOTA", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["departamento", "BOGOTA D.C.", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["pais", "COLOMBIA", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["telefono", "6017050000", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["celular", "3150000000", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["correo_principal", "notificaciones@empresa.com", "Administrativa", "sistema@empresa.com", "2026-11-15"],
    ["ciiu", "6120", "Fiscal", "sistema@empresa.com", "2026-11-15"],
    ["origen_fondos", "Servicio telecomunicaciones inalambricas", "Contabilidad", "sistema@empresa.com", "2026-11-15"],
    ["contacto_nombre", "ANDREA COORDINADORA", "Comercial", "sistema@empresa.com", "2026-11-15"],
    ["contacto_cargo", "Analista de Proveedores", "Comercial", "sistema@empresa.com", "2026-11-15"],
    ["activos_totales", 4086449000000, "Contabilidad", "sistema@empresa.com", "2026-11-15"],
    ["pasivos_totales", 5517599000000, "Contabilidad", "sistema@empresa.com", "2026-11-15"],
    ["patrimonio", -1431150000000, "Contabilidad", "sistema@empresa.com", "2026-11-15"],
    ["ingresos_totales", 733953000000, "Contabilidad", "sistema@empresa.com", "2026-11-15"],
    ["ganancia_perdida", 125734000000, "Contabilidad", "sistema@empresa.com", "2026-11-15"],
    ["rep_legal_calidad", "Principal", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["rep_legal_nombre_completo", "CARLOS RODRIGUEZ", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["rep_legal_primer_nombre", "CARLOS", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["rep_legal_primer_apellido", "RODRIGUEZ", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["rep_legal_tipo_id", "CC", "Jurídica", "sistema@empresa.com", "2026-11-15"],
    ["rep_legal_id", "79123456", "Jurídica", "sistema@empresa.com", "2026-11-15"]
]

@st.cache_resource(show_spinner=False)
def _obtener_cliente(ruta_credenciales: str):
    if "gcp_service_account" in st.secrets:
        info_creds = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info_creds, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(ruta_credenciales, scopes=SCOPES)
    return gspread.authorize(creds)

def conectar_sheet(ruta_credenciales: str, nombre_sheet: str):
    cliente = _obtener_cliente(ruta_credenciales)
    return cliente.open(nombre_sheet)

def sincronizar_respuestas_forms_a_maestra(ruta_credenciales: str, nombre_sheet: str) -> int:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    ws_maestra = sh.worksheet("Maestra")
    filas_maestra = ws_maestra.get_all_values()
    mapa_filas = {str(f[0]).strip().lower(): idx + 1 for idx, f in enumerate(filas_maestra) if f and f[0]}
    lote_actualizaciones = []

    for ws in sh.worksheets():
        titulo_low = ws.title.lower()
        if "respuestas de formulario" in titulo_low or "form responses" in titulo_low:
            datos_form = ws.get_all_values()
            if len(datos_form) < 2:
                continue
            encabezados = datos_form[0]
            ultima_respuesta = datos_form[-1]
            fecha_resp = ultima_respuesta[0] if ultima_respuesta else datetime.datetime.now().strftime("%Y-%m-%d")
            correo_resp = ws.title

            for col_idx, pregunta in enumerate(encabezados):
                if col_idx == 0:
                    continue
                if "correo" in pregunta.lower() and col_idx < len(ultima_respuesta) and "@" in ultima_respuesta[col_idx]:
                    correo_resp = ultima_respuesta[col_idx]
                    continue

                clave_limpia = pregunta.strip().lower().replace(" ", "_")
                valor_resp = ultima_respuesta[col_idx].strip() if col_idx < len(ultima_respuesta) else ""
                valor_final = valor_resp if valor_resp != "" else "VACIO"

                if clave_limpia in mapa_filas:
                    fila_num = mapa_filas[clave_limpia]
                    lote_actualizaciones.append({
                        "range": f"B{fila_num}:E{fila_num}",
                        "values": [[valor_final, "Google Forms", correo_resp, fecha_resp]]
                    })

    if lote_actualizaciones:
        ws_maestra.batch_update(lote_actualizaciones)
    st.cache_data.clear()
    return len(lote_actualizaciones)

def cargar_datos_completos_maestra(ruta_credenciales: str, nombre_sheet: str) -> None:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Maestra")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Maestra", rows=100, cols=10)
    ws.clear()
    ws.update("A1", DATOS_COMPLETOS_EMPRESA)
    st.cache_data.clear()

@st.cache_data(ttl=300, show_spinner=False)
def obtener_maestra(ruta_credenciales: str, nombre_sheet: str) -> dict:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Maestra")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Maestra", rows=100, cols=10)
        ws.update("A1", DATOS_COMPLETOS_EMPRESA)

    filas = ws.get_all_values()
    if len(filas) <= 1:
        ws.update("A1", DATOS_COMPLETOS_EMPRESA)
        filas = ws.get_all_values()

    datos = {}
    actualizaciones_vacias = []
    ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    for idx, fila in enumerate(filas[1:], start=2):
        clave = str(fila[0]).strip() if len(fila) > 0 else ""
        if not clave:
            continue
        valor = str(fila[1]).strip() if len(fila) > 1 else ""
        area = str(fila[2]).strip() if len(fila) > 2 else ""
        autor = str(fila[3]).strip() if len(fila) > 3 else ""
        fecha = str(fila[4]).strip() if len(fila) > 4 else ""

        if valor == "" or area == "" or autor == "" or fecha == "":
            valor_nuevo = valor if valor != "" else "VACIO"
            area_nueva = area if area != "" else "VACIO"
            autor_nuevo = autor if autor != "" else "VACIO"
            fecha_nueva = fecha if fecha != "" else ahora
            actualizaciones_vacias.append({
                "range": f"B{idx}:E{idx}",
                "values": [[valor_nuevo, area_nueva, autor_nuevo, fecha_nueva]]
            })
            valor = valor_nuevo

        datos[clave] = valor

    if actualizaciones_vacias:
        ws.batch_update(actualizaciones_vacias)

    return datos

@st.cache_data(ttl=300, show_spinner=False)
def obtener_mapa_sheets(ruta_credenciales: str, nombre_sheet: str, nombre_plantilla: str) -> dict | None:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Mapas")
    except gspread.WorksheetNotFound:
        return None

    filas = ws.get_all_values()
    if len(filas) <= 1:
        return None

    encabezados = [h.strip().lower() for h in filas[0]]
    if "clave" in encabezados and "celda_o_coordenada" in encabezados:
        campos = []
        for fila in filas[1:]:
            if len(fila) >= 5 and fila[0].strip().lower() == nombre_plantilla.strip().lower():
                item = {
                    "clave": fila[1],
                    "etiqueta": fila[2],
                }
                ubicacion_hoja = fila[3]
                coordenada = fila[4]
                if coordenada.startswith("["):
                    item["pagina"] = int(ubicacion_hoja) if ubicacion_hoja.isdigit() else 0
                    item["bbox"] = json.loads(coordenada)
                    item["tipo"] = fila[5] if len(fila) > 5 and fila[5] in ("texto", "checkbox") else "texto"
                    item["opcion"] = fila[6] if len(fila) > 6 and fila[6] != "VACIO" else None
                else:
                    item["hoja"] = ubicacion_hoja
                    item["celda"] = coordenada
                    item["columna_csv"] = coordenada
                campos.append(item)
        return {"campos": campos} if campos else None
    return None

def guardar_mapa_sheets(
    ruta_credenciales: str,
    nombre_sheet: str,
    nombre_plantilla: str,
    mapa: dict,
    datos_maestra: dict
) -> None:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Mapas")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Mapas", rows=500, cols=8)

    encabezado_esperado = [
        "plantilla", "clave", "etiqueta", "hoja_o_pagina",
        "celda_o_coordenada", "tipo", "estado_en_maestra", "valor_actual"
    ]
    filas_existentes = ws.get_all_values()

    filas_conservadas = [encabezado_esperado]
    if len(filas_existentes) > 1 and filas_existentes[0] == encabezado_esperado:
        for f in filas_existentes[1:]:
            if f and f[0].strip().lower() != nombre_plantilla.strip().lower():
                filas_conservadas.append(f)

    for c in mapa.get("campos", []):
        clave = str(c.get("clave", "VACIO")).strip()
        etiqueta = str(c.get("etiqueta", "VACIO")).strip()
        hoja_pag = str(c.get("hoja") if c.get("hoja") is not None else c.get("pagina", "0"))
        coord = str(c.get("celda") or c.get("columna_csv") or json.dumps(c.get("bbox", [])))
        tipo = str(c.get("tipo", "celda"))
        val_tabla = str(c.get("valor_a_escribir", "")).strip()
        val_maestra = val_tabla if val_tabla not in ("", "VACIO", "None") else str(datos_maestra.get(clave, "VACIO")).strip()
        estado = "VACIO (Falta en Maestra)" if val_maestra in ("", "VACIO", "None") else "LLENO"

        filas_conservadas.append([
            nombre_plantilla, clave, etiqueta, hoja_pag, coord, tipo, estado, val_maestra
        ])

    ws.clear()
    ws.update("A1", filas_conservadas)
    st.cache_data.clear()

def registrar_auditoria_sheets(
    ruta_credenciales: str,
    nombre_sheet: str,
    usuario: str,
    proveedor: str,
    formato: str,
    ruta_salida: pathlib.Path,
    campos_llenados: int = 0,
    lista_vacios: list[str] | None = None
) -> int:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Auditoria")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Auditoria", rows=500, cols=10)

    encabezados = [
        "consecutivo", "fecha", "usuario", "proveedor", "formato",
        "nombre_archivo", "campos_llenados", "campos_vacios_faltantes", "hash_sha256"
    ]
    filas_actuales = ws.get_all_values()
    if not filas_actuales:
        ws.append_row(encabezados)
        filas_actuales = [encabezados]
    elif len(filas_actuales[0]) < 9:
        ws.update("A1:I1", [encabezados])

    consecutivo = max(1, len(filas_actuales))
    sha256 = hashlib.sha256(ruta_salida.read_bytes()).hexdigest()
    fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    texto_vacios = ", ".join(lista_vacios) if lista_vacios else "NINGUNO (100% LLENO)"

    ws.append_row([
        consecutivo,
        fecha,
        usuario or "VACIO",
        proveedor or "VACIO",
        formato or "VACIO",
        ruta_salida.name,
        campos_llenados,
        texto_vacios,
        sha256
    ])
    return consecutivo

def actualizar_maestra_desde_tabla(ruta_credenciales: str, nombre_sheet: str, nuevos_valores: dict, usuario: str) -> None:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    ws = sh.worksheet("Maestra")
    filas = ws.get_all_values()
    mapa_filas = {str(f[0]).strip(): idx + 1 for idx, f in enumerate(filas) if f and f[0]}
    ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    lote = []
    filas_nuevas = []
    for clave, valor in nuevos_valores.items():
        clave_str = str(clave).strip()
        val_str = str(valor).strip()
        if not clave_str:
            continue
        val_final = val_str if val_str != "" else "VACIO"
        if clave_str in mapa_filas:
            fila_num = mapa_filas[clave_str]
            lote.append({
                "range": f"B{fila_num}:E{fila_num}",
                "values": [[val_final, "Portal Web", usuario or "web", ahora]]
            })
        else:
            filas_nuevas.append([clave_str, val_final, "Portal Web", usuario or "web", ahora])

    if lote:
        ws.batch_update(lote)
    if filas_nuevas:
        ws.append_rows(filas_nuevas)
    st.cache_data.clear()