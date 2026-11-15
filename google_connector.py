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

def _limpiar_numero(valor) -> float:
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).replace("$", "").replace(" ", "").strip()
    if not texto:
        return 0.0
    if texto.count(".") > 1:
        texto = texto.replace(".", "")
    texto = texto.replace(",", ".")
    return float(texto)

def conectar_sheet(ruta_credenciales: str, nombre_sheet: str):
    if "gcp_service_account" in st.secrets:
        info_creds = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info_creds, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(ruta_credenciales, scopes=SCOPES)
    cliente = gspread.authorize(creds)
    return cliente.open(nombre_sheet)

def obtener_maestra(ruta_credenciales: str, nombre_sheet: str) -> dict:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    ws = sh.worksheet("Maestra")
    registros = ws.get_all_records()
    datos = {str(fila["clave"]).strip(): fila["valor"] for fila in registros if fila.get("clave")}

    activos = _limpiar_numero(datos.get("activos_totales", 0))
    pasivos = _limpiar_numero(datos.get("pasivos_totales", 0))
    patrimonio = _limpiar_numero(datos.get("patrimonio", 0))

    if abs((activos - pasivos) - patrimonio) > 1:
        raise ValueError(
            f"Inconsistencia contable en Google Sheets: Activos ({activos:,.0f}) - "
            f"Pasivos ({pasivos:,.0f}) no coincide con Patrimonio ({patrimonio:,.0f})."
        )
    return datos

def obtener_mapa_sheets(ruta_credenciales: str, nombre_sheet: str, nombre_plantilla: str) -> dict | None:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Mapas")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Mapas", rows=200, cols=3)
        ws.append_row(["plantilla", "json_mapa", "fecha_actualizacion"])
        return None

    registros = ws.get_all_records()
    for fila in registros:
        if str(fila.get("plantilla", "")).strip().lower() == nombre_plantilla.strip().lower():
            return json.loads(fila["json_mapa"])
    return None

def guardar_mapa_sheets(ruta_credenciales: str, nombre_sheet: str, nombre_plantilla: str, mapa: dict) -> None:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    try:
        ws = sh.worksheet("Mapas")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Mapas", rows=200, cols=3)
        ws.append_row(["plantilla", "json_mapa", "fecha_actualizacion"])

    celdas = ws.col_values(1)
    fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    json_texto = json.dumps(mapa, ensure_ascii=False)

    for idx, valor_celda in enumerate(celdas, start=1):
        if valor_celda.strip().lower() == nombre_plantilla.strip().lower():
            ws.update_cell(idx, 2, json_texto)
            ws.update_cell(idx, 3, fecha)
            return

    ws.append_row([nombre_plantilla, json_texto, fecha])

def registrar_auditoria_sheets(
    ruta_credenciales: str,
    nombre_sheet: str,
    usuario: str,
    proveedor: str,
    formato: str,
    ruta_salida: pathlib.Path
) -> int:
    sh = conectar_sheet(ruta_credenciales, nombre_sheet)
    ws = sh.worksheet("Auditoria")
    filas_actuales = len(ws.get_all_values())
    consecutivo = max(1, filas_actuales)

    sha256 = hashlib.sha256(ruta_salida.read_bytes()).hexdigest()
    fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ws.append_row([consecutivo, fecha, usuario, proveedor, formato, ruta_salida.name, sha256])
    return consecutivo