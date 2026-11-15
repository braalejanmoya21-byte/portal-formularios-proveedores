import json
import os
import pathlib
import openpyxl
import pandas as pd
import streamlit as st
from docx import Document
from google import genai
from google.genai import types

MAPA_AVALFIDUCIARIA_BASE = {
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

# Mapa exacto verificado celda por celda para oficial_GJU-F053 V5 VF 1.xlsx
MAPA_GJU_F053_BASE = {
    "campos": [
        {"clave": "tipo_solicitud", "etiqueta": "2-Tipo de solicitud", "hoja": "GJU-F053", "celda": "H24"},
        {"clave": "tipo_contraparte", "etiqueta": "3-Tipo de contraparte", "hoja": "GJU-F053", "celda": "L24"},
        {"clave": "tipo_persona", "etiqueta": "4-Tipo de persona", "hoja": "GJU-F053", "celda": "D33"},
        {"clave": "tipo_id", "etiqueta": "5-Tipo ID", "hoja": "GJU-F053", "celda": "H33"},
        {"clave": "nit", "etiqueta": "6-Número ID", "hoja": "GJU-F053", "celda": "M33"},
        {"clave": "razon_social", "etiqueta": "7-Nombre o razón social", "hoja": "GJU-F053", "celda": "E35"},
        {"clave": "tipo_organizacion", "etiqueta": "8-Tipo de Organización", "hoja": "GJU-F053", "celda": "L35"},
        {"clave": "direccion", "etiqueta": "9-Domicilio principal", "hoja": "GJU-F053", "celda": "D37"},
        {"clave": "ciudad", "etiqueta": "10-Ciudad/Municipio", "hoja": "GJU-F053", "celda": "K37"},
        {"clave": "departamento", "etiqueta": "11-Departamento", "hoja": "GJU-F053", "celda": "D39"},
        {"clave": "pais", "etiqueta": "12-País", "hoja": "GJU-F053", "celda": "H39"},
        {"clave": "telefono", "etiqueta": "13-Teléfono", "hoja": "GJU-F053", "celda": "M39"},
        {"clave": "ciiu", "etiqueta": "14-CIUU", "hoja": "GJU-F053", "celda": "E42"},
        {"clave": "contacto_nombre", "etiqueta": "29-Nombre persona de contacto", "hoja": "GJU-F053", "celda": "F54"},
        {"clave": "contacto_cargo", "etiqueta": "30-Cargo persona de contacto", "hoja": "GJU-F053", "celda": "L54"},
        {"clave": "activos_totales", "etiqueta": "40-Activos Totales", "hoja": "GJU-F053", "celda": "G71"},
        {"clave": "pasivos_totales", "etiqueta": "41-Pasivos Totales", "hoja": "GJU-F053", "celda": "L71"},
        {"clave": "patrimonio", "etiqueta": "42-Patrimonio (Total)", "hoja": "GJU-F053", "celda": "D73"},
        {"clave": "ingresos_totales", "etiqueta": "43-Ingreso (Total)", "hoja": "GJU-F053", "celda": "H73"},
        {"clave": "ganancia_perdida", "etiqueta": "44-Ganancia (Pérdida)", "hoja": "GJU-F053", "celda": "L73"},
        {"clave": "rep_legal_calidad", "etiqueta": "65-Calidad Rep Legal", "hoja": "GJU-F053", "celda": "B161"},
        {"clave": "rep_legal_nombre_completo", "etiqueta": "66-Nombres y apellidos Rep Legal", "hoja": "GJU-F053", "celda": "C161"},
        {"clave": "rep_legal_tipo_id", "etiqueta": "67-Tipo ID Rep Legal", "hoja": "GJU-F053", "celda": "G161"},
        {"clave": "rep_legal_id", "etiqueta": "68-Número ID Rep Legal", "hoja": "GJU-F053", "celda": "H161"},
        {"clave": "pais", "etiqueta": "69-País de expedición Rep Legal", "hoja": "GJU-F053", "celda": "J161"},
        {"clave": "direccion", "etiqueta": "70-Dirección Rep Legal", "hoja": "GJU-F053", "celda": "L161"},
        {"clave": "telefono", "etiqueta": "71-Teléfono Rep Legal", "hoja": "GJU-F053", "celda": "N161"}
    ]
}

def _autodetectar_respaldo_local(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    ext = ruta_archivo.suffix.lower()
    nombre = ruta_archivo.name.lower()

    if "avalfiduciaria" in nombre and ext == ".pdf":
        return MAPA_AVALFIDUCIARIA_BASE
    if "gju-f053" in nombre and ext == ".xlsx":
        return MAPA_GJU_F053_BASE

    if ext == ".xlsx":
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        if "GJU-F053" in wb.sheetnames:
            return MAPA_GJU_F053_BASE

        ws = wb.active
        # Mapa de rangos combinados para saltar siempre fuera de la etiqueta combinada
        merged_lookup = {}
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    merged_lookup[(r, c)] = rng

        sinonimos = {
            "razon_social": ["razón social", "razon social", "nombre de la empresa"],
            "nit": ["nit", "6-número id", "identificación tributaria"],
            "ciiu": ["ciiu", "ciuu", "actividad económica"],
            "direccion": ["domicilio principal", "dirección", "direccion"],
            "ciudad": ["ciudad", "municipio"],
            "telefono": ["teléfono", "telefono"],
            "correo_principal": ["correo", "email"],
            "activos_totales": ["activos totales", "total activos"],
            "pasivos_totales": ["pasivos totales", "total pasivos"],
            "patrimonio": ["patrimonio"],
            "ingresos_totales": ["ingreso (total)", "ingresos", "ventas anuales"]
        }

        campos_detectados = []
        claves_usadas = set()
        for fila in ws.iter_rows(max_row=180, max_col=18):
            for celda in fila:
                if celda.value is None:
                    continue
                texto = str(celda.value).strip()
                # Ignorar párrafos largos de instrucciones
                if len(texto) > 65:
                    continue
                texto_lower = texto.lower()
                for clave, palabras in sinonimos.items():
                    if clave in claves_usadas:
                        continue
                    if any(p in texto_lower for p in palabras):
                        rng = merged_lookup.get((celda.row, celda.column))
                        col_derecha = (rng.max_col + 1) if rng else (celda.column + 1)
                        fila_abajo = (rng.max_row + 1) if rng else (celda.row + 1)
                        celda_der = ws.cell(row=celda.row, column=col_derecha)
                        celda_abj = ws.cell(row=fila_abajo, column=celda.column)
                        destino = celda_der.coordinate if celda_der.value is None else celda_abj.coordinate
                        campos_detectados.append({
                            "clave": clave,
                            "etiqueta": texto,
                            "hoja": ws.title,
                            "celda": destino
                        })
                        claves_usadas.add(clave)
                        break
        return {"campos": campos_detectados}

    elif ext == ".csv":
        try:
            df = pd.read_csv(ruta_archivo, nrows=2, sep=None, engine="python")
            columnas = list(df.columns)
        except Exception:
            columnas = []
        if len(columnas) > 1:
            return {"campos": [{"clave": "razon_social", "etiqueta": col, "columna_csv": col} for col in columnas]}
        return {"campos": [{"clave": k, "etiqueta": k, "columna_csv": k} for k in claves_maestra]}

    return MAPA_AVALFIDUCIARIA_BASE

def generar_mapa_con_gemini(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    ext = ruta_archivo.suffix.lower()
    nombre = ruta_archivo.name.lower()

    if "avalfiduciaria" in nombre and ext == ".pdf":
        return MAPA_AVALFIDUCIARIA_BASE
    if "gju-f053" in nombre and ext == ".xlsx":
        return MAPA_GJU_F053_BASE

    return _autodetectar_respaldo_local(ruta_archivo, claves_maestra)