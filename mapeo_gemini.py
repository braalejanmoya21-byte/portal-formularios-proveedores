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

def extraer_estructura(ruta: pathlib.Path) -> str:
    ext = ruta.suffix.lower()
    if ext == ".xlsx":
        wb = openpyxl.load_workbook(ruta, data_only=False)
        lineas = []
        for hoja in wb.sheetnames:
            ws = wb[hoja]
            for fila in ws.iter_rows(max_row=150, max_col=20, values_only=False):
                for celda in fila:
                    if celda.value is not None:
                        lineas.append(f"Hoja '{hoja}' | Celda {celda.coordinate}: {celda.value}")
        return "\n".join(lineas[:800])
    elif ext == ".docx":
        doc = Document(ruta)
        textos = [f"Párrafo {i}: {p.text}" for i, p in enumerate(doc.paragraphs) if p.text.strip()]
        for t_idx, tabla in enumerate(doc.tables):
            for r_idx, fila in enumerate(tabla.rows):
                celdas = [f"Col {c_idx}: {c.text.strip()}" for c_idx, c in enumerate(fila.cells)]
                textos.append(f"Tabla {t_idx} Fila {r_idx} -> " + " | ".join(celdas))
        return "\n".join(textos)
    elif ext == ".csv":
        df = pd.read_csv(ruta, nrows=5, sep=None, engine="python")
        return f"Columnas del archivo CSV: {list(df.columns)}"
    return ""

def _autodetectar_respaldo_local(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    ext = ruta_archivo.suffix.lower()
    if "avalfiduciaria" in ruta_archivo.name.lower() and ext == ".pdf":
        return MAPA_AVALFIDUCIARIA_BASE

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
        if len(columnas) > 1:
            return {"campos": [{"clave": "razon_social", "etiqueta": col, "columna_csv": col} for col in columnas]}
        return {"campos": [{"clave": k, "etiqueta": k, "columna_csv": k} for k in claves_maestra]}

    elif ext == ".docx":
        return {"campos": [{"clave": "razon_social", "etiqueta": "Razón Social", "ubicacion": "tabla", "tabla_idx": 0, "fila_idx": 0, "col_idx": 1}]}

    return MAPA_AVALFIDUCIARIA_BASE

def generar_mapa_con_gemini(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    ext = ruta_archivo.suffix.lower()

    # Si es el PDF de Avalfiduciaria, usamos directamente el mapa calibrado sin gastar cuota de IA
    if "avalfiduciaria" in ruta_archivo.name.lower() and ext == ".pdf":
        return MAPA_AVALFIDUCIARIA_BASE

    api_key = os.environ.get("GEMINI_API_KEY") or (st.secrets.get("GEMINI_API_KEY") if hasattr(st, "secrets") else None)
    if not api_key or "TuClaveDeGeminiAqui" in str(api_key):
        st.warning("Usando escáner automático local (configura una GEMINI_API_KEY válida en Secrets para usar IA).")
        return _autodetectar_respaldo_local(ruta_archivo, claves_maestra)

    prompt = f"""
    Eres un experto en automatización de formularios de proveedores.
    Tienes estas claves disponibles en la Base Maestra: {json.dumps(claves_maestra, ensure_ascii=False)}
    
    Analiza el formulario adjunto ({ext}) y devuelve un JSON con la clave "campos" (lista de objetos).
    Según el formato, cada objeto debe tener:
    - "clave": nombre exacto de la clave de la Base Maestra.
    - "etiqueta": nombre del campo en el formulario.
    - Si es .xlsx: incluye "hoja" y "celda" (la celda vacía donde debe escribirse la respuesta, NUNCA la celda de la etiqueta ni una celda con fórmula).
    - Si es .docx: incluye "ubicacion" ("tabla" o "parrafo"), "tabla_idx", "fila_idx", "col_idx" (la celda vacía donde va el valor).
    - Si es .pdf: incluye "pagina" (empezando en 0), "tipo" ("texto" o "checkbox"), "opcion" (si es checkbox) y "bbox" [x0, y0, x1, y1] en escala de 0 a 1000 indicando el espacio en blanco donde se debe escribir.
    - Si es .csv: incluye "columna_csv" con el nombre exacto de la columna del CSV.
    """

    try:
        client = genai.Client(api_key=api_key)
        if ext == ".pdf":
            contenido = [
                types.Part.from_bytes(data=ruta_archivo.read_bytes(), mime_type="application/pdf"),
                prompt
            ]
        else:
            estructura = extraer_estructura(ruta_archivo)
            contenido = [f"Estructura del archivo:\n{estructura}", prompt]

        respuesta = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=contenido,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.0
            )
        )
        return json.loads(respuesta.text)
    except Exception as e:
        st.warning(f"Se activó el detector de estructura local automático (Detalle API: {type(e).__name__}).")
        return _autodetectar_respaldo_local(ruta_archivo, claves_maestra)