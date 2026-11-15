import json
import os
import pathlib
import openpyxl
import pandas as pd
from docx import Document
from google import genai
from google.genai import types

def extraer_estructura(ruta: pathlib.Path) -> str:
    ext = ruta.suffix.lower()
    if ext == ".xlsx":
        wb = openpyxl.load_workbook(ruta, data_only=False)
        lineas = []
        for hoja in wb.sheetnames:
            ws = wb[hoja]
            for fila in ws.iter_rows(values_only=False):
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

def generar_mapa_con_gemini(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
    ext = ruta_archivo.suffix.lower()
    
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
    
    if ext == ".pdf":
        contenido = [
            types.Part.from_bytes(data=ruta_archivo.read_bytes(), mime_type="application/pdf"),
            prompt
        ]
    else:
        estructura = extraer_estructura(ruta_archivo)
        contenido = [f"Estructura del archivo:\n{estructura}", prompt]
        
    respuesta = client.models.generate_content(
        model="gemini-2.5-pro",
        contents=contenido,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.0
        )
    )
    return json.loads(respuesta.text)