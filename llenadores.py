import io
import json
import openpyxl
from openpyxl.cell.cell import MergedCell
import pandas as pd
from docx import Document
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

def _obtener_valor_campo(campo: dict, datos: dict):
    """Prioriza el valor editado directamente en la columna 'valor_a_escribir' de la tabla."""
    val_tabla = campo.get("valor_a_escribir")
    if val_tabla is not None and str(val_tabla).strip() not in ("", "VACIO", "None"):
        return val_tabla
    val_maestra = datos.get(str(campo.get("clave", "")).strip())
    if val_maestra is not None and str(val_maestra).strip() not in ("", "VACIO", "None"):
        return val_maestra
    return None

def llenar_excel(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    wb = openpyxl.load_workbook(ruta_entrada)
    for campo in mapa.get("campos", []):
        valor = _obtener_valor_campo(campo, datos)
        celda_coord = str(campo.get("celda", "")).strip()
        if valor is None or not celda_coord:
            continue

        nombre_hoja = str(campo.get("hoja", "")).strip()
        if nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
        else:
            hojas_minus = {h.lower(): h for h in wb.sheetnames}
            ws = wb[hojas_minus[nombre_hoja.lower()]] if nombre_hoja.lower() in hojas_minus else wb.active

        try:
            celda = ws[celda_coord]
        except Exception:
            continue

        if isinstance(celda, MergedCell):
            for rango in ws.merged_cells.ranges:
                if celda.coordinate in rango:
                    celda = ws.cell(row=rango.min_row, column=rango.min_col)
                    break

        if isinstance(celda.value, str) and celda.value.startswith("="):
            continue
        celda.value = valor
    wb.save(ruta_salida)

def llenar_word(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    doc = Document(ruta_entrada)
    for campo in mapa.get("campos", []):
        valor = _obtener_valor_campo(campo, datos)
        if valor is None:
            continue
        try:
            if campo.get("ubicacion") == "tabla":
                t_idx = int(campo.get("tabla_idx", 0))
                f_idx = int(campo.get("fila_idx", 0))
                c_idx = int(campo.get("col_idx", 0))
                doc.tables[t_idx].rows[f_idx].cells[c_idx].text = str(valor)
            elif campo.get("ubicacion") == "parrafo":
                p_idx = int(campo.get("parrafo_idx", 0))
                doc.paragraphs[p_idx].add_run(f" {valor}")
        except (IndexError, ValueError):
            continue
    doc.save(ruta_salida)

def llenar_pdf(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    escritor = PdfWriter(clone_from=ruta_entrada)
    campos = mapa.get("campos", [])

    for num_pag, pagina in enumerate(escritor.pages):
        campos_pagina = [f for f in campos if int(f.get("pagina", 0) or 0) == num_pag]
        if not campos_pagina:
            continue

        W, H = float(pagina.mediabox.width), float(pagina.mediabox.height)
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(W, H))
        dibujo_algo = False

        for campo in campos_pagina:
            valor = _obtener_valor_campo(campo, datos)
            bbox_raw = campo.get("bbox")
            if valor is None or not bbox_raw:
                continue

            if isinstance(bbox_raw, str):
                try:
                    bbox = json.loads(bbox_raw)
                except Exception:
                    continue
            else:
                bbox = list(bbox_raw)

            if len(bbox) != 4:
                continue
            x0, y0, x1, y1 = [float(v) for v in bbox]

            if str(campo.get("tipo", "texto")).lower() == "checkbox":
                opcion_esperada = str(campo.get("opcion") or "").strip().lower()
                val_str = str(valor).strip().lower()
                # Marca con X si el usuario escribe X, Si, True o el nombre de la opción (ej. NIT)
                if val_str in ("x", "si", "sí", "true", "1", opcion_esperada):
                    c.setFont("Helvetica-Bold", 9)
                    c.drawCentredString(((x0 + x1) / 2) / 1000 * W, H - (((y0 + y1) / 2) / 1000 * H) - 3, "X")
                    dibujo_algo = True
            else:
                c.setFont("Helvetica", 8)
                c.drawString((x0 / 1000) * W + 2, H - ((y1 / 1000) * H) + 3, str(valor))
                dibujo_algo = True

        if dibujo_algo:
            c.showPage()
            c.save()
            packet.seek(0)
            capa_nueva = PdfReader(packet)
            pagina.merge_page(capa_nueva.pages[0])

    with open(ruta_salida, "wb") as f:
        escritor.write(f)

def llenar_csv(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    try:
        df = pd.read_csv(ruta_entrada, sep=None, engine="python")
    except Exception:
        df = pd.DataFrame()

    if len(df.columns) > 1:
        nueva_fila = {col: "" for col in df.columns}
        for campo in mapa.get("campos", []):
            col_destino = str(campo.get("columna_csv") or campo.get("celda") or "").strip()
            valor = _obtener_valor_campo(campo, datos)
            if valor is None:
                valor = ""
            if col_destino in nueva_fila:
                nueva_fila[col_destino] = valor
            else:
                for col_real in df.columns:
                    if col_real.strip().lower() == col_destino.lower():
                        nueva_fila[col_real] = valor
        df_salida = pd.DataFrame([nueva_fila])
    else:
        filas = []
        for campo in mapa.get("campos", []):
            clave = campo.get("clave", "")
            etiqueta = campo.get("columna_csv") or campo.get("etiqueta") or clave
            valor = _obtener_valor_campo(campo, datos)
            filas.append({"Campo": etiqueta, "Valor": valor if valor is not None else ""})
        df_salida = pd.DataFrame(filas)

    df_salida.to_csv(ruta_salida, index=False, sep=";", encoding="utf-8-sig")