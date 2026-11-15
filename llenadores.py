import io
import openpyxl
from openpyxl.cell.cell import MergedCell
import pandas as pd
from docx import Document
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

def llenar_excel(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    wb = openpyxl.load_workbook(ruta_entrada)
    for campo in mapa.get("campos", []):
        clave = campo.get("clave")
        valor = datos.get(clave)
        celda_coord = str(campo.get("celda", "")).strip()
        if valor is None or not celda_coord:
            continue
            
        # Buscar la hoja indicada; si no existe, usar la hoja principal activa del Excel
        nombre_hoja = str(campo.get("hoja", "")).strip()
        if nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
        else:
            hojas_minus = {h.lower(): h for h in wb.sheetnames}
            if nombre_hoja.lower() in hojas_minus:
                ws = wb[hojas_minus[nombre_hoja.lower()]]
            else:
                ws = wb.active

        try:
            celda = ws[celda_coord]
        except Exception:
            continue

        # Si la celda es parte de una combinación, escribir en la esquina superior izquierda
        if isinstance(celda, MergedCell):
            for rango in ws.merged_cells.ranges:
                if celda.coordinate in rango:
                    celda = ws.cell(row=rango.min_row, column=rango.min_col)
                    break

        # Respetar celdas que ya contienen fórmulas como =TODAY()
        if isinstance(celda.value, str) and celda.value.startswith("="):
            continue
        celda.value = valor
    wb.save(ruta_salida)

def llenar_word(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    doc = Document(ruta_entrada)
    for campo in mapa.get("campos", []):
        valor = datos.get(campo.get("clave"))
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
        except IndexError:
            continue
    doc.save(ruta_salida)

def llenar_pdf(ruta_entrada, ruta_salida, mapa: dict, datos: dict):
    escritor = PdfWriter(clone_from=ruta_entrada)
    campos = mapa.get("campos", [])
    
    for num_pag, pagina in enumerate(escritor.pages):
        campos_pagina = [f for f in campos if int(f.get("pagina", -1)) == num_pag]
        if not campos_pagina:
            continue
            
        W, H = float(pagina.mediabox.width), float(pagina.mediabox.height)
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(W, H))
        dibujo_algo = False
        
        for campo in campos_pagina:
            valor = datos.get(campo.get("clave"))
            if valor in (None, "") or not campo.get("bbox"):
                continue
            x0, y0, x1, y1 = campo["bbox"]
            if campo.get("tipo") == "checkbox":
                if str(valor).strip().lower() == str(campo.get("opcion", "")).strip().lower():
                    c.setFont("Helvetica-Bold", 9)
                    c.drawCentredString(((x0 + x1) / 2) / 1000 * W, H - (((y0 + y1) / 2) / 1000 * H) - 3, "X")
                    dibujo_algo = True
            else:
                c.setFont("Helvetica", 8)
                c.drawString((x0 / 1000) * W + 2, H - ((y1 / 1000) * H) + 2, str(valor))
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

    # Si el CSV tiene columnas definidas, llenamos una fila debajo de esos encabezados
    if len(df.columns) > 1:
        nueva_fila = {col: "" for col in df.columns}
        for campo in mapa.get("campos", []):
            col_destino = str(campo.get("columna_csv", "")).strip()
            valor = datos.get(campo.get("clave"), "")
            if col_destino in nueva_fila:
                nueva_fila[col_destino] = valor
            else:
                for col_real in df.columns:
                    if col_real.strip().lower() == col_destino.lower():
                        nueva_fila[col_real] = valor
        df_salida = pd.DataFrame([nueva_fila])
    else:
        # Si el CSV es vertical o de una sola columna, exportamos formato Clave;Valor
        filas = []
        for campo in mapa.get("campos", []):
            clave = campo.get("clave", "")
            etiqueta = campo.get("columna_csv") or campo.get("etiqueta") or clave
            filas.append({"Campo": etiqueta, "Valor": datos.get(clave, "")})
        df_salida = pd.DataFrame(filas)

    df_salida.to_csv(ruta_salida, index=False, sep=";", encoding="utf-8-sig")