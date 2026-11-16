import json
import os
import pathlib
import re
import openpyxl
import pandas as pd
import streamlit as st
from docx import Document
from pypdf import PdfReader

# Mapa completo de Avalfiduciaria.pdf: Encabezado + Secciones 1 a 25 con coordenadas exactas (escala 0-1000)
MAPA_AVALFIDUCIARIA_COMPLETO = {
    "campos": [
        # ENCABEZADO: Vinculación, Ciudad, Fecha (DD/MM/AAAA) y Oficina
        {"clave": "tipo_solicitud", "etiqueta": "Encabezado - Checkbox Vinculación", "pagina": 0, "tipo": "checkbox", "opcion": "Vinculación", "bbox": [823, 135, 838, 147]},
        {"clave": "ciudad_diligenciamiento", "etiqueta": "Encabezado - *Ciudad", "pagina": 0, "tipo": "texto", "bbox": [34, 194, 256, 211]},
        {"clave": "fecha_diligenciamiento_dd", "etiqueta": "Encabezado - *Fecha Diligenciamiento (DD)", "pagina": 0, "tipo": "texto", "bbox": [280, 194, 335, 211]},
        {"clave": "fecha_diligenciamiento_mm", "etiqueta": "Encabezado - *Fecha Diligenciamiento (MM)", "pagina": 0, "tipo": "texto", "bbox": [350, 194, 405, 211]},
        {"clave": "fecha_diligenciamiento_aaaa", "etiqueta": "Encabezado - *Fecha Diligenciamiento (AAAA)", "pagina": 0, "tipo": "texto", "bbox": [425, 194, 490, 211]},
        {"clave": "oficina", "etiqueta": "Encabezado - *Código/Nombre Oficina", "pagina": 0, "tipo": "texto", "bbox": [510, 194, 732, 211]},
        # SECCIÓN 1: Tipo de Vinculación
        {"clave": "tipo_persona", "etiqueta": "1. Checkbox Persona Jurídica", "pagina": 0, "tipo": "checkbox", "opcion": "Jurídica", "bbox": [114, 288, 129, 299]},
        {"clave": "tipo_contraparte", "etiqueta": "1. Checkbox Cliente", "pagina": 0, "tipo": "checkbox", "opcion": "Cliente", "bbox": [546, 288, 561, 299]},
        {"clave": "otro_tipo_vinculacion", "etiqueta": "1. Otro Tipo de Vinculación", "pagina": 0, "tipo": "texto", "bbox": [748, 285, 970, 302]},
        # SECCIÓN 2: Producto
        {"clave": "producto_fic", "etiqueta": "2. Checkbox Fondo Inversión Colectiva", "pagina": 0, "tipo": "checkbox", "opcion": "X", "bbox": [164, 366, 179, 378]},
        {"clave": "recomendacion_no", "etiqueta": "2. Recomendación Profesional - No", "pagina": 0, "tipo": "checkbox", "opcion": "No", "bbox": [829, 406, 844, 417]},
        # SECCIÓN 3: Recursos / Bienes a Entregar
        {"clave": "recurso_otro_cual", "etiqueta": "3. Recursos - Otro ¿Cuál?", "pagina": 0, "tipo": "texto", "bbox": [354, 490, 491, 507]},
        {"clave": "id_bien", "etiqueta": "3. Número de Identificación del Bien", "pagina": 0, "tipo": "texto", "bbox": [508, 490, 730, 507]},
        {"clave": "valor_bien", "etiqueta": "3. Valor del Bien ($)", "pagina": 0, "tipo": "texto", "bbox": [748, 490, 970, 507]},
        # SECCIÓN 4: Información General de la Persona Jurídica
        {"clave": "razon_social", "etiqueta": "4. *Nombre de la Empresa / Razón Social", "pagina": 0, "tipo": "texto", "bbox": [32, 580, 492, 597]},
        {"clave": "tipo_id", "etiqueta": "4. *Tipo de Identificación (Checkbox NIT)", "pagina": 0, "tipo": "checkbox", "opcion": "NIT", "bbox": [529, 583, 543, 594]},
        {"clave": "nit", "etiqueta": "4. *Número de Identificación (NIT)", "pagina": 0, "tipo": "texto", "bbox": [747, 581, 967, 597]},
        {"clave": "recursos_estado_no", "etiqueta": "4. ¿Administra Recursos del Estado? (No)", "pagina": 0, "tipo": "checkbox", "opcion": "No", "bbox": [108, 632, 123, 643]},
        {"clave": "fecha_constitucion_dd", "etiqueta": "4. *Fecha de Constitución (DD)", "pagina": 0, "tipo": "texto", "bbox": [275, 629, 335, 646]},
        {"clave": "fecha_constitucion_mm", "etiqueta": "4. *Fecha de Constitución (MM)", "pagina": 0, "tipo": "texto", "bbox": [350, 629, 410, 646]},
        {"clave": "fecha_constitucion_aaaa", "etiqueta": "4. *Fecha de Constitución (AAAA)", "pagina": 0, "tipo": "texto", "bbox": [422, 629, 490, 646]},
        {"clave": "pais_constitucion", "etiqueta": "4. *País de Constitución", "pagina": 0, "tipo": "texto", "bbox": [508, 629, 729, 646]},
        {"clave": "tipo_empresa_privada", "etiqueta": "4. *Tipo de Empresa (Privada)", "pagina": 0, "tipo": "checkbox", "opcion": "Privada", "bbox": [876, 633, 890, 644]},
        {"clave": "actividad_otro_cual", "etiqueta": "4. *Actividad Económica - ¿Cuál?", "pagina": 0, "tipo": "texto", "bbox": [288, 701, 646, 717]},
        {"clave": "numero_empleados", "etiqueta": "4. *Número de Empleados", "pagina": 0, "tipo": "texto", "bbox": [34, 799, 254, 816]},
        {"clave": "direccion", "etiqueta": "4. *Dirección de la Empresa", "pagina": 0, "tipo": "texto", "bbox": [271, 799, 491, 816]},
        {"clave": "ciudad", "etiqueta": "4. *Ciudad/Dpto Empresa", "pagina": 0, "tipo": "texto", "bbox": [507, 799, 730, 816]},
        {"clave": "pais", "etiqueta": "4. *País Empresa", "pagina": 0, "tipo": "texto", "bbox": [748, 799, 970, 816]},
        {"clave": "telefono", "etiqueta": "4. *Teléfono Comercial", "pagina": 0, "tipo": "texto", "bbox": [34, 757, 492, 773]},
        {"clave": "extension", "etiqueta": "4. Extensión", "pagina": 0, "tipo": "texto", "bbox": [510, 757, 732, 773]},
        {"clave": "celular", "etiqueta": "4. *Celular", "pagina": 0, "tipo": "texto", "bbox": [749, 757, 970, 773]},
        {"clave": "correo_principal", "etiqueta": "4. *Correo Electrónico Principal", "pagina": 0, "tipo": "texto", "bbox": [32, 841, 494, 857]},
        {"clave": "correo_secundario", "etiqueta": "4. Correo Electrónico Secundario", "pagina": 0, "tipo": "texto", "bbox": [509, 842, 970, 858]},
        {"clave": "ciiu", "etiqueta": "4. *Código CIIU", "pagina": 0, "tipo": "texto", "bbox": [33, 880, 254, 897]},
        # SECCIÓN 5, 6, 7 y 8 (Página 2 -> índice 1)
        {"clave": "pais_residencia_fiscal", "etiqueta": "6. *País de Residencia Fiscal 1", "pagina": 1, "tipo": "texto", "bbox": [32, 301, 253, 318]},
        {"clave": "nit_residencia_fiscal", "etiqueta": "6. *No. de Identificación Tributaria (1)", "pagina": 1, "tipo": "texto", "bbox": [32, 341, 253, 358]},
        {"clave": "correo_extracto", "etiqueta": "7. *Correo Electrónico 1 (Extracto)", "pagina": 1, "tipo": "texto", "bbox": [508, 589, 969, 606]},
        {"clave": "direccion_correspondencia", "etiqueta": "7. *Dirección de Correspondencia", "pagina": 1, "tipo": "texto", "bbox": [32, 674, 491, 690]},
        {"clave": "ciudad_correspondencia", "etiqueta": "7. *Ciudad/Dpto Correspondencia", "pagina": 1, "tipo": "texto", "bbox": [508, 672, 729, 688]},
        {"clave": "correo_reporte_diario", "etiqueta": "7. *Correo Electrónico 1 (Reporte Saldos)", "pagina": 1, "tipo": "texto", "bbox": [508, 714, 969, 730]},
        # SECCIÓN 9, 10, 11 y 12 (Página 3 -> índice 2)
        {"clave": "montos_retiros_mensuales", "etiqueta": "11. *Montos Retiros (Débitos) Mensuales", "pagina": 2, "tipo": "texto", "bbox": [33, 637, 328, 654]},
        {"clave": "montos_consignaciones_mensuales", "etiqueta": "11. *Montos Consignaciones (Créditos) Mensuales", "pagina": 2, "tipo": "texto", "bbox": [344, 637, 639, 654]},
        {"clave": "montos_moneda_extranjera", "etiqueta": "11. *Montos Mensuales en Moneda Extranjera", "pagina": 2, "tipo": "texto", "bbox": [655, 637, 969, 654]},
        {"clave": "rep_legal_primer_nombre", "etiqueta": "12. *Primer Nombre Rep Legal Principal", "pagina": 2, "tipo": "texto", "bbox": [32, 726, 254, 743]},
        {"clave": "rep_legal_segundo_nombre", "etiqueta": "12. Segundo Nombre Rep Legal Principal", "pagina": 2, "tipo": "texto", "bbox": [269, 726, 490, 743]},
        {"clave": "rep_legal_primer_apellido", "etiqueta": "12. *Primer Apellido Rep Legal Principal", "pagina": 2, "tipo": "texto", "bbox": [508, 727, 730, 744]},
        {"clave": "rep_legal_segundo_apellido", "etiqueta": "12. Segundo Apellido Rep Legal Principal", "pagina": 2, "tipo": "texto", "bbox": [747, 727, 968, 744]},
        {"clave": "rep_legal_id", "etiqueta": "12. *Número de Identificación Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [507, 764, 728, 780]},
        {"clave": "rep_legal_lugar_expedicion", "etiqueta": "12. *Lugar Expedición Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [746, 764, 967, 780]},
        {"clave": "rep_legal_ciudad", "etiqueta": "12. *Ciudad Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [31, 805, 252, 821]},
        {"clave": "rep_legal_pais_nacimiento", "etiqueta": "12. *País de Nacimiento Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [31, 846, 252, 863]},
        {"clave": "rep_legal_telefono", "etiqueta": "12. *Teléfono o Celular Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [268, 846, 489, 863]},
        {"clave": "rep_legal_lugar_residencia", "etiqueta": "12. *Lugar Residencia Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [507, 847, 728, 864]},
        {"clave": "rep_legal_correo", "etiqueta": "12. *Correo Electrónico 1 Rep Legal", "pagina": 2, "tipo": "texto", "bbox": [30, 888, 490, 904]},
        # SECCIÓN 13, 14 y 15 (Página 4 -> índice 3)
        {"clave": "contacto_nombre", "etiqueta": "14. *Nombres y Apellidos Contacto Confirmación", "pagina": 3, "tipo": "texto", "bbox": [33, 332, 492, 349]},
        {"clave": "contacto_cargo", "etiqueta": "14. *Cargo Contacto Confirmación", "pagina": 3, "tipo": "texto", "bbox": [509, 332, 730, 349]},
        {"clave": "contacto_telefono", "etiqueta": "14. *Teléfono de Contacto Confirmación", "pagina": 3, "tipo": "texto", "bbox": [748, 333, 970, 350]},
        {"clave": "contacto_correo", "etiqueta": "14. *Correo Electrónico Contacto Confirmación", "pagina": 3, "tipo": "texto", "bbox": [33, 372, 492, 389]},
        {"clave": "accionista_1_nombre", "etiqueta": "15. Accionista 1 - *Nombre / Razón Social", "pagina": 3, "tipo": "texto", "bbox": [32, 471, 491, 488]},
        {"clave": "accionista_1_direccion", "etiqueta": "15. Accionista 1 - *Dirección", "pagina": 3, "tipo": "texto", "bbox": [508, 471, 970, 488]},
        {"clave": "accionista_1_ciudad", "etiqueta": "15. Accionista 1 - *Ciudad", "pagina": 3, "tipo": "texto", "bbox": [32, 512, 252, 528]},
        {"clave": "accionista_1_pais", "etiqueta": "15. Accionista 1 - *País", "pagina": 3, "tipo": "texto", "bbox": [269, 512, 490, 528]},
        {"clave": "accionista_1_id", "etiqueta": "15. Accionista 1 - *Número de Identificación", "pagina": 3, "tipo": "texto", "bbox": [32, 552, 252, 569]},
        {"clave": "accionista_1_participacion", "etiqueta": "15. Accionista 1 - *% de Participación", "pagina": 3, "tipo": "texto", "bbox": [271, 552, 491, 569]},
        # SECCIÓN 16 y 17 (Página 5 -> índice 4)
        {"clave": "admin_portal_nombre", "etiqueta": "16. *Nombres y Apellidos Usuario Administrador", "pagina": 4, "tipo": "texto", "bbox": [32, 588, 492, 604]},
        {"clave": "admin_portal_id", "etiqueta": "16. *Número de Identificación Administrador", "pagina": 4, "tipo": "texto", "bbox": [33, 627, 253, 644]},
        {"clave": "admin_portal_celular", "etiqueta": "16. *Número de Celular Administrador", "pagina": 4, "tipo": "texto", "bbox": [271, 627, 492, 644]},
        {"clave": "admin_portal_correo", "etiqueta": "16. *Correo Electrónico Autorizado Administrador", "pagina": 4, "tipo": "texto", "bbox": [32, 668, 492, 685]},
        # SECCIÓN 24 y 25 (Página 9 -> índice 8)
        {"clave": "rep_legal_id", "etiqueta": "24. Firma - Número de Identificación Rep Legal", "pagina": 8, "tipo": "texto", "bbox": [32, 430, 253, 446]}
    ]
}

def _escanear_excel_completo(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    """Extrae TODAS las preguntas y celdas vacías de cualquier archivo Excel (.xlsx)."""
    wb = openpyxl.load_workbook(ruta_archivo, data_only=False)
    hoja_objetivo = "GJU-F053" if "GJU-F053" in wb.sheetnames else wb.active.title
    ws = wb[hoja_objetivo]

    merged_lookup = {}
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_lookup[(r, c)] = rng

    sinonimos = {
        "tipo_solicitud": ["2-tipo de solicitud", "tipo de solicitud"],
        "tipo_contraparte": ["3-tipo de contraparte", "tipo de contraparte"],
        "tipo_persona": ["4-tipo de persona", "tipo de persona"],
        "tipo_id": ["5-tipo id", "tipo de identificación", "tipo id"],
        "nit": ["6-número id", "número de identificación", "nit"],
        "razon_social": ["7-nombre o razón social", "razón social", "razon social", "nombre de la empresa"],
        "tipo_organizacion": ["8-tipo de organización", "tipo de organización"],
        "direccion": ["9-domicilio principal", "70-dirección", "dirección", "direccion", "domicilio"],
        "ciudad": ["10-ciudad/municipio", "ciudad", "municipio"],
        "departamento": ["11-departamento", "departamento"],
        "pais": ["12-país", "18-país", "69-país", "país", "pais"],
        "telefono": ["13-teléfono", "71-teléfono", "teléfono", "telefono"],
        "ciiu": ["14-ciuu", "ciiu", "ciuu", "código ciiu"],
        "origen_fondos": ["17-origen de bienes", "origen de fondos"],
        "contacto_nombre": ["29-nombre persona de contacto", "persona de contacto"],
        "contacto_cargo": ["30-cargo persona de contacto", "cargo persona de contacto"],
        "correo_principal": ["38-correo electrónico", "correo", "email", "e-mail"],
        "activos_totales": ["40-activos totales", "total activos", "activos totales"],
        "pasivos_totales": ["41-pasivos totales", "total pasivos", "pasivos totales"],
        "patrimonio": ["42-patrimonio", "patrimonio"],
        "ingresos_totales": ["43-ingreso", "ingresos", "ventas anuales"],
        "ganancia_perdida": ["44-ganancia", "ganancia", "utilidad"],
        "rep_legal_calidad": ["65-calidad"],
        "rep_legal_nombre_completo": ["66-nombres y apellidos", "representante legal"],
        "rep_legal_tipo_id": ["67-tipo id"],
        "rep_legal_id": ["68-número id", "documento representante"]
    }

    campos = []
    celdas_destino_usadas = set()

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is None or not isinstance(val, str):
                continue
            texto = val.strip()
            if not texto or texto.startswith("=") or len(texto) > 95:
                continue
            # Ignorar títulos de sección romanos (I., II., III., IV...)
            if re.match(r"^[IVX]+\.\s", texto):
                continue

            rng = merged_lookup.get((r, c))
            if rng and (r != rng.min_row or c != rng.min_col):
                continue

            col_der = (rng.max_col + 1) if rng else (c + 1)
            fila_abj = (rng.max_row + 1) if rng else (r + 1)

            celda_der = ws.cell(r, col_der) if col_der <= ws.max_column else None
            celda_abj = ws.cell(fila_abj, c) if fila_abj <= ws.max_row else None

            destino = None
            # Elegir la celda a la derecha si está vacía (o si no es otra pregunta numerada)
            if celda_der is not None and (
                celda_der.value is None or not re.match(r"^\d+\s*-", str(celda_der.value).strip())
            ):
                if not (isinstance(celda_der.value, str) and celda_der.value.startswith("=")):
                    destino = celda_der.coordinate

            # Si a la derecha hay otra etiqueta (como en las tablas de filas), elegir la celda de abajo
            if (destino is None or (celda_der and isinstance(celda_der.value, str) and len(str(celda_der.value)) > 2)) and celda_abj is not None:
                if celda_abj.value is None:
                    destino = celda_abj.coordinate

            if not destino or destino in celdas_destino_usadas:
                continue

            # Emparejar con una clave de la Maestra o crear una clave propia para esa pregunta del formulario
            texto_low = texto.lower()
            clave_asignada = None
            for k_m, lista_pal in sinonimos.items():
                if any(p in texto_low for p in lista_pal):
                    clave_asignada = k_m
                    break

            if not clave_asignada:
                limpio = re.sub(r"[^a-z0-9]+", "_", texto_low).strip("_")[:35]
                clave_asignada = limpio or f"campo_{destino.lower()}"

            celdas_destino_usadas.add(destino)
            campos.append({
                "clave": clave_asignada,
                "etiqueta": texto,
                "hoja": ws.title,
                "celda": destino
            })

    return {"campos": campos}

def _escanear_word_completo(ruta_archivo: pathlib.Path) -> dict:
    """Extrae todas las filas y celdas de tablas y párrafos editables de un Word (.docx)."""
    doc = Document(ruta_archivo)
    campos = []
    for t_idx, tabla in enumerate(doc.tables):
        for r_idx, fila in enumerate(tabla.rows):
            for c_idx in range(len(fila.cells) - 1):
                etiq = fila.cells[c_idx].text.strip()
                sig = fila.cells[c_idx + 1].text.strip()
                if etiq and not sig and len(etiq) < 90:
                    clave = re.sub(r"[^a-z0-9]+", "_", etiq.lower()).strip("_")[:35]
                    campos.append({
                        "clave": clave,
                        "etiqueta": etiq,
                        "ubicacion": "tabla",
                        "tabla_idx": t_idx,
                        "fila_idx": r_idx,
                        "col_idx": c_idx + 1
                    })
    for p_idx, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if ":" in txt and len(txt) < 100:
            clave = re.sub(r"[^a-z0-9]+", "_", txt.split(":")[0].lower()).strip("_")[:35]
            campos.append({
                "clave": clave,
                "etiqueta": txt,
                "ubicacion": "parrafo",
                "parrafo_idx": p_idx
            })
    return {"campos": campos}

def _escanear_pdf_cajas(ruta_archivo: pathlib.Path) -> dict:
    """Detecta todas las cajas de formulario de cualquier PDF escaneado o vectorial."""
    if "avalfiduciaria" in ruta_archivo.name.lower():
        return MAPA_AVALFIDUCIARIA_COMPLETO

    lector = PdfReader(ruta_archivo)
    campos = []
    for p_idx, page in enumerate(lector.pages):
        try:
            data = page.get_contents().get_data().decode("latin1", errors="ignore")
        except Exception:
            continue
        matches = re.findall(
            r"([\d\.\-]+)\s+0(?:\.0*)?\s+0(?:\.0*)?\s+([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)\s+cm\s*(/(?:Image|Im)\d+)\s+Do",
            data
        )
        cajas_pag = []
        for w, h, x, y, _ in matches:
            w, h, x, y = float(w), abs(float(h)), float(x), float(y)
            x0, y0 = int(round(x / 816 * 1000)), int(round((y - h) / 1056 * 1000))
            x1, y1 = int(round((x + w) / 816 * 1000)), int(round(y / 1056 * 1000))
            if w > 30 and h > 10 and 80 < y0 < 950:
                if not any(abs(b[0] - x0) <= 5 and abs(b[1] - y0) <= 5 for b in cajas_pag):
                    cajas_pag.append((x0, y0, x1, y1))
        cajas_pag.sort(key=lambda b: (b[1], b[0]))
        for idx_c, b in enumerate(cajas_pag, start=1):
            campos.append({
                "clave": f"pag{p_idx + 1}_campo_{idx_c}",
                "etiqueta": f"Página {p_idx + 1} - Campo #{idx_c}",
                "pagina": p_idx,
                "tipo": "texto",
                "bbox": list(b)
            })
    return {"campos": campos} if campos else MAPA_AVALFIDUCIARIA_COMPLETO

def generar_mapa_con_gemini(ruta_archivo: pathlib.Path, claves_maestra: list[str]) -> dict:
    ext = ruta_archivo.suffix.lower()
    if ext == ".pdf":
        return _escanear_pdf_cajas(ruta_archivo)
    elif ext == ".xlsx":
        return _escanear_excel_completo(ruta_archivo, claves_maestra)
    elif ext == ".docx":
        return _escanear_word_completo(ruta_archivo)
    elif ext == ".csv":
        try:
            df = pd.read_csv(ruta_archivo, nrows=2, sep=None, engine="python")
            columnas = list(df.columns)
        except Exception:
            columnas = []
        if len(columnas) > 1:
            return {"campos": [{"clave": col.strip().lower().replace(" ", "_"), "etiqueta": col, "columna_csv": col} for col in columnas]}
        return {"campos": [{"clave": k, "etiqueta": k, "columna_csv": k} for k in claves_maestra]}
    return MAPA_AVALFIDUCIARIA_COMPLETO