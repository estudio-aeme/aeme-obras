"""
cert_module.py — Módulo de certificación ObraManager
Estrategia: duplica el XML de la hoja anterior directamente en el ZIP del xlsx.
Preserva 100% el formato, fórmulas y estilos originales.
"""
import os, io, json, zipfile, re, requests
from datetime import datetime

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

DRIVE = {
    "durlock":      {"xlsx": "1oU_5BEzovYxkzkspgF3_NJvr_H71QXO8", "folder": "1S7jMjkeILwL2CqG7GasEL8ksLRc7pu0F"},
    "electricidad": {"xlsx": "1_3nTIIu1FsCHYH9FG3vznip97GpPVy6w", "folder": "1CFSIvx02W95erSUJAtrOPYxHwkvZxX2E"},
    "sanitarias":   {"xlsx": "1F5vFqXlC10FpjL9yylC3y474lwPkjh4S", "folder": "1AufrPH_vzRURf7N6ks5yhqcfXWVEzJse"},
    "pre_aa":       {"xlsx": "1BhWZWJt4LDU1BU-P2hgpDDixkFxDppqd", "folder": "1d7gODVbgNxw6UB06KdO_fXpp7JlNbSCk"},
    "herreria":     {"xlsx": "1ZV5Q9pQLnudPk8wA04uqD6qhVODOYO-4", "folder": "1Tj20i2ry813fsWx3anDbV0afRgS0mQu5"},
}

# Mapa fijo de filas por piso/tarea (estructura real del xlsx de Durlock)
DURLOCK_ROWS = {
    "PLANTA BAJA": {"Armado de estructura según plano": 11, "Emplacado general": 12, "Masillado e iluminacion": 13},
    "PISO 1°":     {"Armado de estructura según plano": 15, "Emplacado general": 16, "Masillado e iluminacion": 17},
    "PISO 2°":     {"Armado de estructura según plano": 19, "Emplacado general": 20, "Masillado e iluminacion": 21},
    "PISO 3°":     {"Armado de estructura según plano": 23, "Emplacado general": 24, "Masillado e iluminacion": 25},
    "PISO 4°":     {"Armado de estructura según plano": 27, "Emplacado general": 28, "Masillado e iluminacion": 29},
    "PISO 5°":     {"Armado de estructura según plano": 31, "Emplacado general": 32, "Masillado e iluminacion": 33},
    "PISO 6°":     {"Armado de estructura según plano": 35, "Emplacado general": 36, "Masillado e iluminacion": 37},
    "PISO 7°":     {"Armado de estructura según plano": 39, "Emplacado general": 40, "Masillado e iluminacion": 41},
    "PISO 8°":     {"Armado de estructura según plano": 43, "Emplacado general": 44, "Masillado e iluminacion": 45},
    "PISO 9°":     {"Armado de estructura según plano": 47, "Emplacado general": 48, "Masillado e iluminacion": 49},
    "PISO 10°":    {"Armado de estructura según plano": 51, "Emplacado general": 52, "Masillado e iluminacion": 53},
    "PISO 11°":    {"Armado de estructura según plano": 55, "Emplacado general": 56, "Masillado e iluminacion": 57},
    "PISO 12°":    {"Armado de estructura según plano": 59, "Emplacado general": 60, "Masillado e iluminacion": 61},
    "PISO 13°":    {"Armado de estructura según plano": 63, "Emplacado general": 64, "Masillado e iluminacion": 65},
    "PISO 14°":    {"Armado de estructura según plano": 67, "Emplacado general": 68, "Masillado e iluminacion": 69},
}

PISO_ALIAS = {
    "PLANTA BAJA (LOSA EXISTENTE)": "PLANTA BAJA", "PLANTA BAJA": "PLANTA BAJA",
}
for i in range(1, 15):
    PISO_ALIAS[f"PISO {i}°"] = f"PISO {i}°"
    PISO_ALIAS[f"PISO {i}"] = f"PISO {i}°"


def _token():
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise Exception("GOOGLE_SERVICE_ACCOUNT_JSON no configurado")
    from google.oauth2 import service_account
    from google.auth.transport.requests import Request
    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json), scopes=["https://www.googleapis.com/auth/drive"])
    creds.refresh(Request())
    return creds.token


def _download(file_id):
    r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media",
                     headers={"Authorization": f"Bearer {_token()}"})
    r.raise_for_status()
    return r.content


def _update_drive(file_id, data, mime):
    r = requests.patch(
        f"https://www.googleapis.com/upload/drive/v3/files/{file_id}?uploadType=media",
        headers={"Authorization": f"Bearer {_token()}", "Content-Type": mime}, data=data)
    r.raise_for_status()
    return r.json()


def _upload_new(folder_id, name, data, mime):
    token = _token()
    r = requests.post(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers={"Authorization": f"Bearer {token}"},
        files={"metadata": (None, json.dumps({"name": name, "parents": [folder_id]}), "application/json"),
               "file": (name, data, mime)})
    r.raise_for_status()
    return r.json().get("id")


def _excel_serial(dt):
    """Convierte datetime a número serial de Excel."""
    from datetime import date
    base = date(1899, 12, 30)
    return (dt.date() - base).days


def _set_cell_value(row_xml, col_letter, new_value):
    """Reemplaza el valor de una celda específica en el XML de la fila."""
    cell_ref = f"{col_letter}{re.search(r'r=\"(\\d+)\"', row_xml).group(1)}"
    # Patrón para la celda con valor
    pattern = rf'(<c r="{cell_ref}"[^>]*>)(?:<f>[^<]*</f>)?(<v>[^<]*</v>|<v/>)'
    replacement = rf'\g<1><v>{new_value}</v>'
    return re.sub(pattern, replacement, row_xml)


def _get_cell_value(row_xml, col_letter, row_num):
    """Obtiene el valor de una celda del XML."""
    cell_ref = f"{col_letter}{row_num}"
    m = re.search(rf'<c r="{cell_ref}"[^>]*>(?:<f>[^<]*</f>)?<v>([^<]*)</v>', row_xml)
    return float(m.group(1)) if m else 0.0


def _modify_sheet_xml(sheet_xml, num_cert, fecha_serial, avances_por_fila):
    """
    Modifica el XML de la hoja duplicada:
    - M1: número de cert
    - M2: fecha
    - Para cada fila: H = nuevo anterior, J = nuevo acumulado
    """
    lines = sheet_xml

    # Reemplazar M1 (número cert)
    lines = re.sub(
        r'(<c r="M1"[^>]*>)(?:<f>[^<]*</f>)?<v>[^<]*</v>',
        rf'\g<1><v>{num_cert}</v>', lines)

    # Reemplazar M2 (fecha)
    lines = re.sub(
        r'(<c r="M2"[^>]*>)(?:<f>[^<]*</f>)?<v>[^<]*</v>',
        rf'\g<1><v>{fecha_serial}</v>', lines)

    # Reemplazar valores H y J por cada fila
    for row_num, (nuevo_h, nuevo_j) in avances_por_fila.items():
        # Col H
        lines = re.sub(
            rf'(<c r="H{row_num}"[^>]*>)(?:<f>[^<]*</f>)?<v>[^<]*</v>',
            rf'\g<1><v>{nuevo_h}</v>', lines)
        # Col J
        lines = re.sub(
            rf'(<c r="J{row_num}"[^>]*>)(?:<f>[^<]*</f>)?<v>[^<]*</v>',
            rf'\g<1><v>{nuevo_j}</v>', lines)

    return lines


def _duplicar_xlsx_con_nueva_hoja(xlsx_bytes, src_sheet_name, new_sheet_name,
                                   num_cert, fecha_serial, avances_por_fila):
    """
    Duplica el xlsx añadiendo una nueva hoja copiada del sheet anterior,
    modificando solo los valores editables. Retorna el nuevo xlsx como bytes.
    """
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), 'r') as zin:
        all_files = zin.namelist()
        
        # Leer workbook.xml para mapear sheet name → archivo
        wb_xml = zin.read('xl/workbook.xml').decode('utf-8')
        rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        
        # Encontrar el rId y archivo de la hoja fuente
        # Encontrar todos los sheets en workbook.xml
        sheets_in_wb = re.findall(r'<sheet name="([^"]+)"[^/]*r:id="([^"]+)"', wb_xml)
        src_rid = None
        for sname, rid in sheets_in_wb:
            if sname == src_sheet_name:
                src_rid = rid
                break
        
        if not src_rid:
            raise Exception(f"No encontré la hoja '{src_sheet_name}'")
        
        # Encontrar el archivo de la hoja fuente
        # Buscar target — puede venir antes o después del Id
        src_file_match = re.search(rf'Id="{src_rid}"[^>]*Target="([^"]+)"', rels_xml)
        if not src_file_match:
            src_file_match = re.search(rf'Target="([^"]+)"[^>]*Id="{src_rid}"', rels_xml)
        if not src_file_match:
            raise Exception(f"No encontré el archivo para rId={src_rid}")
        
        src_target = src_file_match.group(1)
        # Normalizar path (puede ser relativo o absoluto)
        if src_target.startswith('/'):
            src_path = src_target[1:]
        elif src_target.startswith('xl/'):
            src_path = src_target
        else:
            src_path = f"xl/{src_target}"
        
        print(f"Hoja fuente: {src_path}")
        
        # Leer el XML de la hoja fuente
        src_xml = zin.read(src_path).decode('utf-8')
        
        # Modificar el XML con los nuevos valores
        new_xml = _modify_sheet_xml(src_xml, num_cert, fecha_serial, avances_por_fila)
        
        # Determinar el próximo número de hoja
        existing_sheet_nums = [int(re.search(r'sheet(\d+)\.xml', f).group(1))
                               for f in all_files if re.match(r'xl/worksheets/sheet\d+\.xml', f)]
        next_num = max(existing_sheet_nums) + 1
        new_file = f"xl/worksheets/sheet{next_num}.xml"
        new_rid = f"rId{next_num + 10}"  # rId único
        
        # Construir nuevo ZIP
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in all_files:
                data = zin.read(item)
                
                if item == 'xl/workbook.xml':
                    # Agregar la nueva hoja antes de </sheets>
                    decoded = data.decode('utf-8')
                    new_sheet_entry = f'<sheet name="{new_sheet_name}" sheetId="{next_num}" r:id="{new_rid}"/>'
                    decoded = decoded.replace('</sheets>', f'{new_sheet_entry}</sheets>')
                    data = decoded.encode('utf-8')
                
                elif item == 'xl/_rels/workbook.xml.rels':
                    decoded = data.decode('utf-8')
                    new_rel = f'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet{next_num}.xml" Id="{new_rid}"/>'
                    decoded = decoded.replace('</Relationships>', f'{new_rel}</Relationships>')
                    data = decoded.encode('utf-8')
                
                elif item == '[Content_Types].xml':
                    decoded = data.decode('utf-8')
                    new_ct = f'<Override PartName="/xl/worksheets/sheet{next_num}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                    decoded = decoded.replace('</Types>', f'{new_ct}</Types>')
                    data = decoded.encode('utf-8')
                
                zout.writestr(item, data)
            
            # Agregar la nueva hoja
            zout.writestr(new_file, new_xml.encode('utf-8'))
        
        return out.getvalue()


def _parse_avances(mensaje):
    system = """Parsea certificados de Durlock MO para Santiago del Estero.
Respondé SOLO JSON válido sin texto extra:
{"fecha": "dd/mm/aaaa o null", "avances": {"PISO 8°": {"Armado de estructura según plano": 10}, "PISO 9°": {"Armado de estructura según plano": 10}}}
Pisos: PLANTA BAJA, PISO 1° a PISO 14°
Tareas: "Armado de estructura según plano", "Emplacado general", "Masillado e iluminacion"
% = avance ACTUAL del certificado (no acumulado).
"estructura" = "Armado de estructura según plano"."""

    headers = {"Content-Type": "application/json", "x-api-key": ANTHROPIC_API_KEY,
               "anthropic-version": "2023-06-01"}
    body = {"model": "claude-haiku-4-5-20251001", "max_tokens": 500,
            "system": system, "messages": [{"role": "user", "content": mensaje}]}
    r = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=body)
    text = r.json()["content"][0]["text"].strip()
    text = re.sub(r"^```json|^```|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


def _generar_pdf(num_cert, fecha_dt, lineas_avance, monto_cert):
    """Genera PDF con reportlab o fpdf2."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER

        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm,
                                leftMargin=2*cm, rightMargin=2*cm)
        styles = getSampleStyleSheet()
        story = []

        # Header
        h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontSize=16, alignment=TA_CENTER, spaceAfter=4)
        h2 = ParagraphStyle('h2', parent=styles['Normal'], fontSize=11, alignment=TA_CENTER, spaceAfter=12)
        story.append(Paragraph(f"CERTIFICADO N°{num_cert}", h1))
        story.append(Paragraph(f"Durlock MO — Constitución, Santiago del Estero", h2))
        story.append(Paragraph(f"Fecha: {fecha_dt.strftime('%d/%m/%Y')}", h2))
        story.append(Spacer(1, 0.5*cm))

        # Tabla de avances
        data = [["PISO / TAREA", "% ACTUAL"]]
        for linea in lineas_avance:
            linea = linea.strip().lstrip('•').strip()
            parts = linea.rsplit(":", 1)
            if len(parts) == 2:
                data.append([parts[0].strip(), parts[1].strip()])

        t = Table(data, colWidths=[13*cm, 3*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F2D3D')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7F7F7')]),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (0, -1), 8),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.5*cm))

        # Monto
        total_s = ParagraphStyle('total', parent=styles['Normal'], fontSize=12, spaceAfter=4)
        story.append(Paragraph(f"<b>Monto certificado estimado:</b>  ${monto_cert:,.0f}", total_s))
        story.append(Spacer(1, 2*cm))

        # Firmas
        firma_data = [["_________________________", "_________________________"],
                      ["Dirección de Obra", "Contratista — Julio Cabrera"]]
        ft = Table(firma_data, colWidths=[8*cm, 8*cm])
        ft.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTSIZE', (0, 0), (-1, -1), 10),
                                ('TOPPADDING', (0, 0), (-1, -1), 4)]))
        story.append(ft)

        doc.build(story)
        return buf.getvalue()

    except Exception as e:
        print(f"PDF error: {e}")
        try:
            from fpdf import FPDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 16)
            pdf.cell(0, 10, f"CERTIFICADO N{chr(176)}{num_cert} - DURLOCK MO", ln=True, align="C")
            pdf.set_font("Helvetica", "", 11)
            pdf.cell(0, 8, f"Santiago del Estero  |  {fecha_dt.strftime('%d/%m/%Y')}", ln=True, align="C")
            pdf.ln(5)
            for linea in lineas_avance:
                pdf.cell(0, 7, linea.encode('latin-1', 'replace').decode('latin-1'), ln=True)
            pdf.ln(5)
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 8, f"Monto certificado: ${monto_cert:,.0f}", ln=True)
            return bytes(pdf.output())
        except Exception as e2:
            print(f"PDF fpdf2 error: {e2}")
            return None


def certificar_durlock(mensaje, num_whatsapp=None):
    try:
        # 1. Parsear avances
        parsed = _parse_avances(mensaje)
        avances_raw = parsed.get("avances", {})
        fecha_str = parsed.get("fecha") or datetime.now().strftime("%d/%m/%Y")

        if not avances_raw:
            return "❌ No identifiqué avances. Ejemplo: 'piso 8 estructura 10%, piso 9 estructura 10%'", False

        # Normalizar pisos
        avances = {}
        for piso_raw, tareas in avances_raw.items():
            piso_norm = PISO_ALIAS.get(piso_raw.strip(), piso_raw.strip())
            avances[piso_norm] = tareas

        try:
            fecha_dt = datetime.strptime(fecha_str, "%d/%m/%Y")
        except:
            fecha_dt = datetime.now()
        fecha_serial = _excel_serial(fecha_dt)

        # 2. Descargar xlsx
        xlsx_bytes = _download(DRIVE["durlock"]["xlsx"])

        # 3. Detectar última hoja de cert para copiar
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            wb_xml = zf.read('xl/workbook.xml').decode('utf-8')
            rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')

        cert_sheets = re.findall(r'<sheet name="(CERT N°(\d+))"', wb_xml)
        if cert_sheets:
            cert_nums_found = [(name, int(n)) for name, n in cert_sheets]
            cert_nums_found.sort(key=lambda x: x[1])
            last_cert_name, last_cert_num = cert_nums_found[-1]
            num_cert = last_cert_num + 1
        else:
            last_cert_name = None
            num_cert = 1

        new_sheet_name = f"CERT N°{num_cert}"
        print(f"Generando {new_sheet_name} basado en {last_cert_name}")

        # 4. Leer acumulados del cert anterior para calcular nuevos
        # Necesitamos leer los J values del último cert para saber el anterior
        prev_j_values = {}
        if last_cert_name:
            with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
                # Encontrar el archivo del último cert
                sheets_in_wb = re.findall(r'<sheet name="([^"]+)"[^/]*r:id="([^"]+)"', wb_xml)
                src_rid = next((rid for name, rid in sheets_in_wb if name == last_cert_name), None)
                if src_rid:
                    # Buscar target del rId en rels (puede tener / al inicio)
                    src_match = re.search(rf'Id="{src_rid}"[^>]*Target="([^"]+)"', rels_xml)
                    if not src_match:
                        src_match = re.search(rf'Target="([^"]+)"[^>]*Id="{src_rid}"', rels_xml)
                    if src_match:
                        src_target = src_match.group(1)
                        if src_target.startswith('/'):
                            src_path = src_target[1:]
                        elif src_target.startswith('xl/'):
                            src_path = src_target
                        else:
                            src_path = f"xl/{src_target}"
                        prev_xml = zf.read(src_path).decode('utf-8')

                        # Leer todos los J values
                        for piso, tareas_map in DURLOCK_ROWS.items():
                            for tarea, row_num in tareas_map.items():
                                j_match = re.search(
                                    rf'<c r="J{row_num}"[^>]*>(?:<f>[^<]*</f>)?<v>([^<]*)</v>',
                                    prev_xml)
                                if j_match:
                                    prev_j_values[row_num] = float(j_match.group(1))

        # 5. Construir mapa de cambios: {row_num: (nuevo_H, nuevo_J)}
        avances_por_fila = {}
        lineas_avance = []
        monto_cert_est = 0

        for piso, tareas in avances.items():
            if piso not in DURLOCK_ROWS:
                continue
            for tarea_raw, pct_act in tareas.items():
                tarea_key = next((t for t in DURLOCK_ROWS[piso]
                                  if tarea_raw.lower() in t.lower() or t.lower() in tarea_raw.lower()), None)
                if not tarea_key:
                    continue

                row_num = DURLOCK_ROWS[piso][tarea_key]
                pct_act_dec = pct_act / 100.0
                nuevo_h = prev_j_values.get(row_num, 0.0)
                nuevo_j = min(1.0, nuevo_h + pct_act_dec)
                avances_por_fila[row_num] = (nuevo_h, nuevo_j)

                tarea_corta = tarea_key.replace("según plano", "").replace("e iluminacion", "").strip()
                lineas_avance.append(f"  • {piso} — {tarea_corta}: {pct_act}%")

        # Para pisos NO mencionados: H = J = acumulado anterior (sin avance nuevo)
        for piso, tareas_map in DURLOCK_ROWS.items():
            for tarea, row_num in tareas_map.items():
                if row_num not in avances_por_fila:
                    prev_j = prev_j_values.get(row_num, 0.0)
                    avances_por_fila[row_num] = (prev_j, prev_j)

        # 6. Duplicar xlsx con nueva hoja
        xlsx_nuevo = _duplicar_xlsx_con_nueva_hoja(
            xlsx_bytes, last_cert_name, new_sheet_name,
            num_cert, fecha_serial, avances_por_fila
        )

        # 7. Subir xlsx al Drive
        _update_drive(DRIVE["durlock"]["xlsx"], xlsx_nuevo,
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 8. Generar y subir PDF
        pdf_str = ""
        try:
            pdf_bytes = _generar_pdf(num_cert, fecha_dt, lineas_avance, monto_cert_est)
            if pdf_bytes:
                pdf_nombre = f"Cs_SDE_DURLOCK_CERT_N{num_cert}_{fecha_dt.strftime('%Y%m%d')}.pdf"
                fid = _upload_new(DRIVE["durlock"]["folder"], pdf_nombre, pdf_bytes, "application/pdf")
                if fid:
                    pdf_str = f"\n📎 PDF: https://drive.google.com/file/d/{fid}/view"
        except Exception as pe:
            print(f"PDF skip: {pe}")

        # 9. Respuesta
        respuesta = (
            f"✅ *{new_sheet_name} — Durlock interno*\n"
            f"📅 {fecha_dt.strftime('%d/%m/%Y')}\n\n"
            f"*Avances:*\n" + "\n".join(lineas_avance) + "\n\n"
            f"📊 Excel actualizado en Drive"
            + pdf_str
        )
        return respuesta, True

    except Exception as e:
        import traceback
        print(f"CERT ERROR: {traceback.format_exc()}")
        return f"❌ Error: {str(e)}", False
