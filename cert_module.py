"""
cert_module.py — Módulo de certificación ObraManager
Estrategia: copia el último cert y solo modifica los valores editables.
Así preserva 100% el formato, fórmulas y estructura original.
"""
import os, io, json, copy, requests, tempfile, subprocess
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import importlib.util

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

# ─── IDs Drive ────────────────────────────────────────────────────────────────
DRIVE = {
    "durlock":      {"xlsx": "1oU_5BEzovYxkzkspgF3_NJvr_H71QXO8", "folder": "1S7jMjkeILwL2CqG7GasEL8ksLRc7pu0F"},
    "electricidad": {"xlsx": "1_3nTIIu1FsCHYH9FG3vznip97GpPVy6w", "folder": "1CFSIvx02W95erSUJAtrOPYxHwkvZxX2E"},
    "sanitarias":   {"xlsx": "1F5vFqXlC10FpjL9yylC3y474lwPkjh4S", "folder": "1AufrPH_vzRURf7N6ks5yhqcfXWVEzJse"},
    "pre_aa":       {"xlsx": "1BhWZWJt4LDU1BU-P2hgpDDixkFxDppqd", "folder": "1d7gODVbgNxw6UB06KdO_fXpp7JlNbSCk"},
    "herreria":     {"xlsx": "1ZV5Q9pQLnudPk8wA04uqD6qhVODOYO-4", "folder": "1Tj20i2ry813fsWx3anDbV0afRgS0mQu5"},
}

# Mapa fijo: piso -> fila en la hoja de cert (basado en estructura real del xlsx)
DURLOCK_ROW_MAP = {
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
# Alias para normalizar nombres de piso
PISO_ALIAS = {
    "PLANTA BAJA (LOSA EXISTENTE)": "PLANTA BAJA",
    "PLANTA BAJA": "PLANTA BAJA",
}
for i in range(1, 15):
    PISO_ALIAS[f"PISO {i}°"] = f"PISO {i}°"
    PISO_ALIAS[f"PISO {i}"] = f"PISO {i}°"

# Columnas en la hoja de cert
COL_H = 8   # % ANTERIOR (traer del acumulado anterior)
COL_I = 9   # % ACTUAL (fórmula =J-H, no editar)
COL_J = 10  # % ACUMULADO (editar: anterior + actual)
COL_M1 = 13 # Número de certificado
COL_M2 = 13 # Fecha


def _token():
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise Exception("GOOGLE_SERVICE_ACCOUNT_JSON no configurado")
    from google.oauth2 import service_account
    from google.auth.transport.requests import Request
    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json),
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    creds.refresh(Request())
    return creds.token


def _download(file_id):
    r = requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media",
        headers={"Authorization": f"Bearer {_token()}"}
    )
    r.raise_for_status()
    return r.content


def _update(file_id, content_bytes, mime_type):
    r = requests.patch(
        f"https://www.googleapis.com/upload/drive/v3/files/{file_id}?uploadType=media",
        headers={"Authorization": f"Bearer {_token()}", "Content-Type": mime_type},
        data=content_bytes
    )
    r.raise_for_status()
    return r.json()


def _upload_new(folder_id, filename, content_bytes, mime_type):
    token = _token()
    metadata = {"name": filename, "parents": [folder_id]}
    r = requests.post(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers={"Authorization": f"Bearer {token}"},
        files={
            "metadata": (None, json.dumps(metadata), "application/json"),
            "file": (filename, content_bytes, mime_type),
        }
    )
    r.raise_for_status()
    return r.json().get("id")


def _parse_avances(mensaje, contrato):
    system = f"""Parsea certificados de construcción para {contrato}.
Respondé SOLO JSON válido sin texto extra:
{{
  "fecha": "dd/mm/aaaa o null",
  "avances": {{
    "PISO 8°": {{"Armado de estructura según plano": 10}},
    "PISO 9°": {{"Armado de estructura según plano": 10}}
  }}
}}
Pisos válidos: PLANTA BAJA, PISO 1° a PISO 14°
Tareas: "Armado de estructura según plano", "Emplacado general", "Masillado e iluminacion"
Los % son el avance ACTUAL del certificado (no el acumulado).
Si dice solo "estructura" se refiere a "Armado de estructura según plano"."""

    headers = {"Content-Type": "application/json", "x-api-key": ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01"}
    body = {"model": "claude-haiku-4-5-20251001", "max_tokens": 500, "system": system,
            "messages": [{"role": "user", "content": mensaje}]}
    r = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=body)
    import re
    text = r.json()["content"][0]["text"].strip()
    text = re.sub(r"^```json|^```|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


def _copiar_hoja(wb, src_name, dst_name):
    """Copia una hoja preservando todo el contenido y formato."""
    from copy import copy
    src = wb[src_name]
    dst = wb.copy_worksheet(src)
    dst.title = dst_name
    return dst


def certificar_durlock(mensaje, num_whatsapp=None):
    try:
        # 1. Parsear avances con IA
        parsed = _parse_avances(mensaje, "Durlock MO")
        avances_raw = parsed.get("avances", {})
        fecha_str = parsed.get("fecha") or datetime.now().strftime("%d/%m/%Y")

        if not avances_raw:
            return "❌ No identifiqué avances. Ejemplo: 'piso 8 estructura 10%, piso 9 estructura 10%'", False

        # Normalizar nombres de piso
        avances = {}
        for piso_raw, tareas in avances_raw.items():
            piso_norm = PISO_ALIAS.get(piso_raw.strip(), piso_raw.strip())
            avances[piso_norm] = tareas

        # 2. Descargar xlsx
        xlsx_bytes = _download(DRIVE["durlock"]["xlsx"])
        wb = load_workbook(io.BytesIO(xlsx_bytes))

        # Determinar número de cert
        cert_sheets = [s for s in wb.sheetnames if s.startswith("CERT N°")]
        cert_nums = []
        for s in cert_sheets:
            try: cert_nums.append(int(s.replace("CERT N°", "")))
            except: pass
        num_cert = max(cert_nums) + 1 if cert_nums else 1
        ultimo_cert = f"CERT N°{max(cert_nums)}" if cert_nums else None

        # 3. Copiar última hoja de cert como base
        if ultimo_cert and ultimo_cert in wb.sheetnames:
            nueva_hoja = wb.copy_worksheet(wb[ultimo_cert])
            nueva_hoja.title = f"CERT N°{num_cert}"
        else:
            nueva_hoja = wb.create_sheet(f"CERT N°{num_cert}")

        ws = nueva_hoja

        # 4. Actualizar número y fecha
        ws.cell(1, COL_M1).value = num_cert
        try:
            fecha_dt = datetime.strptime(fecha_str, "%d/%m/%Y")
        except:
            fecha_dt = datetime.now()
        ws.cell(2, COL_M2).value = fecha_dt

        # 5. Para cada piso/tarea: actualizar H (anterior) y J (acumulado)
        monto_cert = 0
        lineas_avance = []

        for piso, tareas in avances.items():
            if piso not in DURLOCK_ROW_MAP:
                continue
            for tarea_raw, pct_act in tareas.items():
                # Buscar la tarea exacta
                tarea_key = None
                for t in DURLOCK_ROW_MAP[piso]:
                    if tarea_raw.lower() in t.lower() or t.lower() in tarea_raw.lower():
                        tarea_key = t
                        break
                if not tarea_key:
                    continue

                row = DURLOCK_ROW_MAP[piso][tarea_key]
                pct_act_dec = pct_act / 100.0

                # Leer el acumulado anterior (J del cert anterior)
                pct_ant = 0.0
                if ultimo_cert and ultimo_cert in wb.sheetnames:
                    prev_ws = wb[ultimo_cert]
                    j_val = prev_ws.cell(row, COL_J).value
                    if isinstance(j_val, (int, float)):
                        pct_ant = float(j_val)

                pct_acum = min(1.0, pct_ant + pct_act_dec)

                # Actualizar: H = anterior, J = nuevo acumulado
                ws.cell(row, COL_H).value = pct_ant
                ws.cell(row, COL_J).value = pct_acum

                # El % actual (col I) es fórmula =J-H, no tocar
                # Pero calculamos el importe para el resumen
                f_val = ws.cell(row, 6).value  # col F = valor total
                if isinstance(f_val, (int, float)) and f_val:
                    imp_act = round(f_val * pct_act_dec)
                    monto_cert += imp_act

                tarea_corta = tarea_key.replace("según plano", "").replace("e iluminacion", "").strip()
                lineas_avance.append(f"  • {piso} — {tarea_corta}: {pct_act}%")

        # 6. Resetear todos los demás % actuales a 0
        # (poner J = H para pisos no mencionados → actual queda 0)
        for piso, tareas_map in DURLOCK_ROW_MAP.items():
            if piso in avances:
                continue  # ya lo procesamos
            for tarea, row in tareas_map.items():
                # Si no se certificó este piso/tarea, el acumulado queda igual que el anterior
                pct_ant = 0.0
                if ultimo_cert and ultimo_cert in wb.sheetnames:
                    prev_ws = wb[ultimo_cert]
                    j_val = prev_ws.cell(row, COL_J).value
                    if isinstance(j_val, (int, float)):
                        pct_ant = float(j_val)
                ws.cell(row, COL_H).value = pct_ant
                ws.cell(row, COL_J).value = pct_ant  # acumulado = anterior (sin avance nuevo)

        # 7. Guardar xlsx
        out = io.BytesIO()
        wb.save(out)
        xlsx_nuevo = out.getvalue()

        # 8. Subir xlsx al Drive
        _update(DRIVE["durlock"]["xlsx"], xlsx_nuevo,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 9. Generar PDF con WeasyPrint o reportlab como fallback
        pdf_bytes = _generar_pdf_cert(ws, num_cert, fecha_dt, avances, monto_cert)
        pdf_url = None
        if pdf_bytes:
            pdf_nombre = f"Cs_SDE_DURLOCK_CERT_N{num_cert}_{fecha_dt.strftime('%Y%m%d')}.pdf"
            file_id = _upload_new(DRIVE["durlock"]["folder"], pdf_nombre, pdf_bytes, "application/pdf")
            if file_id:
                pdf_url = f"https://drive.google.com/file/d/{file_id}/view"

        # 10. Respuesta
        pdf_str = f"\n📎 PDF: {pdf_url}" if pdf_url else "\n📎 PDF: no disponible en este entorno"
        respuesta = (
            f"✅ *CERT N°{num_cert} — Durlock interno*\n"
            f"📅 {fecha_dt.strftime('%d/%m/%Y')}\n\n"
            f"*Avances:*\n" + "\n".join(lineas_avance) + "\n\n"
            f"💵 Monto cert est.: *${monto_cert:,.0f}*\n"
            f"📊 Excel actualizado en Drive"
            + pdf_str
        )
        return respuesta, True

    except Exception as e:
        import traceback
        print(f"CERT ERROR: {traceback.format_exc()}")
        return f"❌ Error: {str(e)}", False


def _generar_pdf_cert(ws, num_cert, fecha_dt, avances, monto_cert):
    """Genera un PDF limpio del certificado usando reportlab."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT

        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm,
                                leftMargin=1.5*cm, rightMargin=1.5*cm)
        styles = getSampleStyleSheet()
        story = []

        # Título
        title_style = ParagraphStyle('title', parent=styles['Heading1'], fontSize=14,
                                     alignment=TA_CENTER, spaceAfter=6)
        story.append(Paragraph(f"CERTIFICADO N°{num_cert} — DURLOCK MO", title_style))
        story.append(Paragraph(f"Constitución — Santiago del Estero  |  {fecha_dt.strftime('%d/%m/%Y')}", 
                               ParagraphStyle('sub', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER)))
        story.append(Spacer(1, 0.4*cm))

        # Tabla de avances
        headers = ["PISO / TAREA", "ANT. %", "ACT. %", "ACUM. %", "IMPORTE ACT."]
        data = [headers]

        for piso, tareas_map in DURLOCK_ROW_MAP.items():
            piso_added = False
            for tarea, row in tareas_map.items():
                h_val = ws.cell(row, COL_H).value or 0
                j_val = ws.cell(row, COL_J).value or 0
                i_val = j_val - h_val  # actual = acumulado - anterior
                f_val = ws.cell(row, 6).value or 0
                imp_act = f_val * i_val if isinstance(f_val, (int, float)) else 0

                # Solo mostrar filas con algún valor
                if h_val == 0 and j_val == 0:
                    continue

                piso_label = piso if not piso_added else ""
                piso_added = True
                tarea_corta = tarea.replace("según plano", "").replace("e iluminacion", "").strip()
                data.append([
                    f"{piso_label}\n{tarea_corta}" if piso_label else tarea_corta,
                    f"{h_val*100:.1f}%",
                    f"{i_val*100:.1f}%",
                    f"{j_val*100:.1f}%",
                    f"${imp_act:,.0f}" if imp_act else "-"
                ])

        if len(data) > 1:
            col_widths = [7*cm, 2*cm, 2*cm, 2*cm, 3.5*cm]
            t = Table(data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2D3D')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F5F5F5')]),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            story.append(t)

        story.append(Spacer(1, 0.5*cm))

        # Totales
        total_style = ParagraphStyle('total', parent=styles['Normal'], fontSize=11, spaceAfter=4)
        story.append(Paragraph(f"<b>Monto certificado:</b>  ${monto_cert:,.0f}", total_style))

        # Firmas
        story.append(Spacer(1, 2*cm))
        firma_data = [["_________________________", "_________________________"],
                      ["Dirección de Obra", "Contratista"]]
        ft = Table(firma_data, colWidths=[8*cm, 8*cm])
        ft.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTSIZE', (0,0), (-1,-1), 10)]))
        story.append(ft)

        doc.build(story)
        return buf.getvalue()

    except ImportError:
        print("reportlab no disponible, intentando con fpdf2")
        return _generar_pdf_fpdf(num_cert, fecha_dt, avances, monto_cert)
    except Exception as e:
        print(f"PDF reportlab error: {e}")
        return None


def _generar_pdf_fpdf(num_cert, fecha_dt, avances, monto_cert):
    """Fallback PDF con fpdf2."""
    try:
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 10, f"CERTIFICADO N\xb0{num_cert} - DURLOCK MO", ln=True, align="C")
        pdf.set_font("Helvetica", "", 11)
        pdf.cell(0, 8, f"Santiago del Estero  |  {fecha_dt.strftime('%d/%m/%Y')}", ln=True, align="C")
        pdf.ln(5)
        for piso, tareas in avances.items():
            for tarea, pct in tareas.items():
                tarea_c = tarea.replace("según plano","").replace("e iluminacion","").strip()
                pdf.cell(0, 7, f"{piso} - {tarea_c}: {pct}%", ln=True)
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, f"Monto certificado: ${monto_cert:,.0f}", ln=True)
        return bytes(pdf.output())
    except Exception as e:
        print(f"PDF fpdf2 error: {e}")
        return None
