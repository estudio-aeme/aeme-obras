"""
cert_module.py — Módulo de certificación ObraManager
Estrategia: usa openpyxl para manipular el .xlsx de forma robusta.
Versión simplificada: solo actualiza Excel y devuelve el link (sin PDF).
"""
import os, io, json, re, requests
from datetime import datetime, date

import openpyxl

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


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de Google Drive
# ─────────────────────────────────────────────────────────────────────────────
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


# ─────────────────────────────────────────────────────────────────────────────
# Manipulación del xlsx con openpyxl
# ─────────────────────────────────────────────────────────────────────────────
def _excel_serial(dt):
    """Convierte datetime a número serial de Excel."""
    base = date(1899, 12, 30)
    return (dt.date() - base).days


def _leer_acumulados_anteriores(ws_anterior):
    """Lee los valores de la columna J (acumulado) del cert anterior por cada fila relevante."""
    prev_j = {}
    if ws_anterior is None:
        return prev_j

    for piso, tareas_map in DURLOCK_ROWS.items():
        for tarea, row_num in tareas_map.items():
            cell = ws_anterior.cell(row=row_num, column=10)  # col J = 10
            try:
                val = cell.value
                if isinstance(val, (int, float)):
                    prev_j[row_num] = float(val)
                elif isinstance(val, str) and val.strip():
                    s = val.strip().replace('%', '').replace(',', '.')
                    try:
                        num = float(s)
                        if num > 1:
                            num = num / 100.0
                        prev_j[row_num] = num
                    except ValueError:
                        prev_j[row_num] = 0.0
                else:
                    prev_j[row_num] = 0.0
            except Exception:
                prev_j[row_num] = 0.0

    return prev_j


def _aplicar_avances_a_hoja(ws, num_cert, fecha_serial, avances_por_fila):
    """
    Modifica la hoja:
    - M1: número de cert
    - M2: fecha (serial)
    - Columnas H y J de cada fila: nuevos valores
    Conserva fórmulas y estilos en el resto.
    """
    ws.cell(row=1, column=13).value = num_cert       # M1
    ws.cell(row=2, column=13).value = fecha_serial   # M2

    for row_num, (nuevo_h, nuevo_j) in avances_por_fila.items():
        ws.cell(row=row_num, column=8).value = nuevo_h    # col H = 8
        ws.cell(row=row_num, column=10).value = nuevo_j   # col J = 10


# ─────────────────────────────────────────────────────────────────────────────
# Parser de avances con LLM
# ─────────────────────────────────────────────────────────────────────────────
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


# ─────────────────────────────────────────────────────────────────────────────
# Función principal
# ─────────────────────────────────────────────────────────────────────────────
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

        # 2. Descargar xlsx desde Drive
        xlsx_bytes = _download(DRIVE["durlock"]["xlsx"])
        print(f"[DEBUG] xlsx descargado: {len(xlsx_bytes)} bytes")

        # 3. Abrir con openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        print(f"[DEBUG] Hojas en workbook: {wb.sheetnames}")

        # 4. Detectar última hoja CERT
        cert_sheets = []
        for sname in wb.sheetnames:
            m = re.match(r'^CERT N°(\d+)$', sname.strip())
            if m:
                cert_sheets.append((sname, int(m.group(1))))

        if cert_sheets:
            cert_sheets.sort(key=lambda x: x[1])
            last_cert_name, last_cert_num = cert_sheets[-1]
            num_cert = last_cert_num + 1
        else:
            last_cert_name = None
            num_cert = 1

        new_sheet_name = f"CERT N°{num_cert}"
        print(f"[DEBUG] Generando '{new_sheet_name}' basado en '{last_cert_name}'")

        # 5. Leer acumulados del cert anterior (col J)
        prev_j_values = {}
        if last_cert_name:
            ws_prev = wb[last_cert_name]
            prev_j_values = _leer_acumulados_anteriores(ws_prev)
            print(f"[DEBUG] Acumulados leidos del cert anterior: {len(prev_j_values)} filas")

        # 6. Construir mapa de cambios: {row_num: (nuevo_H, nuevo_J)}
        avances_por_fila = {}
        lineas_avance = []

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

        # Pisos NO mencionados: H = J = acumulado anterior (sin avance nuevo)
        for piso, tareas_map in DURLOCK_ROWS.items():
            for tarea, row_num in tareas_map.items():
                if row_num not in avances_por_fila:
                    prev_j = prev_j_values.get(row_num, 0.0)
                    avances_por_fila[row_num] = (prev_j, prev_j)

        # 7. Duplicar hoja con openpyxl (preserva formato, fórmulas y estilos)
        if last_cert_name:
            src_ws = wb[last_cert_name]
            new_ws = wb.copy_worksheet(src_ws)
            new_ws.title = new_sheet_name
        else:
            new_ws = wb.create_sheet(title=new_sheet_name)

        print(f"[DEBUG] Hoja duplicada: '{new_ws.title}'")

        # 8. Aplicar valores en la hoja nueva
        _aplicar_avances_a_hoja(new_ws, num_cert, fecha_serial, avances_por_fila)
        print(f"[DEBUG] Valores aplicados en {len(avances_por_fila)} filas")

        # 9. Guardar a bytes
        out = io.BytesIO()
        wb.save(out)
        xlsx_nuevo = out.getvalue()
        print(f"[DEBUG] xlsx generado: {len(xlsx_nuevo)} bytes")

        # 10. Subir xlsx al Drive (sobrescribe)
        _update_drive(DRIVE["durlock"]["xlsx"], xlsx_nuevo,
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        print(f"[DEBUG] xlsx subido a Drive OK")

        # 11. Construir link al Excel
        excel_link = f"https://docs.google.com/spreadsheets/d/{DRIVE['durlock']['xlsx']}/edit"

        # 12. Respuesta
        respuesta = (
            f"✅ *{new_sheet_name} — Durlock interno*\n"
            f"📅 {fecha_dt.strftime('%d/%m/%Y')}\n\n"
            f"*Avances:*\n" + "\n".join(lineas_avance) + "\n\n"
            f"📊 Excel actualizado:\n{excel_link}"
        )
        return respuesta, True

    except Exception as e:
        import traceback
        print(f"CERT ERROR: {traceback.format_exc()}")
        return f"❌ Error: {str(e)}", False
