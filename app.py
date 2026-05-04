import os, json, re
from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import psycopg2
import requests
from datetime import datetime

app = Flask(__name__)
DATABASE_URL = os.environ.get("DATABASE_URL")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

# ─────────────────────────────────────────────────────────────────────────────
# DB
# ─────────────────────────────────────────────────────────────────────────────
def get_db():
    return psycopg2.connect(DATABASE_URL, sslmode="require")

def init_db():
    conn = get_db(); cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS obras (
        id SERIAL PRIMARY KEY, nombre TEXT, bloque TEXT, estado TEXT DEFAULT 'activa');
    CREATE TABLE IF NOT EXISTS contratos (
        id SERIAL PRIMARY KEY, obra_id INTEGER, descripcion TEXT, proveedor TEXT,
        tipo TEXT DEFAULT 'MO', estado TEXT DEFAULT 'activo',
        presup_compra BIGINT DEFAULT 0, pagado BIGINT DEFAULT 0, cargas BIGINT DEFAULT 0,
        presup_venta BIGINT DEFAULT 0, cobrado_orig BIGINT DEFAULT 0, cac_cobrado BIGINT DEFAULT 0,
        cac_pct_est REAL DEFAULT 0.35, updated_at TIMESTAMP DEFAULT NOW());
    CREATE TABLE IF NOT EXISTS movimientos (
        id SERIAL PRIMARY KEY, contrato_id INTEGER, obra_id INTEGER,
        fecha TIMESTAMP DEFAULT NOW(), tipo TEXT, monto BIGINT, nota TEXT, usuario TEXT DEFAULT 'whatsapp');
    CREATE TABLE IF NOT EXISTS desacopios (
        id SERIAL PRIMARY KEY, obra_id INTEGER, fecha TIMESTAMP DEFAULT NOW(),
        proveedor TEXT, rubro TEXT, monto BIGINT, factura TEXT, nota TEXT);
    CREATE TABLE IF NOT EXISTS certificados_semanales (
        id SERIAL PRIMARY KEY, obra_id INTEGER, fecha DATE, descripcion TEXT,
        contrato_id INTEGER, tipo TEXT, monto BIGINT,
        incluye_cac BOOLEAN DEFAULT FALSE, monto_cac BIGINT DEFAULT 0, nota TEXT);
    """)
    conn.commit(); cur.close(); conn.close()

# ─────────────────────────────────────────────────────────────────────────────
# SEED — contratos SDE con datos reales del Drive
# ─────────────────────────────────────────────────────────────────────────────
def seed_sde():
    conn = get_db(); cur = conn.cursor()

    cur.execute("SELECT id FROM obras WHERE bloque='Santiago del Estero' LIMIT 1")
    row = cur.fetchone()
    if row:
        obra_id = row[0]
    else:
        cur.execute("INSERT INTO obras(nombre,bloque) VALUES('Constitución','Santiago del Estero') RETURNING id")
        obra_id = cur.fetchone()[0]

    cur.execute("DELETE FROM contratos WHERE obra_id=%s", (obra_id,))

    contratos = [
        # (desc, proveedor, estado, presupC, pagado, cargas, presupV, cobrado, cac, cac_pct)
        ("Albañilería",           "Joel Benitez",      "activo",    430000000, 286232463, 75878237, 1214627824, 592717300, 113850202, 0.35),
        ("Instalación eléctrica", "Pablo Fidi",        "activo",    215115672,  89652089,         0,  600000000, 281857322,   6041290, 0.35),
        ("Sanitarias",            "Ricardo Sequeira",  "activo",    219600000,  66000000,         0,  600000000, 335652282,   7759794, 0.35),
        ("Pre-instalación AA",    "Ricardo Sequeira",  "activo",            0,  47800000,         0,  119500000, 117075362,         0, 0.35),
        ("Herrería moldes balcones","Leo Gallardo",    "finalizado",  1740000,   1740000,         0,    6090000,   6090000,         0, 0.0),
        ("Durlock MO",            "Julio Cabrera",     "activo",     44000000,   3814940,         0,  119400000,  11800000,         0, 0.35),
        # Finalizados
        ("Hormigón",              "Joel Benitez",      "finalizado",180000000, 182444134, 49807249,  495000000, 495000000,         0, 0.0),
        ("Alb. adic. — obrador",  "Joel Benitez",      "finalizado",  4500000,   4500000,         0,    9620000,   9620000,         0, 0.0),
        ("Alb. adic. — protecciones","Joel Benitez",   "finalizado",  3500000,   3500000,         0,    8640000,   8640000,         0, 0.0),
        ("Alb. adic. — cerco divisor","Joel Benitez",  "finalizado",  1300000,   1300000,         0,    3500000,   3500000,         0, 0.0),
        ("Ayuda de gremios",      "Joel Benitez",      "finalizado",        0,         0,         0,   77810000,  77810000,         0, 0.0),
        ("Pintura medianeras",    "Joel Benitez",      "finalizado",  2000000,   2000000,         0,    3000000,   3000000,         0, 0.0),
        ("Revoque medianeras",    "Joaquín",           "finalizado",  6000000,   6000000,         0,   18999433,  18999433,         0, 0.0),
    ]

    for d in contratos:
        cur.execute("""
            INSERT INTO contratos
            (obra_id,descripcion,proveedor,estado,
             presup_compra,pagado,cargas,
             presup_venta,cobrado_orig,cac_cobrado,cac_pct_est)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (obra_id, d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9]))

    conn.commit(); cur.close(); conn.close()
    return f"✅ {len(contratos)} contratos SDE cargados."

# ─────────────────────────────────────────────────────────────────────────────
# CONSULTAS
# ─────────────────────────────────────────────────────────────────────────────
def get_obra_sde():
    conn = get_db(); cur = conn.cursor()
    cur.execute("SELECT id FROM obras WHERE bloque='Santiago del Estero' LIMIT 1")
    row = cur.fetchone(); cur.close(); conn.close()
    return row[0] if row else None

def resumen_obra():
    obra_id = get_obra_sde()
    if not obra_id: return "❌ Obra SDE no encontrada."
    conn = get_db(); cur = conn.cursor()
    cur.execute("""
        SELECT descripcion, estado,
               pagado+cargas as total_compra, presup_compra,
               cobrado_orig+cac_cobrado as total_venta, presup_venta,
               (cobrado_orig+cac_cobrado)-(pagado+cargas) as utilidad
        FROM contratos WHERE obra_id=%s ORDER BY estado DESC, descripcion
    """, (obra_id,))
    rows = cur.fetchall()
    cur.close(); conn.close()

    activos = [r for r in rows if r[1]=='activo']
    tot_C = sum(r[2] for r in activos)
    tot_V = sum(r[4] for r in activos)
    util  = tot_V - tot_C

    lines = ["📊 *Constitución — Santiago del Estero*\n"]
    lines.append("*Contratos activos:*")
    for r in activos:
        pct_C = int(r[2]/r[3]*100) if r[3] else 0
        pct_V = int(r[4]/r[5]*100) if r[5] else 0
        util_r = r[6] or 0
        lines.append(f"  • {r[0]}: util=${util_r/1e6:.1f}M | pago {pct_C}% | cobro {pct_V}%")
    lines.append(f"\n💰 Utilidad acum activos: *${util/1e6:.1f}M*")
    lines.append(f"📥 Total cobrado: ${tot_V/1e6:.1f}M  |  📤 Total pagado: ${tot_C/1e6:.1f}M")
    return "\n".join(lines)

def resumen_contrato(keywords):
    obra_id = get_obra_sde()
    if not obra_id: return "❌ Obra no encontrada."
    conn = get_db(); cur = conn.cursor()
    kw = f"%{keywords.lower()}%"
    cur.execute("""
        SELECT descripcion, proveedor, estado,
               presup_compra, pagado, cargas,
               presup_venta, cobrado_orig, cac_cobrado, cac_pct_est
        FROM contratos
        WHERE obra_id=%s AND LOWER(descripcion) LIKE %s
        LIMIT 1
    """, (obra_id, kw))
    r = cur.fetchone(); cur.close(); conn.close()
    if not r: return f"❌ No encontré contrato con '{keywords}'."

    desc,prov,est,pC,pag,car,pV,cob,cac,cac_pct = r
    tot_C = pag+car
    tot_V = cob+cac
    saldo_C = max(0, pC-pag) if pC else 0
    saldo_base = max(0, pV-cob)
    cac_proy = saldo_base * cac_pct
    total_cobrar = saldo_base + cac_proy
    util_act = tot_V - tot_C
    util_proy = total_cobrar - saldo_C

    return (
        f"📋 *{desc}* ({prov}) — {est.upper()}\n\n"
        f"*COMPRA* (presup: ${pC/1e6:.1f}M)\n"
        f"  Pagado: ${pag/1e6:.1f}M  Cargas: ${car/1e6:.1f}M\n"
        f"  ▸ Saldo a pagar: *${saldo_C/1e6:.1f}M*\n\n"
        f"*VENTA* (presup: ${pV/1e6:.1f}M)\n"
        f"  Cobrado: ${cob/1e6:.1f}M  CAC: ${cac/1e6:.1f}M\n"
        f"  ▸ Saldo a cobrar (base): *${saldo_base/1e6:.1f}M*\n\n"
        f"*UTILIDAD ACTUAL:* ${util_act/1e6:.1f}M\n"
        f"*PROYECCIÓN futura ({int(cac_pct*100)}% CAC):*\n"
        f"  Cobrar: ${total_cobrar/1e6:.1f}M  Pagar: ${saldo_C/1e6:.1f}M\n"
        f"  ▸ Util. proyectada: *${util_proy/1e6:.1f}M*"
    )

def saldos_proveedor(proveedor_kw):
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    kw = f"%{proveedor_kw.lower()}%"
    cur.execute("""
        SELECT descripcion, presup_compra, pagado, cargas
        FROM contratos
        WHERE obra_id=%s AND LOWER(proveedor) LIKE %s AND estado='activo'
    """, (obra_id, kw))
    rows = cur.fetchall(); cur.close(); conn.close()
    if not rows: return f"❌ No hay contratos activos para '{proveedor_kw}'."
    total_saldo = 0
    lines = [f"💳 Saldos pendientes a pagar:"]
    for r in rows:
        saldo = max(0, (r[1] or 0) - r[2])
        total_saldo += saldo
        lines.append(f"  • {r[0]}: ${saldo/1e6:.1f}M")
    lines.append(f"\n*Total: ${total_saldo/1e6:.1f}M*")
    return "\n".join(lines)

def saldo_cobrar_cliente():
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    cur.execute("""
        SELECT descripcion, presup_venta, cobrado_orig, cac_cobrado, cac_pct_est
        FROM contratos WHERE obra_id=%s AND estado='activo'
    """, (obra_id,))
    rows = cur.fetchall(); cur.close(); conn.close()
    total = 0
    lines = ["💰 Saldo por cobrar del cliente (activos):"]
    for r in rows:
        desc,pV,cob,cac,pct = r
        saldo_base = max(0, pV - cob)
        cac_proy = saldo_base * pct
        total_cobrar = saldo_base + cac_proy
        total += total_cobrar
        if total_cobrar > 0:
            lines.append(f"  • {desc}: ${total_cobrar/1e6:.1f}M (base ${saldo_base/1e6:.1f}M + CAC ${cac_proy/1e6:.1f}M)")
    lines.append(f"\n*Total a cobrar: ${total/1e6:.1f}M*")
    return "\n".join(lines)

def ultimos_certificados(n=5):
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    cur.execute("""
        SELECT fecha, descripcion, tipo, monto, incluye_cac, monto_cac
        FROM certificados_semanales WHERE obra_id=%s
        ORDER BY fecha DESC LIMIT %s
    """, (obra_id, n))
    rows = cur.fetchall(); cur.close(); conn.close()
    if not rows: return "📋 No hay certificados registrados aún."
    lines = [f"📋 Últimos {n} certificados SDE:"]
    for r in rows:
        cac_str = f" (+CAC ${r[5]/1e6:.1f}M)" if r[4] else ""
        lines.append(f"  {r[0].strftime('%d/%m')} | {r[1]} | {r[2]} | ${r[3]/1e6:.1f}M{cac_str}")
    return "\n".join(lines)

def ultimos_desacopios(n=5):
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    cur.execute("""
        SELECT fecha, proveedor, rubro, monto, nota
        FROM desacopios WHERE obra_id=%s
        ORDER BY fecha DESC LIMIT %s
    """, (obra_id, n))
    rows = cur.fetchall(); cur.close(); conn.close()
    if not rows: return "📦 No hay desacopios registrados aún."
    lines = [f"📦 Últimos {n} desacopios SDE:"]
    total = 0
    for r in rows:
        rubro = f" ({r[2]})" if r[2] else ""
        lines.append(f"  {r[0].strftime('%d/%m')} | {r[1]}{rubro} | ${r[3]/1e6:.1f}M")
        total += r[3]
    lines.append(f"*Total: ${total/1e6:.1f}M*")
    return "\n".join(lines)

# ─────────────────────────────────────────────────────────────────────────────
# REGISTRO DE MOVIMIENTOS
# ─────────────────────────────────────────────────────────────────────────────
def registrar_movimiento(tipo, monto, contrato_kw, nota=""):
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    kw = f"%{contrato_kw.lower()}%"
    cur.execute("SELECT id,descripcion,pagado,cobrado_orig,cac_cobrado,cargas FROM contratos WHERE obra_id=%s AND LOWER(descripcion) LIKE %s LIMIT 1", (obra_id, kw))
    c = cur.fetchone()
    if not c:
        cur.close(); conn.close()
        return None, f"❌ No encontré contrato con '{contrato_kw}'."

    cid, desc, pagado, cobrado, cac, cargas = c

    if tipo == "pago_proveedor":
        cur.execute("UPDATE contratos SET pagado=pagado+%s, updated_at=NOW() WHERE id=%s", (monto, cid))
        cur.execute("INSERT INTO movimientos(contrato_id,obra_id,tipo,monto,nota) VALUES(%s,%s,%s,%s,%s)", (cid,obra_id,"pago_proveedor",monto,nota))
        nuevo = pagado + monto
        msg = f"✅ Pago registrado\n📋 {desc}\n💸 ${monto/1e6:.2f}M\n📊 Total pagado: ${nuevo/1e6:.2f}M"

    elif tipo == "cobro_cliente":
        cur.execute("UPDATE contratos SET cobrado_orig=cobrado_orig+%s, updated_at=NOW() WHERE id=%s", (monto, cid))
        cur.execute("INSERT INTO movimientos(contrato_id,obra_id,tipo,monto,nota) VALUES(%s,%s,%s,%s,%s)", (cid,obra_id,"cobro_cliente",monto,nota))
        nuevo = cobrado + monto
        msg = f"✅ Cobro registrado\n📋 {desc}\n💰 ${monto/1e6:.2f}M\n📊 Total cobrado: ${nuevo/1e6:.2f}M"

    elif tipo == "cac_cobrado":
        cur.execute("UPDATE contratos SET cac_cobrado=cac_cobrado+%s, updated_at=NOW() WHERE id=%s", (monto, cid))
        cur.execute("INSERT INTO movimientos(contrato_id,obra_id,tipo,monto,nota) VALUES(%s,%s,%s,%s,%s)", (cid,obra_id,"cac_cobrado",monto,nota))
        nuevo = cac + monto
        msg = f"✅ CAC registrado\n📋 {desc}\n📈 ${monto/1e6:.2f}M\n📊 CAC acum: ${nuevo/1e6:.2f}M"

    conn.commit(); cur.close(); conn.close()
    return cid, msg

def registrar_desacopio(proveedor, monto, rubro="", nota="", factura=""):
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    cur.execute("INSERT INTO desacopios(obra_id,proveedor,rubro,monto,factura,nota) VALUES(%s,%s,%s,%s,%s,%s)",
                (obra_id, proveedor, rubro, monto, factura, nota))
    conn.commit(); cur.close(); conn.close()
    rub_str = f" ({rubro})" if rubro else ""
    return f"📦 Desacopio registrado\n🏪 {proveedor}{rub_str}\n💵 ${monto/1e6:.2f}M"

def registrar_certificado(fecha_str, descripcion, monto, tipo, contrato_kw="", incluye_cac=False, monto_cac=0):
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    try:
        fecha = datetime.strptime(fecha_str, "%d/%m/%Y").date()
    except:
        fecha = datetime.now().date()
    contrato_id = None
    if contrato_kw:
        kw = f"%{contrato_kw.lower()}%"
        cur.execute("SELECT id FROM contratos WHERE obra_id=%s AND LOWER(descripcion) LIKE %s LIMIT 1", (obra_id, kw))
        row = cur.fetchone()
        if row: contrato_id = row[0]
    cur.execute("""
        INSERT INTO certificados_semanales(obra_id,fecha,descripcion,contrato_id,tipo,monto,incluye_cac,monto_cac)
        VALUES(%s,%s,%s,%s,%s,%s,%s,%s)
    """, (obra_id, fecha, descripcion, contrato_id, tipo, monto, incluye_cac, monto_cac))
    conn.commit(); cur.close(); conn.close()
    return f"📋 Certificado registrado\n📅 {fecha}\n📌 {descripcion} | {tipo}\n💵 ${monto/1e6:.2f}M"

# ─────────────────────────────────────────────────────────────────────────────
# IA — PROCESAMIENTO DE INTENCIÓN
# ─────────────────────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """Sos el asistente de ObraManager para la constructora Aeme Obras.
Procesás mensajes de WhatsApp sobre la obra Constitución — Santiago del Estero (SDE).

INTENCIONES posibles:
- "resumen_obra": pide resumen general, utilidad, cómo va la obra
- "consulta_contrato": consulta un contrato específico. Extraé "contrato" (keywords)
- "saldo_cobrar": quiere saber cuánto le falta cobrar al cliente
- "saldo_pagar": quiere saber cuánto le falta pagar a un proveedor. Extraé "proveedor"
- "certificados": últimos certificados semanales registrados
- "desacopios": últimos desacopios de materiales
- "registrar_pago": pago a proveedor. Extraé "monto" (número), "contrato" (keywords), "nota"
- "registrar_cobro": cobro del cliente. Extraé "monto", "contrato", "nota"
- "registrar_cac": CAC cobrado al cliente. Extraé "monto", "contrato"
- "registrar_desacopio": compra de materiales. Extraé "proveedor", "monto", "rubro", "factura", "nota"
- "registrar_certificado": cert semanal. Extraé "fecha"(dd/mm/aaaa), "descripcion", "monto", "tipo"(compra/venta), "contrato", "incluye_cac", "monto_cac"
- "confirmacion": el usuario confirma con "sí", "si", "dale", "confirmo", "ok"
- "cancelacion": el usuario cancela con "no", "cancelar", "no registres"
- "otro": cualquier otra cosa

Respondé SIEMPRE con JSON válido, sin texto antes ni después:
{"intencion": "registrar_pago", "monto": 5000000, "contrato": "albañilería", "nota": "cert N°73"}
{"intencion": "registrar_cobro", "monto": 12000000, "contrato": "electricidad", "nota": ""}
{"intencion": "registrar_cac", "monto": 1200000, "contrato": "sanitarias"}
{"intencion": "registrar_desacopio", "proveedor": "NOVA", "monto": 3500000, "rubro": "materiales eléctricos", "factura": "0001-00001234", "nota": ""}
{"intencion": "registrar_certificado", "fecha": "29/04/2026", "descripcion": "Albañilería N°72", "monto": 41174238, "tipo": "compra", "contrato": "albañilería", "incluye_cac": false, "monto_cac": 0}
{"intencion": "consulta_contrato", "contrato": "electricidad"}
{"intencion": "saldo_pagar", "proveedor": "Fidi"}
{"intencion": "resumen_obra"}
{"intencion": "confirmacion"}
{"intencion": "otro", "respuesta": "No entendí. Podés decirme: 'Pagué $5M albañilería', 'Cobré $8M electricidad', 'Desacopio NOVA $3M materiales', 'Cómo va la obra'"}

Extraé montos siempre como número entero (sin signos ni puntos).
Hablá en español rioplatense."""

# Estado temporal de confirmaciones pendientes
pendientes = {}

def procesar_con_ia(mensaje):
    headers = {"Content-Type": "application/json",
               "x-api-key": ANTHROPIC_API_KEY,
               "anthropic-version": "2023-06-01"}
    body = {"model": "claude-sonnet-4-20250514", "max_tokens": 400,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": mensaje}]}
    r = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=body)
    if r.status_code != 200:
        return {"intencion": "otro", "respuesta": "Error al procesar el mensaje."}
    text = r.json()["content"][0]["text"].strip()
    text = re.sub(r"^```json|^```|```$", "", text, flags=re.MULTILINE).strip()
    try:
        return json.loads(text)
    except:
        return {"intencion": "otro", "respuesta": text}

def monto_fmt(m):
    if m >= 1_000_000: return f"${m/1_000_000:.2f}M"
    return f"${m:,}"

def procesar_mensaje(remitente, mensaje):
    global pendientes
    d = procesar_con_ia(mensaje)
    intent = d.get("intencion", "otro")

    # ── Confirmación de operación pendiente ───────────────────────────────────
    if intent == "confirmacion" and remitente in pendientes:
        op = pendientes.pop(remitente)
        if op["tipo"] == "pago":
            _, msg = registrar_movimiento("pago_proveedor", op["monto"], op["contrato"], op.get("nota",""))
        elif op["tipo"] == "cobro":
            _, msg = registrar_movimiento("cobro_cliente", op["monto"], op["contrato"], op.get("nota",""))
        elif op["tipo"] == "cac":
            _, msg = registrar_movimiento("cac_cobrado", op["monto"], op["contrato"])
        elif op["tipo"] == "desacopio":
            msg = registrar_desacopio(op["proveedor"], op["monto"], op.get("rubro",""), op.get("nota",""), op.get("factura",""))
        elif op["tipo"] == "certificado":
            msg = registrar_certificado(op["fecha"], op["descripcion"], op["monto"], op["tipo_cert"], op.get("contrato",""), op.get("incluye_cac",False), op.get("monto_cac",0))
        else:
            msg = "✅ Registrado."
        return msg

    if intent == "cancelacion":
        pendientes.pop(remitente, None)
        return "❌ Cancelado. ¿Qué más necesitás?"

    # ── Consultas ─────────────────────────────────────────────────────────────
    if intent == "resumen_obra":
        return resumen_obra()

    if intent == "consulta_contrato":
        return resumen_contrato(d.get("contrato",""))

    if intent == "saldo_cobrar":
        return saldo_cobrar_cliente()

    if intent == "saldo_pagar":
        return saldos_proveedor(d.get("proveedor",""))

    if intent == "certificados":
        return ultimos_certificados()

    if intent == "desacopios":
        return ultimos_desacopios()

    # ── Registros con confirmación previa ────────────────────────────────────
    if intent == "registrar_pago":
        monto = d.get("monto",0); contrato = d.get("contrato",""); nota = d.get("nota","")
        pendientes[remitente] = {"tipo":"pago","monto":monto,"contrato":contrato,"nota":nota}
        return f"❓ Confirmás el pago de *{monto_fmt(monto)}* al contrato *{contrato}*?\n(Respondé *sí* para confirmar o *no* para cancelar)"

    if intent == "registrar_cobro":
        monto = d.get("monto",0); contrato = d.get("contrato",""); nota = d.get("nota","")
        pendientes[remitente] = {"tipo":"cobro","monto":monto,"contrato":contrato,"nota":nota}
        return f"❓ Confirmás el cobro de *{monto_fmt(monto)}* del contrato *{contrato}*?\n(Respondé *sí* para confirmar)"

    if intent == "registrar_cac":
        monto = d.get("monto",0); contrato = d.get("contrato","")
        pendientes[remitente] = {"tipo":"cac","monto":monto,"contrato":contrato}
        return f"❓ Confirmás CAC cobrado de *{monto_fmt(monto)}* en *{contrato}*?\n(Respondé *sí* para confirmar)"

    if intent == "registrar_desacopio":
        monto = d.get("monto",0); prov = d.get("proveedor",""); rubro = d.get("rubro","")
        pendientes[remitente] = {"tipo":"desacopio","monto":monto,"proveedor":prov,"rubro":rubro,"nota":d.get("nota",""),"factura":d.get("factura","")}
        rub_str = f" ({rubro})" if rubro else ""
        return f"❓ Confirmás desacopio en *{prov}*{rub_str} por *{monto_fmt(monto)}*?\n(Respondé *sí* para confirmar)"

    if intent == "registrar_certificado":
        monto = d.get("monto",0); desc = d.get("descripcion",""); fecha = d.get("fecha","hoy")
        tipo_cert = d.get("tipo","compra"); contrato = d.get("contrato","")
        incluye_cac = d.get("incluye_cac", False); monto_cac = d.get("monto_cac",0)
        pendientes[remitente] = {"tipo":"certificado","monto":monto,"descripcion":desc,"fecha":fecha,"tipo_cert":tipo_cert,"contrato":contrato,"incluye_cac":incluye_cac,"monto_cac":monto_cac}
        cac_str = f" + CAC ${monto_cac/1e6:.2f}M" if incluye_cac else ""
        return f"❓ Confirmás certificado *{desc}* del {fecha} por *{monto_fmt(monto)}*{cac_str} ({tipo_cert})?\n(Respondé *sí* para confirmar)"

    # ── Otro ──────────────────────────────────────────────────────────────────
    return d.get("respuesta", "No entendí. Podés preguntarme:\n• *Cómo va la obra*\n• *Saldo electricidad*\n• *Pagué $5M albañilería*\n• *Cobré $8M sanitarias*\n• *Desacopio NOVA $3M materiales eléctricos*\n• *Cert albañilería N°72 $41M*")

# ─────────────────────────────────────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    mensaje  = request.form.get("Body", "")
    remitente = request.form.get("From", "unknown")
    respuesta = procesar_mensaje(remitente, mensaje)
    resp = MessagingResponse()
    resp.message(respuesta)
    return str(resp)

@app.route("/")
def home():
    return "🏗️ ObraManager SDE — activo"

@app.route("/init")
def route_init():
    try:
        init_db()
        return "✅ Tablas creadas."
    except Exception as e:
        return f"❌ {e}"

@app.route("/seed")
def route_seed():
    try:
        return seed_sde()
    except Exception as e:
        return f"❌ {e}"

@app.route("/resumen")
def route_resumen():
    return resumen_obra().replace("\n","<br>")

@app.route("/datos")
def route_datos():
    obra_id = get_obra_sde()
    conn = get_db(); cur = conn.cursor()
    cur.execute("SELECT descripcion,proveedor,estado,pagado+cargas,cobrado_orig+cac_cobrado FROM contratos WHERE obra_id=%s ORDER BY estado,descripcion", (obra_id,))
    rows = cur.fetchall(); cur.close(); conn.close()
    out = "<table border=1 cellpadding=5>"
    out += "<tr><th>Contrato</th><th>Proveedor</th><th>Estado</th><th>Erogado</th><th>Cobrado</th><th>Utilidad</th></tr>"
    for r in rows:
        util = (r[4] or 0)-(r[3] or 0)
        out += f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>${r[3]/1e6:.1f}M</td><td>${r[4]/1e6:.1f}M</td><td>${util/1e6:.1f}M</td></tr>"
    out += "</table>"
    return out

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
