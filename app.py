from flask import Flask, request, jsonify
from twilio.twiml.messaging_response import MessagingResponse
import os
import json
import requests
from datetime import datetime
import pg8000.native

app = Flask(__name__)
DATABASE_URL = os.environ.get("DATABASE_URL", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# Estado de conversación en memoria (para confirmaciones pendientes)
pendientes = {}

def parse_db_url(url):
    url = url.replace("postgresql://", "")
    user_pass, rest = url.split("@")
    user, password = user_pass.split(":")
    host_db = rest.split("/")
    return user, password, host_db[0], host_db[1]

def get_db():
    user, password, host, db = parse_db_url(DATABASE_URL)
    return pg8000.native.Connection(user, password=password, host=host, database=db, ssl_context=True)

def init_db():
    conn = get_db()
    # Tabla movimientos generales
    conn.run("""CREATE TABLE IF NOT EXISTS movimientos (
        id SERIAL PRIMARY KEY, fecha TIMESTAMP DEFAULT NOW(),
        tipo VARCHAR(20), obra VARCHAR(100), monto NUMERIC(15,2),
        descripcion TEXT, proveedor VARCHAR(100), remitente VARCHAR(50))""")
    # Tabla acopios generales
    conn.run("""CREATE TABLE IF NOT EXISTS acopios (
        id SERIAL PRIMARY KEY, fecha TIMESTAMP DEFAULT NOW(),
        material VARCHAR(100), cantidad NUMERIC(10,2), unidad VARCHAR(20),
        obra VARCHAR(100), tipo VARCHAR(20) DEFAULT 'ingreso', remitente VARCHAR(50))""")
    # Tabla contratos de COMPRA (internos)
    conn.run("""CREATE TABLE IF NOT EXISTS contratos_compra (
        id SERIAL PRIMARY KEY,
        obra VARCHAR(100),
        bloque VARCHAR(20),
        descripcion VARCHAR(200),
        proveedor VARCHAR(100),
        presupuesto NUMERIC(15,2),
        pagado NUMERIC(15,2) DEFAULT 0,
        cargas NUMERIC(15,2) DEFAULT 0,
        activo BOOLEAN DEFAULT TRUE)""")
    # Tabla contratos de VENTA (cliente)
    conn.run("""CREATE TABLE IF NOT EXISTS contratos_venta (
        id SERIAL PRIMARY KEY,
        obra VARCHAR(100),
        bloque VARCHAR(20),
        descripcion VARCHAR(200),
        presupuesto NUMERIC(15,2),
        cobrado NUMERIC(15,2) DEFAULT 0,
        cobrado_cac NUMERIC(15,2) DEFAULT 0,
        activo BOOLEAN DEFAULT TRUE)""")
    # Tabla movimientos de contratos
    conn.run("""CREATE TABLE IF NOT EXISTS movimientos_contratos (
        id SERIAL PRIMARY KEY,
        fecha TIMESTAMP DEFAULT NOW(),
        tipo VARCHAR(10),
        contrato_id INTEGER,
        obra VARCHAR(100),
        bloque VARCHAR(20),
        descripcion TEXT,
        proveedor VARCHAR(100),
        monto NUMERIC(15,2),
        es_cac BOOLEAN DEFAULT FALSE,
        remitente VARCHAR(50),
        observacion TEXT)""")
    conn.close()

def cargar_contratos_iniciales():
    """Carga los contratos de Constitución si no existen"""
    conn = get_db()
    count = conn.run("SELECT COUNT(*) FROM contratos_compra")[0][0]
    if count == 0:
        # CONTRATOS COMPRA — SAN JOSÉ
        compras_sj = [
            ("Constitución","San José","ALBAÑILERÍA","Victor Rolon / Miguel Brandell",104339494,119970000,0),
            ("Constitución","San José","INSTALACIÓN ELÉCTRICA","Pablo Fidi",49555648,43900000,0),
            ("Constitución","San José","CIELORRASOS DURLOCK","Hugo Esquivel",14500000,13719768,228375),
            ("Constitución","San José","INSTALACIÓN SANITARIA","Ricardo Sequeira",77500000,77500000,0),
            ("Constitución","San José","IMPERMEABILIZACIÓN","Edgar",8878000,8905000,0),
            ("Constitución","San José","COLOCACIÓN REVESTIMIENTOS","Miguel Brandell",17826419,16793419,0),
            ("Constitución","San José","PINTURA","Luis Contreras",26800000,32960000,0),
            ("Constitución","San José","PICADO SUBSUELO + FUNDACIONES","Leandro Ortega",10356000,10356000,952),
            ("Constitución","San José","PRE-INSTALACIÓN AA","Jean Pierre",22176000,22180000,0),
            ("Constitución","San José","HORMIGÓN SUBSUELO","Juan Soloa",1750000,1700000,50000),
            ("Constitución","San José","ALBAÑILERÍA ADICIONAL","Joel Benitez",105723781,124665574,0),
            ("Constitución","San José","PLANTA BAJA VARIOS","Leandro Ortega",6181000,6180000,1000),
            ("Constitución","San José","HERRERÍA GENERAL","Leonardo Gallardo",26339640,26566640,0),
            ("Constitución","San José","HERRERÍA AZOTEA","Leonardo Gallardo",1948500,1948500,0),
            ("Constitución","San José","DURLOCK ADICIONAL","Ever + Julio",2275000,3895000,0),
            ("Constitución","San José","ILUMINACIÓN SUBSUELO","Pablo Fidi",1114350,1114350,0),
            ("Constitución","San José","PORTONES DE CHAPA","Sergio",11150000,11150000,0),
            ("Constitución","San José","COLOCACIÓN INTERTRABADOS","Flores + Joel Benitez",0,1800000,0),
            ("Constitución","San José","MEDIANERAS JARDÍN","Joaquín",2000000,2000000,0),
        ]
        # CONTRATOS COMPRA — SDE
        compras_sde = [
            ("Constitución","Santiago del Estero","LIMPIEZA + PROTECCIONES","Miguel Brandell",2700000,2700000,0),
            ("Constitución","Santiago del Estero","HORMIGÓN","Joel Benitez",180000000,182444134,9542503),
            ("Constitución","Santiago del Estero","BOCAS EN LOSAS","Pablo Fidi",13275000,13275000,0),
            ("Constitución","Santiago del Estero","ALBAÑILERÍA","Joel Benitez",430000000,286232463,75878237),
            ("Constitución","Santiago del Estero","VESTUARIOS DE OBRA","Sequeira + Joel Benitez",0,8300000,0),
            ("Constitución","Santiago del Estero","AYUDA DE GREMIOS GRAL","Joel Benitez",0,8000000,0),
            ("Constitución","Santiago del Estero","PISO HORMIGÓN SUBSUELO","Cristian Mereles",0,5430000,0),
        ]
        for r in compras_sj + compras_sde:
            conn.run("INSERT INTO contratos_compra (obra,bloque,descripcion,proveedor,presupuesto,pagado,cargas) VALUES (:o,:b,:d,:p,:pr,:pa,:c)",
                o=r[0],b=r[1],d=r[2],p=r[3],pr=r[4],pa=r[5],c=r[6])

    count_v = conn.run("SELECT COUNT(*) FROM contratos_venta")[0][0]
    if count_v == 0:
        # CONTRATOS VENTA — SAN JOSÉ
        ventas_sj = [
            ("Constitución","San José","ALBAÑILERÍA",230000000,230000000,0),
            ("Constitución","San José","INSTALACIÓN ELÉCTRICA",94061982,94061982,0),
            ("Constitución","San José","CIELORRASOS DURLOCK",33450000,33450000,0),
            ("Constitución","San José","INSTALACIÓN SANITARIA",147000000,145482581,0),
            ("Constitución","San José","IMPERMEABILIZACIÓN",18873010,18873010,0),
            ("Constitución","San José","COLOCACIÓN REVESTIMIENTOS",35722618,35722618,0),
            ("Constitución","San José","PINTURA",56413643,56413643,0),
            ("Constitución","San José","PICADO SUBSUELO + FUNDACIONES",14850000,14850000,0),
            ("Constitución","San José","PRE-INSTALACIÓN AA",44263700,44263700,0),
            ("Constitución","San José","ALBAÑILERÍA ADICIONAL",93190000,93190000,0),
            ("Constitución","San José","HERRERÍA GENERAL",47400000,42660000,0),
            ("Constitución","San José","HERRERÍA AZOTEA",4970000,4970000,0),
            ("Constitución","San José","DURLOCK ADICIONAL",16600000,16600000,0),
            ("Constitución","San José","ILUMINACIÓN SUBSUELO",6570870,6570870,0),
            ("Constitución","San José","PORTONES DE CHAPA",17250000,17250000,0),
            ("Constitución","San José","ADICIONAL ELÉCTRICO",20277475,20277475,0),
            ("Constitución","San José","ADICIONAL PINTURA",11804000,11804000,0),
            ("Constitución","San José","MEDIANERAS JARDÍN",9499717,9499717,0),
        ]
        # CONTRATOS VENTA — SDE
        ventas_sde = [
            ("Constitución","Santiago del Estero","LIMPIEZA + PROTECCIONES",5500000,5500000,0),
            ("Constitución","Santiago del Estero","HORMIGÓN",495400000,495400000,0),
            ("Constitución","Santiago del Estero","BOCAS EN LOSAS",28000000,28000000,0),
            ("Constitución","Santiago del Estero","ALBAÑILERÍA",1214627824,592717300,113850202),
            ("Constitución","Santiago del Estero","ILUMINACIÓN DE OBRA",12190000,12190000,0),
            ("Constitución","Santiago del Estero","VESTUARIOS DE OBRA",9620000,9620000,0),
            ("Constitución","Santiago del Estero","AYUDA DE GREMIOS GRAL",8760000,8760000,0),
            ("Constitución","Santiago del Estero","AYUDA DE GREMIOS SEMANAL",11680000,11680000,0),
        ]
        for r in ventas_sj + ventas_sde:
            conn.run("INSERT INTO contratos_venta (obra,bloque,descripcion,presupuesto,cobrado,cobrado_cac) VALUES (:o,:b,:d,:pr,:c,:cac)",
                o=r[0],b=r[1],d=r[2],pr=r[3],c=r[4],cac=r[5])
    conn.close()

def fmt_ars(n):
    return f"${float(n):,.0f}".replace(",",".")

def buscar_contrato_compra(texto):
    """Busca contrato de compra por descripción o proveedor"""
    conn = get_db()
    rows = conn.run("SELECT id, obra, bloque, descripcion, proveedor, presupuesto, pagado, cargas FROM contratos_compra WHERE activo=TRUE ORDER BY id")
    conn.close()
    texto_lower = texto.lower()
    mejor = None
    mejor_score = 0
    for row in rows:
        id_, obra, bloque, desc, prov, ppto, pagado, cargas = row
        score = 0
        for palabra in texto_lower.split():
            if len(palabra) > 3:
                if palabra in desc.lower(): score += 3
                if palabra in prov.lower(): score += 2
                if palabra in bloque.lower(): score += 1
        if score > mejor_score:
            mejor_score = score
            mejor = row
    return mejor if mejor_score >= 3 else None

def buscar_contrato_venta(texto):
    """Busca contrato de venta por descripción"""
    conn = get_db()
    rows = conn.run("SELECT id, obra, bloque, descripcion, presupuesto, cobrado, cobrado_cac FROM contratos_venta WHERE activo=TRUE ORDER BY id")
    conn.close()
    texto_lower = texto.lower()
    mejor = None
    mejor_score = 0
    for row in rows:
        id_, obra, bloque, desc, ppto, cobrado, cob_cac = row
        score = 0
        for palabra in texto_lower.split():
            if len(palabra) > 3:
                if palabra in desc.lower(): score += 3
                if palabra in bloque.lower(): score += 1
        if score > mejor_score:
            mejor_score = score
            mejor = row
    return mejor if mejor_score >= 3 else None

SYSTEM_PROMPT = """Sos el asistente de ObraManager para la constructora Aeme Obras (Julián y Julieta).

Manejás contratos de COMPRA (pagos a proveedores) y VENTA (cobros del cliente).

Cuando el usuario quiere registrar un pago o cobro sobre un contrato, extraé:
- tipo: "pago_proveedor" o "cobro_cliente"
- monto: número
- contrato: palabras clave del contrato
- es_cac: true si menciona CAC/ajuste/índice

Cuando el usuario pide listar contratos:
- solo_pendientes: true si dice "activos", "vigentes", "con saldo", "pendientes", "abiertos", "en curso"
- tipo: "compra" si dice compra/proveedores/pagos/interno. "venta" si dice venta/cliente/cobros. Si no especifica, "compra"
- bloque: "Santiago del Estero" si menciona SDE/Santiago/Santiago del Estero. "San José" si menciona San Jose/SJ. null si no especifica bloque

Respondé SIEMPRE en JSON:
{"intencion": "desacopio", "tipo": "pago_proveedor", "monto": 5000000, "contrato": "albañilería SDE", "es_cac": false}
{"intencion": "desacopio", "tipo": "cobro_cliente", "monto": 3000000, "contrato": "hormigón", "es_cac": false}
{"intencion": "consulta_contrato", "contrato": "albañilería SDE"}
{"intencion": "listar_contratos", "tipo": "compra", "solo_pendientes": true, "bloque": "Santiago del Estero"}
{"intencion": "listar_contratos", "tipo": "venta", "solo_pendientes": false, "bloque": null}
{"intencion": "otro", "respuesta": "texto de respuesta directa"}

Para consultas generales respondé con intencion "otro".
Hablá en español rioplatense."""

def procesar_mensaje(mensaje, remitente):
    # ¿Hay confirmación pendiente?
    if remitente in pendientes:
        pendiente = pendientes[remitente]
        if mensaje.strip().upper() in ["SI", "SÍ", "S", "YES", "OK", "CONFIRMAR"]:
            # Ejecutar acción pendiente
            resultado = ejecutar_pendiente(pendiente, remitente)
            del pendientes[remitente]
            return resultado
        elif mensaje.strip().upper() in ["NO", "N", "CANCELAR", "CANCEL"]:
            del pendientes[remitente]
            return "❌ Cancelado. ¿En qué más puedo ayudarte?"
        else:
            return f"Respondé *SI* para confirmar o *NO* para cancelar.\n\n{pendiente['resumen']}"

    # Procesar con IA
    try:
        headers = {
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        }
        payload = {
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 300,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": mensaje}]
        }
        response = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=payload, timeout=15)
        
        if response.status_code != 200:
            return "Error al procesar el mensaje."
        
        texto = response.json()["content"][0]["text"].strip()
        
        # Parsear JSON
        try:
            if "{" in texto:
                inicio = texto.index("{")
                fin = texto.rindex("}") + 1
                data = json.loads(texto[inicio:fin])
            else:
                return texto
        except:
            return texto

        intencion = data.get("intencion", "otro")

        # ── DESACOPIO ──
        if intencion == "desacopio":
            tipo = data.get("tipo")
            monto = float(data.get("monto", 0))
            contrato_texto = data.get("contrato", "")
            es_cac = data.get("es_cac", False)

            if tipo == "pago_proveedor":
                contrato = buscar_contrato_compra(contrato_texto + " " + mensaje)
                if not contrato:
                    return f"⚠️ No encontré el contrato de COMPRA para '{contrato_texto}'. Podés decirme el nombre exacto o pedirme que liste los contratos."
                id_, obra, bloque, desc, prov, ppto, pagado, cargas = contrato
                nuevo_pagado = float(pagado) + monto
                pendiente_pago = float(ppto) - nuevo_pagado
                resumen = (f"💳 *PAGO A PROVEEDOR*\n"
                          f"Contrato: {desc}\n"
                          f"Proveedor: {prov}\n"
                          f"Bloque: {bloque}\n"
                          f"Monto: {fmt_ars(monto)}\n"
                          f"Pagado hasta ahora: {fmt_ars(pagado)}\n"
                          f"Nuevo total pagado: {fmt_ars(nuevo_pagado)}\n"
                          f"Pendiente ppto: {fmt_ars(pendiente_pago)}\n\n"
                          f"¿Confirmás? Respondé *SI* o *NO*")
                pendientes[remitente] = {
                    "tipo": "pago_proveedor",
                    "contrato_id": id_,
                    "monto": monto,
                    "desc": desc,
                    "prov": prov,
                    "bloque": bloque,
                    "obra": obra,
                    "es_cac": es_cac,
                    "resumen": resumen
                }
                return resumen

            elif tipo == "cobro_cliente":
                contrato = buscar_contrato_venta(contrato_texto + " " + mensaje)
                if not contrato:
                    return f"⚠️ No encontré el contrato de VENTA para '{contrato_texto}'. Pedime que liste los contratos."
                id_, obra, bloque, desc, ppto, cobrado, cob_cac = contrato
                nuevo_cobrado = float(cobrado) + (0 if es_cac else monto)
                nuevo_cac = float(cob_cac) + (monto if es_cac else 0)
                total = nuevo_cobrado + nuevo_cac
                pendiente_cobro = float(ppto) - total
                pct = (total / float(ppto) * 100) if float(ppto) > 0 else 0
                resumen = (f"💰 *COBRO DEL CLIENTE*\n"
                          f"Contrato: {desc}\n"
                          f"Bloque: {bloque}\n"
                          f"Monto: {fmt_ars(monto)}{' (CAC)' if es_cac else ''}\n"
                          f"Cobrado hasta ahora: {fmt_ars(cobrado + cob_cac)}\n"
                          f"Nuevo total cobrado: {fmt_ars(total)}\n"
                          f"Avance: {pct:.1f}% del presupuesto\n"
                          f"Pendiente: {fmt_ars(pendiente_cobro)}\n\n"
                          f"¿Confirmás? Respondé *SI* o *NO*")
                pendientes[remitente] = {
                    "tipo": "cobro_cliente",
                    "contrato_id": id_,
                    "monto": monto,
                    "desc": desc,
                    "bloque": bloque,
                    "obra": obra,
                    "es_cac": es_cac,
                    "resumen": resumen
                }
                return resumen

        # ── CONSULTA CONTRATO ──
        elif intencion == "consulta_contrato":
            contrato_texto = data.get("contrato", "")
            compra = buscar_contrato_compra(contrato_texto + " " + mensaje)
            venta = buscar_contrato_venta(contrato_texto + " " + mensaje)
            resp = ""
            if compra:
                id_, obra, bloque, desc, prov, ppto, pagado, cargas = compra
                pendiente = float(ppto) - float(pagado)
                pct = (float(pagado)/float(ppto)*100) if float(ppto)>0 else 0
                resp += f"🔵 *COMPRA — {desc}* ({bloque})\nProv: {prov}\nPpto: {fmt_ars(ppto)}\nPagado: {fmt_ars(pagado)} ({pct:.0f}%)\nPendiente: {fmt_ars(pendiente)}\n\n"
            if venta:
                id_, obra, bloque, desc, ppto, cobrado, cob_cac = venta
                total = float(cobrado)+float(cob_cac)
                pendiente = float(ppto) - total
                pct = (total/float(ppto)*100) if float(ppto)>0 else 0
                resp += f"🟢 *VENTA — {desc}* ({bloque})\nPpto cliente: {fmt_ars(ppto)}\nCobrado: {fmt_ars(total)} ({pct:.0f}%)\nPendiente: {fmt_ars(pendiente)}"
            return resp if resp else "No encontré contratos para esa búsqueda."

        # ── LISTAR CONTRATOS ──
        elif intencion == "listar_contratos":
            tipo_lista = data.get("tipo", "compra")
            msg_low = mensaje.lower()

            # Filtro de bloque — detectado del texto directamente
            bloque_filtro = None
            if any(w in msg_low for w in ["sde","santiago del estero","santiago"]):
                bloque_filtro = "Santiago del Estero"
            elif any(w in msg_low for w in ["san jose","san josé","sj"]):
                bloque_filtro = "San José"

            # Filtro solo pendientes
            solo_pendientes = data.get("solo_pendientes", False) or any(w in msg_low for w in ["activo","vigente","pendiente","saldo","abierto","en curso"])

            conn = get_db()
            if tipo_lista == "compra":
                if bloque_filtro:
                    rows = conn.run("SELECT bloque, descripcion, proveedor, presupuesto, pagado FROM contratos_compra WHERE activo=TRUE AND bloque=:b ORDER BY id", b=bloque_filtro)
                else:
                    rows = conn.run("SELECT bloque, descripcion, proveedor, presupuesto, pagado FROM contratos_compra WHERE activo=TRUE ORDER BY bloque, id")
                conn.close()
                titulo = f" — {bloque_filtro}" if bloque_filtro else ""
                resp = f"🔵 *COMPRA{titulo}" + (" CON SALDO*
" if solo_pendientes else "*
")
                bloque_ant = ""
                n = 0
                for row in rows:
                    bl, desc, prov, ppto, pagado = row
                    pendiente = float(ppto) - float(pagado)
                    pct = (float(pagado)/float(ppto)*100) if float(ppto)>0 else 0
                    if solo_pendientes and pendiente <= 0:
                        continue
                    if not bloque_filtro and bl != bloque_ant:
                        resp += f"
📍 *{bl}*
"
                        bloque_ant = bl
                    resp += f"  • {desc} ({prov[:18]})
    {pct:.0f}% pagado | Pendiente: {fmt_ars(pendiente)}
"
                    n += 1
                if n == 0:
                    resp += "
No hay contratos con saldo pendiente."
            else:
                if bloque_filtro:
                    rows = conn.run("SELECT bloque, descripcion, presupuesto, cobrado, cobrado_cac FROM contratos_venta WHERE activo=TRUE AND bloque=:b ORDER BY id", b=bloque_filtro)
                else:
                    rows = conn.run("SELECT bloque, descripcion, presupuesto, cobrado, cobrado_cac FROM contratos_venta WHERE activo=TRUE ORDER BY bloque, id")
                conn.close()
                titulo = f" — {bloque_filtro}" if bloque_filtro else ""
                resp = f"🟢 *VENTA{titulo}" + (" CON SALDO*
" if solo_pendientes else "*
")
                bloque_ant = ""
                n = 0
                for row in rows:
                    bl, desc, ppto, cobrado, cob_cac = row
                    total = float(cobrado)+float(cob_cac)
                    pendiente = float(ppto) - total
                    pct = (total/float(ppto)*100) if float(ppto)>0 else 0
                    if solo_pendientes and pendiente <= 0:
                        continue
                    if not bloque_filtro and bl != bloque_ant:
                        resp += f"
📍 *{bl}*
"
                        bloque_ant = bl
                    resp += f"  • {desc}
    {pct:.0f}% cobrado | Pendiente: {fmt_ars(pendiente)}
"
                    n += 1
                if n == 0:
                    resp += "
No hay contratos con saldo pendiente."
            return resp

        # ── OTRO ──
        else:
            return data.get("respuesta", texto)

    except Exception as e:
        return f"Error: {str(e)}"

def ejecutar_pendiente(pendiente, remitente):
    try:
        conn = get_db()
        tipo = pendiente["tipo"]
        monto = pendiente["monto"]
        
        if tipo == "pago_proveedor":
            conn.run("UPDATE contratos_compra SET pagado = pagado + :m WHERE id = :id",
                m=monto, id=pendiente["contrato_id"])
            conn.run("""INSERT INTO movimientos_contratos 
                (tipo, contrato_id, obra, bloque, descripcion, proveedor, monto, es_cac, remitente)
                VALUES ('compra', :cid, :o, :b, :d, :p, :m, :cac, :rem)""",
                cid=pendiente["contrato_id"], o=pendiente["obra"], b=pendiente["bloque"],
                d=pendiente["desc"], p=pendiente["prov"], m=monto,
                cac=pendiente.get("es_cac", False), rem=remitente)
            conn.close()
            return f"✅ *Pago registrado*\n{fmt_ars(monto)} → {pendiente['desc']} ({pendiente['bloque']})"

        elif tipo == "cobro_cliente":
            es_cac = pendiente.get("es_cac", False)
            if es_cac:
                conn.run("UPDATE contratos_venta SET cobrado_cac = cobrado_cac + :m WHERE id = :id",
                    m=monto, id=pendiente["contrato_id"])
            else:
                conn.run("UPDATE contratos_venta SET cobrado = cobrado + :m WHERE id = :id",
                    m=monto, id=pendiente["contrato_id"])
            conn.run("""INSERT INTO movimientos_contratos 
                (tipo, contrato_id, obra, bloque, descripcion, monto, es_cac, remitente)
                VALUES ('venta', :cid, :o, :b, :d, :m, :cac, :rem)""",
                cid=pendiente["contrato_id"], o=pendiente["obra"], b=pendiente["bloque"],
                d=pendiente["desc"], m=monto, cac=es_cac, rem=remitente)
            conn.close()
            return f"✅ *Cobro registrado*\n{fmt_ars(monto)}{' (CAC)' if es_cac else ''} → {pendiente['desc']} ({pendiente['bloque']})"
    except Exception as e:
        return f"❌ Error al guardar: {str(e)}"

@app.route("/webhook", methods=["POST"])
def webhook():
    mensaje = request.form.get("Body", "")
    remitente = request.form.get("From", "")
    respuesta = procesar_mensaje(mensaje, remitente)
    resp = MessagingResponse()
    resp.message(respuesta)
    return str(resp)

@app.route("/")
def home():
    return "🏗️ ObraManager Bot activo!"

@app.route("/dashboard")
def dashboard():
    return open("dashboard.html").read()

@app.route("/constitucion")
def constitucion():
    return open("constitucion_dashboard.html").read()

@app.route("/datos")
def ver_datos():
    try:
        conn = get_db()
        movs = conn.run("SELECT * FROM movimientos ORDER BY fecha DESC LIMIT 100")
        acops = conn.run("SELECT * FROM acopios ORDER BY fecha DESC LIMIT 100")
        conn.close()
        return jsonify({"movimientos": [list(r) for r in movs], "acopios": [list(r) for r in acops]})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/resumen")
def resumen():
    try:
        conn = get_db()
        obras = conn.run("SELECT obra, SUM(CASE WHEN tipo='ingreso' THEN monto ELSE 0 END), SUM(CASE WHEN tipo='egreso' THEN monto ELSE 0 END), SUM(CASE WHEN tipo='ingreso' THEN monto ELSE -monto END) FROM movimientos GROUP BY obra ORDER BY obra")
        stocks = conn.run("SELECT obra, material, unidad, SUM(CASE WHEN tipo='acopio_ingreso' THEN cantidad ELSE -cantidad END) FROM acopios GROUP BY obra, material, unidad ORDER BY obra, material")
        conn.close()
        return jsonify({"obras": [list(r) for r in obras], "stocks": [list(r) for r in stocks]})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/contratos")
def ver_contratos():
    try:
        conn = get_db()
        compras = conn.run("SELECT id, obra, bloque, descripcion, proveedor, presupuesto, pagado, cargas FROM contratos_compra WHERE activo=TRUE ORDER BY bloque, id")
        ventas = conn.run("SELECT id, obra, bloque, descripcion, presupuesto, cobrado, cobrado_cac FROM contratos_venta WHERE activo=TRUE ORDER BY bloque, id")
        movs = conn.run("SELECT * FROM movimientos_contratos ORDER BY fecha DESC LIMIT 200")
        conn.close()
        return jsonify({
            "compra": [{"id":r[0],"obra":r[1],"bloque":r[2],"desc":r[3],"prov":r[4],"ppto":float(r[5]),"pagado":float(r[6]),"cargas":float(r[7]),"pendiente":float(r[5])-float(r[6])} for r in compras],
            "venta": [{"id":r[0],"obra":r[1],"bloque":r[2],"desc":r[3],"ppto":float(r[4]),"cobrado":float(r[5]),"cac":float(r[6]),"total":float(r[5])+float(r[6]),"pendiente":float(r[4])-float(r[5])-float(r[6])} for r in ventas],
            "movimientos": [list(r) for r in movs]
        })
    except Exception as e:
        return jsonify({"error": str(e)})


@app.route("/reset-sde")
def reset_sde():
    try:
        conn = get_db()
        conn.run("DELETE FROM contratos_compra WHERE bloque='Santiago del Estero'")
        conn.run("DELETE FROM contratos_venta WHERE bloque='Santiago del Estero'")
        nuevos_compra = [
            ("Constitución","Santiago del Estero","ALBAÑILERÍA","Joel Benitez",430000000,286232463,75878237),
            ("Constitución","Santiago del Estero","INSTALACIÓN ELÉCTRICA","Pablo Fidi",600000000,38220519,0),
            ("Constitución","Santiago del Estero","SANITARIAS","Ricardo Sequeira",600000000,68310219,0),
            ("Constitución","Santiago del Estero","PRE-INSTALACIÓN AA","Ricardo Sequeira",119500000,47800000,0),
            ("Constitución","Santiago del Estero","HERRERÍA MOLDES BALCONES","Leo Gallardo",6090000,6090000,0),
            ("Constitución","Santiago del Estero","DURLOCK MO","Salazar",200000000,2050000,0),
        ]
        nuevos_venta = [
            ("Constitución","Santiago del Estero","ALBAÑILERÍA",1214627824,592717300,113850202),
            ("Constitución","Santiago del Estero","INSTALACIÓN ELÉCTRICA",600000000,38220519,0),
            ("Constitución","Santiago del Estero","SANITARIAS",600000000,68310219,0),
            ("Constitución","Santiago del Estero","PRE-INSTALACIÓN AA",119500000,47800000,0),
            ("Constitución","Santiago del Estero","HERRERÍA MOLDES BALCONES",6090000,6090000,0),
            ("Constitución","Santiago del Estero","DURLOCK MO",200000000,2050000,0),
        ]
        for r in nuevos_compra:
            conn.run("INSERT INTO contratos_compra (obra,bloque,descripcion,proveedor,presupuesto,pagado,cargas) VALUES (:o,:b,:d,:p,:pr,:pa,:c)",
                o=r[0],b=r[1],d=r[2],p=r[3],pr=r[4],pa=r[5],c=r[6])
        for r in nuevos_venta:
            conn.run("INSERT INTO contratos_venta (obra,bloque,descripcion,presupuesto,cobrado,cobrado_cac) VALUES (:o,:b,:d,:pr,:c,:cac)",
                o=r[0],b=r[1],d=r[2],pr=r[3],c=r[4],cac=r[5])
        conn.close()
        return "✅ Contratos SDE actualizados correctamente con datos reales de Drive."
    except Exception as e:
        return f"❌ Error: {str(e)}"

# Inicializar
try:
    init_db()
    cargar_contratos_iniciales()
except Exception as e:
    print(f"Init error: {e}")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
