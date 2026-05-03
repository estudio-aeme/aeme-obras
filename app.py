from flask import Flask, request, jsonify
from twilio.twiml.messaging_response import MessagingResponse
import os
import json
import requests
from datetime import datetime
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

def get_db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

def init_db():
    """Crea las tablas si no existen"""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS movimientos (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMP DEFAULT NOW(),
            tipo VARCHAR(20),
            obra VARCHAR(100),
            monto NUMERIC(15,2),
            descripcion TEXT,
            proveedor VARCHAR(100),
            remitente VARCHAR(50)
        );
        CREATE TABLE IF NOT EXISTS acopios (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMP DEFAULT NOW(),
            material VARCHAR(100),
            cantidad NUMERIC(10,2),
            unidad VARCHAR(20),
            obra VARCHAR(100),
            tipo VARCHAR(20) DEFAULT 'ingreso',
            remitente VARCHAR(50)
        );
    """)
    conn.commit()
    cur.close()
    conn.close()

def procesar_con_ia(mensaje, remitente):
    try:
        headers = {
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        }

        system_prompt = """Sos el asistente de ObraManager para la constructora Aeme Obras.
Ayudás a Julián y Julieta a gestionar obras de construcción vía WhatsApp.

Podés hacer estas acciones:
1. REGISTRAR un gasto → ACCION:{"tipo":"egreso","obra":"nombre","monto":1000,"descripcion":"cemento","proveedor":"opcional"}
2. REGISTRAR un ingreso → ACCION:{"tipo":"ingreso","obra":"nombre","monto":5000,"descripcion":"anticipo cliente"}
3. REGISTRAR acopio (material que entra) → ACCION:{"tipo":"acopio_ingreso","material":"cemento","cantidad":50,"unidad":"bolsas","obra":"nombre"}
4. REGISTRAR desacopio (material que sale) → ACCION:{"tipo":"acopio_egreso","material":"cemento","cantidad":10,"unidad":"bolsas","obra":"nombre"}
5. CONSULTAR saldo de una obra → respondé con los datos de la base

Cuando registres algo incluí el JSON exactamente así con el prefijo ACCION: al inicio.
Para consultas respondé en texto amigable.
Hablá en español rioplatense, claro y conciso. Máximo 3 oraciones."""

        payload = {
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 400,
            "system": system_prompt,
            "messages": [{"role": "user", "content": mensaje}]
        }

        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers,
            json=payload,
            timeout=15
        )

        if response.status_code == 200:
            respuesta = response.json()["content"][0]["text"]

            if "ACCION:" in respuesta:
                try:
                    inicio = respuesta.index("ACCION:") + 7
                    fin = respuesta.index("}", inicio) + 1
                    accion = json.loads(respuesta[inicio:fin])
                    guardar_en_db(accion, remitente)
                    texto = respuesta[fin:].strip() or "¡Registrado!"
                    return f"✅ {texto}"
                except Exception as e:
                    pass

            return respuesta
        else:
            return f"Error al procesar ({response.status_code})"

    except Exception as e:
        return f"Error: {str(e)}"

def guardar_en_db(accion, remitente):
    conn = get_db()
    cur = conn.cursor()
    tipo = accion.get("tipo", "")

    if tipo in ["egreso", "ingreso"]:
        cur.execute("""
            INSERT INTO movimientos (tipo, obra, monto, descripcion, proveedor, remitente)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            tipo,
            accion.get("obra", "General"),
            accion.get("monto", 0),
            accion.get("descripcion", ""),
            accion.get("proveedor", ""),
            remitente
        ))
    elif tipo in ["acopio_ingreso", "acopio_egreso"]:
        cur.execute("""
            INSERT INTO acopios (material, cantidad, unidad, obra, tipo, remitente)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            accion.get("material", ""),
            accion.get("cantidad", 0),
            accion.get("unidad", ""),
            accion.get("obra", "General"),
            tipo,
            remitente
        ))

    conn.commit()
    cur.close()
    conn.close()

@app.route("/webhook", methods=["POST"])
def webhook():
    mensaje = request.form.get("Body", "")
    remitente = request.form.get("From", "")
    respuesta = procesar_con_ia(mensaje, remitente)
    resp = MessagingResponse()
    resp.message(respuesta)
    return str(resp)

@app.route("/")
def home():
    return "🏗️ ObraManager Bot activo!"

@app.route("/datos")
def ver_datos():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM movimientos ORDER BY fecha DESC LIMIT 100")
        movimientos = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT * FROM acopios ORDER BY fecha DESC LIMIT 100")
        acopios = [dict(r) for r in cur.fetchall()]
        cur.close()
        conn.close()
        return jsonify({"movimientos": movimientos, "acopios": acopios})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/resumen")
def resumen():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            SELECT obra,
                SUM(CASE WHEN tipo='ingreso' THEN monto ELSE 0 END) as total_ingresos,
                SUM(CASE WHEN tipo='egreso' THEN monto ELSE 0 END) as total_egresos,
                SUM(CASE WHEN tipo='ingreso' THEN monto ELSE -monto END) as saldo
            FROM movimientos GROUP BY obra ORDER BY obra
        """)
        obras = [dict(r) for r in cur.fetchall()]
        cur.execute("""
            SELECT obra, material, unidad,
                SUM(CASE WHEN tipo='acopio_ingreso' THEN cantidad ELSE -cantidad END) as stock
            FROM acopios GROUP BY obra, material, unidad ORDER BY obra, material
        """)
        stocks = [dict(r) for r in cur.fetchall()]
        cur.close()
        conn.close()
        return jsonify({"resumen_por_obra": obras, "stock_materiales": stocks})
    except Exception as e:
        return jsonify({"error": str(e)})

# Inicializar DB al arrancar
with app.app_context():
    try:
        init_db()
    except:
        pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
