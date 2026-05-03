from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import os
import json
import requests
from datetime import datetime

app = Flask(__name__)

# Base de datos simple en memoria
datos = {
    "movimientos": [],
    "acopios": []
}

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

def procesar_con_ia(mensaje):
    """Procesa el mensaje con Claude via API REST"""
    try:
        headers = {
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        }
        
        system_prompt = """Sos el asistente de ObraManager para la constructora Aeme Obras.
Ayudás a Julián y Julieta a gestionar obras de construcción vía WhatsApp.

Podés hacer estas acciones:
1. CONSULTAR saldo de caja por obra
2. REGISTRAR un gasto o ingreso  
3. CONSULTAR stock de materiales
4. REGISTRAR retiro de materiales

Cuando el usuario quiera registrar algo, incluí un JSON en tu respuesta así:
ACCION:{"tipo": "gasto", "obra": "nombre", "monto": 1000, "descripcion": "cemento"}
ACCION:{"tipo": "ingreso", "obra": "nombre", "monto": 5000, "descripcion": "anticipo"}

Para consultas, respondé directamente en texto amigable.
Hablá en español rioplatense, de forma clara y concisa. Máximo 3 oraciones."""

        payload = {
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 300,
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
            data = response.json()
            respuesta = data["content"][0]["text"]
            
            # Detectar y ejecutar acciones
            if "ACCION:" in respuesta:
                try:
                    inicio = respuesta.index("ACCION:") + 7
                    fin = respuesta.index("}", inicio) + 1
                    json_str = respuesta[inicio:fin]
                    accion = json.loads(json_str)
                    ejecutar_accion(accion)
                    respuesta = respuesta.replace(f"ACCION:{json_str}", "").strip()
                    return f"✅ Registrado! {respuesta}"
                except:
                    pass
            
            return respuesta
        else:
            return f"Error al procesar: {response.status_code}"
            
    except Exception as e:
        return f"Error: {str(e)}"

def ejecutar_accion(accion):
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    tipo = accion.get("tipo", "")
    
    if tipo in ["gasto", "egreso"]:
        datos["movimientos"].append({
            "fecha": ahora,
            "tipo": "egreso",
            "obra": accion.get("obra", "General"),
            "monto": accion.get("monto", 0),
            "descripcion": accion.get("descripcion", "")
        })
    elif tipo in ["ingreso"]:
        datos["movimientos"].append({
            "fecha": ahora,
            "tipo": "ingreso", 
            "obra": accion.get("obra", "General"),
            "monto": accion.get("monto", 0),
            "descripcion": accion.get("descripcion", "")
        })

@app.route("/webhook", methods=["POST"])
def webhook():
    mensaje_entrante = request.form.get("Body", "")
    remitente = request.form.get("From", "")
    print(f"Mensaje de {remitente}: {mensaje_entrante}")
    
    respuesta = procesar_con_ia(mensaje_entrante)
    
    resp = MessagingResponse()
    resp.message(respuesta)
    return str(resp)

@app.route("/")
def home():
    return "🏗️ ObraManager Bot activo!"

@app.route("/datos")
def ver_datos():
    return json.dumps(datos, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
