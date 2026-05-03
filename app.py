from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import os
import json
from datetime import datetime
import anthropic

app = Flask(__name__)

# Cliente de Anthropic
claude_client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# Base de datos simple en memoria (en producción usarías una base de datos real)
datos = {
    "movimientos": [],
    "acopios": []
}

SYSTEM_PROMPT = """Sos el asistente de ObraManager para la constructora Aeme Obras.
Ayudás a Julián y Julieta a gestionar obras de construcción vía WhatsApp.

Podés hacer estas acciones:
1. CONSULTAR saldo de caja por obra
2. REGISTRAR un gasto o ingreso
3. CONSULTAR stock de materiales (acopio)
4. REGISTRAR retiro de materiales

Cuando el usuario quiera registrar algo, respondé con un JSON así:
{"accion": "registrar_gasto", "obra": "nombre", "monto": 1000, "descripcion": "cemento", "proveedor": "Corralón XYZ"}
{"accion": "registrar_ingreso", "obra": "nombre", "monto": 5000, "descripcion": "anticipo cliente"}
{"accion": "registrar_acopio", "material": "cemento", "cantidad": 50, "unidad": "bolsas", "obra": "nombre"}

Para consultas, respondé directamente en texto amigable.
Siempre hablá en español rioplatense, de forma clara y concisa.
Si no entendés algo, pedí que te lo aclaren."""

def procesar_con_ia(mensaje, historial=[]):
    """Procesa el mensaje con Claude"""
    try:
        messages = historial + [{"role": "user", "content": mensaje}]
        
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=500,
            system=SYSTEM_PROMPT,
            messages=messages
        )
        
        respuesta = response.content[0].text
        
        # Intentar detectar si es un JSON de acción
        try:
            if "{" in respuesta and "accion" in respuesta:
                inicio = respuesta.index("{")
                fin = respuesta.rindex("}") + 1
                json_str = respuesta[inicio:fin]
                accion = json.loads(json_str)
                ejecutar_accion(accion)
                # Respuesta al usuario después de ejecutar
                return f"✅ ¡Registrado! {respuesta[:inicio].strip()}"
        except:
            pass
            
        return respuesta
        
    except Exception as e:
        return f"Ocurrió un error: {str(e)}"

def ejecutar_accion(accion):
    """Ejecuta una acción sobre los datos"""
    tipo = accion.get("accion")
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    if tipo == "registrar_gasto":
        datos["movimientos"].append({
            "fecha": ahora,
            "tipo": "egreso",
            "obra": accion.get("obra", "General"),
            "monto": accion.get("monto", 0),
            "descripcion": accion.get("descripcion", ""),
            "proveedor": accion.get("proveedor", "")
        })
    elif tipo == "registrar_ingreso":
        datos["movimientos"].append({
            "fecha": ahora,
            "tipo": "ingreso",
            "obra": accion.get("obra", "General"),
            "monto": accion.get("monto", 0),
            "descripcion": accion.get("descripcion", "")
        })
    elif tipo == "registrar_acopio":
        datos["acopios"].append({
            "fecha": ahora,
            "material": accion.get("material", ""),
            "cantidad": accion.get("cantidad", 0),
            "unidad": accion.get("unidad", ""),
            "obra": accion.get("obra", "General")
        })

@app.route("/webhook", methods=["POST"])
def webhook():
    """Recibe mensajes de WhatsApp via Twilio"""
    mensaje_entrante = request.form.get("Body", "")
    remitente = request.form.get("From", "")
    
    print(f"Mensaje de {remitente}: {mensaje_entrante}")
    
    # Procesar con IA
    respuesta = procesar_con_ia(mensaje_entrante)
    
    # Responder via Twilio
    resp = MessagingResponse()
    resp.message(respuesta)
    
    return str(resp)

@app.route("/")
def home():
    return "🏗️ ObraManager Bot activo!"

@app.route("/datos")
def ver_datos():
    """Endpoint para ver los datos registrados"""
    return json.dumps(datos, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
