import json
import os
import uuid
import html
import base64
import hashlib
from datetime import datetime, timezone
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
import streamlit as st
import streamlit.components.v1 as components
from pypdf import PdfReader


# ============================================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ============================================================

st.set_page_config(
    page_title="ORIA",
    page_icon="✨",
    layout="wide",
)


# ============================================================
# 2. ALMACENAMIENTO: CUENTAS, NUBE (SUPABASE) Y COPIA LOCAL
# ============================================================

# Cada usuario se identifica por su correo (si ha iniciado sesión con
# Google) o, si el login aún no está configurado, por un identificador
# de invitado guardado en la URL. Los datos (chats + memoria) se guardan:
#   - en Supabase (base de datos en la nube) si está configurado, para
#     que la cuenta funcione igual en cualquier dispositivo y no se
#     borre cuando Streamlit reinicia el servidor;
#   - o, si no, en archivos JSON locales (se pueden perder al reiniciar).
CARPETA_USUARIOS = "usuarios"
CARPETA_IMAGENES = "imagenes_generadas"
TABLA_SUPABASE = "oria_usuarios"

# Cuántos mensajes recientes mandamos como contexto a Groq.
# Evita que las conversaciones muy largas se coman el contexto
# (y la cuota gratuita) del modelo.
MAX_MENSAJES_HISTORIAL = 20


def login_configurado():
    """True si en Secrets existe la sección [auth] (login con Google)."""
    try:
        return "auth" in st.secrets
    except Exception:
        return False


def _config_supabase():
    """Devuelve (url, key) de Supabase, o (None, None) si no está
    configurado en Secrets."""
    try:
        url = str(st.secrets["SUPABASE_URL"]).strip().strip('"').rstrip("/")
        key = str(st.secrets["SUPABASE_KEY"]).strip().strip('"')

        if url and key:
            return url, key

    except Exception:
        pass

    return None, None


def supabase_activo():
    url, key = _config_supabase()
    return bool(url and key)


def _headers_supabase(key, extra=None):
    headers = {
        "apikey": key,
        "Content-Type": "application/json",
    }

    # Las claves antiguas (service_role) son JWT y también van en
    # Authorization; las nuevas (sb_secret_...) solo en apikey.
    if key.startswith("eyJ"):
        headers["Authorization"] = f"Bearer {key}"

    if extra:
        headers.update(extra)

    return headers


def obtener_uid():
    """Modo invitado (solo si el login con Google aún no está
    configurado): da a cada visitante un identificador propio,
    guardado como parámetro en la URL."""

    if "uid" not in st.query_params:
        nuevo_uid = uuid.uuid4().hex[:10]
        st.query_params["uid"] = nuevo_uid
        st.rerun()

    return st.query_params["uid"]


def obtener_identidad():
    """Decide quién es el usuario actual. Si el login con Google está
    configurado, obliga a iniciar sesión (y para la ejecución mostrando
    la pantalla de acceso si aún no lo ha hecho). Devuelve un diccionario
    con id, nombre y si es invitado."""

    if login_configurado():

        if not getattr(st.user, "is_logged_in", False):
            mostrar_pantalla_login()
            st.stop()

        email = str(getattr(st.user, "email", "") or "").strip().lower()
        nombre = str(getattr(st.user, "name", "") or "").strip() or email

        if not email:
            st.error(
                "Google no ha devuelto tu correo. Cierra sesión e "
                "inténtalo de nuevo."
            )
            st.button("Cerrar sesión", on_click=st.logout)
            st.stop()

        return {"id": email, "nombre": nombre, "invitado": False}

    return {
        "id": f"invitado:{obtener_uid()}",
        "nombre": "Invitado",
        "invitado": True,
    }


def _ruta_datos_usuario(usuario_id):
    os.makedirs(CARPETA_USUARIOS, exist_ok=True)
    nombre = hashlib.sha256(usuario_id.encode("utf-8")).hexdigest()[:24]
    return os.path.join(CARPETA_USUARIOS, f"{nombre}.json")


def _cargar_local(usuario_id):
    ruta = _ruta_datos_usuario(usuario_id)

    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)

                if isinstance(datos, dict):
                    datos.setdefault("chats", {})
                    datos.setdefault("memoria", "")
                    return datos

        except Exception:
            pass

    return {"chats": {}, "memoria": ""}


def _guardar_local(usuario_id, chats, memoria):
    ruta = _ruta_datos_usuario(usuario_id)

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(
            {"chats": chats, "memoria": memoria},
            f,
            ensure_ascii=False,
            indent=2,
        )


def cargar_datos_usuario(usuario_id):
    """Carga chats y memoria del usuario. Devuelve (datos, ok).
    Si falla la nube devuelve ok=False para NO seguir adelante con
    datos vacíos (así no se sobrescribe lo que ya tenía guardado)."""

    url, key = _config_supabase()

    if not (url and key):
        return _cargar_local(usuario_id), True

    try:
        respuesta = requests.get(
            f"{url}/rest/v1/{TABLA_SUPABASE}",
            headers=_headers_supabase(key),
            params={
                "email": f"eq.{usuario_id}",
                "select": "chats,memoria",
            },
            timeout=20,
        )

        if respuesta.status_code != 200:
            return None, False

        filas = respuesta.json()

        if not filas:
            return {"chats": {}, "memoria": ""}, True

        fila = filas[0]
        chats = fila.get("chats") or {}
        memoria = fila.get("memoria") or ""

        if not isinstance(chats, dict):
            chats = {}

        return {"chats": chats, "memoria": memoria}, True

    except Exception:
        return None, False


def guardar_datos_usuario(usuario_id, chats, memoria):
    """Guarda chats y memoria del usuario (nube si está configurada,
    y si no, archivo local)."""

    url, key = _config_supabase()

    if not (url and key):
        try:
            _guardar_local(usuario_id, chats, memoria)
        except Exception as e:
            st.error(f"Error al guardar tus datos: {e}")
        return

    try:
        respuesta = requests.post(
            f"{url}/rest/v1/{TABLA_SUPABASE}",
            headers=_headers_supabase(
                key,
                {"Prefer": "resolution=merge-duplicates,return=minimal"},
            ),
            json={
                "email": usuario_id,
                "chats": chats,
                "memoria": memoria,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            },
            timeout=20,
        )

        if respuesta.status_code not in (200, 201, 204):
            st.error(
                "No se han podido guardar tus datos en la nube "
                f"(error {respuesta.status_code})."
            )

    except Exception as e:
        st.error(f"No se han podido guardar tus datos en la nube: {e}")


# ============================================================
# 3. FECHA ACTUAL
# ============================================================

# Usamos hora de España.
ahora = datetime.now(ZoneInfo("Europe/Madrid"))

dias = [
    "lunes",
    "martes",
    "miércoles",
    "jueves",
    "viernes",
    "sábado",
    "domingo",
]

meses = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]

fecha_hoy_str = (
    f"{dias[ahora.weekday()]}, "
    f"{ahora.day} de "
    f"{meses[ahora.month - 1]} de "
    f"{ahora.year}"
)


# ============================================================
# 4. CONFIGURACIÓN DE GROQ
# ============================================================

# Modelo de texto (rápido, para la conversación normal).
MODELO_GROQ = "openai/gpt-oss-120b"

# Modelo con visión (el multimodal vigente del catálogo self-serve
# de Groq a fecha de hoy). Se usa solo cuando el usuario adjunta
# una imagen. Groq cambia este modelo de vez en cuando: si en el
# futuro vuelve a dar error 404, comprueba el modelo actual en
# https://console.groq.com/docs/vision y actualízalo aquí.
MODELO_GROQ_VISION = "qwen/qwen3.8-27b"

URL_GROQ = "https://api.groq.com/openai/v1/chat/completions"


def obtener_api_key():
    """Obtiene y limpia la API key desde Streamlit Secrets."""

    try:
        if "GROQ_API_KEY" not in st.secrets:
            return None

        raw_key = str(st.secrets["GROQ_API_KEY"])

        api_key = (
            raw_key
            .strip()
            .strip('"')
            .strip("'")
            .strip()
        )

        return api_key

    except Exception:
        return None


# ============================================================
# 5. UTILIDADES PARA IMÁGENES Y PDF
# ============================================================

def imagen_a_data_uri(archivo_subido):
    """Convierte un archivo de imagen subido en un data URI base64,
    formato que necesita la API de Groq para 'ver' la imagen."""

    bytes_imagen = archivo_subido.getvalue()
    b64 = base64.b64encode(bytes_imagen).decode("utf-8")
    tipo_mime = archivo_subido.type or "image/png"

    return f"data:{tipo_mime};base64,{b64}"


def extraer_texto_pdf(archivo_subido):
    """Extrae el texto de un PDF subido. Devuelve (texto, error)."""

    try:
        lector = PdfReader(archivo_subido)

        texto = ""
        for pagina in lector.pages:
            texto += (pagina.extract_text() or "") + "\n"

        texto = texto.strip()

        if not texto:
            return None, (
                "No se ha podido extraer texto de este PDF "
                "(puede que sea un documento escaneado, es decir, "
                "imágenes sin texto real)."
            )

        # Recortamos para no disparar el consumo de tokens.
        if len(texto) > 12000:
            texto = texto[:12000] + "\n\n[...documento truncado por longitud...]"

        return texto, None

    except Exception as e:
        return None, str(e)


def obtener_pollinations_key():
    """API key gratuita opcional de Pollinations.ai (auth.pollinations.ai).
    Sin ella, las imágenes se generan igual, pero con la marca de agua
    y límites más estrictos (nivel anónimo)."""

    try:
        if "POLLINATIONS_API_KEY" not in st.secrets:
            return None

        return str(st.secrets["POLLINATIONS_API_KEY"]).strip()

    except Exception:
        return None


def mejorar_prompt_imagen(prompt_simple):
    """Usa el modelo de texto de Groq para convertir una idea corta
    del usuario en una descripción rica y detallada, lo que mejora
    mucho la calidad de la imagen generada. Si algo falla, se queda
    con el prompt original del usuario."""

    api_key = obtener_api_key()

    if not api_key or not api_key.startswith("gsk_"):
        return prompt_simple

    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

        payload = {
            "model": MODELO_GROQ,
            "stream": False,
            "temperature": 0.8,
            "max_completion_tokens": 300,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "Eres un experto escribiendo prompts para modelos "
                        "de generación de imágenes (tipo Flux). Convierte "
                        "la idea del usuario en una descripción muy "
                        "detallada y visual, en una sola frase larga: "
                        "sujeto, estilo artístico, iluminación, encuadre, "
                        "ambiente y calidad (ej. 'fotografía realista', "
                        "'8k', 'cinematográfico', 'alto detalle'...). "
                        "Responde ÚNICAMENTE con el prompt final, sin "
                        "comillas, explicaciones ni texto adicional."
                    ),
                },
                {"role": "user", "content": prompt_simple},
            ],
        }

        respuesta = requests.post(
            URL_GROQ, headers=headers, json=payload, timeout=20
        )

        if respuesta.status_code == 200:
            texto = (
                respuesta.json()
                .get("choices", [{}])[0]
                .get("message", {})
                .get("content", "")
                .strip()
            )

            if texto:
                return texto

        return prompt_simple

    except Exception:
        return prompt_simple


def generar_imagen_ia(prompt_imagen, intentos=2):
    """Genera una imagen a partir de un texto usando Pollinations.ai
    (servicio gratuito). Si hay una POLLINATIONS_API_KEY configurada
    en Secrets, se usa para quitar la marca de agua y tener más
    estabilidad; si no, funciona igualmente en modo anónimo.
    Devuelve (bytes, error)."""

    token = obtener_pollinations_key()

    prompt_codificado = quote(prompt_imagen)
    semilla = uuid.uuid4().int % 1_000_000

    url = (
        f"https://image.pollinations.ai/prompt/{prompt_codificado}"
        f"?model=flux&width=1024&height=1024"
        f"&nologo=true&enhance=true&seed={semilla}"
    )

    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    ultimo_error = None

    for _ in range(intentos):

        try:
            respuesta = requests.get(url, headers=headers, timeout=90)

            if (
                respuesta.status_code == 200
                and respuesta.headers.get("content-type", "").startswith("image")
            ):
                return respuesta.content, None

            ultimo_error = (
                f"El generador de imágenes devolvió el error "
                f"{respuesta.status_code}."
            )

        except requests.exceptions.Timeout:
            ultimo_error = "El generador de imágenes ha tardado demasiado."

        except Exception as e:
            ultimo_error = str(e)

    return None, ultimo_error


def transcribir_audio(audio_bytes, nombre_archivo="grabacion.wav"):
    """Transcribe un audio a texto usando Whisper (incluido en tu misma
    cuenta de Groq, sin nada nuevo que configurar). Devuelve (texto, error)."""

    api_key = obtener_api_key()

    if not api_key or not api_key.startswith("gsk_"):
        return None, "No se ha encontrado una GROQ_API_KEY válida."

    try:
        url = "https://api.groq.com/openai/v1/audio/transcriptions"

        headers = {"Authorization": f"Bearer {api_key}"}

        archivos = {
            "file": (nombre_archivo, audio_bytes, "audio/wav"),
        }

        datos = {
            "model": "whisper-large-v3-turbo",
            "language": "es",
            "response_format": "json",
        }

        respuesta = requests.post(
            url, headers=headers, files=archivos, data=datos, timeout=60
        )

        if respuesta.status_code == 200:
            texto = respuesta.json().get("text", "").strip()

            if texto:
                return texto, None

            return None, "No se ha detectado voz en el audio."

        try:
            error_data = respuesta.json()
            error_message = error_data.get("error", {}).get(
                "message", respuesta.text
            )
        except Exception:
            error_message = respuesta.text

        return None, f"Error {respuesta.status_code}: {error_message}"

    except requests.exceptions.Timeout:
        return None, "La transcripción ha tardado demasiado."

    except Exception as e:
        return None, str(e)


# ============================================================
# 6. FUNCIÓN PRINCIPAL DE LA IA
# ============================================================

def obtener_respuesta_ia_stream(
    prompt_usuario,
    historial_mensajes=None,
    imagen_data_uri=None,
    memoria_texto="",
):
    """
    Envía la conversación a Groq y devuelve la respuesta
    progresivamente mediante streaming.

    Si se pasa `imagen_data_uri`, la consulta se manda al modelo
    con visión (MODELO_GROQ_VISION) junto con la imagen adjunta.
    """

    # --------------------------------------------------------
    # Comprobar API KEY
    # --------------------------------------------------------

    api_key = obtener_api_key()

    if not api_key:
        yield (
            "⚠️ **No se ha encontrado `GROQ_API_KEY`.**\n\n"
            "Ve a **Streamlit → Settings → Secrets** y comprueba "
            "que tengas configurado:\n\n"
            "```toml\n"
            'GROQ_API_KEY = "TU_API_KEY"\n'
            "```"
        )
        return

    if not api_key.startswith("gsk_"):
        yield (
            "⚠️ **La API key de Groq no parece válida.**\n\n"
            "Debe comenzar por `gsk_`."
        )
        return

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    # --------------------------------------------------------
    # System prompt (incluye memoria si existe)
    # --------------------------------------------------------

    instrucciones_sistema = (
        "Eres ORIA, una asistente virtual inteligente, "
        "rápida, natural, amable y precisa. "
        "Puedes ayudar con tecnología, programación, "
        "deportes, fútbol, estudios, informática, "
        "cultura, entretenimiento y muchos otros temas. "
        "También puedes analizar imágenes y documentos PDF "
        "que el usuario te adjunte, y generar imágenes cuando te "
        "lo pidan. "
        f"La fecha actual es {fecha_hoy_str}. "
        "Responde siempre en español, salvo que el usuario "
        "pida expresamente otro idioma. "
        "Explica las cosas de forma clara y útil."
    )

    if memoria_texto and memoria_texto.strip():
        instrucciones_sistema += (
            "\n\nDatos que debes recordar sobre este usuario "
            "(úsalos de forma natural cuando sea relevante, sin "
            "recitarlos innecesariamente):\n"
            f"{memoria_texto.strip()}"
        )

    messages = [
        {
            "role": "system",
            "content": instrucciones_sistema,
        }
    ]

    # --------------------------------------------------------
    # Modo visión: un único turno con la imagen adjunta
    # --------------------------------------------------------

    if imagen_data_uri:

        modelo_usar = MODELO_GROQ_VISION

        messages.append(
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt_usuario},
                    {
                        "type": "image_url",
                        "image_url": {"url": imagen_data_uri},
                    },
                ],
            }
        )

    # --------------------------------------------------------
    # Modo texto normal: con historial recortado por caracteres
    # --------------------------------------------------------

    else:

        modelo_usar = MODELO_GROQ

        # El plan gratuito de Groq tiene un límite de tokens por
        # minuto bastante ajustado. Truncamos mensajes individuales
        # muy largos (p. ej. un código pegado entero) y además nos
        # quedamos solo con los últimos mensajes hasta llenar un
        # presupuesto de caracteres, para no reventar ese límite.
        LIMITE_CARACTERES_POR_MENSAJE = 6000
        PRESUPUESTO_CARACTERES_HISTORIAL = 11000

        if historial_mensajes:

            recientes = historial_mensajes[-MAX_MENSAJES_HISTORIAL:]

            mensajes_preparados = []

            for msg in recientes:

                if not isinstance(msg, dict):
                    continue

                if "role" not in msg or "content" not in msg:
                    continue

                # Los mensajes de tipo "imagen" (generadas por ORIA)
                # no se pueden mandar como texto; los resumimos.
                if msg.get("type") == "image":
                    if msg["role"] == "assistant":
                        mensajes_preparados.append(
                            {
                                "role": "assistant",
                                "content": (
                                    "[Generé una imagen a partir de: "
                                    f"{msg.get('prompt', '')}]"
                                ),
                            }
                        )
                    continue

                if msg["role"] == "user":
                    role = "user"

                elif msg["role"] == "assistant":
                    role = "assistant"

                else:
                    continue

                contenido = str(msg["content"])

                if len(contenido) > LIMITE_CARACTERES_POR_MENSAJE:
                    contenido = (
                        contenido[:LIMITE_CARACTERES_POR_MENSAJE]
                        + "\n\n[...mensaje truncado por longitud...]"
                    )

                mensajes_preparados.append(
                    {"role": role, "content": contenido}
                )

            # Recorremos de más reciente a más antiguo llenando el
            # presupuesto de caracteres. Siempre dejamos pasar al
            # menos el último mensaje, aunque él solo se acerque al
            # límite, para no dejar la conversación sin nada que
            # mandar.
            mensajes_incluidos = []
            caracteres_acumulados = 0

            for msg in reversed(mensajes_preparados):

                longitud = len(msg["content"])

                if (
                    mensajes_incluidos
                    and caracteres_acumulados + longitud
                    > PRESUPUESTO_CARACTERES_HISTORIAL
                ):
                    break

                caracteres_acumulados += longitud
                mensajes_incluidos.insert(0, msg)

            messages.extend(mensajes_incluidos)

        # Si por alguna razón no tenemos historial,
        # añadimos el mensaje actual.
        if not historial_mensajes:
            messages.append(
                {
                    "role": "user",
                    "content": prompt_usuario,
                }
            )

    # --------------------------------------------------------
    # Payload
    # --------------------------------------------------------

    payload = {
        "model": modelo_usar,
        "messages": messages,
        "stream": True,
        "temperature": 0.7,
        "max_completion_tokens": 1024,
    }

    # --------------------------------------------------------
    # Petición
    # --------------------------------------------------------

    try:

        response = requests.post(
            URL_GROQ,
            headers=headers,
            json=payload,
            stream=True,
            timeout=60,
        )

        # ====================================================
        # RESPUESTA CORRECTA
        # ====================================================

        if response.status_code == 200:

            for line in response.iter_lines():

                if not line:
                    continue

                try:

                    line_str = line.decode("utf-8")

                except UnicodeDecodeError:
                    continue

                # Solo procesamos líneas SSE.
                if not line_str.startswith("data: "):
                    continue

                data_json = line_str[6:]

                # Fin del streaming.
                if data_json.strip() == "[DONE]":
                    break

                try:

                    chunk_obj = json.loads(data_json)

                    choices = chunk_obj.get("choices", [])

                    if not choices:
                        continue

                    delta = choices[0].get("delta", {})

                    content = delta.get("content")

                    if content:
                        yield content

                except json.JSONDecodeError:
                    continue

                except (KeyError, IndexError, TypeError):
                    continue

            return

        # ====================================================
        # API KEY INCORRECTA
        # ====================================================

        if response.status_code == 401:

            yield (
                "⚠️ **Error 401: API key no válida.**\n\n"
                "Comprueba que la nueva API key de Groq esté "
                "correctamente colocada en "
                "**Streamlit → Settings → Secrets**."
            )
            return

        # ====================================================
        # SIN PERMISOS
        # ====================================================

        if response.status_code == 403:

            yield (
                "⚠️ **Error 403: acceso denegado por Groq.**\n\n"
                "La clave existe, pero Groq está rechazando "
                "el acceso a la petición."
            )
            return

        # ====================================================
        # MODELO NO DISPONIBLE
        # ====================================================

        if response.status_code == 404:

            try:
                error_data = response.json()

                error_message = (
                    error_data
                    .get("error", {})
                    .get(
                        "message",
                        "El modelo no está disponible."
                    )
                )

            except Exception:
                error_message = response.text

            yield (
                "⚠️ **Error 404 de Groq.**\n\n"
                f"{error_message}\n\n"
                f"Modelo utilizado: `{modelo_usar}`"
            )

            return

        # ====================================================
        # MENSAJE DEMASIADO GRANDE (413)
        # ====================================================

        if response.status_code == 413:

            yield (
                "⚠️ **La conversación se ha quedado demasiado larga "
                "para el plan gratuito de Groq en este momento "
                "(demasiados tokens por minuto).**\n\n"
                "Prueba a pulsar '➕ Nueva conversación' para empezar "
                "de cero, o espera un minuto y vuelve a intentarlo."
            )

            return

        # ====================================================
        # OTROS ERRORES
        # ====================================================

        try:

            error_data = response.json()

            error_message = (
                error_data
                .get("error", {})
                .get(
                    "message",
                    response.text
                )
            )

        except Exception:

            error_message = response.text

        yield (
            f"⚠️ **Error de Groq ({response.status_code})**\n\n"
            f"{error_message}"
        )

    # ========================================================
    # TIMEOUT
    # ========================================================

    except requests.exceptions.Timeout:

        yield (
            "⚠️ **Groq ha tardado demasiado en responder.**\n\n"
            "Vuelve a intentarlo."
        )

    # ========================================================
    # ERROR DE CONEXIÓN
    # ========================================================

    except requests.exceptions.ConnectionError:

        yield (
            "⚠️ **No se ha podido conectar con Groq.**\n\n"
            "Comprueba la conexión del servidor."
        )

    # ========================================================
    # OTRO ERROR
    # ========================================================

    except requests.exceptions.RequestException as e:

        yield (
            "⚠️ **Error de conexión con Groq:**\n\n"
            f"{str(e)}"
        )

    except Exception as e:

        yield (
            "⚠️ **Se ha producido un error inesperado:**\n\n"
            f"{str(e)}"
        )


# ============================================================
# 7. CSS
# ============================================================

css_code = """
<style>

/* Oculta la barra de herramientas de Streamlit (menú, Share,
GitHub, editar...) para que la app se vea como un producto propio,
no como un proyecto de Streamlit. Nota: si eres tú quien ha
publicado la app y entras logueado en Streamlit Cloud, es posible
que sigas viendo el aviso "Manage app" abajo a la derecha — eso lo
pone la plataforma por fuera de la app y solo tú lo ves, no tus
amigos ni familiares. */
#MainMenu,
header[data-testid="stHeader"],
[data-testid="stToolbar"],
footer {
    visibility: hidden;
    height: 0;
}

[data-testid="stChatMessageAvatarUser"],
[data-testid="stChatMessageAvatarAssistant"] {
    display: none !important;
}

.stChatMessage {
    background-color: transparent !important;
    border: none !important;
    padding: 4px 0px !important;
}

.user-bubble-container {
    display: flex;
    justify-content: flex-end;
    width: 100%;
    margin-bottom: 12px;
}

.user-bubble {
    background-color: #F5F5F4;
    color: #1A1A1A;
    padding: 12px 18px;
    border-radius: 20px 20px 4px 20px;
    max-width: 75%;
    word-wrap: break-word;
    overflow-wrap: break-word;
    box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    font-family:
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        Roboto,
        sans-serif;
    font-size: 1rem;
}

.oria-logo-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 10px;
}

.stSidebar .stButton > button {
    border-radius: 8px !important;
}

/* Ajustes para pantallas de móvil */
@media (max-width: 640px) {

    .user-bubble {
        max-width: 88%;
        padding: 10px 14px;
        font-size: 0.95rem;
    }

    h3 {
        font-size: 1.1rem !important;
    }

    /* Los interruptores en columna en vez de apretujados uno
    junto al otro en pantallas estrechas. */
    div[data-testid="stHorizontalBlock"] {
        flex-wrap: wrap;
    }
}

</style>
"""

st.markdown(
    css_code,
    unsafe_allow_html=True,
)


# ============================================================
# 8. SESSION STATE
# ============================================================

def mostrar_pantalla_login():
    """Pantalla de acceso: logo + botón para entrar con Google."""

    st.markdown("<br><br><br>", unsafe_allow_html=True)

    _c1, _c2, _c3 = st.columns([1, 2, 1])

    with _c2:

        st.markdown(
            """
            <div class="oria-logo-container">
                <svg width="60" height="60" viewBox="0 0 56 56"
                     xmlns="http://www.w3.org/2000/svg">
                    <circle cx="28" cy="28" r="25" fill="none"
                            stroke="#2B2B31" stroke-width="2.5"/>
                    <circle cx="28" cy="28" r="9" fill="#2B2B31"/>
                    <circle cx="45" cy="13" r="3.5" fill="#2B2B31"/>
                </svg>
                <h1 style="text-align:center;font-size:3rem;font-weight:700;
                           letter-spacing:0.05em;margin:0;color:#1A1A1A;">
                    ORIA
                </h1>
            </div>
            <h3 style="text-align:center;color:#666;font-weight:400;">
                Inicia sesión para continuar
            </h3>
            """,
            unsafe_allow_html=True,
        )

        st.button(
            "Continuar con Google",
            on_click=st.login,
            type="primary",
            use_container_width=True,
        )

        st.caption(
            "Tus conversaciones y tu memoria quedan guardadas en tu "
            "cuenta y las tendrás en cualquier dispositivo."
        )


# Identidad del usuario (si el login está configurado y aún no ha
# entrado, aquí se muestra la pantalla de acceso y se detiene todo).
_identidad = obtener_identidad()

# Cargamos sus datos al entrar o si cambia de cuenta.
if st.session_state.get("usuario_id") != _identidad["id"]:

    _datos_usuario, _ok = cargar_datos_usuario(_identidad["id"])

    if not _ok:
        st.error(
            "No se han podido cargar tus datos en este momento. "
            "Recarga la página en unos segundos. (Por seguridad no "
            "continúo para no sobrescribir tus conversaciones.)"
        )
        st.stop()

    st.session_state.usuario_id = _identidad["id"]
    st.session_state.usuario_nombre = _identidad["nombre"]
    st.session_state.es_invitado = _identidad["invitado"]
    st.session_state.chats = _datos_usuario["chats"]
    st.session_state.memoria = _datos_usuario["memoria"]
    st.session_state.current_chat_id = None

if "current_chat_id" not in st.session_state:
    st.session_state.current_chat_id = None


def guardar_todo():
    """Guarda las conversaciones y la memoria del usuario actual."""
    guardar_datos_usuario(
        st.session_state.usuario_id,
        st.session_state.chats,
        st.session_state.memoria,
    )


# ============================================================
# 9. BARRA LATERAL
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
            <svg width="26" height="26" viewBox="0 0 56 56"
                 xmlns="http://www.w3.org/2000/svg">
                <circle cx="28" cy="28" r="25" fill="none"
                        stroke="#2B2B31" stroke-width="3"/>
                <circle cx="28" cy="28" r="9" fill="#2B2B31"/>
                <circle cx="45" cy="13" r="3.5" fill="#2B2B31"/>
            </svg>
            <span style="font-size:1.3rem;font-weight:700;letter-spacing:0.04em;">ORIA</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Conversaciones")

    # --------------------------------------------
    # NUEVA CONVERSACIÓN
    # --------------------------------------------

    if st.button(
        "➕ Nueva conversación",
        use_container_width=True,
        type="primary",
    ):

        st.session_state.current_chat_id = None

        st.rerun()

    st.markdown("---")

    # --------------------------------------------
    # LISTA DE CONVERSACIONES
    # --------------------------------------------

    chats_a_borrar = []

    for cid, chat_info in reversed(
        list(st.session_state.chats.items())
    ):

        col_btn, col_del = st.columns(
            [0.82, 0.18]
        )

        # Determinar botón seleccionado.
        if cid == st.session_state.current_chat_id:
            btn_type = "primary"
        else:
            btn_type = "secondary"

        # ----------------------------------------
        # BOTÓN CHAT
        # ----------------------------------------

        with col_btn:

            titulo_chat = chat_info.get(
                "title",
                "Nueva conversación",
            )

            if st.button(
                f"💬 {titulo_chat}",
                key=f"chat_{cid}",
                use_container_width=True,
                type=btn_type,
            ):

                st.session_state.current_chat_id = cid

                st.rerun()

        # ----------------------------------------
        # BOTÓN BORRAR
        # ----------------------------------------

        with col_del:

            if st.button(
                "🗑️",
                key=f"del_{cid}",
            ):

                chats_a_borrar.append(cid)

    # --------------------------------------------
    # BORRAR CHATS
    # --------------------------------------------

    if chats_a_borrar:

        for cid in chats_a_borrar:

            if cid in st.session_state.chats:

                del st.session_state.chats[cid]

            if (
                st.session_state.current_chat_id
                == cid
            ):

                st.session_state.current_chat_id = None

        guardar_todo()

        st.rerun()

    # --------------------------------------------
    # MEMORIA DE ORIA
    # --------------------------------------------

    st.markdown("---")

    with st.expander("🧠 Memoria de ORIA"):

        st.caption(
            "Escribe aquí datos que quieras que ORIA recuerde "
            "siempre (tu nombre, tus preferencias, tu contexto...). "
            "Se incluirán en todas las conversaciones."
        )

        nueva_memoria = st.text_area(
            "Datos a recordar",
            value=st.session_state.memoria,
            height=150,
            label_visibility="collapsed",
        )

        if st.button("Guardar memoria", use_container_width=True):
            st.session_state.memoria = nueva_memoria
            guardar_todo()
            st.success("Memoria guardada.")

    # --------------------------------------------
    # CALIDAD DE LAS IMÁGENES (OPCIONAL)
    # --------------------------------------------

    with st.expander("🎨 Sobre la calidad de las imágenes"):

        st.caption(
            "ORIA ya mejora automáticamente tu descripción antes de "
            "generar la imagen (más detalle, mejor estilo). Se usa "
            "el generador gratuito y anónimo de Pollinations.ai, que "
            "no requiere cuenta ni pago, pero por eso incluye una "
            "pequeña marca de agua y a veces algún error puntual bajo "
            "mucha demanda — es el límite normal de una herramienta "
            "100% gratuita. Si alguna vez quieres quitarlo del todo, "
            "existen servicios de pago (OpenAI, Google) con mejor "
            "calidad y sin marca de agua."
        )

    # --------------------------------------------
    # COMPARTIR ORIA
    # --------------------------------------------

    with st.expander("🔗 Compartir ORIA"):
        st.caption(
            "Comparte el enlace de esta página. Cada persona entra "
            "con su propia cuenta y tiene su historial y memoria "
            "separados del tuyo. En el móvil pueden usar 'Añadir a "
            "pantalla de inicio' para que funcione como una app."
        )

    # --------------------------------------------
    # CUENTA
    # --------------------------------------------

    st.markdown("---")

    if st.session_state.get("es_invitado"):
        st.caption(
            "👤 Modo invitado: el inicio de sesión con Google aún no "
            "está configurado."
        )
    else:
        st.caption(f"👤 {st.session_state.get('usuario_nombre', '')}")
        st.button(
            "Cerrar sesión",
            on_click=st.logout,
            use_container_width=True,
        )

    if supabase_activo():
        st.caption("☁️ Datos guardados en la nube")
    else:
        st.caption(
            "⚠️ Datos guardados solo en el servidor (pueden perderse "
            "al reiniciar). Configura Supabase para guardarlos "
            "en la nube."
        )


# ============================================================
# 10. OBTENER CHAT ACTUAL
# ============================================================

if (
    st.session_state.current_chat_id
    and
    st.session_state.current_chat_id
    in st.session_state.chats
):

    mensajes_actuales = (
        st.session_state
        .chats[
            st.session_state.current_chat_id
        ]
        .get("messages", [])
    )

else:

    mensajes_actuales = []


# ============================================================
# 11. PANTALLA INICIAL
# ============================================================

if len(mensajes_actuales) == 0:

    st.markdown(
        "<br><br><br>",
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(
        [1, 2, 1]
    )

    with col2:

        st.markdown(
            """
            <div class="oria-logo-container">
                <svg width="60" height="60" viewBox="0 0 56 56"
                     xmlns="http://www.w3.org/2000/svg">
                    <circle cx="28" cy="28" r="25" fill="none"
                            stroke="#2B2B31" stroke-width="2.5"/>
                    <circle cx="28" cy="28" r="9" fill="#2B2B31"/>
                    <circle cx="45" cy="13" r="3.5" fill="#2B2B31"/>
                </svg>
                <h1 style="
                    text-align: center;
                    font-size: 3rem;
                    font-weight: 700;
                    letter-spacing: 0.05em;
                    margin: 0;
                    color: #1A1A1A;
                ">
                    ORIA
                </h1>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <h3 style="
                text-align: center;
                color: #666;
                font-weight: 400;
            ">
                ¿En qué te puedo ayudar hoy?
            </h3>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            "<br>",
            unsafe_allow_html=True,
        )


# ============================================================
# 12. MOSTRAR MENSAJES ANTERIORES
# ============================================================

for indice_mensaje, message in enumerate(mensajes_actuales):

    role = message.get("role")
    tipo = message.get("type", "text")
    content = message.get("content", "")

    # --------------------------------------------
    # MENSAJE DEL USUARIO
    # --------------------------------------------

    if role == "user":

        # Escapamos HTML para evitar que el contenido
        # introduzca etiquetas HTML directamente.
        contenido_seguro = html.escape(
            str(content)
        ).replace("\n", "<br>")

        content_html = (
            '<div class="user-bubble-container">'
            '<div class="user-bubble">'
            f"{contenido_seguro}"
            "</div>"
            "</div>"
        )

        st.markdown(
            content_html,
            unsafe_allow_html=True,
        )

    # --------------------------------------------
    # MENSAJE DE LA IA (TEXTO)
    # --------------------------------------------

    elif role == "assistant" and tipo == "text":

        with st.chat_message("assistant"):

            st.markdown(
                str(content)
            )

            # Botones de acción bajo la respuesta: escuchar en voz
            # (sintetizador del navegador, gratis) y copiar al
            # portapapeles (API del navegador, gratis). No necesitan
            # ninguna llamada al servidor.
            texto_js = json.dumps(str(content))
            id_boton_copiar = f"copiar_{indice_mensaje}"

            components.html(
                f"""
                <div style="margin-top:-6px;display:flex;gap:8px;">
                  <button onclick='
                    window.speechSynthesis.cancel();
                    var u = new SpeechSynthesisUtterance({texto_js});
                    u.lang = "es-ES";
                    window.speechSynthesis.speak(u);
                  ' style="
                    background:#F5F5F4;border:none;border-radius:14px;
                    padding:4px 12px;font-size:0.78rem;cursor:pointer;
                    color:#333;
                  ">🔊 Escuchar</button>

                  <button id="{id_boton_copiar}" onclick='
                    navigator.clipboard.writeText({texto_js});
                    var b = document.getElementById("{id_boton_copiar}");
                    b.innerText = "✅ Copiado";
                    setTimeout(function() {{ b.innerText = "📋 Copiar"; }}, 1500);
                  ' style="
                    background:#F5F5F4;border:none;border-radius:14px;
                    padding:4px 12px;font-size:0.78rem;cursor:pointer;
                    color:#333;
                  ">📋 Copiar</button>
                </div>
                """,
                height=36,
            )


    # --------------------------------------------
    # MENSAJE DE LA IA (IMAGEN GENERADA)
    # --------------------------------------------

    elif role == "assistant" and tipo == "image":

        with st.chat_message("assistant"):

            if os.path.exists(content):
                st.image(
                    content,
                    caption=message.get("prompt", ""),
                )
            else:
                st.markdown(
                    "⚠️ *(La imagen generada ya no está disponible)*"
                )


# ============================================================
# 13. INPUT DEL CHAT
# ============================================================

col_toggle_img, col_toggle_voz, col_relleno = st.columns([0.16, 0.14, 0.70])

with col_toggle_img:
    modo_imagen = st.toggle(
        "🎨 Imagen",
        key="modo_imagen",
        help=(
            "Actívalo y escribe lo que quieras que ORIA dibuje. "
            "Desactívalo para volver al chat normal."
        ),
    )

with col_toggle_voz:
    modo_voz = st.toggle(
        "🎤 Voz",
        key="modo_voz",
        help=(
            "Graba tu pregunta con el micrófono y ORIA te responde "
            "en texto. Pulsa 🔊 en su respuesta para escucharla en voz."
        ),
    )

# --------------------------------------------------------------
# MODO VOZ: grabar y transcribir automáticamente
# --------------------------------------------------------------

texto_por_voz = None

if modo_voz:

    audio_grabado = st.audio_input("Pulsa para grabar tu pregunta")

    if audio_grabado is not None:

        audio_bytes = audio_grabado.getvalue()
        audio_hash = hash(audio_bytes)

        if st.session_state.get("ultimo_audio_procesado") != audio_hash:

            with st.spinner("Transcribiendo tu voz..."):
                texto_transcrito, error_audio = transcribir_audio(audio_bytes)

            if error_audio:
                st.error(f"No se pudo transcribir el audio: {error_audio}")
            elif texto_transcrito:
                st.session_state.ultimo_audio_procesado = audio_hash
                texto_por_voz = texto_transcrito

entrada = st.chat_input(
    "Pregunta a ORIA, o adjunta una imagen/PDF...",
    accept_file=True,
    file_type=["png", "jpg", "jpeg", "pdf"],
)


if entrada or texto_por_voz:

    if entrada:
        user_text = (entrada.text or "").strip()
        archivo_adjunto = entrada.files[0] if entrada.files else None
    else:
        user_text = texto_por_voz.strip()
        archivo_adjunto = None

    if not user_text and not archivo_adjunto:
        st.stop()

    # ========================================================
    # CREAR NUEVA CONVERSACIÓN SI NO EXISTE
    # ========================================================

    if (
        not st.session_state.current_chat_id
        or
        st.session_state.current_chat_id
        not in st.session_state.chats
    ):

        nuevo_id = str(
            uuid.uuid4()
        )

        titulo_base = user_text if user_text else (
            f"Archivo: {archivo_adjunto.name}" if archivo_adjunto else "Nueva conversación"
        )

        titulo = titulo_base.strip()

        if len(titulo) > 26:
            titulo = titulo[:26] + "..."

        titulo = titulo.capitalize()

        st.session_state.chats[
            nuevo_id
        ] = {
            "title": titulo,
            "messages": [],
        }

        st.session_state.current_chat_id = (
            nuevo_id
        )

    # ========================================================
    # GUARDAR MENSAJE DEL USUARIO
    # ========================================================

    chat_actual = (
        st.session_state
        .chats[
            st.session_state.current_chat_id
        ]
    )

    texto_mostrado = user_text
    if archivo_adjunto:
        etiqueta_archivo = f"📎 {archivo_adjunto.name}"
        texto_mostrado = (
            f"{user_text}\n\n{etiqueta_archivo}"
            if user_text else etiqueta_archivo
        )

    chat_actual["messages"].append(
        {
            "role": "user",
            "type": "text",
            "content": texto_mostrado,
        }
    )

    guardar_todo()

    # ========================================================
    # MOSTRAR MENSAJE DEL USUARIO
    # ========================================================

    user_text_html = html.escape(
        texto_mostrado
    ).replace("\n", "<br>")

    content_html = (
        '<div class="user-bubble-container">'
        '<div class="user-bubble">'
        f"{user_text_html}"
        "</div>"
        "</div>"
    )

    st.markdown(
        content_html,
        unsafe_allow_html=True,
    )

    # ========================================================
    # GENERAR RESPUESTA DE ORIA
    # ========================================================

    respuesta_final = None

    with st.chat_message("assistant"):

        # ----------------------------------------------------
        # MODO: GENERAR IMAGEN
        # ----------------------------------------------------

        if modo_imagen and user_text:

            with st.spinner("Puliendo la descripción..."):
                prompt_mejorado = mejorar_prompt_imagen(user_text)

            with st.spinner("Generando imagen..."):
                imagen_bytes, error = generar_imagen_ia(prompt_mejorado)

            if error:
                st.error(f"No se pudo generar la imagen: {error}")
                respuesta_final = {
                    "role": "assistant",
                    "type": "text",
                    "content": f"⚠️ No he podido generar la imagen: {error}",
                }
            else:
                os.makedirs(CARPETA_IMAGENES, exist_ok=True)
                nombre_archivo = os.path.join(
                    CARPETA_IMAGENES, f"{uuid.uuid4()}.png"
                )

                with open(nombre_archivo, "wb") as f:
                    f.write(imagen_bytes)

                st.image(imagen_bytes, caption=user_text)

                respuesta_final = {
                    "role": "assistant",
                    "type": "image",
                    "content": nombre_archivo,
                    "prompt": user_text,
                }

        # ----------------------------------------------------
        # MODO: PDF ADJUNTO
        # ----------------------------------------------------

        elif archivo_adjunto and archivo_adjunto.type == "application/pdf":

            with st.spinner("Leyendo el PDF..."):
                texto_pdf, error = extraer_texto_pdf(archivo_adjunto)

            if error:
                st.error(error)
                respuesta_final = {
                    "role": "assistant",
                    "type": "text",
                    "content": f"⚠️ No he podido leer el PDF: {error}",
                }
            else:
                pregunta = user_text if user_text else (
                    "Resume este documento y destaca los puntos clave."
                )

                prompt_aumentado = (
                    f"El usuario ha subido un PDF llamado "
                    f"'{archivo_adjunto.name}'. Este es el contenido "
                    f"extraído del documento:\n\n{texto_pdf}\n\n"
                    f"Petición del usuario sobre el documento: {pregunta}"
                )

                respuesta_texto = st.write_stream(
                    obtener_respuesta_ia_stream(
                        prompt_aumentado,
                        historial_mensajes=None,
                        memoria_texto=st.session_state.memoria,
                    )
                )

                respuesta_final = {
                    "role": "assistant",
                    "type": "text",
                    "content": respuesta_texto,
                }

        # ----------------------------------------------------
        # MODO: IMAGEN ADJUNTA (visión)
        # ----------------------------------------------------

        elif archivo_adjunto and archivo_adjunto.type in (
            "image/png", "image/jpeg", "image/jpg"
        ):

            data_uri = imagen_a_data_uri(archivo_adjunto)

            pregunta = user_text if user_text else (
                "Describe esta imagen y explica qué ves con detalle."
            )

            respuesta_texto = st.write_stream(
                obtener_respuesta_ia_stream(
                    pregunta,
                    historial_mensajes=None,
                    imagen_data_uri=data_uri,
                    memoria_texto=st.session_state.memoria,
                )
            )

            respuesta_final = {
                "role": "assistant",
                "type": "text",
                "content": respuesta_texto,
            }

        # ----------------------------------------------------
        # MODO: CHAT NORMAL
        # ----------------------------------------------------

        else:

            respuesta_texto = st.write_stream(
                obtener_respuesta_ia_stream(
                    user_text,
                    historial_mensajes=chat_actual["messages"],
                    memoria_texto=st.session_state.memoria,
                )
            )

            respuesta_final = {
                "role": "assistant",
                "type": "text",
                "content": respuesta_texto,
            }

    # ========================================================
    # GUARDAR RESPUESTA DE LA IA
    # ========================================================

    chat_actual["messages"].append(respuesta_final)

    guardar_todo()

    # ========================================================
    # RECARGAR
    # ========================================================

    st.rerun()
