import json
import os
import uuid
import html
import base64
import hashlib
import re
from datetime import datetime, timezone
from urllib.parse import quote, urlparse
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


_CACHE_CLIENTE_SUPABASE = {}


def obtener_cliente_supabase_auth():
    """Cliente de supabase-py, solo para el login por email (código de
    un solo uso). El resto de datos (chats y memoria) se sigue
    leyendo/escribiendo con requests, como ya tenías, para no tocar
    lo que ya funciona.

    Importante: solo guardamos en caché el cliente cuando SÍ se ha
    podido crear. Si en algún momento devolviera None (por ejemplo,
    justo después de desplegar, antes de pegar las claves en
    Secrets) y usáramos @st.cache_resource, ese None quedaría
    guardado para siempre mientras la app siga encendida, aunque
    luego añadas las claves correctamente — parecería que "no
    funciona" sin ningún motivo aparente. Así, cada vez que falte
    algo lo volvemos a comprobar."""

    if "cliente" in _CACHE_CLIENTE_SUPABASE:
        return _CACHE_CLIENTE_SUPABASE["cliente"]

    url, key = _config_supabase()

    if not (url and key):
        return None

    try:
        from supabase import create_client
        cliente = create_client(url, key)
        _CACHE_CLIENTE_SUPABASE["cliente"] = cliente
        return cliente
    except Exception as e:
        st.session_state["_error_supabase_auth"] = (
            f"{type(e).__name__}: {e}"
        )
        return None


def enviar_codigo_login(email):
    """Pide a Supabase que mande un código de un solo uso a ese
    correo. Devuelve (ok, error)."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return False, "El login por correo no está configurado."

    try:
        cliente.auth.sign_in_with_otp(
            {
                "email": email,
                "options": {"should_create_user": True},
            }
        )
        return True, None

    except Exception as e:
        return False, str(e)


def verificar_codigo_login(email, codigo):
    """Comprueba el código de 6 dígitos. Devuelve (email_confirmado, error)."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return None, "El login por correo no está configurado."

    try:
        resultado = cliente.auth.verify_otp(
            {"email": email, "token": codigo.strip(), "type": "email"}
        )

        usuario = getattr(resultado, "user", None)
        email_confirmado = getattr(usuario, "email", None) if usuario else None

        if email_confirmado:
            return email_confirmado.strip().lower(), None

        return None, "Código incorrecto."

    except Exception as e:
        mensaje = str(e)

        if "expired" in mensaje.lower() or "invalid" in mensaje.lower():
            mensaje = "El código no es válido o ha caducado. Pide uno nuevo."

        return None, mensaje


def _traducir_error_auth(mensaje):
    """Traduce los mensajes de error más habituales de Supabase Auth."""

    m = mensaje.lower()

    if "already registered" in m or "already exists" in m:
        return "Ya existe una cuenta con ese correo. Prueba a iniciar sesión."
    if "invalid login credentials" in m:
        return "Correo o contraseña incorrectos."
    if "email not confirmed" in m:
        return "Todavía no has confirmado tu correo. Revisa tu bandeja de entrada."
    if "password" in m and ("6 char" in m or "at least" in m or "short" in m):
        return "La contraseña debe tener al menos 6 caracteres."
    if "expired" in m or "invalid" in m and "token" in m:
        return "El código no es válido o ha caducado. Pide uno nuevo."
    if "rate limit" in m or "429" in m:
        return "Has hecho demasiados intentos seguidos. Espera un minuto."

    return mensaje


def registrar_usuario(email, password):
    """Crea la cuenta y dispara el correo de confirmación de Supabase.
    Devuelve (ok, error)."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return False, "El registro no está configurado."

    try:
        resultado = cliente.auth.sign_up({"email": email, "password": password})

        usuario = getattr(resultado, "user", None)
        identidades = getattr(usuario, "identities", None) if usuario else None

        # Supabase, por seguridad, no siempre avisa si el correo ya
        # existía: cuando pasa esto, devuelve un usuario sin
        # identidades nuevas asociadas.
        if usuario is not None and identidades == []:
            return False, (
                "Ya existe una cuenta con ese correo. Prueba a iniciar "
                "sesión o a recuperar tu contraseña."
            )

        return True, None

    except Exception as e:
        return False, _traducir_error_auth(str(e))


def reenviar_confirmacion(email):
    """Vuelve a mandar el correo de confirmación de la cuenta."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return False, "No configurado."

    try:
        cliente.auth.resend({"type": "signup", "email": email})
        return True, None
    except Exception as e:
        return False, _traducir_error_auth(str(e))


def iniciar_sesion_password(email, password):
    """Entra con correo + contraseña. Devuelve (email, error)."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return None, "El login no está configurado."

    try:
        resultado = cliente.auth.sign_in_with_password(
            {"email": email, "password": password}
        )

        usuario = getattr(resultado, "user", None)
        email_confirmado = getattr(usuario, "email", None) if usuario else None

        if email_confirmado:
            return email_confirmado.strip().lower(), None

        return None, "No se ha podido iniciar sesión."

    except Exception as e:
        return None, _traducir_error_auth(str(e))


def solicitar_recuperacion(email):
    """Manda el código para restablecer la contraseña."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return False, "No configurado."

    try:
        cliente.auth.reset_password_email(email)
        return True, None
    except Exception as e:
        return False, _traducir_error_auth(str(e))


def verificar_codigo_recuperacion(email, codigo):
    """Comprueba el código de recuperación. Deja la sesión abierta en
    el cliente para poder cambiar la contraseña justo después."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return None, "No configurado."

    try:
        resultado = cliente.auth.verify_otp(
            {"email": email, "token": codigo.strip(), "type": "recovery"}
        )

        usuario = getattr(resultado, "user", None)
        email_confirmado = getattr(usuario, "email", None) if usuario else None

        if email_confirmado:
            return email_confirmado.strip().lower(), None

        return None, "Código incorrecto."

    except Exception as e:
        return None, _traducir_error_auth(str(e))


def guardar_password_nueva(password_nueva):
    """Cambia la contraseña usando la sesión de recuperación activa."""

    cliente = obtener_cliente_supabase_auth()

    if not cliente:
        return False, "No configurado."

    try:
        cliente.auth.update_user({"password": password_nueva})
        return True, None
    except Exception as e:
        return False, _traducir_error_auth(str(e))


def borrar_cuenta_usuario(usuario_id):
    """Borra los datos guardados (chats y memoria) de este usuario,
    en la nube y localmente. La cuenta de acceso en sí (el correo
    registrado en Supabase Auth) no se borra desde aquí."""

    url, key = _config_supabase()

    if url and key:
        try:
            requests.delete(
                f"{url}/rest/v1/{TABLA_SUPABASE}",
                headers=_headers_supabase(key),
                params={"email": f"eq.{usuario_id}"},
                timeout=20,
            )
        except Exception:
            pass

    try:
        ruta = _ruta_datos_usuario(usuario_id)
        if os.path.exists(ruta):
            os.remove(ruta)
    except Exception:
        pass


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
    """Decide quién es el usuario actual. Si hay algún método de
    inicio de sesión configurado (Google y/o correo con Supabase),
    obliga a entrar por uno de ellos antes de seguir. Devuelve un
    diccionario con id, nombre y si es invitado."""

    # 1) Ya ha entrado por correo (código verificado en esta sesión).
    email_por_correo = st.session_state.get("email_login_verificado")

    if email_por_correo:
        return {"id": email_por_correo, "nombre": email_por_correo, "invitado": False}

    # 2) Ya ha entrado con Google.
    if login_configurado() and getattr(st.user, "is_logged_in", False):

        email = str(getattr(st.user, "email", "") or "").strip().lower()
        nombre = str(getattr(st.user, "name", "") or "").strip() or email

        if not email:
            st.error(
                "Google no ha devuelto tu correo. Cierra sesión e "
                "inténtalo de nuevo."
            )
            st.button("Cerrar sesión", on_click=st.logout)
            st.stop()

        st.session_state.metodo_login = "google"
        return {"id": email, "nombre": nombre, "invitado": False}

    # 3) Hay algún método configurado, pero aún no ha entrado por
    #    ninguno: mostramos la pantalla de acceso y paramos aquí.
    if login_configurado() or obtener_cliente_supabase_auth():
        mostrar_pantalla_login()
        st.stop()

    # 4) Nada configurado todavía: modo invitado (como hasta ahora).
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

hora_hoy_str = ahora.strftime("%H:%M")


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
# 5B. BÚSQUEDA WEB CON FUENTES (Tavily, plan gratuito)
# ============================================================

# Frases que indican que la pregunta depende de datos actuales
# (deportes, noticias, precios...). Si aparecen, ORIA busca en la web
# automáticamente; además el interruptor "🌐 Web" fuerza la búsqueda.
PATRON_ACTUALIDAD = re.compile(
    r"(partido|juega|juegan|jugó|jugaron|resultado|marcador|"
    r"clasificaci[oó]n|en directo|en vivo|qui[eé]n gan[oó]|"
    r"cu[aá]ndo juega|a qu[eé] hora|noticias|[uú]ltima hora|"
    r"[uú]ltimas noticias|ahora mismo|actualmente|cotizaci[oó]n|"
    r"bitcoin|el tiempo en|previsi[oó]n del tiempo|cartelera|estreno)",
    re.IGNORECASE,
)


def obtener_tavily_key():
    """API key gratuita de Tavily (tavily.com) para buscar en la web."""

    try:
        if "TAVILY_API_KEY" not in st.secrets:
            return None

        clave = (
            str(st.secrets["TAVILY_API_KEY"])
            .strip()
            .strip('"')
            .strip("'")
            .strip()
        )

        return clave or None

    except Exception:
        return None


def necesita_busqueda(texto):
    """True si el mensaje parece depender de información actual."""
    return bool(PATRON_ACTUALIDAD.search(texto or ""))


def buscar_en_web(consulta):
    """Busca en la web con Tavily. Devuelve (resultados, error).
    Cada resultado es {"titulo", "url", "texto"}. Si el error es
    "no_configurado", falta TAVILY_API_KEY en Secrets."""

    clave = obtener_tavily_key()

    if not clave:
        return [], "no_configurado"

    es_actualidad = necesita_busqueda(consulta)

    consulta_final = consulta.strip()[:300]

    if es_actualidad:
        # Añadimos la fecha para que salgan resultados de hoy.
        consulta_final += f" {ahora.day} {meses[ahora.month - 1]} {ahora.year}"

    headers = {
        "Authorization": f"Bearer {clave}",
        "Content-Type": "application/json",
    }

    intentos = [{"topic": "news", "time_range": "week"}, {"topic": "general"}]

    if not es_actualidad:
        intentos = [{"topic": "general"}]

    ultimo_error = None

    for extra in intentos:

        payload = {
            "query": consulta_final,
            "search_depth": "basic",
            "max_results": 5,
            "include_answer": False,
        }
        payload.update(extra)

        try:
            respuesta = requests.post(
                "https://api.tavily.com/search",
                headers=headers,
                json=payload,
                timeout=20,
            )

            if respuesta.status_code != 200:
                ultimo_error = f"error_{respuesta.status_code}"
                continue

            resultados = []

            for r in respuesta.json().get("results", [])[:5]:

                url = r.get("url")

                if not url:
                    continue

                resultados.append(
                    {
                        "titulo": (r.get("title") or url)[:120],
                        "url": url,
                        "texto": (r.get("content") or "")[:600],
                    }
                )

            if resultados:
                return resultados, None

            ultimo_error = "sin_resultados"

        except Exception as e:
            ultimo_error = str(e)

    return [], ultimo_error


def formatear_contexto_web(resultados):
    """Convierte los resultados en el texto numerado que ve el modelo."""

    bloques = []

    for i, r in enumerate(resultados, 1):
        bloques.append(f"[{i}] {r['titulo']} ({r['url']})\n{r['texto']}")

    return "\n\n".join(bloques)


# ============================================================
# 6. FUNCIÓN PRINCIPAL DE LA IA
# ============================================================

def obtener_respuesta_ia_stream(
    prompt_usuario,
    historial_mensajes=None,
    imagen_data_uri=None,
    memoria_texto="",
    contexto_web="",
    web_fallida=False,
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
        f"La fecha actual es {fecha_hoy_str} y la hora en España "
        f"son las {hora_hoy_str}. "
        "Responde siempre en español, salvo que el usuario "
        "pida expresamente otro idioma. "
        "Explica las cosas de forma clara y útil. "
        "Nunca inventes marcadores, resultados, noticias, horarios "
        "ni precios: solo puedes darlos si aparecen en los "
        "resultados de búsqueda web que se te proporcionan."
    )

    if contexto_web:
        instrucciones_sistema += (
            "\n\nSe ha hecho una búsqueda web ahora mismo. Usa estos "
            "resultados para responder con datos actuales y cita las "
            "fuentes en el texto con su número entre corchetes, por "
            "ejemplo [1] o [2]. Si los resultados no responden con "
            "claridad o se contradicen, dilo con honestidad en vez de "
            "inventar. Los horarios conviértelos a hora de España. "
            "Resultados:\n\n" + contexto_web
        )

    elif web_fallida:
        instrucciones_sistema += (
            "\n\nSe intentó buscar en la web pero no ha sido posible. "
            "Si la pregunta depende de datos de última hora, dilo "
            "claramente y no inventes nada."
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
        PRESUPUESTO_CARACTERES_HISTORIAL = 6000 if contexto_web else 11000

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
[data-testid="stToolbar"],
footer {
    visibility: hidden;
    height: 0;
}

/* La cabecera en sí la dejamos "existir" (transparente, sin borde)
en lugar de ocultarla del todo: dentro de ella vive el botón que
abre la barra lateral en el móvil (flecha »). Si se oculta toda la
cabecera, ese botón desaparece con ella y en el móvil no hay forma
de abrir las conversaciones ni la cuenta — por eso solo la
"vaciamos" visualmente. */
header[data-testid="stHeader"] {
    background: transparent !important;
    box-shadow: none !important;
}

/* El botón de abrir/cerrar la barra lateral tiene que verse bien y
ser fácil de tocar en el móvil (el nombre interno de este elemento
ha cambiado entre versiones de Streamlit, por eso cubrimos los dos). */
[data-testid="stSidebarCollapsedControl"],
[data-testid="collapsedControl"] {
    visibility: visible !important;
    opacity: 1 !important;
    z-index: 1000 !important;
    background-color: #FFFFFF !important;
    border: 1px solid #E2E2DD !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08) !important;
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

/* Logo + nombre: centrados en el mismo eje */
.oria-brand {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 14px;
    width: 100%;
    text-align: center;
}

.oria-brand svg {
    display: block;
    margin: 0 auto;
}

.oria-brand-name {
    font-size: 3rem;
    font-weight: 700;
    line-height: 1;
    letter-spacing: 0.14em;
    /* compensa el espacio que deja el letter-spacing tras la última
    letra, para que el texto quede centrado de verdad */
    padding-left: 0.14em;
    color: #1A1A1A;
}

.oria-tagline {
    width: 100%;
    text-align: center;
    color: #666;
    font-size: 1.35rem;
    font-weight: 400;
    margin-top: 18px;
}

/* Interruptores Imagen / Voz / Web: fijos justo encima del cuadro
de texto, alineados a la izquierda */
.st-key-barra_modos {
    position: fixed;
    bottom: 7.4rem;
    z-index: 999;
    display: flex !important;
    flex-direction: row !important;
    align-items: center;
    gap: 1.4rem !important;
    width: auto !important;
    background: transparent;
}

.st-key-barra_modos > div {
    width: auto !important;
    flex: 0 0 auto !important;
}

/* Deja hueco abajo para que la barra fija no tape los mensajes */
[data-testid="stMainBlockContainer"] {
    padding-bottom: 10rem !important;
}

/* ------------------------------------------------------------
Barra lateral limpia, estilo Claude: fondo ligeramente distinto
al del chat, botones sin caja ni sombra (solo texto + icono, con
un resaltado suave al pasar el ratón o cuando están seleccionados)
y una fila de cuenta fija abajo del todo.
------------------------------------------------------------ */

[data-testid="stSidebar"] {
    background-color: #FAFAF9 !important;
    border-right: 1px solid #ECECE9;
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 0.8rem;
}

.stSidebar .stButton > button {
    border: none !important;
    box-shadow: none !important;
    background-color: transparent !important;
    color: #3A3A38 !important;
    text-align: left !important;
    justify-content: flex-start !important;
    font-weight: 400 !important;
    border-radius: 8px !important;
    padding: 0.45rem 0.6rem !important;
    transition: background-color 0.1s ease-in-out;
}

.stSidebar .stButton > button:hover {
    background-color: #EFEFEB !important;
    color: #1A1A1A !important;
}

.stSidebar .stButton > button[kind="primary"] {
    background-color: #E8E7E2 !important;
    color: #1A1A1A !important;
    font-weight: 500 !important;
}

.stSidebar .stButton > button[kind="primary"]:hover {
    background-color: #E0DFDA !important;
}

/* "Nueva conversación": un poco de contorno para que destaque
como acción principal, sin ser un botón pesado. */
.st-key-boton_nueva_conversacion .stButton > button {
    border: 1px solid #E2E2DD !important;
    background-color: #FFFFFF !important;
    font-weight: 500 !important;
    margin-bottom: 0.3rem;
}

.st-key-boton_nueva_conversacion .stButton > button:hover {
    background-color: #F3F3F0 !important;
}

/* Botón de borrar cada conversación: pequeño e icónico, no roba
protagonismo a la lista. */
.st-key-lista_chats .stButton > button {
    padding: 0.45rem 0.3rem !important;
    color: #B5B5AF !important;
}

.st-key-lista_chats .stButton > button:hover {
    color: #C0392B !important;
    background-color: #F6E9E7 !important;
}

/* Fila de cuenta, fija en la parte inferior de la barra lateral. */
.st-key-fila_cuenta {
    border-top: 1px solid #ECECE9;
    padding-top: 0.6rem;
    margin-top: 0.4rem;
}

.st-key-fila_cuenta .stButton > button {
    font-weight: 500 !important;
}

.oria-avatar-circulo {
    width: 28px;
    height: 28px;
    min-width: 28px;
    border-radius: 50%;
    background-color: #2B2B31;
    color: #FFFFFF;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.8rem;
    font-weight: 600;
    text-transform: uppercase;
}

/* Barra lateral propia para el móvil: en vez de depender del botón
interno de Streamlit para abrir/cerrar la barra lateral (no
funciona igual en todas las versiones), la controlamos nosotros
del todo con una clase en <body>. En ordenador no se toca nada de
esto: la barra lateral ya sale abierta sola. */
#oria-boton-menu {
    display: none;
}

#oria-fondo-sidebar {
    display: none;
}

@media (max-width: 768px) {

    #oria-boton-menu {
        display: flex;
        position: fixed;
        top: 0.7rem;
        left: 0.7rem;
        z-index: 1000001;
        width: 40px;
        height: 40px;
        border-radius: 10px;
        background-color: #FFFFFF;
        border: 1px solid #E2E2DD;
        box-shadow: 0 1px 4px rgba(0,0,0,0.12);
        align-items: center;
        justify-content: center;
        cursor: pointer;
        font-size: 1.2rem;
        color: #2B2B31;
        -webkit-tap-highlight-color: transparent;
    }

    /* Deja hueco para que el botón no tape el logo/título de arriba */
    [data-testid="stMainBlockContainer"] {
        padding-top: 3.6rem !important;
    }

    /* La barra lateral, oculta fuera de la pantalla por defecto y
    fija (para que se quede encima de todo, como un cajón). */
    [data-testid="stSidebar"] {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        height: 100vh !important;
        width: 85vw !important;
        max-width: 320px !important;
        z-index: 1000000 !important;
        box-shadow: 2px 0 16px rgba(0,0,0,0.18);
        transform: translateX(-100%);
        transition: transform 0.22s ease-in-out;
    }

    /* Cuando <body> lleva la clase "oria-sidebar-abierta" (la pone
    nuestro botón ☰), la barra lateral entra en pantalla. */
    body.oria-sidebar-abierta [data-testid="stSidebar"] {
        transform: translateX(0) !important;
    }

    /* Fondo oscuro detrás, para poder cerrar tocando fuera. */
    body.oria-sidebar-abierta #oria-fondo-sidebar {
        display: block;
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.4);
        z-index: 999999;
    }
}

/* Ajustes para pantallas de móvil */
@media (max-width: 640px) {

    .user-bubble {
        max-width: 88%;
        padding: 10px 14px;
        font-size: 0.95rem;
    }

    .oria-brand-name {
        font-size: 2.4rem;
    }

    .oria-tagline {
        font-size: 1.1rem;
    }

    .st-key-barra_modos {
        bottom: 5.6rem;
        gap: 1rem !important;
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

# Barra lateral propia para el móvil (no depende de ningún control
# interno de Streamlit, que en algunas versiones no aparece o no se
# puede activar por CSS): un botón ☰ que añade/quita una clase en
# <body>, y un fondo oscuro para poder cerrarla tocando fuera.
#
# Importante: esto NO se hace con st.markdown (el "onclick" dentro
# de HTML puesto con st.markdown puede quedar bloqueado por el
# navegador y por eso no funcionaba). Se hace con
# streamlit.components.v1.html, que sí ejecuta JavaScript de
# verdad. El script vive en un iframe aparte, así que entra en el
# documento real de la página (window.parent.document) y crea ahí
# el botón y el fondo directamente — comprobando primero que no
# existan ya, para no duplicarlos cada vez que la app se actualiza.
components.html(
    """
    <script>
    (function () {
        var doc = window.parent.document;

        var boton = doc.getElementById("oria-boton-menu");
        if (!boton) {
            boton = doc.createElement("div");
            boton.id = "oria-boton-menu";
            boton.innerHTML = "&#9776;";
            doc.body.appendChild(boton);
        }
        boton.onclick = function () {
            doc.body.classList.toggle("oria-sidebar-abierta");
        };

        var fondo = doc.getElementById("oria-fondo-sidebar");
        if (!fondo) {
            fondo = doc.createElement("div");
            fondo.id = "oria-fondo-sidebar";
            doc.body.appendChild(fondo);
        }
        fondo.onclick = function () {
            doc.body.classList.remove("oria-sidebar-abierta");
        };
    })();
    </script>
    """,
    height=0,
)


# ============================================================
# 8. SESSION STATE
# ============================================================

LOGO_HTML = """<div class="oria-brand"><svg width="64" height="64" viewBox="0 0 56 56" xmlns="http://www.w3.org/2000/svg"><circle cx="28" cy="28" r="25" fill="none" stroke="#2B2B31" stroke-width="2.5"/><circle cx="28" cy="28" r="9" fill="#2B2B31"/><circle cx="45" cy="13" r="3.5" fill="#2B2B31"/></svg><div class="oria-brand-name">ORIA</div></div>"""


TEXTO_POLITICA_PRIVACIDAD = """
### Política de privacidad de ORIA

*Última actualización: septiembre de 2026. Este texto es un modelo
orientativo redactado para este proyecto personal y no sustituye el
asesoramiento de un profesional legal.*

**Quién trata tus datos.** ORIA es un proyecto personal, no una
empresa. La persona que ha desplegado esta ORIA es responsable de
los datos que se guardan.

**Qué datos se guardan.** Tu correo electrónico (para identificar tu
cuenta), el historial de tus conversaciones con ORIA, cualquier
imagen o documento PDF que adjuntes mientras hablas con ella, y las
notas que le pidas que recuerde ("Memoria de ORIA").

**Para qué se usan.** Únicamente para que ORIA pueda responderte y
para que tu historial esté disponible la próxima vez que entres,
desde cualquier dispositivo.

**Con quién se comparten.** Tus mensajes se envían a Groq (para
generar las respuestas de texto, voz e imágenes) y, si inicias
sesión por correo, a Supabase (donde se guarda tu cuenta y tu
historial). Ningún dato se vende ni se usa con fines publicitarios.

**Cuánto tiempo se guardan.** Hasta que borres una conversación
concreta o elimines tu cuenta desde la barra lateral.

**Tus derechos.** Puedes borrar tus conversaciones una a una, o
eliminar toda tu cuenta y tus datos en cualquier momento desde
**Cuenta → Eliminar mi cuenta y mis datos**, en la barra lateral.
"""

TEXTO_POLITICA_COOKIES = """
### Política de cookies de ORIA

ORIA usa únicamente las cookies técnicas necesarias para mantener tu
sesión iniciada mientras usas la aplicación (por ejemplo, para
recordar que ya iniciaste sesión sin pedírtelo en cada mensaje). No
se usan cookies de publicidad ni de seguimiento entre otras webs.
"""


def mostrar_pantalla_login():
    """Pantalla de acceso: logo, pestañas de Iniciar sesión / Crear
    cuenta (correo + contraseña con Supabase) y, si está configurado,
    entrar con Google."""

    st.markdown("<br><br>", unsafe_allow_html=True)

    _c1, _c2, _c3 = st.columns([1, 2, 1])

    with _c2:

        st.markdown(LOGO_HTML, unsafe_allow_html=True)

        st.markdown(
            '<div class="oria-tagline">Te damos la bienvenida a ORIA</div>',
            unsafe_allow_html=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)

        correo_disponible = bool(obtener_cliente_supabase_auth())
        pantalla = st.session_state.get("pantalla_login", "login")

        # ------------------------------------------------------
        # RECUPERAR CONTRASEÑA (tiene su propio flujo, aparte de
        # las pestañas)
        # ------------------------------------------------------

        if correo_disponible and pantalla.startswith("recuperar"):
            _mostrar_recuperacion_password(pantalla)

        elif correo_disponible:

            tab_login, tab_registro = st.tabs(
                ["Iniciar sesión", "Crear cuenta"]
            )

            with tab_login:
                _mostrar_form_login()

            with tab_registro:
                _mostrar_form_registro()

        elif not login_configurado():
            st.info(
                "El inicio de sesión aún no está configurado. Revisa "
                "los Secrets de la app."
            )

        elif st.session_state.get("_error_supabase_auth"):
            st.warning(
                "⚠️ El login por correo no ha podido activarse:\n\n"
                f"`{st.session_state['_error_supabase_auth']}`\n\n"
                "Comprueba que `requirements.txt` en GitHub tiene la "
                "línea `supabase>=2.9.0` y que la última versión se "
                "ha vuelto a subir junto con `app.py`."
            )

        else:
            # No hay excepción (el paquete supabase-py está bien
            # instalado), así que si no tenemos cliente es porque no
            # se han encontrado las claves en Secrets. Lo decimos con
            # detalle para no tener que ir a ciegas.
            try:
                _tiene_url = "SUPABASE_URL" in st.secrets
            except Exception:
                _tiene_url = False
            try:
                _tiene_key = "SUPABASE_KEY" in st.secrets
            except Exception:
                _tiene_key = False

            st.warning(
                "⚠️ El login por correo (crear cuenta / contraseña) "
                "todavía no está activo porque no encuentro las claves "
                "de Supabase en los Secrets de la app:\n\n"
                f"- `SUPABASE_URL`: {'✅ encontrada' if _tiene_url else '❌ no encontrada'}\n"
                f"- `SUPABASE_KEY`: {'✅ encontrada' if _tiene_key else '❌ no encontrada'}\n\n"
                "Ve a Streamlit Cloud → tu app (arriba a la derecha, "
                "menú ⋮) → **Settings → Secrets** y comprueba que están "
                "escritas así, tal cual, **antes** de la línea "
                "`[auth]`:\n\n"
                "```\n"
                'SUPABASE_URL = "https://xxxxxxxx.supabase.co"\n'
                'SUPABASE_KEY = "tu_clave_de_supabase"\n'
                "```\n\n"
                "Después de guardarlas, reinicia la app (menú ⋮ → "
                "**Reboot app**)."
            )

        # ------------------------------------------------------
        # GOOGLE
        # ------------------------------------------------------

        if login_configurado() and pantalla == "login":

            if correo_disponible:
                st.markdown(
                    '<p style="text-align:center;color:#999;margin:14px 0;">o</p>',
                    unsafe_allow_html=True,
                )

            st.button(
                "Continuar con Google",
                on_click=st.login,
                use_container_width=True,
            )

        st.caption(
            "Tus conversaciones y tu memoria quedan guardadas en tu "
            "cuenta y las tendrás en cualquier dispositivo."
        )


def _mostrar_form_login():

    with st.form("form_login", clear_on_submit=False):

        email = st.text_input("Correo electrónico", key="li_email")
        password = st.text_input(
            "Contraseña", type="password", key="li_password"
        )

        entrar = st.form_submit_button(
            "Entrar", type="primary", use_container_width=True
        )

    if entrar:

        email = email.strip().lower()

        if not email or not password:
            st.error("Escribe tu correo y tu contraseña.")
        else:
            with st.spinner("Comprobando..."):
                email_ok, error = iniciar_sesion_password(email, password)

            if email_ok:
                st.session_state.email_login_verificado = email_ok
                st.session_state.metodo_login = "email"
                st.rerun()
            else:
                st.error(error)

                if error and "confirmado" in error:
                    if st.button("Reenviar correo de confirmación"):
                        reenviar_confirmacion(email)
                        st.success("Correo reenviado. Revisa tu bandeja.")

    if st.button("¿Has olvidado tu contraseña?", use_container_width=True):
        st.session_state.pantalla_login = "recuperar_pedir"
        st.rerun()


def _mostrar_form_registro():

    if st.session_state.get("registro_hecho"):
        st.success(
            f"Te hemos enviado un correo de confirmación a "
            f"**{st.session_state.registro_hecho}**. Ábrelo y confirma "
            "tu cuenta, y después inicia sesión en la otra pestaña."
        )
        if st.button("Volver", key="volver_tras_registro"):
            del st.session_state["registro_hecho"]
            st.rerun()
        return

    with st.form("form_registro", clear_on_submit=False):

        email = st.text_input("Correo electrónico", key="re_email")
        password = st.text_input(
            "Contraseña", type="password", key="re_password",
            help="Al menos 6 caracteres.",
        )
        password2 = st.text_input(
            "Repite la contraseña", type="password", key="re_password2"
        )

        with st.expander("Política de privacidad y cookies"):
            st.markdown(TEXTO_POLITICA_PRIVACIDAD)
            st.markdown(TEXTO_POLITICA_COOKIES)

        acepto = st.checkbox(
            "He leído y acepto la política de privacidad y de cookies"
        )

        crear = st.form_submit_button(
            "Crear cuenta", type="primary", use_container_width=True
        )

    if crear:

        email = email.strip().lower()

        if "@" not in email or "." not in email:
            st.error("Escribe un correo electrónico válido.")
        elif len(password) < 6:
            st.error("La contraseña debe tener al menos 6 caracteres.")
        elif password != password2:
            st.error("Las dos contraseñas no coinciden.")
        elif not acepto:
            st.error(
                "Tienes que aceptar la política de privacidad y "
                "cookies para crear la cuenta."
            )
        else:
            with st.spinner("Creando la cuenta..."):
                ok, error = registrar_usuario(email, password)

            if ok:
                st.session_state.registro_hecho = email
                st.rerun()
            else:
                st.error(error)


def _mostrar_recuperacion_password(pantalla):

    if pantalla == "recuperar_pedir":

        with st.form("form_recuperar_pedir"):
            email = st.text_input("Tu correo electrónico")
            enviar = st.form_submit_button(
                "Enviar código", type="primary", use_container_width=True
            )

        if enviar:
            email = email.strip().lower()
            with st.spinner("Enviando..."):
                ok, error = solicitar_recuperacion(email)

            if ok:
                st.session_state.recuperar_email = email
                st.session_state.pantalla_login = "recuperar_verificar"
                st.rerun()
            else:
                st.error(error)

    elif pantalla == "recuperar_verificar":

        email = st.session_state.get("recuperar_email", "")
        st.success(f"Te hemos enviado un código a **{email}**.")

        with st.form("form_recuperar_verificar"):
            codigo = st.text_input("Código de 6 dígitos", max_chars=6)
            verificar = st.form_submit_button(
                "Verificar código", type="primary", use_container_width=True
            )

        if verificar:
            with st.spinner("Comprobando..."):
                email_ok, error = verificar_codigo_recuperacion(email, codigo)

            if email_ok:
                st.session_state.pantalla_login = "recuperar_nueva"
                st.rerun()
            else:
                st.error(error)

    elif pantalla == "recuperar_nueva":

        st.info("Escribe tu nueva contraseña.")

        with st.form("form_recuperar_nueva"):
            password = st.text_input(
                "Nueva contraseña", type="password", help="Al menos 6 caracteres."
            )
            password2 = st.text_input("Repite la contraseña", type="password")
            guardar = st.form_submit_button(
                "Guardar nueva contraseña",
                type="primary",
                use_container_width=True,
            )

        if guardar:

            if len(password) < 6:
                st.error("La contraseña debe tener al menos 6 caracteres.")
            elif password != password2:
                st.error("Las dos contraseñas no coinciden.")
            else:
                with st.spinner("Guardando..."):
                    ok, error = guardar_password_nueva(password)

                if ok:
                    st.session_state.email_login_verificado = (
                        st.session_state.get("recuperar_email")
                    )
                    st.session_state.metodo_login = "email"
                    for _clave in (
                        "pantalla_login",
                        "recuperar_email",
                    ):
                        st.session_state.pop(_clave, None)
                    st.success("Contraseña actualizada. Entrando...")
                    st.rerun()
                else:
                    st.error(error)

    if st.button("Cancelar", use_container_width=True):
        for _clave in ("pantalla_login", "recuperar_email"):
            st.session_state.pop(_clave, None)
        st.rerun()


# Identidad del usuario (si el login está configurado y aún no ha
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


def _cerrar_sesion():
    if st.session_state.get("metodo_login") == "google":
        st.logout()
    else:
        for _clave in (
            "email_login_verificado",
            "metodo_login",
            "usuario_id",
            "usuario_nombre",
            "es_invitado",
            "chats",
            "memoria",
            "current_chat_id",
        ):
            st.session_state.pop(_clave, None)


@st.dialog("Ajustes")
def _mostrar_ajustes():
    """Todo lo que no es 'chatear' vive aquí: memoria, estado de las
    integraciones, cómo compartir ORIA y la cuenta. Así la barra
    lateral se queda solo con lo esencial (como en Claude) y esto
    se abre aparte, cuando de verdad se necesita."""

    tab_memoria, tab_cuenta, tab_info = st.tabs(
        ["🧠 Memoria", "👤 Cuenta", "ℹ️ Estado"]
    )

    # --------------------------------------------
    # MEMORIA DE ORIA
    # --------------------------------------------

    with tab_memoria:

        st.caption(
            "Escribe aquí datos que quieras que ORIA recuerde siempre "
            "(tu nombre, tus preferencias, tu contexto...). Se "
            "incluirán en todas las conversaciones."
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
    # CUENTA
    # --------------------------------------------

    with tab_cuenta:

        if st.session_state.get("es_invitado"):
            st.caption(
                "👤 Modo invitado: el inicio de sesión aún no está "
                "configurado."
            )
        else:
            st.caption(
                f"Sesión iniciada como "
                f"**{st.session_state.get('usuario_nombre', '')}**."
            )

            st.button(
                "Cerrar sesión",
                on_click=_cerrar_sesion,
                use_container_width=True,
            )

            st.markdown("")

            with st.expander("⚠️ Eliminar mi cuenta y mis datos"):
                st.caption(
                    "Borra todas tus conversaciones y tu memoria de "
                    "forma permanente. No se puede deshacer."
                )
                if st.checkbox(
                    "Sí, quiero eliminar todos mis datos",
                    key="confirmar_borrado",
                ):
                    if st.button(
                        "Eliminar definitivamente",
                        type="primary",
                        use_container_width=True,
                    ):
                        borrar_cuenta_usuario(
                            st.session_state.usuario_id
                        )
                        _cerrar_sesion()
                        st.success("Tus datos se han eliminado.")
                        st.rerun()

        st.markdown("---")
        st.caption(
            "Comparte el enlace de esta página con quien quieras: "
            "cada persona entra con su propia cuenta y tiene su "
            "historial y memoria separados del tuyo. En el móvil "
            "pueden usar 'Añadir a pantalla de inicio' para que "
            "funcione como una app."
        )

    # --------------------------------------------
    # ESTADO DE LAS INTEGRACIONES
    # --------------------------------------------

    with tab_info:

        if obtener_tavily_key():
            st.caption("🌐 Búsqueda web activa")
        else:
            st.caption("🌐 Búsqueda web sin configurar")

        if supabase_activo():
            st.caption("☁️ Datos guardados en la nube")
        else:
            st.caption(
                "⚠️ Datos guardados solo en el servidor (pueden "
                "perderse al reiniciar). Configura Supabase para "
                "guardarlos en la nube."
            )

        st.markdown("---")

        st.caption(
            "ORIA mejora automáticamente tu descripción antes de "
            "generar una imagen. Se usa el generador gratuito y "
            "anónimo de Pollinations.ai, que no requiere cuenta ni "
            "pago, pero por eso incluye una pequeña marca de agua y "
            "a veces algún error puntual bajo mucha demanda — es el "
            "límite normal de una herramienta 100% gratuita."
        )


with st.sidebar:

    st.markdown(
        """
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
            <svg width="24" height="24" viewBox="0 0 56 56"
                 xmlns="http://www.w3.org/2000/svg">
                <circle cx="28" cy="28" r="25" fill="none"
                        stroke="#2B2B31" stroke-width="3"/>
                <circle cx="28" cy="28" r="9" fill="#2B2B31"/>
                <circle cx="45" cy="13" r="3.5" fill="#2B2B31"/>
            </svg>
            <span style="font-size:1.2rem;font-weight:700;letter-spacing:0.04em;">ORIA</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------
    # NUEVA CONVERSACIÓN
    # --------------------------------------------

    with st.container(key="boton_nueva_conversacion"):
        if st.button(
            "＋  Nueva conversación",
            use_container_width=True,
        ):
            st.session_state.current_chat_id = None
            st.rerun()

    # --------------------------------------------
    # LISTA DE CONVERSACIONES
    # --------------------------------------------

    chats_a_borrar = []

    with st.container(key="lista_chats"):

        for cid, chat_info in reversed(
            list(st.session_state.chats.items())
        ):

            col_btn, col_del = st.columns([0.85, 0.15])

            # Determinar botón seleccionado.
            if cid == st.session_state.current_chat_id:
                btn_type = "primary"
            else:
                btn_type = "secondary"

            with col_btn:

                titulo_chat = chat_info.get(
                    "title",
                    "Nueva conversación",
                )

                if st.button(
                    titulo_chat,
                    key=f"chat_{cid}",
                    use_container_width=True,
                    type=btn_type,
                ):
                    st.session_state.current_chat_id = cid
                    st.rerun()

            with col_del:

                if st.button("✕", key=f"del_{cid}"):
                    chats_a_borrar.append(cid)

    if chats_a_borrar:

        for cid in chats_a_borrar:

            if cid in st.session_state.chats:
                del st.session_state.chats[cid]

            if st.session_state.current_chat_id == cid:
                st.session_state.current_chat_id = None

        guardar_todo()
        st.rerun()

    # --------------------------------------------
    # CUENTA (fila fija abajo, abre los ajustes)
    # --------------------------------------------

    with st.container(key="fila_cuenta"):

        nombre_usuario = (
            "Invitado"
            if st.session_state.get("es_invitado")
            else st.session_state.get("usuario_nombre", "")
        )
        inicial = (nombre_usuario or "?").strip()[:1] or "?"

        col_avatar, col_nombre = st.columns([0.18, 0.82])

        with col_avatar:
            st.markdown(
                f'<div class="oria-avatar-circulo">{inicial}</div>',
                unsafe_allow_html=True,
            )

        with col_nombre:
            if st.button(
                nombre_usuario,
                key="abrir_ajustes",
                use_container_width=True,
            ):
                _mostrar_ajustes()


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

        st.markdown(LOGO_HTML, unsafe_allow_html=True)

        st.markdown(
            '<div class="oria-tagline">¿En qué te puedo ayudar hoy?</div>',
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

            # Fuentes consultadas en la web (si las hubo)
            fuentes = message.get("sources") or []

            if fuentes:
                with st.expander(f"🔎 Fuentes ({len(fuentes)})"):
                    for n, fuente in enumerate(fuentes, 1):
                        titulo = (
                            str(fuente.get("title", ""))
                            .replace("[", "(")
                            .replace("]", ")")
                        )
                        url = str(fuente.get("url", ""))
                        dominio = urlparse(url).netloc.replace("www.", "")
                        st.markdown(f"**[{n}]** [{titulo}]({url}) · {dominio}")

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

with st.container(key="barra_modos"):
    modo_imagen = st.toggle("🎨 Imagen", key="modo_imagen")
    modo_voz = st.toggle("🎤 Voz", key="modo_voz")
    modo_web = st.toggle("🌐 Web", key="modo_web")

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

            resultados_web = []
            web_fallida = False

            if modo_web or necesita_busqueda(user_text):

                # Si el mensaje es muy corto ("¿y mañana?"), le
                # sumamos la pregunta anterior para dar contexto.
                consulta = user_text

                if len(user_text) < 25:
                    previos = [
                        m["content"]
                        for m in chat_actual["messages"][:-1]
                        if m.get("role") == "user"
                        and m.get("type", "text") == "text"
                    ]
                    if previos:
                        consulta = f"{previos[-1][:200]} {user_text}"

                with st.spinner("Buscando en la web..."):
                    resultados_web, error_web = buscar_en_web(consulta)

                if not resultados_web:
                    web_fallida = True

                    if error_web == "no_configurado":
                        st.caption(
                            "🌐 La búsqueda web no está configurada "
                            "todavía, así que no puedo confirmar datos "
                            "de última hora."
                        )
                    else:
                        st.caption(
                            "🌐 No he podido consultar la web ahora "
                            "mismo."
                        )

            respuesta_texto = st.write_stream(
                obtener_respuesta_ia_stream(
                    user_text,
                    historial_mensajes=chat_actual["messages"],
                    memoria_texto=st.session_state.memoria,
                    contexto_web=formatear_contexto_web(resultados_web)
                    if resultados_web else "",
                    web_fallida=web_fallida,
                )
            )

            respuesta_final = {
                "role": "assistant",
                "type": "text",
                "content": respuesta_texto,
            }

            if resultados_web:
                respuesta_final["sources"] = [
                    {"title": r["titulo"], "url": r["url"]}
                    for r in resultados_web
                ]

    # ========================================================
    # GUARDAR RESPUESTA DE LA IA
    # ========================================================

    chat_actual["messages"].append(respuesta_final)

    guardar_todo()

    # ========================================================
    # RECARGAR
    # ========================================================

    st.rerun()
