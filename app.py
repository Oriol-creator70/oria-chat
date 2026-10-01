import json
import os
import io
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
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
from reportlab.lib.units import inch as _PDF_INCH
from reportlab.lib.colors import HexColor as _pdf_color
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as _pdf_canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


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
CARPETA_PRESENTACIONES = "presentaciones_generadas"
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


IDIOMA_POR_DEFECTO = "es"

# Idiomas que puede elegir cada usuario para que ORIA le conteste
# siempre en ese idioma (independientemente del idioma en el que
# escriba la pregunta).
IDIOMAS_DISPONIBLES = {
    "es": "Español",
    "en": "English",
    "ca": "Català",
    "fr": "Français",
}

NOMBRE_IDIOMA_PARA_PROMPT = {
    "es": "español",
    "en": "English",
    "ca": "català",
    "fr": "français",
}

# Idioma que usa el sintetizador de voz del navegador (botón
# "Escuchar") para pronunciar bien según el idioma elegido.
CODIGO_VOZ_NAVEGADOR = {
    "es": "es-ES",
    "en": "en-US",
    "ca": "ca-ES",
    "fr": "fr-FR",
}

# ------------------------------------------------------------
# Textos de la interfaz (todo lo que NO es la respuesta de la IA:
# botones, pestañas, avisos...) en los idiomas disponibles. La
# pantalla de inicio de sesión se queda en español (todavía no hay
# ninguna cuenta de la que leer el idioma elegido); en cuanto el
# usuario entra, toda la interfaz pasa a hablar en su idioma.
# ------------------------------------------------------------
TEXTOS = {
    "es": {
        "tagline_bienvenida": "¿En qué te puedo ayudar hoy?",
        "nueva_conversacion": "＋  Nueva conversación",
        "conversacion_sin_titulo": "Nueva conversación",
        "archivo_titulo": "Archivo: {nombre}",
        "ajustes_titulo": "Ajustes",
        "tab_idioma": "Idioma",
        "tab_memoria": "Memoria",
        "tab_cuenta": "Cuenta",
        "tab_estado": "Estado",
        "idioma_caption": (
            "Elige en qué idioma quieres que te responda ORIA. Puedes "
            "escribirle en cualquier idioma: ella siempre te "
            "contestará en el que elijas aquí."
        ),
        "idioma_cambiado": "Idioma cambiado a {idioma}.",
        "memoria_caption": (
            "Escribe aquí datos que quieras que ORIA recuerde siempre "
            "(tu nombre, tus preferencias, tu contexto...). Se "
            "incluirán en todas las conversaciones."
        ),
        "guardar_memoria": "Guardar memoria",
        "memoria_guardada": "Memoria guardada.",
        "cuenta_invitado": (
            "Modo invitado: el inicio de sesión aún no está "
            "configurado."
        ),
        "cuenta_sesion_como": "Sesión iniciada como **{nombre}**.",
        "cerrar_sesion": "Cerrar sesión",
        "eliminar_cuenta_titulo": "Eliminar mi cuenta y mis datos",
        "eliminar_cuenta_caption": (
            "Borra todas tus conversaciones y tu memoria de forma "
            "permanente. No se puede deshacer."
        ),
        "eliminar_cuenta_checkbox": "Sí, quiero eliminar todos mis datos",
        "eliminar_definitivamente": "Eliminar definitivamente",
        "datos_eliminados": "Tus datos se han eliminado.",
        "compartir_caption": (
            "Comparte el enlace de esta página con quien quieras: "
            "cada persona entra con su propia cuenta y tiene su "
            "historial y memoria separados del tuyo. En el móvil "
            "pueden usar 'Añadir a pantalla de inicio' para que "
            "funcione como una app."
        ),
        "estado_web_activa": "Búsqueda web activa",
        "estado_web_inactiva": "Búsqueda web sin configurar",
        "estado_nube": "Datos guardados en la nube",
        "estado_servidor": (
            "Datos guardados solo en el servidor (pueden perderse "
            "al reiniciar). Configura Supabase para guardarlos en la "
            "nube."
        ),
        "estado_calidad_imagen": (
            "ORIA mejora automáticamente tu descripción antes de "
            "generar una imagen. Se usa el generador gratuito y "
            "anónimo de Pollinations.ai, que no requiere cuenta ni "
            "pago, pero por eso incluye una pequeña marca de agua y "
            "a veces algún error puntual bajo mucha demanda — es el "
            "límite normal de una herramienta 100% gratuita."
        ),
        "chat_placeholder": "Pregunta a ORIA, o adjunta una imagen/PDF...",
        "boton_escuchar": "Escuchar",
        "boton_copiar": "Copiar",
        "boton_copiado": "Copiado",
        "fuentes": "Fuentes ({n})",
        "imagen_no_disponible": "*(La imagen generada ya no está disponible)*",
        "resumen_pdf_defecto": "Resume este documento y destaca los puntos clave.",
        "describe_imagen_defecto": "Describe esta imagen y explica qué ves con detalle.",
        "error_generar_imagen": "No se pudo generar la imagen: {error}",
        "buscando_web": "Buscando en la web...",
        "puliendo_descripcion": "Puliendo la descripción...",
        "generando_imagen": "Generando imagen...",
        "no_pude_generar_imagen": "No he podido generar la imagen: {error}",
        "leyendo_pdf": "Leyendo el PDF...",
        "no_pude_leer_pdf": "No he podido leer el PDF: {error}",
        "generando_contenido_presentacion": "Redactando el contenido de la presentación...",
        "creando_pptx": "Montando la presentación y buscando imágenes...",
        "creando_exportables": "Generando el PDF y la vista previa...",
        "error_generar_presentacion": "No se pudo generar la presentación: {error}",
        "no_pude_generar_presentacion": "No he podido generar la presentación: {error}",
        "presentacion_generada": "He creado tu presentación: **{titulo}** ({n} diapositivas). Aquí la tienes, puedes verla antes de descargarla.",
        "exportar_como": "Exportar como",
        "descargar_presentacion": "PowerPoint (.pptx)",
        "descargar_pdf": "PDF",
        "pptx_no_disponible": "*(Esta presentación ya no está disponible)*",
        "presentacion_sin_titulo": "Presentación",
        "web_no_configurada": (
            "La búsqueda web no está configurada todavía, así que "
            "no puedo confirmar datos de última hora."
        ),
        "web_no_disponible": "No he podido consultar la web ahora mismo.",
        "error_413": (
            "**La conversación se ha quedado demasiado larga para "
            "el plan gratuito de Groq en este momento (demasiados "
            "tokens por minuto).**\n\n"
            "Prueba a pulsar 'Nueva conversación' para empezar de "
            "cero, o espera un minuto y vuelve a intentarlo."
        ),
        "invitado": "Invitado",
    },
    "en": {
        "tagline_bienvenida": "What can I help you with today?",
        "nueva_conversacion": "＋  New chat",
        "conversacion_sin_titulo": "New chat",
        "archivo_titulo": "File: {nombre}",
        "ajustes_titulo": "Settings",
        "tab_idioma": "Language",
        "tab_memoria": "Memory",
        "tab_cuenta": "Account",
        "tab_estado": "Status",
        "idioma_caption": (
            "Choose the language you want ORIA to reply in. You can "
            "write to it in any language: it will always answer in "
            "the one you choose here."
        ),
        "idioma_cambiado": "Language changed to {idioma}.",
        "memoria_caption": (
            "Write here anything you want ORIA to always remember "
            "(your name, your preferences, your context...). It will "
            "be included in every conversation."
        ),
        "guardar_memoria": "Save memory",
        "memoria_guardada": "Memory saved.",
        "cuenta_invitado": (
            "Guest mode: sign-in isn't configured yet."
        ),
        "cuenta_sesion_como": "Signed in as **{nombre}**.",
        "cerrar_sesion": "Sign out",
        "eliminar_cuenta_titulo": "Delete my account and data",
        "eliminar_cuenta_caption": (
            "Permanently deletes all your conversations and your "
            "memory. This can't be undone."
        ),
        "eliminar_cuenta_checkbox": "Yes, I want to delete all my data",
        "eliminar_definitivamente": "Delete permanently",
        "datos_eliminados": "Your data has been deleted.",
        "compartir_caption": (
            "Share this page's link with anyone: each person signs "
            "in with their own account and has their history and "
            "memory kept separate from yours. On mobile they can use "
            "'Add to home screen' so it works like an app."
        ),
        "estado_web_activa": "Web search active",
        "estado_web_inactiva": "Web search not configured",
        "estado_nube": "Data saved in the cloud",
        "estado_servidor": (
            "Data saved only on the server (may be lost on "
            "restart). Configure Supabase to save it to the cloud."
        ),
        "estado_calidad_imagen": (
            "ORIA automatically improves your description before "
            "generating an image. It uses Pollinations.ai's free, "
            "anonymous generator, which needs no account or payment, "
            "but that's also why it includes a small watermark and "
            "occasionally an error under heavy demand — the normal "
            "limit of a 100% free tool."
        ),
        "chat_placeholder": "Ask ORIA, or attach an image/PDF...",
        "boton_escuchar": "Listen",
        "boton_copiar": "Copy",
        "boton_copiado": "Copied",
        "fuentes": "Sources ({n})",
        "imagen_no_disponible": "*(This generated image is no longer available)*",
        "resumen_pdf_defecto": "Summarize this document and highlight the key points.",
        "describe_imagen_defecto": "Describe this image and explain what you see in detail.",
        "error_generar_imagen": "Couldn't generate the image: {error}",
        "buscando_web": "Searching the web...",
        "puliendo_descripcion": "Polishing the description...",
        "generando_imagen": "Generating image...",
        "no_pude_generar_imagen": "I couldn't generate the image: {error}",
        "leyendo_pdf": "Reading the PDF...",
        "no_pude_leer_pdf": "I couldn't read the PDF: {error}",
        "generando_contenido_presentacion": "Writing the presentation content...",
        "creando_pptx": "Putting the slides together and finding images...",
        "creando_exportables": "Generating the PDF and the preview...",
        "error_generar_presentacion": "Couldn't generate the presentation: {error}",
        "no_pude_generar_presentacion": "I couldn't generate the presentation: {error}",
        "presentacion_generada": "I've created your presentation: **{titulo}** ({n} slides). Here it is, you can preview it before downloading.",
        "exportar_como": "Export as",
        "descargar_presentacion": "PowerPoint (.pptx)",
        "descargar_pdf": "PDF",
        "pptx_no_disponible": "*(This presentation is no longer available)*",
        "presentacion_sin_titulo": "Presentation",
        "web_no_configurada": (
            "Web search isn't configured yet, so I can't confirm "
            "up-to-the-minute data."
        ),
        "web_no_disponible": "I couldn't check the web right now.",
        "error_413": (
            "**This conversation has gotten too long for Groq's "
            "free plan right now (too many tokens per minute).**\n\n"
            "Try tapping 'New chat' to start fresh, or wait a "
            "minute and try again."
        ),
        "invitado": "Guest",
    },
    "ca": {
        "tagline_bienvenida": "En què et puc ajudar avui?",
        "nueva_conversacion": "＋  Nova conversa",
        "conversacion_sin_titulo": "Nova conversa",
        "archivo_titulo": "Fitxer: {nombre}",
        "ajustes_titulo": "Ajustos",
        "tab_idioma": "Idioma",
        "tab_memoria": "Memòria",
        "tab_cuenta": "Compte",
        "tab_estado": "Estat",
        "idioma_caption": (
            "Tria en quin idioma vols que et respongui ORIA. Li pots "
            "escriure en qualsevol idioma: sempre et contestarà en "
            "el que triïs aquí."
        ),
        "idioma_cambiado": "Idioma canviat a {idioma}.",
        "memoria_caption": (
            "Escriu aquí dades que vulguis que ORIA recordi sempre "
            "(el teu nom, les teves preferències, el teu context...). "
            "S'inclouran a totes les converses."
        ),
        "guardar_memoria": "Desar memòria",
        "memoria_guardada": "Memòria desada.",
        "cuenta_invitado": (
            "Mode convidat: l'inici de sessió encara no està "
            "configurat."
        ),
        "cuenta_sesion_como": "Sessió iniciada com a **{nombre}**.",
        "cerrar_sesion": "Tancar sessió",
        "eliminar_cuenta_titulo": "Eliminar el meu compte i les meves dades",
        "eliminar_cuenta_caption": (
            "Esborra totes les teves converses i la teva memòria de "
            "forma permanent. No es pot desfer."
        ),
        "eliminar_cuenta_checkbox": "Sí, vull eliminar totes les meves dades",
        "eliminar_definitivamente": "Eliminar definitivament",
        "datos_eliminados": "Les teves dades s'han eliminat.",
        "compartir_caption": (
            "Comparteix l'enllaç d'aquesta pàgina amb qui vulguis: "
            "cada persona entra amb el seu propi compte i té el seu "
            "historial i memòria separats del teu. Al mòbil poden "
            "usar 'Afegeix a la pantalla d'inici' perquè funcioni "
            "com una app."
        ),
        "estado_web_activa": "Cerca web activa",
        "estado_web_inactiva": "Cerca web sense configurar",
        "estado_nube": "Dades desades al núvol",
        "estado_servidor": (
            "Dades desades només al servidor (es poden perdre en "
            "reiniciar). Configura Supabase per desar-les al núvol."
        ),
        "estado_calidad_imagen": (
            "ORIA millora automàticament la teva descripció abans de "
            "generar una imatge. S'utilitza el generador gratuït i "
            "anònim de Pollinations.ai, que no requereix compte ni "
            "pagament, però per això inclou una petita marca d'aigua "
            "i de vegades algun error puntual sota molta demanda — "
            "és el límit normal d'una eina 100% gratuïta."
        ),
        "chat_placeholder": "Pregunta a l'ORIA, o adjunta una imatge/PDF...",
        "boton_escuchar": "Escoltar",
        "boton_copiar": "Copiar",
        "boton_copiado": "Copiat",
        "fuentes": "Fonts ({n})",
        "imagen_no_disponible": "*(La imatge generada ja no està disponible)*",
        "resumen_pdf_defecto": "Resumeix aquest document i destaca'n els punts clau.",
        "describe_imagen_defecto": "Descriu aquesta imatge i explica amb detall què hi veus.",
        "error_generar_imagen": "No s'ha pogut generar la imatge: {error}",
        "buscando_web": "Cercant a la web...",
        "puliendo_descripcion": "Polint la descripció...",
        "generando_imagen": "Generant la imatge...",
        "no_pude_generar_imagen": "No he pogut generar la imatge: {error}",
        "leyendo_pdf": "Llegint el PDF...",
        "no_pude_leer_pdf": "No he pogut llegir el PDF: {error}",
        "generando_contenido_presentacion": "Redactant el contingut de la presentació...",
        "creando_pptx": "Muntant la presentació i cercant imatges...",
        "creando_exportables": "Generant el PDF i la vista prèvia...",
        "error_generar_presentacion": "No s'ha pogut generar la presentació: {error}",
        "no_pude_generar_presentacion": "No he pogut generar la presentació: {error}",
        "presentacion_generada": "He creat la teva presentació: **{titulo}** ({n} diapositives). Aquí la tens, la pots veure abans de descarregar-la.",
        "exportar_como": "Exportar com a",
        "descargar_presentacion": "PowerPoint (.pptx)",
        "descargar_pdf": "PDF",
        "pptx_no_disponible": "*(Aquesta presentació ja no està disponible)*",
        "presentacion_sin_titulo": "Presentació",
        "web_no_configurada": (
            "La cerca web encara no està configurada, així que no "
            "puc confirmar dades de darrera hora."
        ),
        "web_no_disponible": "No he pogut consultar la web ara mateix.",
        "error_413": (
            "**La conversa s'ha quedat massa llarga per al pla "
            "gratuït de Groq en aquest moment (massa tokens per "
            "minut).**\n\n"
            "Prova de prémer 'Nova conversa' per començar de zero, "
            "o espera un minut i torna-ho a provar."
        ),
        "invitado": "Convidat",
    },
    "fr": {
        "tagline_bienvenida": "Comment puis-je t'aider aujourd'hui ?",
        "nueva_conversacion": "＋  Nouvelle conversation",
        "conversacion_sin_titulo": "Nouvelle conversation",
        "archivo_titulo": "Fichier : {nombre}",
        "ajustes_titulo": "Paramètres",
        "tab_idioma": "Langue",
        "tab_memoria": "Mémoire",
        "tab_cuenta": "Compte",
        "tab_estado": "État",
        "idioma_caption": (
            "Choisis la langue dans laquelle tu veux qu'ORIA te "
            "réponde. Tu peux lui écrire dans n'importe quelle "
            "langue : elle te répondra toujours dans celle choisie "
            "ici."
        ),
        "idioma_cambiado": "Langue changée en {idioma}.",
        "memoria_caption": (
            "Écris ici ce que tu veux qu'ORIA se rappelle toujours "
            "(ton nom, tes préférences, ton contexte...). Ce sera "
            "inclus dans toutes les conversations."
        ),
        "guardar_memoria": "Enregistrer la mémoire",
        "memoria_guardada": "Mémoire enregistrée.",
        "cuenta_invitado": (
            "Mode invité : la connexion n'est pas encore configurée."
        ),
        "cuenta_sesion_como": "Connecté en tant que **{nombre}**.",
        "cerrar_sesion": "Se déconnecter",
        "eliminar_cuenta_titulo": "Supprimer mon compte et mes données",
        "eliminar_cuenta_caption": (
            "Supprime définitivement toutes tes conversations et ta "
            "mémoire. Cette action est irréversible."
        ),
        "eliminar_cuenta_checkbox": "Oui, je veux supprimer toutes mes données",
        "eliminar_definitivamente": "Supprimer définitivement",
        "datos_eliminados": "Tes données ont été supprimées.",
        "compartir_caption": (
            "Partage le lien de cette page avec qui tu veux : chaque "
            "personne se connecte avec son propre compte et garde "
            "son historique et sa mémoire séparés des tiens. Sur "
            "mobile, on peut utiliser 'Ajouter à l'écran d'accueil' "
            "pour que ça fonctionne comme une appli."
        ),
        "estado_web_activa": "Recherche web active",
        "estado_web_inactiva": "Recherche web non configurée",
        "estado_nube": "Données enregistrées dans le cloud",
        "estado_servidor": (
            "Données enregistrées seulement sur le serveur "
            "(peuvent être perdues au redémarrage). Configure "
            "Supabase pour les enregistrer dans le cloud."
        ),
        "estado_calidad_imagen": (
            "ORIA améliore automatiquement ta description avant de "
            "générer une image. Le générateur gratuit et anonyme de "
            "Pollinations.ai est utilisé, sans compte ni paiement, "
            "mais c'est aussi pour ça qu'il inclut un petit filigrane "
            "et parfois une erreur ponctuelle en cas de forte "
            "demande — la limite normale d'un outil 100% gratuit."
        ),
        "chat_placeholder": "Demande à ORIA, ou joins une image/PDF...",
        "boton_escuchar": "Écouter",
        "boton_copiar": "Copier",
        "boton_copiado": "Copié",
        "fuentes": "Sources ({n})",
        "imagen_no_disponible": "*(Cette image générée n'est plus disponible)*",
        "resumen_pdf_defecto": "Résume ce document et souligne les points clés.",
        "describe_imagen_defecto": "Décris cette image et explique en détail ce que tu vois.",
        "error_generar_imagen": "Impossible de générer l'image : {error}",
        "buscando_web": "Recherche sur le web...",
        "puliendo_descripcion": "Amélioration de la description...",
        "generando_imagen": "Génération de l'image...",
        "no_pude_generar_imagen": "Je n'ai pas pu générer l'image : {error}",
        "leyendo_pdf": "Lecture du PDF...",
        "no_pude_leer_pdf": "Je n'ai pas pu lire le PDF : {error}",
        "generando_contenido_presentacion": "Rédaction du contenu de la présentation...",
        "creando_pptx": "Assemblage de la présentation et recherche d'images...",
        "creando_exportables": "Génération du PDF et de l'aperçu...",
        "error_generar_presentacion": "Impossible de générer la présentation : {error}",
        "no_pude_generar_presentacion": "Je n'ai pas pu générer la présentation : {error}",
        "presentacion_generada": "J'ai créé ta présentation : **{titulo}** ({n} diapositives). La voici, tu peux la prévisualiser avant de la télécharger.",
        "exportar_como": "Exporter en",
        "descargar_presentacion": "PowerPoint (.pptx)",
        "descargar_pdf": "PDF",
        "pptx_no_disponible": "*(Cette présentation n'est plus disponible)*",
        "presentacion_sin_titulo": "Présentation",
        "web_no_configurada": (
            "La recherche web n'est pas encore configurée, je ne "
            "peux donc pas confirmer les données de dernière minute."
        ),
        "web_no_disponible": "Je n'ai pas pu consulter le web pour le moment.",
        "error_413": (
            "**La conversation est devenue trop longue pour le "
            "forfait gratuit de Groq en ce moment (trop de tokens "
            "par minute).**\n\n"
            "Essaie d'appuyer sur 'Nouvelle conversation' pour "
            "repartir de zéro, ou attends une minute et réessaie."
        ),
        "invitado": "Invité",
    },
}


def t(clave, idioma=None, **kwargs):
    """Devuelve el texto de la interfaz en el idioma elegido por el
    usuario (o en español si aún no hay ninguno elegido, como en la
    pantalla de inicio de sesión). Si la clave falta en ese idioma,
    cae en español antes que mostrar un hueco en blanco.

    Se puede pasar `idioma` explícitamente para los sitios donde
    todavía no hay contexto de `st.session_state` (por ejemplo,
    dentro del generador que llama a Groq)."""

    if idioma is None:
        idioma = st.session_state.get("idioma", IDIOMA_POR_DEFECTO)

    texto = TEXTOS.get(idioma, TEXTOS[IDIOMA_POR_DEFECTO]).get(clave)

    if texto is None:
        texto = TEXTOS[IDIOMA_POR_DEFECTO].get(clave, clave)

    return texto.format(**kwargs) if kwargs else texto


def _cargar_local(usuario_id):
    ruta = _ruta_datos_usuario(usuario_id)

    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)

                if isinstance(datos, dict):
                    datos.setdefault("chats", {})
                    datos.setdefault("memoria", "")
                    datos.setdefault("idioma", IDIOMA_POR_DEFECTO)
                    return datos

        except Exception:
            pass

    return {"chats": {}, "memoria": "", "idioma": IDIOMA_POR_DEFECTO}


def _guardar_local(usuario_id, chats, memoria, idioma):
    ruta = _ruta_datos_usuario(usuario_id)

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(
            {"chats": chats, "memoria": memoria, "idioma": idioma},
            f,
            ensure_ascii=False,
            indent=2,
        )


def cargar_datos_usuario(usuario_id):
    """Carga chats, memoria e idioma del usuario. Devuelve (datos, ok).
    Si falla la nube devuelve ok=False para NO seguir adelante con
    datos vacíos (así no se sobrescribe lo que ya tenía guardado).

    Nota sobre "idioma": si la columna todavía no existe en la tabla
    de Supabase (por ejemplo, porque la cuenta se creó antes de
    añadir esta función), Supabase devuelve un error solo por pedir
    esa columna en "select". Por eso, si el primer intento falla,
    reintentamos sin pedir "idioma" y usamos el valor por defecto —
    así una cuenta antigua no se queda sin poder entrar."""

    url, key = _config_supabase()

    if not (url and key):
        return _cargar_local(usuario_id), True

    try:
        respuesta = requests.get(
            f"{url}/rest/v1/{TABLA_SUPABASE}",
            headers=_headers_supabase(key),
            params={
                "email": f"eq.{usuario_id}",
                "select": "chats,memoria,idioma",
            },
            timeout=20,
        )

        if respuesta.status_code != 200:
            # Puede que la columna "idioma" no exista todavía en la
            # tabla. Reintentamos sin ella para no bloquear al usuario.
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
            return (
                {"chats": {}, "memoria": "", "idioma": IDIOMA_POR_DEFECTO},
                True,
            )

        fila = filas[0]
        chats = fila.get("chats") or {}
        memoria = fila.get("memoria") or ""
        idioma = fila.get("idioma") or IDIOMA_POR_DEFECTO

        if not isinstance(chats, dict):
            chats = {}

        return {"chats": chats, "memoria": memoria, "idioma": idioma}, True

    except Exception:
        return None, False


def guardar_datos_usuario(usuario_id, chats, memoria, idioma=IDIOMA_POR_DEFECTO):
    """Guarda chats, memoria e idioma del usuario (nube si está
    configurada, y si no, archivo local)."""

    url, key = _config_supabase()

    if not (url and key):
        try:
            _guardar_local(usuario_id, chats, memoria, idioma)
        except Exception as e:
            st.error(f"Error al guardar tus datos: {e}")
        return

    datos_a_guardar = {
        "email": usuario_id,
        "chats": chats,
        "memoria": memoria,
        "idioma": idioma,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }

    try:
        respuesta = requests.post(
            f"{url}/rest/v1/{TABLA_SUPABASE}",
            headers=_headers_supabase(
                key,
                {"Prefer": "resolution=merge-duplicates,return=minimal"},
            ),
            json=datos_a_guardar,
            timeout=20,
        )

        if respuesta.status_code not in (200, 201, 204):
            # Si el fallo es porque la columna "idioma" no existe
            # aún en Supabase, reintentamos sin ella para no perder
            # el resto de los datos del usuario.
            datos_sin_idioma = dict(datos_a_guardar)
            datos_sin_idioma.pop("idioma", None)

            respuesta = requests.post(
                f"{url}/rest/v1/{TABLA_SUPABASE}",
                headers=_headers_supabase(
                    key,
                    {"Prefer": "resolution=merge-duplicates,return=minimal"},
                ),
                json=datos_sin_idioma,
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
                        "de generación de imágenes (tipo Flux). Tu trabajo "
                        "es EXPANDIR la idea del usuario en una descripción "
                        "muy detallada y visual, en una sola frase larga, "
                        "añadiendo estilo artístico, iluminación, encuadre, "
                        "ambiente y calidad (ej. 'fotografía realista', "
                        "'8k', 'cinematográfico', 'alto detalle'...).\n\n"
                        "Regla más importante, por encima de cualquier "
                        "otra: nunca cambies, quites ni añadas elementos "
                        "concretos que el usuario haya pedido explícitamente "
                        "(número de personas u objetos, colores, texto "
                        "exacto que deba aparecer, marcas, animales, "
                        "acciones...). Solo añades detalle alrededor de "
                        "exactamente lo que ha pedido, nunca inventas cosas "
                        "nuevas que cambien el resultado ni te tomas "
                        "libertades creativas con lo que ya es específico. "
                        "Si el usuario ya ha sido muy concreto, limítate a "
                        "pulir el estilo y la calidad, no el contenido.\n\n"
                        "Responde ÚNICAMENTE con el prompt final, en "
                        "inglés (los modelos de imagen entienden mejor "
                        "inglés), sin comillas, explicaciones ni texto "
                        "adicional."
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


def generar_imagen_ia(prompt_imagen, intentos=3):
    """Genera una imagen a partir de un texto usando Pollinations.ai
    (servicio gratuito). Si hay una POLLINATIONS_API_KEY configurada
    en Secrets, se usa para quitar la marca de agua y tener más
    estabilidad; si no, funciona igualmente en modo anónimo.
    Devuelve (bytes, error).

    Nota: NO añadimos "enhance=true" aquí. Ese parámetro le pide a
    Pollinations que reescriba el prompt por su cuenta con su propio
    sistema — pero nosotros ya se lo mandamos muy detallado gracias a
    `mejorar_prompt_imagen`, así que dejar también el suyo activado
    hacía que el resultado final se pareciera menos a lo que pidió el
    usuario (doble reescritura, cada una tirando en una dirección).
    Sin él, la imagen sigue mucho más fielmente nuestro prompt."""

    token = obtener_pollinations_key()

    prompt_codificado = quote(prompt_imagen)
    semilla = uuid.uuid4().int % 1_000_000

    url = (
        f"https://image.pollinations.ai/prompt/{prompt_codificado}"
        f"?model=flux&width=1024&height=1024"
        f"&nologo=true&seed={semilla}"
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


def generar_contenido_presentacion(tema, idioma=IDIOMA_POR_DEFECTO):
    """Le pide a Groq el contenido de una presentación (título,
    subtítulo, categorías/kicker y diapositivas con sus puntos en
    forma de tarjetas {titulo, detalle}) en forma de JSON.
    Devuelve (contenido_dict, error)."""

    api_key = obtener_api_key()

    if not api_key or not api_key.startswith("gsk_"):
        return None, "No se ha encontrado una GROQ_API_KEY válida."

    nombre_idioma = NOMBRE_IDIOMA_PARA_PROMPT.get(
        idioma, NOMBRE_IDIOMA_PARA_PROMPT[IDIOMA_POR_DEFECTO]
    )

    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

        payload = {
            "model": MODELO_GROQ,
            "stream": False,
            "temperature": 0.6,
            "max_completion_tokens": 3800,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "Eres un experto diseñando el CONTENIDO de "
                        "presentaciones profesionales, del nivel de una "
                        "consultora o de un buen diseñador -claras, bien "
                        "organizadas y visuales, con tarjetas en vez de "
                        "listas de viñetas planas-. A partir del tema "
                        "que te da el usuario, genera:\n\n"
                        "1. Un título principal potente, un subtítulo "
                        "breve, y una 'categoria_portada' (una etiqueta "
                        "corta en MAYÚSCULAS, 2-4 palabras, tipo "
                        "'DEPORTE GLOBAL' o 'GUÍA ESTRATÉGICA', que "
                        "resuma el ámbito del tema).\n"
                        "2. Entre 5 y 9 diapositivas de contenido "
                        "(ajusta el número a la complejidad del tema). "
                        "Cada una lleva:\n"
                        "   - 'categoria': etiqueta corta en MAYÚSCULAS "
                        "que resume de qué trata esa diapositiva dentro "
                        "del conjunto (ej. 'HISTORIA Y EVOLUCIÓN').\n"
                        "   - 'titulo': un título corto y concreto, no "
                        "genérico (ej. 'Nacimiento en Acapulco', no "
                        "'Introducción').\n"
                        "   - 'puntos': entre 3 y 4 tarjetas, cada una "
                        "con 'titulo' (una mini-frase impactante de 3-6 "
                        "palabras, con un dato, año o cifra concreta "
                        "cuando exista, ej. '1969: Invención en "
                        "México') y 'detalle' (UNA frase breve que "
                        "amplíe ese punto, nunca un párrafo largo).\n"
                        "No repitas el título principal como "
                        "diapositiva. No uses frases de relleno como "
                        "'En esta diapositiva veremos'. "
                        f"Escribe TODO el contenido en {nombre_idioma}, "
                        "incluidas las categorías.\n\n"
                        "Además, para hacer la presentación más visual, "
                        "decide para la portada y para CADA diapositiva "
                        "si le pega bien una imagen real, un gráfico, o "
                        "ninguna de las dos cosas:\n"
                        "- \"foto\": el contenido representa algo "
                        "fotografiable del mundo real (personas, "
                        "lugares, objetos, naturaleza, negocios, "
                        "tecnología en uso, deportes, comida, etc.).\n"
                        "- \"diagrama\": el contenido es técnico, "
                        "científico, anatómico o histórico y se "
                        "entiende mejor con una ilustración, esquema, "
                        "mapa o imagen de archivo, no con una foto de "
                        "stock genérica.\n"
                        "- \"grafico\": la diapositiva compara cifras "
                        "(evolución en el tiempo, tamaño de mercado, "
                        "porcentajes, ranking entre varias categorías, "
                        "crecimiento año a año, etc.). ES LA OPCIÓN "
                        "PREFERIDA siempre que haya datos numéricos "
                        "comparables, en vez de una foto genérica: un "
                        "gráfico real con las cifras aporta mucho más "
                        "que una foto de archivo. En este caso añade "
                        "también un objeto \"grafico\" con: \"tipo\" "
                        "(\"barras\" para comparar categorías, o "
                        "\"lineas\" para una evolución en el tiempo), "
                        "\"categorias\" (lista de 3 a 6 etiquetas "
                        "cortas: años, países, nombres...), \"valores\" "
                        "(la misma cantidad de números, SOLO números, "
                        "sin texto ni símbolos) y \"sufijo\" (opcional, "
                        "ej. \"%\", \"M\", \"M€\", para mostrar junto a "
                        "cada valor). Usa cifras razonables y conocidas; "
                        "si no las sabes con precisión, da una "
                        "estimación sensata en vez de inventar "
                        "decimales falsos.\n"
                        "- \"ninguna\": el contenido es abstracto o una "
                        "definición y ni una imagen ni un gráfico "
                        "aportarían nada (en ese caso deja "
                        "\"imagen_query\" como cadena vacía). Usa "
                        "\"ninguna\" con sinceridad, no fuerces una "
                        "imagen o un gráfico si no pega.\n"
                        "Para cada imagen que pidas (\"foto\" o "
                        "\"diagrama\"), escribe \"imagen_query\" "
                        "SIEMPRE en inglés (2 a 5 palabras concretas, "
                        "pensadas para buscar en un banco de fotos o en "
                        "una enciclopedia visual), aunque el resto del "
                        "contenido esté en otro idioma. La portada solo "
                        "admite \"foto\" o \"ninguna\" (nunca gráfico "
                        "en la portada).\n\n"
                        "Responde ÚNICAMENTE con un JSON válido, sin "
                        "explicaciones, sin comillas triples ni texto "
                        "adicional antes o después, con exactamente esta "
                        "forma:\n"
                        '{"titulo": "...", "subtitulo": "...", '
                        '"categoria_portada": "...", '
                        '"imagen_portada_query": "...", '
                        '"estilo_imagen_portada": "foto|ninguna", '
                        '"diapositivas": [{"categoria": "...", '
                        '"titulo": "...", "puntos": [{"titulo": "...", '
                        '"detalle": "..."}], "imagen_query": "...", '
                        '"estilo_imagen": "foto|diagrama|grafico|ninguna", '
                        '"grafico": {"tipo": "barras|lineas", '
                        '"categorias": ["...", "..."], '
                        '"valores": [0, 0], "sufijo": "..."}}]}'
                    ),
                },
                {"role": "user", "content": tema},
            ],
        }

        respuesta = requests.post(
            URL_GROQ, headers=headers, json=payload, timeout=45
        )

        if respuesta.status_code != 200:
            return None, (
                f"Groq ha devuelto el error {respuesta.status_code} al "
                "generar el contenido."
            )

        texto = (
            respuesta.json()
            .get("choices", [{}])[0]
            .get("message", {})
            .get("content", "")
            .strip()
        )

        # Por si el modelo envuelve el JSON en ```json ... ``` a pesar
        # de que se lo hemos pedido explícitamente sin eso.
        texto = re.sub(r"^```(?:json)?\s*|\s*```$", "", texto.strip())

        # Si aun así hay texto antes/después del JSON, nos quedamos
        # solo con lo que hay entre la primera "{" y la última "}".
        inicio = texto.find("{")
        fin = texto.rfind("}")
        if inicio != -1 and fin != -1 and fin > inicio:
            texto = texto[inicio:fin + 1]

        try:
            contenido = json.loads(texto)
        except json.JSONDecodeError:
            return None, (
                "No he podido interpretar el contenido generado. "
                "Prueba a pedirlo de nuevo, quizás con un tema más "
                "concreto."
            )

        if not isinstance(contenido, dict) or not contenido.get(
            "diapositivas"
        ):
            return None, "El contenido generado no tiene el formato esperado."

        return contenido, None

    except requests.exceptions.Timeout:
        return None, "Generar el contenido ha tardado demasiado."

    except Exception as e:
        return None, str(e)


def _nombre_archivo_seguro(texto, por_defecto="presentacion"):
    """Limpia un texto (p. ej. el título generado por la IA) para que
    se pueda usar como nombre de archivo descargable en cualquier
    sistema operativo, quitando caracteres problemáticos."""

    limpio = re.sub(r'[\\/:*?"<>|]+', "", str(texto)).strip()
    limpio = re.sub(r"\s+", " ", limpio)

    return limpio[:80] if limpio else por_defecto


# Paleta de colores propia de ORIA para las presentaciones generadas.
# Se usa un diseño hecho a mano (rectángulos + cajas de texto sobre una
# diapositiva en blanco) en vez de las plantillas por defecto de
# PowerPoint, que se ven muy genéricas.
_PPTX_COLOR_FONDO_OSCURO = RGBColor(0x1B, 0x1B, 0x2A)
_PPTX_COLOR_ACENTO = RGBColor(0x6C, 0x5C, 0xE7)
_PPTX_COLOR_TEXTO_OSCURO = RGBColor(0x2B, 0x2B, 0x31)
_PPTX_COLOR_TEXTO_CLARO = RGBColor(0xFF, 0xFF, 0xFF)
_PPTX_COLOR_TEXTO_SECUNDARIO = RGBColor(0x8A, 0x8A, 0x96)
_PPTX_COLOR_SUBTITULO_CLARO = RGBColor(0xC9, 0xC9, 0xDC)
_PPTX_COLOR_LINEA = RGBColor(0xE3, 0xE3, 0xE8)
_PPTX_COLOR_FONDO_CLARO = RGBColor(0xF6, 0xF6, 0xFA)
_PPTX_COLOR_TARJETA_FONDO = RGBColor(0xFF, 0xFF, 0xFF)
_PPTX_FUENTE = "Calibri"


def _pptx_rectangulo(slide, left, top, width, height, color):
    """Añade un rectángulo sólido sin borde ni sombra (bloque de color
    para fondos, barras de acento y líneas divisorias)."""

    forma = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    forma.fill.solid()
    forma.fill.fore_color.rgb = color
    forma.line.fill.background()
    forma.shadow.inherit = False
    return forma


def _pptx_caja_texto(slide, left, top, width, height, anclaje_vertical=None):
    """Añade una caja de texto con ajuste de línea activado."""

    caja = slide.shapes.add_textbox(left, top, width, height)
    marco = caja.text_frame
    marco.word_wrap = True
    if anclaje_vertical is not None:
        marco.vertical_anchor = anclaje_vertical
    return marco


def _pptx_tarjeta_punto(slide, left, top, width, height, titulo, detalle):
    """Dibuja una 'tarjeta' para un punto de la diapositiva -fondo
    blanco, esquinas redondeadas, borde de acento a la izquierda,
    título corto en negrita y una frase de detalle debajo- en vez de
    la viñeta plana de antes."""

    tarjeta = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    try:
        tarjeta.adjustments[0] = 0.07
    except Exception:
        pass
    tarjeta.fill.solid()
    tarjeta.fill.fore_color.rgb = _PPTX_COLOR_TARJETA_FONDO
    tarjeta.line.color.rgb = _PPTX_COLOR_LINEA
    tarjeta.line.width = Pt(0.75)
    tarjeta.shadow.inherit = False

    # Barra de acento pegada al borde izquierdo de la tarjeta.
    _pptx_rectangulo(slide, left, top, Pt(4), height, _PPTX_COLOR_ACENTO)

    margen_h = Inches(0.22)
    margen_v = Inches(0.13)
    marco = _pptx_caja_texto(
        slide,
        left + margen_h,
        top + margen_v,
        width - margen_h * 2,
        height - margen_v * 2,
        anclaje_vertical=MSO_ANCHOR.MIDDLE,
    )

    p_titulo = marco.paragraphs[0]
    p_titulo.text = titulo
    p_titulo.font.size = Pt(15)
    p_titulo.font.bold = True
    p_titulo.font.name = _PPTX_FUENTE
    p_titulo.font.color.rgb = _PPTX_COLOR_TEXTO_OSCURO
    p_titulo.space_after = Pt(3)

    if detalle:
        p_detalle = marco.add_paragraph()
        p_detalle.text = detalle
        p_detalle.font.size = Pt(12)
        p_detalle.font.name = _PPTX_FUENTE
        p_detalle.font.color.rgb = _PPTX_COLOR_TEXTO_SECUNDARIO


def _config_pexels():
    """Devuelve la clave de Pexels desde Secrets, o None si no está
    configurada (es opcional: sin ella, las diapositivas que piden
    una 'foto' se quedan solo con texto)."""
    try:
        clave = str(st.secrets["PEXELS_API_KEY"]).strip().strip('"').strip("'")
        return clave or None
    except Exception:
        return None


def _descargar_imagen_pexels(consulta):
    """Busca una foto real (no generada por IA) en Pexels a partir de
    una consulta en inglés. Devuelve los bytes de la imagen, o None
    si no hay clave configurada, no hay resultados o algo falla."""

    clave = _config_pexels()
    if not clave:
        return None

    try:
        respuesta = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": clave},
            params={"query": consulta, "orientation": "landscape", "per_page": 3},
            timeout=5,
        )
        if respuesta.status_code != 200:
            return None

        fotos = respuesta.json().get("photos", [])
        if not fotos:
            return None

        url_imagen = fotos[0].get("src", {}).get("large")
        if not url_imagen:
            return None

        imagen = requests.get(url_imagen, timeout=5)
        if imagen.status_code != 200:
            return None

        return imagen.content

    except Exception:
        return None


def _config_unsplash():
    """Devuelve la clave de Unsplash desde Secrets, o None si no está
    configurada. Es la alternativa a Pexels (ahora mismo Pexels tiene
    la emisión de claves nuevas pausada): en unsplash.com/developers,
    'New Application', y la 'Access Key' funciona al momento, gratis y
    sin tarjeta, sin esperar ninguna aprobación."""
    try:
        clave = str(st.secrets["UNSPLASH_ACCESS_KEY"]).strip().strip('"').strip("'")
        return clave or None
    except Exception:
        return None


def _descargar_imagen_unsplash(consulta):
    """Busca una foto real (no generada por IA) en Unsplash a partir
    de una consulta en inglés. Devuelve los bytes de la imagen, o
    None si no hay clave configurada, no hay resultados o algo falla."""

    clave = _config_unsplash()
    if not clave:
        return None

    try:
        respuesta = requests.get(
            "https://api.unsplash.com/search/photos",
            headers={"Authorization": f"Client-ID {clave}"},
            params={"query": consulta, "orientation": "landscape", "per_page": 3},
            timeout=5,
        )
        if respuesta.status_code != 200:
            return None

        resultados = respuesta.json().get("results", [])
        if not resultados:
            return None

        url_imagen = resultados[0].get("urls", {}).get("regular")
        if not url_imagen:
            return None

        imagen = requests.get(url_imagen, timeout=5)
        if imagen.status_code != 200:
            return None

        return imagen.content

    except Exception:
        return None


def _descargar_imagen_wikimedia(consulta):
    """Busca una imagen real (foto o ilustración) en Wikimedia
    Commons a partir de una consulta en inglés. No necesita ninguna
    clave de API. Devuelve los bytes de la imagen, o None si no hay
    resultados o algo falla."""

    cabecera = {"User-Agent": "ORIA-App/1.0 (https://oria-chat.streamlit.app/)"}

    try:
        respuesta = requests.get(
            "https://commons.wikimedia.org/w/api.php",
            params={
                "action": "query",
                "generator": "search",
                "gsrsearch": f"{consulta} filetype:bitmap",
                "gsrnamespace": 6,
                "gsrlimit": 5,
                "prop": "imageinfo",
                "iiprop": "url|mime",
                "iiurlwidth": 1200,
                "format": "json",
            },
            headers=cabecera,
            timeout=5,
        )
        if respuesta.status_code != 200:
            return None

        paginas = respuesta.json().get("query", {}).get("pages", {})

        for pagina in paginas.values():
            info = (pagina.get("imageinfo") or [{}])[0]
            mime = info.get("mime", "")
            url_imagen = info.get("thumburl") or info.get("url")

            if url_imagen and mime in ("image/jpeg", "image/png"):
                imagen = requests.get(url_imagen, timeout=5, headers=cabecera)
                if imagen.status_code == 200:
                    return imagen.content

        return None

    except Exception:
        return None


def _obtener_imagen_para_diapositiva(consulta, estilo):
    """Busca una imagen real según el estilo pedido por la IA:
    'foto' -> primero Pexels y, si no hay clave o no encuentra nada,
    Unsplash como alternativa (las dos son opcionales: si ninguna
    está configurada, la diapositiva se queda solo con texto),
    'diagrama' -> Wikimedia Commons (no necesita clave). Devuelve
    bytes de imagen, o None."""

    if not consulta or estilo not in ("foto", "diagrama"):
        return None

    if estilo == "foto":
        return _descargar_imagen_pexels(consulta) or _descargar_imagen_unsplash(consulta)

    return _descargar_imagen_wikimedia(consulta)


_GRAFICO_COLOR_ACENTO = "#6C5CE7"
_GRAFICO_COLOR_TEXTO = "#2B2B31"
_GRAFICO_COLOR_SECUNDARIO = "#8A8A96"
_GRAFICO_COLOR_REJILLA = "#E3E3E8"


def _grafico_formatear_valor(valor, sufijo):
    """Da formato 'humano' a un número para las etiquetas del gráfico
    (sin decimales de más, con coma como separador si hace falta)."""

    try:
        valor = float(valor)
    except (TypeError, ValueError):
        return str(valor)

    if valor == int(valor):
        texto = f"{int(valor):,}".replace(",", ".")
    else:
        texto = f"{valor:,.1f}".replace(",", "·").replace(".", ",").replace("·", ".")

    return f"{texto}{sufijo}"


def _generar_grafico_png(grafico, ancho_px=980, alto_px=760):
    """Genera un gráfico de barras o líneas REAL (con matplotlib, sin
    depender de ningún servicio externo) a partir de las cifras que ha
    dado la IA, con la misma paleta de color que el resto de la
    presentación (morado de acento, texto gris oscuro). Solo una serie
    -sin leyenda, ya que el título ya dice qué se representa- con las
    marcas finas y las etiquetas de valor que recomienda un buen
    diseño de datos. Devuelve los bytes PNG, o None si los datos que
    ha dado la IA no son válidos."""

    if not isinstance(grafico, dict):
        return None

    categorias = [str(c) for c in (grafico.get("categorias") or [])]
    valores_raw = grafico.get("valores") or []

    try:
        valores = [float(v) for v in valores_raw]
    except (TypeError, ValueError):
        return None

    if not categorias or not valores or len(categorias) != len(valores):
        return None

    tipo = grafico.get("tipo") if grafico.get("tipo") in ("barras", "lineas") else "barras"
    sufijo = str(grafico.get("sufijo", "") or "")

    try:
        dpi = 150
        fig, ax = plt.subplots(figsize=(ancho_px / dpi, alto_px / dpi), dpi=dpi)
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        maximo = max(valores)
        minimo = min(valores)
        posiciones = list(range(len(categorias)))

        if tipo == "lineas":
            ax.plot(
                posiciones, valores, color=_GRAFICO_COLOR_ACENTO, linewidth=3,
                solid_capstyle="round", solid_joinstyle="round", marker="o",
                markersize=9, markerfacecolor=_GRAFICO_COLOR_ACENTO,
                markeredgecolor="white", markeredgewidth=2, zorder=3,
            )
            ax.fill_between(
                posiciones, valores, minimo - (maximo - minimo) * 0.15 if maximo != minimo else 0,
                color=_GRAFICO_COLOR_ACENTO, alpha=0.08, zorder=1,
            )
            colchon = (maximo - minimo) * 0.28 or maximo * 0.28 or 1
            ax.set_ylim(minimo - colchon * 0.3, maximo + colchon)
            for i, v in enumerate(valores):
                if i == len(valores) - 1 or v == maximo:
                    ax.annotate(
                        _grafico_formatear_valor(v, sufijo), (i, v),
                        textcoords="offset points", xytext=(0, 13),
                        ha="center", fontsize=14, fontweight="bold",
                        color=_GRAFICO_COLOR_TEXTO,
                    )
        else:
            ax.bar(
                posiciones, valores, color=_GRAFICO_COLOR_ACENTO, width=0.5,
                zorder=3,
            )
            ax.set_ylim(0, maximo * 1.24 if maximo > 0 else 1)
            for i, v in enumerate(valores):
                ax.annotate(
                    _grafico_formatear_valor(v, sufijo), (i, v),
                    textcoords="offset points", xytext=(0, 7),
                    ha="center", fontsize=14, fontweight="bold",
                    color=_GRAFICO_COLOR_TEXTO,
                )

        ax.set_xticks(posiciones)
        ax.set_xticklabels(categorias, fontsize=13, color=_GRAFICO_COLOR_SECUNDARIO)
        ax.tick_params(axis="x", length=0)
        ax.tick_params(axis="y", length=0, labelleft=False)
        for lado in ("top", "right", "left"):
            ax.spines[lado].set_visible(False)
        ax.spines["bottom"].set_color(_GRAFICO_COLOR_REJILLA)
        ax.grid(axis="y", color=_GRAFICO_COLOR_REJILLA, linewidth=1, zorder=0)
        ax.set_axisbelow(True)

        titulo_grafico = str(grafico.get("titulo", "") or "")
        if titulo_grafico:
            ax.set_title(
                titulo_grafico, fontsize=15, fontweight="bold",
                color=_GRAFICO_COLOR_TEXTO, pad=16, loc="left",
            )

        fig.tight_layout(pad=1.3)
        buffer = io.BytesIO()
        fig.savefig(buffer, format="png", dpi=dpi, facecolor="white")
        plt.close(fig)
        return buffer.getvalue()
    except Exception:
        try:
            plt.close(fig)
        except Exception:
            pass
        return None


def obtener_imagenes_presentacion(contenido):
    """Descarga o genera, UNA sola vez, todas las imágenes/gráficos que
    la IA ha pedido para la portada y cada diapositiva, para poder
    reutilizar los mismos bytes en el PPTX, el PDF y la vista previa
    sin repetir el trabajo tres veces. Devuelve un dict:
    {"portada": (bytes, estilo) | None, 0: (bytes, estilo) | None,
    1: ..., ...} (las claves numéricas son el índice de cada
    diapositiva en la lista "diapositivas"). El estilo "grafico" se
    genera localmente con matplotlib en vez de buscarse por internet."""

    cache = {}

    estilo_portada = contenido.get("estilo_imagen_portada")
    datos_portada = _obtener_imagen_para_diapositiva(
        contenido.get("imagen_portada_query", ""), estilo_portada
    )
    cache["portada"] = (datos_portada, estilo_portada) if datos_portada else None

    for indice, diapo in enumerate(contenido.get("diapositivas", [])):
        estilo = diapo.get("estilo_imagen")
        if estilo == "grafico":
            datos = _generar_grafico_png(diapo.get("grafico"))
        else:
            datos = _obtener_imagen_para_diapositiva(diapo.get("imagen_query", ""), estilo)
        cache[indice] = (datos, estilo) if datos else None

    return cache


def _preparar_imagen_recortada(datos_bytes, ancho_caja, alto_caja):
    """Recorta una foto para que encaje EXACTAMENTE en una caja con la
    proporción ancho_caja:alto_caja, sin deformarla (se recortan los
    bordes sobrantes, como en una revista o una web de diseño)."""

    try:
        imagen = Image.open(io.BytesIO(datos_bytes)).convert("RGB")
    except Exception:
        return None

    ancho_img, alto_img = imagen.size
    if ancho_img == 0 or alto_img == 0:
        return None

    relacion_caja = ancho_caja / alto_caja
    relacion_img = ancho_img / alto_img

    if relacion_img > relacion_caja:
        nuevo_ancho = max(int(alto_img * relacion_caja), 1)
        recorte = max((ancho_img - nuevo_ancho) // 2, 0)
        imagen = imagen.crop((recorte, 0, recorte + nuevo_ancho, alto_img))
    else:
        nuevo_alto = max(int(ancho_img / relacion_caja), 1)
        recorte = max((alto_img - nuevo_alto) // 2, 0)
        imagen = imagen.crop((0, recorte, ancho_img, recorte + nuevo_alto))

    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=88)
    buffer.seek(0)
    return buffer


def _pptx_imagen_cubrir(slide, datos_bytes, left, top, ancho_caja, alto_caja):
    """Coloca una foto recortada para llenar exactamente la caja
    indicada (estilo portada de revista), sin bordes ni huecos."""

    buffer = _preparar_imagen_recortada(datos_bytes, ancho_caja, alto_caja)
    if buffer is None:
        return False

    slide.shapes.add_picture(buffer, left, top, width=ancho_caja, height=alto_caja)
    return True


def _pptx_imagen_contener(slide, datos_bytes, left, top, ancho_caja, alto_caja):
    """Coloca una imagen (pensada para diagramas/ilustraciones)
    ajustada DENTRO de la caja sin recortar nada, centrada, para no
    perder etiquetas ni detalles importantes."""

    try:
        imagen = Image.open(io.BytesIO(datos_bytes))
        ancho_img, alto_img = imagen.size
    except Exception:
        return False

    if ancho_img == 0 or alto_img == 0:
        return False

    relacion_caja = ancho_caja / alto_caja
    relacion_img = ancho_img / alto_img

    if relacion_img > relacion_caja:
        ancho_final = ancho_caja
        alto_final = max(int(ancho_caja / relacion_img), 1)
    else:
        alto_final = alto_caja
        ancho_final = max(int(alto_caja * relacion_img), 1)

    left_centrado = int(left + (ancho_caja - ancho_final) / 2)
    top_centrado = int(top + (alto_caja - alto_final) / 2)

    slide.shapes.add_picture(
        io.BytesIO(datos_bytes),
        left_centrado,
        top_centrado,
        width=ancho_final,
        height=alto_final,
    )
    return True


def _normalizar_puntos(puntos_raw):
    """Los puntos vienen como [{"titulo","detalle"}, ...]. Si el
    modelo se despista y devuelve frases sueltas, las adapta al mismo
    formato para que el resto del código no tenga que preocuparse."""

    normalizados = []
    for p in puntos_raw or []:
        if isinstance(p, dict):
            normalizados.append(
                {
                    "titulo": str(p.get("titulo", "")).strip(),
                    "detalle": str(p.get("detalle", "")).strip(),
                }
            )
        elif p:
            normalizados.append({"titulo": str(p).strip(), "detalle": ""})

    return [p for p in normalizados if p["titulo"]][:5]


def crear_pptx(contenido, ruta, cache_imagenes=None):
    """Construye un archivo .pptx con diseño de tarjetas -portada con
    color/foto y categoría (kicker), cabecera con categoría + título,
    una tarjeta por punto (título corto + detalle) y foto o diagrama
    real cuando la IA lo ha pedido- y lo guarda en `ruta`.
    `cache_imagenes` son los bytes ya descargados por
    `obtener_imagenes_presentacion` (si no se pasa, se descargan
    aquí mismo, pero lo normal es pasarlos para no repetir las
    llamadas a Pexels/Wikimedia)."""

    if cache_imagenes is None:
        cache_imagenes = obtener_imagenes_presentacion(contenido)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    ancho = prs.slide_width
    alto = prs.slide_height

    layout_en_blanco = prs.slide_layouts[6]

    # --------------------------------------------
    # Diapositiva de portada
    # --------------------------------------------

    slide = prs.slides.add_slide(layout_en_blanco)

    imagen_portada, _estilo_portada = cache_imagenes.get("portada") or (None, None)
    titulo_txt = str(contenido.get("titulo", ""))
    categoria_portada = str(contenido.get("categoria_portada", "")).upper().strip()

    if imagen_portada:
        # Diseño con foto: panel oscuro a la izquierda con el texto,
        # foto a pantalla completa ocupando el resto (estilo revista).
        ancho_panel = int(ancho * 0.4)

        _pptx_imagen_cubrir(
            slide, imagen_portada, ancho_panel, 0, ancho - ancho_panel, alto
        )
        _pptx_rectangulo(slide, 0, 0, ancho_panel, alto, _PPTX_COLOR_FONDO_OSCURO)
        _pptx_rectangulo(slide, 0, 0, Inches(0.18), alto, _PPTX_COLOR_ACENTO)

        left_panel = Inches(0.8)
        ancho_panel_texto = ancho_panel - Inches(1.4)
        y = Inches(2.15)

        if categoria_portada:
            marco_cat = _pptx_caja_texto(slide, left_panel, y, ancho_panel_texto, Inches(0.4))
            pc = marco_cat.paragraphs[0]
            pc.text = categoria_portada
            pc.font.size = Pt(12)
            pc.font.bold = True
            pc.font.name = _PPTX_FUENTE
            pc.font.color.rgb = _PPTX_COLOR_ACENTO
            y += Inches(0.45)

        tamano_titulo = Pt(36) if len(titulo_txt) <= 40 else Pt(27)

        marco_titulo = _pptx_caja_texto(slide, left_panel, y, ancho_panel_texto, Inches(1.8))
        p = marco_titulo.paragraphs[0]
        p.text = titulo_txt
        p.font.size = tamano_titulo
        p.font.bold = True
        p.font.name = _PPTX_FUENTE
        p.font.color.rgb = _PPTX_COLOR_TEXTO_CLARO

        y += Inches(1.5) if tamano_titulo == Pt(36) else Inches(1.9)

        _pptx_rectangulo(slide, left_panel, y, Inches(1.1), Pt(4), _PPTX_COLOR_ACENTO)
        y += Inches(0.25)

        if contenido.get("subtitulo"):
            marco_sub = _pptx_caja_texto(slide, left_panel, y, ancho_panel_texto, Inches(2))
            ps = marco_sub.paragraphs[0]
            ps.text = str(contenido["subtitulo"])
            ps.font.size = Pt(15)
            ps.font.name = _PPTX_FUENTE
            ps.font.color.rgb = _PPTX_COLOR_SUBTITULO_CLARO

    else:
        # Diseño de siempre (solo color y texto), cuando no se ha
        # pedido imagen para la portada o no se ha encontrado ninguna.
        _pptx_rectangulo(slide, 0, 0, ancho, alto, _PPTX_COLOR_FONDO_OSCURO)
        _pptx_rectangulo(slide, 0, 0, Inches(0.18), alto, _PPTX_COLOR_ACENTO)

        left_full = Inches(1)
        ancho_full = ancho - Inches(2)
        y = Inches(2.2)

        if categoria_portada:
            marco_cat = _pptx_caja_texto(slide, left_full, y, ancho_full, Inches(0.4))
            pc = marco_cat.paragraphs[0]
            pc.text = categoria_portada
            pc.font.size = Pt(13)
            pc.font.bold = True
            pc.font.name = _PPTX_FUENTE
            pc.font.color.rgb = _PPTX_COLOR_ACENTO
            y += Inches(0.5)

        tamano_titulo = Pt(44) if len(titulo_txt) <= 40 else Pt(32)

        marco_titulo = _pptx_caja_texto(slide, left_full, y, ancho_full, Inches(1.8))
        p = marco_titulo.paragraphs[0]
        p.text = titulo_txt
        p.font.size = tamano_titulo
        p.font.bold = True
        p.font.name = _PPTX_FUENTE
        p.font.color.rgb = _PPTX_COLOR_TEXTO_CLARO

        y += Inches(1.55) if tamano_titulo == Pt(44) else Inches(1.95)

        _pptx_rectangulo(slide, left_full, y, Inches(1.1), Pt(4), _PPTX_COLOR_ACENTO)
        y += Inches(0.25)

        if contenido.get("subtitulo"):
            marco_sub = _pptx_caja_texto(slide, left_full, y, ancho_full, Inches(1))
            ps = marco_sub.paragraphs[0]
            ps.text = str(contenido["subtitulo"])
            ps.font.size = Pt(20)
            ps.font.name = _PPTX_FUENTE
            ps.font.color.rgb = _PPTX_COLOR_SUBTITULO_CLARO

    # --------------------------------------------
    # Diapositivas de contenido
    # --------------------------------------------

    diapositivas = contenido.get("diapositivas", [])
    total = len(diapositivas)

    for numero, diapo in enumerate(diapositivas, start=1):

        slide = prs.slides.add_slide(layout_en_blanco)

        # Fondo ligeramente gris para que las tarjetas blancas resalten.
        _pptx_rectangulo(slide, 0, 0, ancho, alto, _PPTX_COLOR_FONDO_CLARO)

        # Cabecera oscura con categoría (kicker) encima del título.
        alto_cabecera = Inches(1.35)
        _pptx_rectangulo(slide, 0, 0, ancho, alto_cabecera, _PPTX_COLOR_FONDO_OSCURO)
        _pptx_rectangulo(slide, 0, 0, Inches(0.18), alto_cabecera, _PPTX_COLOR_ACENTO)

        categoria = str(diapo.get("categoria", "")).upper().strip()
        titulo_diapo = str(diapo.get("titulo", ""))
        tamano_titulo_diapo = Pt(24) if len(titulo_diapo) <= 45 else Pt(19)

        marco_cab = _pptx_caja_texto(
            slide, Inches(0.7), Inches(0.16), ancho - Inches(2.6), Inches(1.05)
        )

        if categoria:
            p_cat = marco_cab.paragraphs[0]
            p_cat.text = categoria
            p_cat.font.size = Pt(11)
            p_cat.font.bold = True
            p_cat.font.name = _PPTX_FUENTE
            p_cat.font.color.rgb = _PPTX_COLOR_ACENTO
            p_cat.space_after = Pt(3)
            p_titulo = marco_cab.add_paragraph()
        else:
            p_titulo = marco_cab.paragraphs[0]

        p_titulo.text = titulo_diapo
        p_titulo.font.size = tamano_titulo_diapo
        p_titulo.font.bold = True
        p_titulo.font.name = _PPTX_FUENTE
        p_titulo.font.color.rgb = _PPTX_COLOR_TEXTO_CLARO

        # Numeración arriba a la derecha.
        marco_num = _pptx_caja_texto(
            slide,
            ancho - Inches(1.8),
            Inches(0.16),
            Inches(1.3),
            alto_cabecera - Inches(0.16),
            anclaje_vertical=MSO_ANCHOR.MIDDLE,
        )
        pn = marco_num.paragraphs[0]
        pn.text = f"{numero:02d} / {total:02d}"
        pn.alignment = PP_ALIGN.RIGHT
        pn.font.size = Pt(12)
        pn.font.name = _PPTX_FUENTE
        pn.font.color.rgb = _PPTX_COLOR_SUBTITULO_CLARO

        # Imagen real (si la IA la ha pedido y se ha encontrado):
        # la diapositiva pasa a un diseño a dos columnas, con las
        # tarjetas a la izquierda y la imagen a la derecha.
        imagen_diapo, estilo_imagen = cache_imagenes.get(numero - 1) or (None, None)

        margen_izq = Inches(0.7)
        margen_der = Inches(0.7)
        top_cuerpo = alto_cabecera + Inches(0.3)
        alto_cuerpo = alto - top_cuerpo - Inches(0.45)

        if imagen_diapo:
            hueco = Inches(0.4)
            espacio_total = ancho - margen_izq - margen_der - hueco

            ancho_tarjetas = int(espacio_total * 0.52)
            ancho_imagen = espacio_total - ancho_tarjetas
            left_imagen = margen_izq + ancho_tarjetas + hueco

            if estilo_imagen == "foto":
                _pptx_imagen_cubrir(
                    slide, imagen_diapo, left_imagen, top_cuerpo, ancho_imagen, alto_cuerpo
                )
            else:
                # Diagrama/ilustración: se ajusta sin recortar para no
                # perder etiquetas ni detalles, sobre un fondo blanco.
                _pptx_rectangulo(
                    slide, left_imagen, top_cuerpo, ancho_imagen, alto_cuerpo,
                    _PPTX_COLOR_TARJETA_FONDO,
                )
                _pptx_imagen_contener(
                    slide, imagen_diapo, left_imagen, top_cuerpo, ancho_imagen, alto_cuerpo
                )

            ancho_tarjeta_final = ancho_tarjetas
        else:
            ancho_tarjeta_final = ancho - margen_izq - margen_der

        # Tarjetas de los puntos, apiladas en la columna de texto.
        puntos = _normalizar_puntos(diapo.get("puntos"))

        if puntos:
            n = len(puntos)
            gap = Inches(0.16)
            alto_tarjeta = (alto_cuerpo - gap * (n - 1)) / n

            # Si hay pocos puntos, las tarjetas no crecen sin límite:
            # se quedan a un tamaño cómodo y el bloque se centra en
            # el espacio disponible, en vez de dejar tarjetas enormes
            # con mucho hueco vacío debajo del texto.
            alto_maxima = Inches(1.85)
            if alto_tarjeta > alto_maxima:
                alto_tarjeta = alto_maxima

            alto_total = alto_tarjeta * n + gap * (n - 1)
            y_tarjeta = top_cuerpo + max(0, int((alto_cuerpo - alto_total) / 2))
            for punto in puntos:
                _pptx_tarjeta_punto(
                    slide,
                    margen_izq,
                    y_tarjeta,
                    ancho_tarjeta_final,
                    alto_tarjeta,
                    punto["titulo"],
                    punto["detalle"],
                )
                y_tarjeta += alto_tarjeta + gap

        # Marca de pie de página.
        marco_pie = _pptx_caja_texto(
            slide, margen_izq, alto - Inches(0.38), Inches(3), Inches(0.32)
        )
        pp_ = marco_pie.paragraphs[0]
        pp_.text = "ORIA"
        pp_.font.size = Pt(10)
        pp_.font.bold = True
        pp_.font.name = _PPTX_FUENTE
        pp_.font.color.rgb = _PPTX_COLOR_TEXTO_SECUNDARIO

    prs.save(ruta)


# --------------------------------------------------------------------
# Exportación a PDF (mismo diseño que crear_pptx, dibujado con reportlab
# en vez de LibreOffice/Chromium para no arriesgar el despliegue
# gratuito en Streamlit Cloud con dependencias de sistema pesadas).
# --------------------------------------------------------------------

_PDF_COLOR_FONDO_OSCURO = _pdf_color("#1B1B2A")
_PDF_COLOR_ACENTO = _pdf_color("#6C5CE7")
_PDF_COLOR_TEXTO_OSCURO = _pdf_color("#2B2B31")
_PDF_COLOR_TEXTO_CLARO = _pdf_color("#FFFFFF")
_PDF_COLOR_TEXTO_SECUNDARIO = _pdf_color("#8A8A96")
_PDF_COLOR_SUBTITULO_CLARO = _pdf_color("#C9C9DC")
_PDF_COLOR_LINEA = _pdf_color("#E3E3E8")
_PDF_COLOR_FONDO_CLARO = _pdf_color("#F6F6FA")
_PDF_COLOR_TARJETA_FONDO = _pdf_color("#FFFFFF")

_PDF_FUENTE = "Helvetica"
_PDF_FUENTE_NEGRITA = "Helvetica-Bold"

_PDF_ANCHO = 13.333 * _PDF_INCH
_PDF_ALTO = 7.5 * _PDF_INCH


def _pdf_ajustar_texto(texto, fuente, tamano, ancho_max):
    """Parte `texto` en líneas que caben en `ancho_max` puntos con esa
    fuente y tamaño (envoltura de palabras, equivalente al ajuste de
    línea automático que ya hace PowerPoint en crear_pptx)."""

    palabras = str(texto or "").split()
    lineas = []
    actual = ""
    for palabra in palabras:
        prueba = (actual + " " + palabra).strip()
        if not actual or stringWidth(prueba, fuente, tamano) <= ancho_max:
            actual = prueba
        else:
            lineas.append(actual)
            actual = palabra
    if actual:
        lineas.append(actual)
    return lineas


def _pdf_rect(c, x, top, ancho, alto, color, radio=None):
    """Dibuja un rectángulo (opcionalmente con esquinas redondeadas)
    usando coordenadas 'desde arriba' como en python-pptx, para no
    tener que pensar en el sistema de coordenadas de reportlab
    (origen abajo a la izquierda) en el resto de la función."""

    c.setFillColor(color)
    y = _PDF_ALTO - top - alto
    if radio:
        c.roundRect(x, y, ancho, alto, radio, fill=1, stroke=0)
    else:
        c.rect(x, y, ancho, alto, fill=1, stroke=0)


def _pdf_texto(c, x, top, texto, fuente, tamano, color, ancho_max=None,
               max_lineas=None, interlineado=None, alineacion="izq"):
    """Dibuja texto (envuelto en varias líneas si se da `ancho_max`)
    empezando en `top` (distancia desde arriba de la página) y
    devuelve el `top` justo debajo del bloque escrito."""

    interlineado = interlineado or tamano * 1.28
    if ancho_max:
        lineas = _pdf_ajustar_texto(texto, fuente, tamano, ancho_max)
    else:
        lineas = [str(texto or "")]
    if max_lineas:
        lineas = lineas[:max_lineas]

    c.setFont(fuente, tamano)
    c.setFillColor(color)
    y_cursor = top
    for linea in lineas:
        baseline = _PDF_ALTO - y_cursor - tamano
        if alineacion == "der" and ancho_max:
            c.drawRightString(x + ancho_max, baseline, linea)
        else:
            c.drawString(x, baseline, linea)
        y_cursor += interlineado
    return y_cursor


def _pdf_medida_imagen_contener(datos_bytes, ancho_caja, alto_caja):
    """Calcula el tamaño con el que debe dibujarse una imagen para que
    quepa entera dentro de la caja sin recortarla (misma lógica que
    _pptx_imagen_contener, para diagramas/ilustraciones)."""

    imagen = Image.open(io.BytesIO(datos_bytes))
    ancho_img, alto_img = imagen.size
    ratio_caja = ancho_caja / alto_caja
    ratio_img = ancho_img / alto_img
    if ratio_img > ratio_caja:
        ancho_dibujo = ancho_caja
        alto_dibujo = ancho_caja / ratio_img
    else:
        alto_dibujo = alto_caja
        ancho_dibujo = alto_caja * ratio_img
    return ancho_dibujo, alto_dibujo


def _pdf_imagen_cubrir(c, datos_bytes, left, top, ancho_caja, alto_caja):
    """Recorta (con PIL, reutilizando el mismo helper que usa el pptx)
    y dibuja una foto rellenando toda la caja, sin dejar huecos."""

    try:
        recortada = _preparar_imagen_recortada(datos_bytes, ancho_caja, alto_caja)
        y = _PDF_ALTO - top - alto_caja
        c.drawImage(
            ImageReader(recortada), left, y, width=ancho_caja, height=alto_caja,
            preserveAspectRatio=False, mask="auto",
        )
    except Exception:
        pass


def _pdf_imagen_contener(c, datos_bytes, left, top, ancho_caja, alto_caja):
    """Dibuja un diagrama/ilustración ajustado dentro de la caja sin
    recortar, centrado, sobre el fondo blanco que ya se ha pintado."""

    try:
        ancho_dibujo, alto_dibujo = _pdf_medida_imagen_contener(
            datos_bytes, ancho_caja, alto_caja
        )
        left_img = left + (ancho_caja - ancho_dibujo) / 2
        top_img = top + (alto_caja - alto_dibujo) / 2
        y = _PDF_ALTO - top_img - alto_dibujo
        c.drawImage(
            ImageReader(io.BytesIO(datos_bytes)), left_img, y,
            width=ancho_dibujo, height=alto_dibujo,
            preserveAspectRatio=True, mask="auto",
        )
    except Exception:
        pass


def _pdf_tarjeta_punto(c, left, top, width, height, titulo, detalle):
    """Dibuja la misma 'tarjeta' de crear_pptx (fondo blanco, esquinas
    redondeadas, barra de acento a la izquierda, título en negrita y
    detalle debajo) pero con reportlab."""

    c.setFillColor(_PDF_COLOR_TARJETA_FONDO)
    c.setStrokeColor(_PDF_COLOR_LINEA)
    c.setLineWidth(0.75)
    y = _PDF_ALTO - top - height
    radio = min(height, width) * 0.07
    c.roundRect(left, y, width, height, radio, fill=1, stroke=1)

    _pdf_rect(c, left, top, 4, height, _PDF_COLOR_ACENTO)

    margen_h = 0.22 * _PDF_INCH
    margen_v = 0.15 * _PDF_INCH
    ancho_texto = width - margen_h * 2

    lineas_titulo = _pdf_ajustar_texto(titulo, _PDF_FUENTE_NEGRITA, 15, ancho_texto)
    lineas_detalle = (
        _pdf_ajustar_texto(detalle, _PDF_FUENTE, 12, ancho_texto) if detalle else []
    )

    alto_bloque = len(lineas_titulo) * 15 * 1.22 + (3 if lineas_detalle else 0) \
        + len(lineas_detalle) * 12 * 1.28
    top_texto = top + max(margen_v, (height - alto_bloque) / 2)

    top_texto = _pdf_texto(
        c, left + margen_h, top_texto, titulo, _PDF_FUENTE_NEGRITA, 15,
        _PDF_COLOR_TEXTO_OSCURO, ancho_max=ancho_texto, interlineado=15 * 1.22,
    )
    if detalle:
        top_texto += 3
        _pdf_texto(
            c, left + margen_h, top_texto, detalle, _PDF_FUENTE, 12,
            _PDF_COLOR_TEXTO_SECUNDARIO, ancho_max=ancho_texto, interlineado=12 * 1.28,
        )


def crear_pdf(contenido, ruta, cache_imagenes=None):
    """Genera un PDF con exactamente el mismo diseño que crear_pptx
    (portada con foto/color y kicker, diapositivas con tarjetas), para
    que el usuario pueda elegir el formato de descarga que prefiera."""

    if cache_imagenes is None:
        cache_imagenes = obtener_imagenes_presentacion(contenido)

    c = _pdf_canvas.Canvas(ruta, pagesize=(_PDF_ANCHO, _PDF_ALTO))

    # --------------------------------------------
    # Portada
    # --------------------------------------------

    imagen_portada, _estilo_portada = cache_imagenes.get("portada") or (None, None)
    titulo_txt = str(contenido.get("titulo", ""))
    categoria_portada = str(contenido.get("categoria_portada", "")).upper().strip()

    if imagen_portada:
        ancho_panel = _PDF_ANCHO * 0.4
        _pdf_imagen_cubrir(c, imagen_portada, ancho_panel, 0, _PDF_ANCHO - ancho_panel, _PDF_ALTO)
        _pdf_rect(c, 0, 0, ancho_panel, _PDF_ALTO, _PDF_COLOR_FONDO_OSCURO)
        _pdf_rect(c, 0, 0, 0.18 * _PDF_INCH, _PDF_ALTO, _PDF_COLOR_ACENTO)

        left_panel = 0.8 * _PDF_INCH
        ancho_panel_texto = ancho_panel - 1.4 * _PDF_INCH
        y = 2.15 * _PDF_INCH

        if categoria_portada:
            y = _pdf_texto(c, left_panel, y, categoria_portada, _PDF_FUENTE_NEGRITA, 12,
                            _PDF_COLOR_ACENTO, ancho_max=ancho_panel_texto)
            y += 0.15 * _PDF_INCH

        tamano_titulo = 36 if len(titulo_txt) <= 40 else 27
        y = _pdf_texto(c, left_panel, y, titulo_txt, _PDF_FUENTE_NEGRITA, tamano_titulo,
                        _PDF_COLOR_TEXTO_CLARO, ancho_max=ancho_panel_texto,
                        interlineado=tamano_titulo * 1.18, max_lineas=3)
        y += 0.2 * _PDF_INCH

        _pdf_rect(c, left_panel, y, 1.1 * _PDF_INCH, 4, _PDF_COLOR_ACENTO)
        y += 0.3 * _PDF_INCH

        if contenido.get("subtitulo"):
            _pdf_texto(c, left_panel, y, str(contenido["subtitulo"]), _PDF_FUENTE, 15,
                        _PDF_COLOR_SUBTITULO_CLARO, ancho_max=ancho_panel_texto, max_lineas=3)
    else:
        _pdf_rect(c, 0, 0, _PDF_ANCHO, _PDF_ALTO, _PDF_COLOR_FONDO_OSCURO)
        _pdf_rect(c, 0, 0, 0.18 * _PDF_INCH, _PDF_ALTO, _PDF_COLOR_ACENTO)

        left_full = 1 * _PDF_INCH
        ancho_full = _PDF_ANCHO - 2 * _PDF_INCH
        y = 2.2 * _PDF_INCH

        if categoria_portada:
            y = _pdf_texto(c, left_full, y, categoria_portada, _PDF_FUENTE_NEGRITA, 13,
                            _PDF_COLOR_ACENTO, ancho_max=ancho_full)
            y += 0.2 * _PDF_INCH

        tamano_titulo = 44 if len(titulo_txt) <= 40 else 32
        y = _pdf_texto(c, left_full, y, titulo_txt, _PDF_FUENTE_NEGRITA, tamano_titulo,
                        _PDF_COLOR_TEXTO_CLARO, ancho_max=ancho_full,
                        interlineado=tamano_titulo * 1.18, max_lineas=3)
        y += 0.2 * _PDF_INCH

        _pdf_rect(c, left_full, y, 1.1 * _PDF_INCH, 4, _PDF_COLOR_ACENTO)
        y += 0.3 * _PDF_INCH

        if contenido.get("subtitulo"):
            _pdf_texto(c, left_full, y, str(contenido["subtitulo"]), _PDF_FUENTE, 20,
                        _PDF_COLOR_SUBTITULO_CLARO, ancho_max=ancho_full, max_lineas=2)

    # --------------------------------------------
    # Diapositivas de contenido
    # --------------------------------------------

    diapositivas = contenido.get("diapositivas", [])
    total = len(diapositivas)

    for numero, diapo in enumerate(diapositivas, start=1):
        c.showPage()

        _pdf_rect(c, 0, 0, _PDF_ANCHO, _PDF_ALTO, _PDF_COLOR_FONDO_CLARO)

        alto_cabecera = 1.35 * _PDF_INCH
        _pdf_rect(c, 0, 0, _PDF_ANCHO, alto_cabecera, _PDF_COLOR_FONDO_OSCURO)
        _pdf_rect(c, 0, 0, 0.18 * _PDF_INCH, alto_cabecera, _PDF_COLOR_ACENTO)

        categoria = str(diapo.get("categoria", "")).upper().strip()
        titulo_diapo = str(diapo.get("titulo", ""))
        tamano_titulo_diapo = 24 if len(titulo_diapo) <= 45 else 19

        ancho_cabecera_texto = _PDF_ANCHO - 2.6 * _PDF_INCH
        y_cab = 0.2 * _PDF_INCH
        if categoria:
            y_cab = _pdf_texto(c, 0.7 * _PDF_INCH, y_cab, categoria, _PDF_FUENTE_NEGRITA, 11,
                                _PDF_COLOR_ACENTO, ancho_max=ancho_cabecera_texto)
            y_cab += 0.05 * _PDF_INCH
        _pdf_texto(c, 0.7 * _PDF_INCH, y_cab, titulo_diapo, _PDF_FUENTE_NEGRITA,
                   tamano_titulo_diapo, _PDF_COLOR_TEXTO_CLARO,
                   ancho_max=ancho_cabecera_texto, max_lineas=2,
                   interlineado=tamano_titulo_diapo * 1.15)

        c.setFont(_PDF_FUENTE, 12)
        c.setFillColor(_PDF_COLOR_SUBTITULO_CLARO)
        c.drawRightString(_PDF_ANCHO - 0.7 * _PDF_INCH, _PDF_ALTO - 0.16 * _PDF_INCH - 12,
                           f"{numero:02d} / {total:02d}")

        imagen_diapo, estilo_imagen = cache_imagenes.get(numero - 1) or (None, None)

        margen_izq = 0.7 * _PDF_INCH
        margen_der = 0.7 * _PDF_INCH
        top_cuerpo = alto_cabecera + 0.3 * _PDF_INCH
        alto_cuerpo = _PDF_ALTO - top_cuerpo - 0.45 * _PDF_INCH

        if imagen_diapo:
            hueco = 0.4 * _PDF_INCH
            espacio_total = _PDF_ANCHO - margen_izq - margen_der - hueco
            ancho_tarjetas = espacio_total * 0.52
            ancho_imagen = espacio_total - ancho_tarjetas
            left_imagen = margen_izq + ancho_tarjetas + hueco

            if estilo_imagen == "foto":
                _pdf_imagen_cubrir(c, imagen_diapo, left_imagen, top_cuerpo, ancho_imagen, alto_cuerpo)
            else:
                _pdf_rect(c, left_imagen, top_cuerpo, ancho_imagen, alto_cuerpo, _PDF_COLOR_TARJETA_FONDO)
                _pdf_imagen_contener(c, imagen_diapo, left_imagen, top_cuerpo, ancho_imagen, alto_cuerpo)

            ancho_tarjeta_final = ancho_tarjetas
        else:
            ancho_tarjeta_final = _PDF_ANCHO - margen_izq - margen_der

        puntos = _normalizar_puntos(diapo.get("puntos"))

        if puntos:
            n = len(puntos)
            gap = 0.16 * _PDF_INCH
            alto_tarjeta = (alto_cuerpo - gap * (n - 1)) / n
            alto_maxima = 1.85 * _PDF_INCH
            if alto_tarjeta > alto_maxima:
                alto_tarjeta = alto_maxima
            alto_total = alto_tarjeta * n + gap * (n - 1)
            y_tarjeta = top_cuerpo + max(0, (alto_cuerpo - alto_total) / 2)
            for punto in puntos:
                _pdf_tarjeta_punto(
                    c, margen_izq, y_tarjeta, ancho_tarjeta_final, alto_tarjeta,
                    punto["titulo"], punto["detalle"],
                )
                y_tarjeta += alto_tarjeta + gap

        c.setFont(_PDF_FUENTE_NEGRITA, 10)
        c.setFillColor(_PDF_COLOR_TEXTO_SECUNDARIO)
        c.drawString(margen_izq, 0.28 * _PDF_INCH, "ORIA")

    c.save()


# --------------------------------------------------------------------
# Vista previa HTML de la presentación (para verla en el chat antes de
# descargarla, con el mismo diseño de tarjetas que crear_pptx/crear_pdf
# y botones de anterior/siguiente, sin depender de nada externo).
# --------------------------------------------------------------------

def _imagen_a_data_uri(datos_bytes):
    """Convierte los bytes de una imagen ya descargada en un data-URI
    JPEG listo para incrustar directamente en el HTML (sin subir nada
    a ningún sitio, todo queda dentro de la propia vista previa)."""

    if not datos_bytes:
        return None
    try:
        imagen = Image.open(io.BytesIO(datos_bytes))
        if imagen.mode != "RGB":
            imagen = imagen.convert("RGB")
        buffer = io.BytesIO()
        imagen.save(buffer, format="JPEG", quality=82)
        codificado = base64.b64encode(buffer.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{codificado}"
    except Exception:
        return None


def generar_html_presentacion(contenido, cache_imagenes):
    """Genera una vista previa HTML autocontenida (portada + tarjetas,
    igual que en el .pptx/.pdf) con navegación anterior/siguiente, para
    que el usuario pueda verla en el chat antes de descargarla."""

    titulo_txt = html.escape(str(contenido.get("titulo", "")))
    subtitulo_txt = html.escape(str(contenido.get("subtitulo", "")))
    categoria_portada = html.escape(
        str(contenido.get("categoria_portada", "")).upper().strip()
    )
    diapositivas = contenido.get("diapositivas", [])
    total_diapos = len(diapositivas) + 1

    imagen_portada, _estilo_portada = cache_imagenes.get("portada") or (None, None)
    uri_portada = _imagen_a_data_uri(imagen_portada)

    if uri_portada:
        bloque_portada = f"""
        <div class="oria-slide oria-cover-foto" data-i="0">
          <div class="oria-panel">
            {f'<div class="oria-kicker">{categoria_portada}</div>' if categoria_portada else ''}
            <div class="oria-titulo-portada">{titulo_txt}</div>
            <div class="oria-barra"></div>
            {f'<div class="oria-subtitulo">{subtitulo_txt}</div>' if subtitulo_txt else ''}
          </div>
          <div class="oria-foto-portada" style="background-image:url('{uri_portada}')"></div>
        </div>"""
    else:
        bloque_portada = f"""
        <div class="oria-slide oria-cover-solido" data-i="0">
          {f'<div class="oria-kicker">{categoria_portada}</div>' if categoria_portada else ''}
          <div class="oria-titulo-portada-solido">{titulo_txt}</div>
          <div class="oria-barra"></div>
          {f'<div class="oria-subtitulo-solido">{subtitulo_txt}</div>' if subtitulo_txt else ''}
        </div>"""

    partes = [bloque_portada]

    for idx, diapo in enumerate(diapositivas, start=1):
        categoria = html.escape(str(diapo.get("categoria", "")).upper().strip())
        titulo_diapo = html.escape(str(diapo.get("titulo", "")))
        puntos = _normalizar_puntos(diapo.get("puntos"))
        imagen_diapo, estilo_imagen = cache_imagenes.get(idx - 1) or (None, None)
        uri_imagen = _imagen_a_data_uri(imagen_diapo)

        tarjetas_html = "".join(
            f"""<div class="oria-tarjeta">
                  <div class="oria-tarjeta-titulo">{html.escape(p['titulo'])}</div>
                  {f'<div class="oria-tarjeta-detalle">{html.escape(p["detalle"])}</div>' if p['detalle'] else ''}
                </div>"""
            for p in puntos
        )

        if uri_imagen:
            clase_fit = "oria-img-cubrir" if estilo_imagen == "foto" else "oria-img-contener"
            columna_imagen = f"""<div class="oria-col-imagen {clase_fit}">
                  <img src="{uri_imagen}" alt="" />
                </div>"""
            clase_cuerpo = "oria-cuerpo-con-imagen"
        else:
            columna_imagen = ""
            clase_cuerpo = "oria-cuerpo-sin-imagen"

        partes.append(f"""
        <div class="oria-slide oria-contenido" data-i="{idx}">
          <div class="oria-cabecera">
            <div class="oria-cabecera-texto">
              {f'<div class="oria-kicker-chico">{categoria}</div>' if categoria else ''}
              <div class="oria-titulo-diapo">{titulo_diapo}</div>
            </div>
            <div class="oria-numero">{idx:02d} / {len(diapositivas):02d}</div>
          </div>
          <div class="oria-cuerpo {clase_cuerpo}">
            <div class="oria-col-tarjetas">{tarjetas_html}</div>
            {columna_imagen}
          </div>
          <div class="oria-pie">ORIA</div>
        </div>""")

    slides_html = "".join(partes)

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; }}
  html, body {{ margin:0; padding:0; background:transparent;
        font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif; }}
  #oria-wrap {{ width:100%; max-width:920px; margin:0 auto; }}
  #oria-marco {{ position:relative; width:100%; aspect-ratio: 13.333 / 7.5;
        background:#fff; border-radius:12px; overflow:hidden;
        box-shadow:0 8px 30px rgba(0,0,0,0.18); }}
  .oria-slide {{ position:absolute; inset:0; display:none; }}
  .oria-slide.oria-activa {{ display:flex; flex-direction:column; }}
  .oria-slide.oria-cover-foto.oria-activa {{ flex-direction:row; }}
  .oria-cover-solido {{ background:#1B1B2A; padding: 8% 6%; justify-content:center; }}
  .oria-cover-solido::before {{ content:""; position:absolute; left:0; top:0;
        bottom:0; width:6px; background:#6C5CE7; }}
  .oria-panel {{ width:40%; height:100%; background:#1B1B2A; padding: 7% 6%;
        position:relative; display:flex; flex-direction:column; justify-content:center; }}
  .oria-panel::before {{ content:""; position:absolute; left:0; top:0; bottom:0;
        width:6px; background:#6C5CE7; }}
  .oria-foto-portada {{ width:60%; height:100%; background-size:cover;
        background-position:center; }}
  .oria-kicker {{ color:#6C5CE7; font-weight:700; font-size: clamp(9px,1.3vw,14px);
        letter-spacing:0.5px; margin-bottom: 6%; }}
  .oria-titulo-portada, .oria-titulo-portada-solido {{ color:#fff; font-weight:800;
        line-height:1.18; font-size: clamp(16px, 2.6vw, 32px); margin-bottom:5%; }}
  .oria-barra {{ width: 46px; height:4px; background:#6C5CE7; margin-bottom: 5%; }}
  .oria-subtitulo, .oria-subtitulo-solido {{ color:#C9C9DC; font-size: clamp(10px,1.15vw,15px);
        line-height:1.4; }}
  .oria-contenido {{ background:#F6F6FA; }}
  .oria-cabecera {{ background:#1B1B2A; position:relative; padding: 2.2% 3%;
        display:flex; justify-content:space-between; align-items:flex-start; }}
  .oria-cabecera::before {{ content:""; position:absolute; left:0; top:0; bottom:0;
        width:6px; background:#6C5CE7; }}
  .oria-kicker-chico {{ color:#6C5CE7; font-weight:700; font-size: clamp(8px,1vw,12px);
        letter-spacing:0.5px; margin-bottom:3%; }}
  .oria-titulo-diapo {{ color:#fff; font-weight:800; font-size: clamp(13px,1.7vw,22px);
        line-height:1.2; }}
  .oria-numero {{ color:#C9C9DC; font-size: clamp(9px,1vw,13px); white-space:nowrap;
        padding-top:2px; }}
  .oria-cuerpo {{ flex:1; display:flex; gap: 2%; padding: 2.6% 3%; overflow:hidden; min-height:0; }}
  .oria-col-tarjetas {{ flex:1; display:flex; flex-direction:column; gap: 2.4%;
        justify-content:center; min-width:0; }}
  .oria-cuerpo-con-imagen .oria-col-tarjetas {{ flex: 0 0 52%; }}
  .oria-col-imagen {{ flex:1; border-radius:6px; overflow:hidden; background:#fff; }}
  .oria-col-imagen img {{ width:100%; height:100%; display:block; }}
  .oria-img-cubrir img {{ object-fit:cover; }}
  .oria-img-contener img {{ object-fit:contain; }}
  .oria-tarjeta {{ background:#fff; border:1px solid #E3E3E8; border-left:4px solid #6C5CE7;
        border-radius:8px; padding: 3% 3.5%; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }}
  .oria-tarjeta-titulo {{ color:#2B2B31; font-weight:700; font-size: clamp(9px,1.05vw,14px);
        margin-bottom:2%; }}
  .oria-tarjeta-detalle {{ color:#8A8A96; font-size: clamp(8px,0.9vw,12px); line-height:1.35; }}
  .oria-pie {{ position:absolute; bottom: 2%; left: 3%; color:#8A8A96; font-weight:700;
        font-size: clamp(7px,0.75vw,10px); }}
  #oria-nav {{ display:flex; align-items:center; justify-content:center; gap:14px; margin-top:12px; }}
  #oria-nav button {{ background:#1B1B2A; color:#fff; border:none; border-radius:6px;
        width:34px; height:34px; font-size:16px; cursor:pointer; display:flex;
        align-items:center; justify-content:center; }}
  #oria-nav button:disabled {{ opacity:0.35; cursor:default; }}
  #oria-nav button:hover:not(:disabled) {{ background:#6C5CE7; }}
  #oria-contador {{ color:#444; font-size:13px; min-width:90px; text-align:center; }}
</style></head>
<body>
  <div id="oria-wrap">
    <div id="oria-marco">{slides_html}</div>
    <div id="oria-nav">
      <button id="oria-btn-prev" onclick="oriaIr(-1)">&#8249;</button>
      <span id="oria-contador"></span>
      <button id="oria-btn-next" onclick="oriaIr(1)">&#8250;</button>
    </div>
  </div>
<script>
(function() {{
  var total = {total_diapos};
  var actual = 0;
  var slides = document.querySelectorAll('#oria-marco .oria-slide');
  function render() {{
    slides.forEach(function(s, i) {{
      s.classList.toggle('oria-activa', i === actual);
    }});
    document.getElementById('oria-contador').textContent = (actual + 1) + ' / ' + total;
    document.getElementById('oria-btn-prev').disabled = actual === 0;
    document.getElementById('oria-btn-next').disabled = actual === total - 1;
  }}
  window.oriaIr = function(delta) {{
    actual = Math.max(0, Math.min(total - 1, actual + delta));
    render();
  }};
  render();
}})();
</script>
</body></html>"""


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
# automáticamente; además el interruptor "Web" fuerza la búsqueda.
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
# 5C. DETECCIÓN AUTOMÁTICA DE INTENCIÓN (sin botones)
# ============================================================

# Herramientas que la IA puede "llamar" ella sola, según lo que haya
# escrito el usuario, en vez de que el usuario tenga que activar un
# interruptor a mano. Esto es lo que se llama "tool calling"/"function
# calling": le describimos a la IA qué acciones existen, y ella decide
# cuál (si alguna) encaja con el mensaje.
_HERRAMIENTAS_INTENCION = [
    {
        "type": "function",
        "function": {
            "name": "generar_imagen",
            "description": (
                "Úsala cuando el usuario pide explícitamente que se "
                "dibuje, genere, cree o haga una imagen, foto, dibujo "
                "o ilustración de algo. No la uses si solo está "
                "hablando de imágenes en general, o si ha adjuntado "
                "él mismo una imagen o un PDF."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "descripcion": {
                        "type": "string",
                        "description": (
                            "Qué imagen hay que generar, limpia de "
                            "frases como 'hazme una imagen de': solo "
                            "la descripción en sí."
                        ),
                    }
                },
                "required": ["descripcion"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "generar_presentacion",
            "description": (
                "Úsala cuando el usuario pide explícitamente una "
                "presentación, unas diapositivas, unos slides o un "
                "PowerPoint sobre un tema."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "tema": {
                        "type": "string",
                        "description": "El tema sobre el que debe tratar la presentación.",
                    }
                },
                "required": ["tema"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "buscar_en_internet",
            "description": (
                "Úsala cuando la pregunta depende de información "
                "actual, reciente o que cambia con el tiempo "
                "(noticias, resultados deportivos, precios, el "
                "tiempo, eventos de hoy, cotizaciones...) que no "
                "puedes saber de memoria porque tu entrenamiento "
                "tiene una fecha límite."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "consulta": {
                        "type": "string",
                        "description": "Qué hay que buscar en internet.",
                    }
                },
                "required": ["consulta"],
            },
        },
    },
]


def detectar_intencion(texto, idioma=IDIOMA_POR_DEFECTO):
    """Le pregunta a Groq (con function calling) qué acción encaja con
    el mensaje del usuario -imagen, presentación o búsqueda web-, para
    activarla automáticamente sin botones. Si algo falla o no aplica
    ninguna, se trata como conversación normal (nunca bloquea el chat).
    Devuelve {"accion": "imagen"|"presentacion"|"web"|"ninguna",
    "parametro": "<texto limpio para esa acción>"}."""

    sin_accion = {"accion": "ninguna", "parametro": texto}

    api_key = obtener_api_key()
    if not api_key or not api_key.startswith("gsk_"):
        return sin_accion

    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

        payload = {
            "model": MODELO_GROQ,
            "stream": False,
            "temperature": 0,
            "max_completion_tokens": 300,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "Decides qué acción necesita el mensaje del "
                        "usuario, sin responderlo todavía. Si pide una "
                        "imagen, llama a generar_imagen. Si pide una "
                        "presentación o PowerPoint, llama a "
                        "generar_presentacion. Si necesita datos "
                        "actuales de internet, llama a "
                        "buscar_en_internet. Si es una conversación "
                        "normal (una pregunta general, pedir una "
                        "explicación, un saludo, charlar...), NO "
                        "llames a ninguna herramienta."
                    ),
                },
                {"role": "user", "content": texto[:2000]},
            ],
            "tools": _HERRAMIENTAS_INTENCION,
            "tool_choice": "auto",
        }

        respuesta = requests.post(
            URL_GROQ, headers=headers, json=payload, timeout=10
        )

        if respuesta.status_code != 200:
            return sin_accion

        mensaje = respuesta.json().get("choices", [{}])[0].get("message", {})
        llamadas = mensaje.get("tool_calls") or []

        if not llamadas:
            return sin_accion

        llamada = llamadas[0].get("function", {})
        nombre = llamada.get("name", "")

        try:
            argumentos = json.loads(llamada.get("arguments") or "{}")
        except json.JSONDecodeError:
            argumentos = {}

        if nombre == "generar_imagen":
            return {
                "accion": "imagen",
                "parametro": argumentos.get("descripcion") or texto,
            }
        if nombre == "generar_presentacion":
            return {
                "accion": "presentacion",
                "parametro": argumentos.get("tema") or texto,
            }
        if nombre == "buscar_en_internet":
            return {
                "accion": "web",
                "parametro": argumentos.get("consulta") or texto,
            }

        return sin_accion

    except Exception:
        return sin_accion


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
    idioma=IDIOMA_POR_DEFECTO,
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
            "**No se ha encontrado `GROQ_API_KEY`.**\n\n"
            "Ve a **Streamlit → Settings → Secrets** y comprueba "
            "que tengas configurado:\n\n"
            "```toml\n"
            'GROQ_API_KEY = "TU_API_KEY"\n'
            "```"
        )
        return

    if not api_key.startswith("gsk_"):
        yield (
            "**La API key de Groq no parece válida.**\n\n"
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

    nombre_idioma = NOMBRE_IDIOMA_PARA_PROMPT.get(
        idioma, NOMBRE_IDIOMA_PARA_PROMPT[IDIOMA_POR_DEFECTO]
    )

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
        f"El usuario ha elegido {nombre_idioma} como idioma de la "
        f"aplicación: responde SIEMPRE en {nombre_idioma}, escriba el "
        "usuario en el idioma que escriba, salvo que te pida "
        "expresamente que le respondas en otro idioma distinto (en "
        "ese caso, sigue esa petición solo para esa respuesta). "
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
                "**Error 401: API key no válida.**\n\n"
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
                "**Error 403: acceso denegado por Groq.**\n\n"
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
                "**Error 404 de Groq.**\n\n"
                f"{error_message}\n\n"
                f"Modelo utilizado: `{modelo_usar}`"
            )

            return

        # ====================================================
        # MENSAJE DEMASIADO GRANDE (413)
        # ====================================================

        if response.status_code == 413:

            yield t("error_413", idioma=idioma)

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
            f"**Error de Groq ({response.status_code})**\n\n"
            f"{error_message}"
        )

    # ========================================================
    # TIMEOUT
    # ========================================================

    except requests.exceptions.Timeout:

        yield (
            "**Groq ha tardado demasiado en responder.**\n\n"
            "Vuelve a intentarlo."
        )

    # ========================================================
    # ERROR DE CONEXIÓN
    # ========================================================

    except requests.exceptions.ConnectionError:

        yield (
            "**No se ha podido conectar con Groq.**\n\n"
            "Comprueba la conexión del servidor."
        )

    # ========================================================
    # OTRO ERROR
    # ========================================================

    except requests.exceptions.RequestException as e:

        yield (
            "**Error de conexión con Groq:**\n\n"
            f"{str(e)}"
        )

    except Exception as e:

        yield (
            "**Se ha producido un error inesperado:**\n\n"
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

/* Deja hueco abajo para que el cuadro de texto no tape los mensajes */
[data-testid="stMainBlockContainer"] {
    padding-bottom: 6rem !important;
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
    nuestro botón ), la barra lateral entra en pantalla. */
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

    /* Los botones en columna en vez de apretujados uno junto al
    otro en pantallas estrechas. */
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
# puede activar por CSS): un botón que añade/quita una clase en
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
                "El login por correo no ha podido activarse:\n\n"
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
                "El login por correo (crear cuenta / contraseña) "
                "todavía no está activo porque no encuentro las claves "
                "de Supabase en los Secrets de la app:\n\n"
                f"- `SUPABASE_URL`: {'encontrada' if _tiene_url else 'no encontrada'}\n"
                f"- `SUPABASE_KEY`: {'encontrada' if _tiene_key else 'no encontrada'}\n\n"
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
    st.session_state.idioma = _datos_usuario.get("idioma", IDIOMA_POR_DEFECTO)
    st.session_state.current_chat_id = None

if "current_chat_id" not in st.session_state:
    st.session_state.current_chat_id = None

if "idioma" not in st.session_state:
    st.session_state.idioma = IDIOMA_POR_DEFECTO


def guardar_todo():
    """Guarda las conversaciones, la memoria y el idioma del usuario
    actual."""
    guardar_datos_usuario(
        st.session_state.usuario_id,
        st.session_state.chats,
        st.session_state.memoria,
        st.session_state.idioma,
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

    tab_idioma, tab_memoria, tab_cuenta, tab_info = st.tabs(
        [t("tab_idioma"), t("tab_memoria"), t("tab_cuenta"), t("tab_estado")]
    )

    # --------------------------------------------
    # IDIOMA DE ORIA
    # --------------------------------------------

    with tab_idioma:

        st.caption(t("idioma_caption"))

        codigos = list(IDIOMAS_DISPONIBLES.keys())
        idioma_actual = st.session_state.get("idioma", IDIOMA_POR_DEFECTO)

        indice_actual = (
            codigos.index(idioma_actual) if idioma_actual in codigos else 0
        )

        nuevo_idioma = st.selectbox(
            "Idioma de ORIA",
            options=codigos,
            index=indice_actual,
            format_func=lambda c: IDIOMAS_DISPONIBLES.get(c, c),
            label_visibility="collapsed",
        )

        if nuevo_idioma != idioma_actual:
            st.session_state.idioma = nuevo_idioma
            guardar_todo()
            st.success(
                t("idioma_cambiado", idioma=IDIOMAS_DISPONIBLES[nuevo_idioma])
            )
            st.rerun()

    # --------------------------------------------
    # MEMORIA DE ORIA
    # --------------------------------------------

    with tab_memoria:

        st.caption(t("memoria_caption"))

        nueva_memoria = st.text_area(
            "Datos a recordar",
            value=st.session_state.memoria,
            height=150,
            label_visibility="collapsed",
        )

        if st.button(t("guardar_memoria"), use_container_width=True):
            st.session_state.memoria = nueva_memoria
            guardar_todo()
            st.success(t("memoria_guardada"))

    # --------------------------------------------
    # CUENTA
    # --------------------------------------------

    with tab_cuenta:

        if st.session_state.get("es_invitado"):
            st.caption(t("cuenta_invitado"))
        else:
            st.caption(
                t(
                    "cuenta_sesion_como",
                    nombre=st.session_state.get("usuario_nombre", ""),
                )
            )

            st.button(
                t("cerrar_sesion"),
                on_click=_cerrar_sesion,
                use_container_width=True,
            )

            st.markdown("")

            with st.expander(t("eliminar_cuenta_titulo")):
                st.caption(t("eliminar_cuenta_caption"))
                if st.checkbox(
                    t("eliminar_cuenta_checkbox"),
                    key="confirmar_borrado",
                ):
                    if st.button(
                        t("eliminar_definitivamente"),
                        type="primary",
                        use_container_width=True,
                    ):
                        borrar_cuenta_usuario(
                            st.session_state.usuario_id
                        )
                        _cerrar_sesion()
                        st.success(t("datos_eliminados"))
                        st.rerun()

        st.markdown("---")
        st.caption(t("compartir_caption"))

    # --------------------------------------------
    # ESTADO DE LAS INTEGRACIONES
    # --------------------------------------------

    with tab_info:

        if obtener_tavily_key():
            st.caption(t("estado_web_activa"))
        else:
            st.caption(t("estado_web_inactiva"))

        if supabase_activo():
            st.caption(t("estado_nube"))
        else:
            st.caption(t("estado_servidor"))

        st.markdown("---")

        st.caption(t("estado_calidad_imagen"))


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
            t("nueva_conversacion"),
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
                    t("conversacion_sin_titulo"),
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

                if st.button("", key=f"del_{cid}"):
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
            t("invitado")
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
            f'<div class="oria-tagline">{t("tagline_bienvenida")}</div>',
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
                with st.expander(t("fuentes", n=len(fuentes))):
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

            idioma_mensaje = st.session_state.get("idioma", IDIOMA_POR_DEFECTO)
            codigo_voz_js = json.dumps(
                CODIGO_VOZ_NAVEGADOR.get(idioma_mensaje, "es-ES")
            )
            texto_copiar_js = json.dumps(t("boton_copiar", idioma=idioma_mensaje))
            texto_copiado_js = json.dumps(t("boton_copiado", idioma=idioma_mensaje))

            components.html(
                f"""
                <div style="margin-top:-6px;display:flex;gap:8px;">
                  <button onclick='
                    window.speechSynthesis.cancel();
                    var u = new SpeechSynthesisUtterance({texto_js});
                    u.lang = {codigo_voz_js};
                    window.speechSynthesis.speak(u);
                  ' style="
                    background:#F5F5F4;border:none;border-radius:14px;
                    padding:4px 12px;font-size:0.78rem;cursor:pointer;
                    color:#333;
                  ">{t("boton_escuchar", idioma=idioma_mensaje)}</button>

                  <button id="{id_boton_copiar}" onclick='
                    navigator.clipboard.writeText({texto_js});
                    var b = document.getElementById("{id_boton_copiar}");
                    b.innerText = {texto_copiado_js};
                    setTimeout(function() {{ b.innerText = {texto_copiar_js}; }}, 1500);
                  ' style="
                    background:#F5F5F4;border:none;border-radius:14px;
                    padding:4px 12px;font-size:0.78rem;cursor:pointer;
                    color:#333;
                  ">{t("boton_copiar", idioma=idioma_mensaje)}</button>
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
                st.markdown(t("imagen_no_disponible"))

    # --------------------------------------------
    # MENSAJE DE LA IA (PRESENTACIÓN GENERADA)
    # --------------------------------------------

    elif role == "assistant" and tipo == "pptx":

        with st.chat_message("assistant"):

            if os.path.exists(content):

                titulo_pptx = message.get(
                    "titulo", t("presentacion_sin_titulo")
                )
                n_diapositivas = message.get("n_diapositivas", 0)
                ruta_pdf = message.get("ruta_pdf")
                ruta_html = message.get("ruta_html")

                st.markdown(
                    t(
                        "presentacion_generada",
                        titulo=titulo_pptx,
                        n=n_diapositivas,
                    )
                )

                if ruta_html and os.path.exists(ruta_html):
                    with open(ruta_html, "r", encoding="utf-8") as f:
                        components.html(f.read(), height=600, scrolling=False)

                hay_pdf = bool(ruta_pdf) and os.path.exists(ruta_pdf)

                def _boton_pptx():
                    with open(content, "rb") as f:
                        st.download_button(
                            t("descargar_presentacion"),
                            data=f.read(),
                            file_name=f"{_nombre_archivo_seguro(titulo_pptx)}.pptx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument"
                                ".presentationml.presentation"
                            ),
                            use_container_width=True,
                            key=f"descargar_pptx_{indice_mensaje}",
                        )

                if hay_pdf:
                    st.caption(t("exportar_como"))
                    col_pptx, col_pdf = st.columns(2)
                    with col_pptx:
                        _boton_pptx()
                    with col_pdf:
                        with open(ruta_pdf, "rb") as f:
                            st.download_button(
                                t("descargar_pdf"),
                                data=f.read(),
                                file_name=f"{_nombre_archivo_seguro(titulo_pptx)}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                                key=f"descargar_pdf_{indice_mensaje}",
                            )
                else:
                    _boton_pptx()
            else:
                st.markdown(t("pptx_no_disponible"))


# ============================================================
# 13. INPUT DEL CHAT
# ============================================================

# Imagen, búsqueda web y presentación ya no son interruptores: ORIA
# decide sola si hacen falta según lo que se escriba (ver
# "detectar_intencion" y su uso más abajo). La voz está desactivada
# de momento (transcribir_audio() se queda sin usar, por si se
# retoma más adelante).

entrada = st.chat_input(
    t("chat_placeholder"),
    accept_file=True,
    file_type=["png", "jpg", "jpeg", "pdf"],
)


if entrada:

    user_text = (entrada.text or "").strip()
    archivo_adjunto = entrada.files[0] if entrada.files else None

    if not user_text and not archivo_adjunto:
        st.stop()

    # ========================================================
    # DETECTAR AUTOMÁTICAMENTE SI HACE FALTA IMAGEN, PRESENTACIÓN
    # O BÚSQUEDA WEB (sin que el usuario tenga que pulsar nada)
    # ========================================================

    intencion = {"accion": "ninguna", "parametro": user_text}

    # Si ya hay un archivo adjunto (PDF o imagen para analizar), ese
    # siempre manda: no tiene sentido "adivinar" otra acción distinta.
    if user_text and not archivo_adjunto:
        intencion = detectar_intencion(user_text, idioma=st.session_state.idioma)

    modo_imagen_auto = intencion["accion"] == "imagen"
    modo_presentacion_auto = intencion["accion"] == "presentacion"
    modo_web_auto = intencion["accion"] == "web"

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
            t("archivo_titulo", nombre=archivo_adjunto.name)
            if archivo_adjunto
            else t("conversacion_sin_titulo")
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
        etiqueta_archivo = f"{archivo_adjunto.name}"
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

        if modo_imagen_auto:

            with st.spinner(t("puliendo_descripcion")):
                prompt_mejorado = mejorar_prompt_imagen(intencion["parametro"])

            with st.spinner(t("generando_imagen")):
                imagen_bytes, error = generar_imagen_ia(prompt_mejorado)

            if error:
                st.error(t("error_generar_imagen", error=error))
                respuesta_final = {
                    "role": "assistant",
                    "type": "text",
                    "content": t("no_pude_generar_imagen", error=error),
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
        # MODO: GENERAR PRESENTACIÓN (.pptx)
        # ----------------------------------------------------

        elif modo_presentacion_auto:

            with st.spinner(t("generando_contenido_presentacion")):
                contenido_pptx, error = generar_contenido_presentacion(
                    intencion["parametro"], idioma=st.session_state.idioma
                )

            if error or not contenido_pptx:
                st.error(t("error_generar_presentacion", error=error))
                respuesta_final = {
                    "role": "assistant",
                    "type": "text",
                    "content": t(
                        "no_pude_generar_presentacion", error=error
                    ),
                }
            else:
                with st.spinner(t("creando_pptx")):
                    os.makedirs(CARPETA_PRESENTACIONES, exist_ok=True)
                    identificador = uuid.uuid4()
                    nombre_archivo = os.path.join(
                        CARPETA_PRESENTACIONES, f"{identificador}.pptx"
                    )
                    ruta_pdf = os.path.join(
                        CARPETA_PRESENTACIONES, f"{identificador}.pdf"
                    )
                    ruta_html = os.path.join(
                        CARPETA_PRESENTACIONES, f"{identificador}.html"
                    )
                    try:
                        # Se descargan las imágenes una sola vez y se
                        # reutilizan en los tres formatos (pptx, pdf y
                        # vista previa) para no llamar dos veces a
                        # Pexels/Wikimedia por la misma diapositiva.
                        cache_imagenes = obtener_imagenes_presentacion(contenido_pptx)
                        crear_pptx(contenido_pptx, nombre_archivo, cache_imagenes=cache_imagenes)
                        error_pptx = None
                    except Exception as e:
                        error_pptx = str(e)

                if error_pptx:
                    st.error(t("error_generar_presentacion", error=error_pptx))
                    respuesta_final = {
                        "role": "assistant",
                        "type": "text",
                        "content": t(
                            "no_pude_generar_presentacion", error=error_pptx
                        ),
                    }
                else:
                    with st.spinner(t("creando_exportables")):
                        try:
                            crear_pdf(contenido_pptx, ruta_pdf, cache_imagenes=cache_imagenes)
                        except Exception:
                            ruta_pdf = None
                        try:
                            html_preview = generar_html_presentacion(
                                contenido_pptx, cache_imagenes
                            )
                            with open(ruta_html, "w", encoding="utf-8") as f:
                                f.write(html_preview)
                        except Exception:
                            ruta_html = None

                    titulo_pptx = contenido_pptx.get(
                        "titulo", t("presentacion_sin_titulo")
                    )
                    n_diapositivas = len(
                        contenido_pptx.get("diapositivas", [])
                    )

                    st.markdown(
                        t(
                            "presentacion_generada",
                            titulo=titulo_pptx,
                            n=n_diapositivas,
                        )
                    )

                    if ruta_html and os.path.exists(ruta_html):
                        with open(ruta_html, "r", encoding="utf-8") as f:
                            components.html(f.read(), height=600, scrolling=False)

                    st.caption(t("exportar_como"))
                    col_pptx, col_pdf = st.columns(2)
                    with col_pptx:
                        with open(nombre_archivo, "rb") as f:
                            st.download_button(
                                t("descargar_presentacion"),
                                data=f.read(),
                                file_name=f"{_nombre_archivo_seguro(titulo_pptx)}.pptx",
                                mime=(
                                    "application/vnd.openxmlformats-officedocument"
                                    ".presentationml.presentation"
                                ),
                                use_container_width=True,
                            )
                    if ruta_pdf and os.path.exists(ruta_pdf):
                        with col_pdf:
                            with open(ruta_pdf, "rb") as f:
                                st.download_button(
                                    t("descargar_pdf"),
                                    data=f.read(),
                                    file_name=f"{_nombre_archivo_seguro(titulo_pptx)}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True,
                                )

                    respuesta_final = {
                        "role": "assistant",
                        "type": "pptx",
                        "content": nombre_archivo,
                        "ruta_pdf": ruta_pdf,
                        "ruta_html": ruta_html,
                        "prompt": user_text,
                        "titulo": titulo_pptx,
                        "n_diapositivas": n_diapositivas,
                    }

        # ----------------------------------------------------
        # MODO: PDF ADJUNTO
        # ----------------------------------------------------

        elif archivo_adjunto and archivo_adjunto.type == "application/pdf":

            with st.spinner(t("leyendo_pdf")):
                texto_pdf, error = extraer_texto_pdf(archivo_adjunto)

            if error:
                st.error(error)
                respuesta_final = {
                    "role": "assistant",
                    "type": "text",
                    "content": t("no_pude_leer_pdf", error=error),
                }
            else:
                pregunta = user_text if user_text else t("resumen_pdf_defecto")

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
                        idioma=st.session_state.idioma,
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

            pregunta = user_text if user_text else t("describe_imagen_defecto")

            respuesta_texto = st.write_stream(
                obtener_respuesta_ia_stream(
                    pregunta,
                    historial_mensajes=None,
                    imagen_data_uri=data_uri,
                    memoria_texto=st.session_state.memoria,
                    idioma=st.session_state.idioma,
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

            if modo_web_auto or necesita_busqueda(user_text):

                # Si el mensaje es muy corto ("¿y mañana?"), le
                # sumamos la pregunta anterior para dar contexto.
                consulta = intencion["parametro"] if modo_web_auto else user_text

                if len(user_text) < 25:
                    previos = [
                        m["content"]
                        for m in chat_actual["messages"][:-1]
                        if m.get("role") == "user"
                        and m.get("type", "text") == "text"
                    ]
                    if previos:
                        consulta = f"{previos[-1][:200]} {user_text}"

                with st.spinner(t("buscando_web")):
                    resultados_web, error_web = buscar_en_web(consulta)

                if not resultados_web:
                    web_fallida = True

                    if error_web == "no_configurado":
                        st.caption(t("web_no_configurada"))
                    else:
                        st.caption(t("web_no_disponible"))

            respuesta_texto = st.write_stream(
                obtener_respuesta_ia_stream(
                    user_text,
                    historial_mensajes=chat_actual["messages"],
                    memoria_texto=st.session_state.memoria,
                    contexto_web=formatear_contexto_web(resultados_web)
                    if resultados_web else "",
                    web_fallida=web_fallida,
                    idioma=st.session_state.idioma,
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
