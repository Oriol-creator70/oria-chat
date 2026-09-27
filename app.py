import json
import os
import uuid
import html
import base64
from datetime import datetime
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
import streamlit as st
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
# 2. BASE DE DATOS LOCAL DE CONVERSACIONES Y MEMORIA
# ============================================================

DB_FILE = "conversaciones.json"
MEMORIA_FILE = "memoria.txt"
CARPETA_IMAGENES = "imagenes_generadas"

# Cuántos mensajes recientes mandamos como contexto a Groq.
# Evita que las conversaciones muy largas se coman el contexto
# (y la cuota gratuita) del modelo.
MAX_MENSAJES_HISTORIAL = 20


def cargar_chats():
    """Carga las conversaciones guardadas."""
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                datos = json.load(f)

                if isinstance(datos, dict):
                    return datos

        except Exception:
            return {}

    return {}


def guardar_chats(chats):
    """Guarda las conversaciones en JSON."""
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(
                chats,
                f,
                ensure_ascii=False,
                indent=2,
            )

    except Exception as e:
        st.error(f"Error al guardar las conversaciones: {e}")


def cargar_memoria():
    """Carga los datos que ORIA debe recordar sobre el usuario."""
    if os.path.exists(MEMORIA_FILE):
        try:
            with open(MEMORIA_FILE, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            return ""
    return ""


def guardar_memoria(texto):
    """Guarda los datos que ORIA debe recordar."""
    try:
        with open(MEMORIA_FILE, "w", encoding="utf-8") as f:
            f.write(texto)
    except Exception as e:
        st.error(f"Error al guardar la memoria: {e}")


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

# Modelo con visión (el único multimodal del catálogo self-serve
# de Groq). Se usa solo cuando el usuario adjunta una imagen.
MODELO_GROQ_VISION = "meta-llama/llama-4-scout-17b-16e-instruct"

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


def generar_imagen_ia(prompt_imagen):
    """Genera una imagen a partir de un texto usando Pollinations.ai
    (servicio gratuito, no requiere API key). Devuelve (bytes, error)."""

    try:
        prompt_codificado = quote(prompt_imagen)
        semilla = uuid.uuid4().int % 1_000_000

        url = (
            f"https://image.pollinations.ai/prompt/{prompt_codificado}"
            f"?width=1024&height=1024&nologo=true&seed={semilla}"
        )

        respuesta = requests.get(url, timeout=90)

        if (
            respuesta.status_code == 200
            and respuesta.headers.get("content-type", "").startswith("image")
        ):
            return respuesta.content, None

        return None, f"El generador de imágenes devolvió el error {respuesta.status_code}."

    except requests.exceptions.Timeout:
        return None, "El generador de imágenes ha tardado demasiado. Inténtalo de nuevo."

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
    # Modo texto normal: con historial (recortado)
    # --------------------------------------------------------

    else:

        modelo_usar = MODELO_GROQ

        if historial_mensajes:

            recientes = historial_mensajes[-MAX_MENSAJES_HISTORIAL:]

            for msg in recientes:

                if not isinstance(msg, dict):
                    continue

                if "role" not in msg or "content" not in msg:
                    continue

                # Los mensajes de tipo "imagen" (generadas por ORIA)
                # no se pueden mandar como texto; los resumimos.
                if msg.get("type") == "image":
                    if msg["role"] == "assistant":
                        messages.append(
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

                messages.append(
                    {
                        "role": role,
                        "content": str(msg["content"]),
                    }
                )

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
    background-color: #f4f4f5;
    color: #0d0d0d;
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

.stSidebar .stButton > button {
    border-radius: 8px !important;
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

if "chats" not in st.session_state:
    st.session_state.chats = cargar_chats()

if "current_chat_id" not in st.session_state:
    st.session_state.current_chat_id = None

if "memoria" not in st.session_state:
    st.session_state.memoria = cargar_memoria()


# ============================================================
# 9. BARRA LATERAL
# ============================================================

with st.sidebar:

    st.title("Conversaciones")

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

        guardar_chats(
            st.session_state.chats
        )

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
            guardar_memoria(nueva_memoria)
            st.session_state.memoria = nueva_memoria
            st.success("Memoria guardada.")


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
            <h1 style="
                text-align: center;
                font-size: 3.5rem;
                font-weight: bold;
            ">
                ORIA
            </h1>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <h3 style="
                text-align: center;
                color: #666;
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

for message in mensajes_actuales:

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

modo_imagen = st.toggle(
    "🎨 Generar una imagen en vez de responder",
    key="modo_imagen",
    help=(
        "Actívalo y escribe lo que quieras que ORIA dibuje. "
        "Desactívalo para volver al chat normal."
    ),
)

entrada = st.chat_input(
    "Pregunta a ORIA, o adjunta una imagen/PDF...",
    accept_file=True,
    file_type=["png", "jpg", "jpeg", "pdf"],
)


if entrada:

    user_text = (entrada.text or "").strip()
    archivo_adjunto = entrada.files[0] if entrada.files else None

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

    guardar_chats(
        st.session_state.chats
    )

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

            with st.spinner("Generando imagen..."):
                imagen_bytes, error = generar_imagen_ia(user_text)

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

    guardar_chats(
        st.session_state.chats
    )

    # ========================================================
    # RECARGAR
    # ========================================================

    st.rerun()
