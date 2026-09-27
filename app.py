import json
import os
import uuid
import html
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
import streamlit as st


# ============================================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ============================================================

st.set_page_config(
    page_title="ORIA",
    page_icon="✨",
    layout="wide",
)


# ============================================================
# 2. BASE DE DATOS LOCAL DE CONVERSACIONES
# ============================================================

DB_FILE = "conversaciones.json"


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

# Modelo actual de Groq.
MODELO_GROQ = "openai/gpt-oss-120b"

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
# 5. FUNCIÓN PRINCIPAL DE LA IA
# ============================================================

def obtener_respuesta_ia_stream(
    prompt_usuario,
    historial_mensajes=None,
):
    """
    Envía la conversación a Groq y devuelve la respuesta
    progresivamente mediante streaming.
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
    # Mensajes
    # --------------------------------------------------------

    messages = [
        {
            "role": "system",
            "content": (
                "Eres ORIA, una asistente virtual inteligente, "
                "rápida, natural, amable y precisa. "
                "Puedes ayudar con tecnología, programación, "
                "deportes, fútbol, estudios, informática, "
                "cultura, entretenimiento y muchos otros temas. "
                f"La fecha actual es {fecha_hoy_str}. "
                "Responde siempre en español, salvo que el usuario "
                "pida expresamente otro idioma. "
                "Explica las cosas de forma clara y útil."
            ),
        }
    ]

    # --------------------------------------------------------
    # Añadir historial
    # --------------------------------------------------------

    if historial_mensajes:
        for msg in historial_mensajes:

            if not isinstance(msg, dict):
                continue

            if "role" not in msg or "content" not in msg:
                continue

            # Solo permitimos user / assistant.
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
        "model": MODELO_GROQ,
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
                f"Modelo utilizado: `{MODELO_GROQ}`"
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
# 6. CSS
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
# 7. SESSION STATE
# ============================================================

if "chats" not in st.session_state:
    st.session_state.chats = cargar_chats()

if "current_chat_id" not in st.session_state:
    st.session_state.current_chat_id = None


# ============================================================
# 8. BARRA LATERAL
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


# ============================================================
# 9. OBTENER CHAT ACTUAL
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
# 10. PANTALLA INICIAL
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
# 11. MOSTRAR MENSAJES ANTERIORES
# ============================================================

for message in mensajes_actuales:

    role = message.get("role")
    content = str(
        message.get("content", "")
    )

    # --------------------------------------------
    # MENSAJE DEL USUARIO
    # --------------------------------------------

    if role == "user":

        # Escapamos HTML para evitar que el contenido
        # introduzca etiquetas HTML directamente.
        contenido_seguro = html.escape(
            content
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
    # MENSAJE DE LA IA
    # --------------------------------------------

    elif role == "assistant":

        with st.chat_message("assistant"):

            st.markdown(
                content
            )


# ============================================================
# 12. INPUT DEL CHAT
# ============================================================

prompt = st.chat_input(
    "Preguntar a ORIA..."
)


if prompt:

    user_text = prompt.strip()

    if not user_text:
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

        titulo = user_text.strip()

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

    chat_actual["messages"].append(
        {
            "role": "user",
            "content": user_text,
        }
    )

    guardar_chats(
        st.session_state.chats
    )

    # ========================================================
    # MOSTRAR MENSAJE DEL USUARIO
    # ========================================================

    user_text_html = html.escape(
        user_text
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

    with st.chat_message("assistant"):

        respuesta_texto = st.write_stream(
            obtener_respuesta_ia_stream(
                user_text,
                chat_actual["messages"],
            )
        )

    # ========================================================
    # GUARDAR RESPUESTA DE LA IA
    # ========================================================

    chat_actual["messages"].append(
        {
            "role": "assistant",
            "content": respuesta_texto,
        }
    )

    guardar_chats(
        st.session_state.chats
    )

    # ========================================================
    # RECARGAR
    # ========================================================

    st.rerun()
