import json
import os
import uuid
from datetime import datetime
import requests
import streamlit as st

# 1. Configuración de la página
st.set_page_config(page_title="ORIA", page_icon="✨", layout="wide")

# 2. Persistencia local en archivo JSON
DB_FILE = "conversaciones.json"


def cargar_chats():
  if os.path.exists(DB_FILE):
    try:
      with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}


def guardar_chats(chats):
  try:
    with open(DB_FILE, "w", encoding="utf-8") as f:
      json.dump(chats, f, ensure_ascii=False, indent=2)
  except Exception as e:
    st.error(f"Error al guardar en disco: {e}")


# 3. Obtención de fecha real del sistema
ahora = datetime.now()
dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
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
fecha_hoy_str = f"{dias[ahora.weekday()]}, {ahora.day} de {meses[ahora.month - 1]} de {ahora.year}"


# 4. Respuesta de IA fluida y sin claves API
def obtener_respuesta_ia_stream(prompt_usuario, historial_mensajes=None):
  try:
    messages = [{
        "role": "system",
        "content": (
            "Eres ORIA, una asistente virtual altamente inteligente,"
            f" atenta y amable. La fecha de hoy es {fecha_hoy_str}. Responde"
            " siempre en español."
        ),
    }]
    if historial_mensajes:
      for msg in historial_mensajes:
        if isinstance(msg, dict) and "role" in msg and "content" in msg:
          role = "user" if msg["role"] == "user" else "assistant"
          messages.append({"role": role, "content": str(msg["content"])})

    payload = {"messages": messages, "model": "openai"}
    response = requests.post(
        "https://text.pollinations.ai/", json=payload, stream=True, timeout=30
    )

    if response.status_code == 200:
      for chunk in response.iter_content(chunk_size=1024, decode_unicode=True):
        if chunk:
          yield chunk
    else:
      yield f"⚠️ Error en el servicio ({response.status_code}): {response.text}"
  except Exception as e:
    yield f"⚠️ No se pudo procesar la respuesta: {str(e)}"


# 5. Estilos CSS Personalizados
css_code = """
<style>
[data-testid="stChatMessageAvatarUser"], [data-testid="stChatMessageAvatarAssistant"] {
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
    box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    font-size: 1rem;
}
.stSidebar .stButton > button {
    border-radius: 8px !important;
}
</style>
"""
st.markdown(css_code, unsafe_allow_html=True)

# 6. Cargar datos en sesión
if "chats" not in st.session_state:
  st.session_state.chats = cargar_chats()

if "current_chat_id" not in st.session_state:
  st.session_state.current_chat_id = None

# 7. Menú Lateral (Barra de Conversaciones)
with st.sidebar:
  st.title("Conversaciones")
  if st.button("➕ Nueva conversación", use_container_width=True, type="primary"):
    st.session_state.current_chat_id = None
    st.rerun()

  st.markdown("---")

  chats_a_borrar = []
  for cid, chat_info in reversed(list(st.session_state.chats.items())):
    col_btn, col_del = st.columns([0.82, 0.18])
    btn_type = (
        "secondary" if cid != st.session_state.current_chat_id else "primary"
    )

    with col_btn:
      if st.button(
          f"💬 {chat_info['title']}",
          key=f"chat_{cid}",
          use_container_width=True,
          type=btn_type,
      ):
        st.session_state.current_chat_id = cid
        st.rerun()
    with col_del:
      if st.button("🗑️", key=f"del_{cid}"):
        chats_a_borrar.append(cid)

  if chats_a_borrar:
    for cid in chats_a_borrar:
      del st.session_state.chats[cid]
      if st.session_state.current_chat_id == cid:
        st.session_state.current_chat_id = None
    guardar_chats(st.session_state.chats)
    st.rerun()

# 8. Obtener mensajes del chat activo
if (
    st.session_state.current_chat_id
    and st.session_state.current_chat_id in st.session_state.chats
):
  mensajes_actuales = st.session_state.chats[st.session_state.current_chat_id][
      "messages"
  ]
else:
  mensajes_actuales = []

# 9. Pantalla inicial centrada si no hay mensajes
if len(mensajes_actuales) == 0:
  st.markdown("<br><br><br>", unsafe_allow_html=True)
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    st.markdown(
        "<h1 style='text-align: center; font-size: 3.5rem; font-weight:"
        " bold;'>ORIA</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h3 style='text-align: center; color: #666;'>¿En qué te puedo ayudar"
        " hoy?</h3>",
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

# 10. Renderizado de Mensajes Anteriores
for message in mensajes_actuales:
  if message["role"] == "user":
    content_html = (
        '<div class="user-bubble-container"><div'
        f' class="user-bubble">{message["content"]}</div></div>'
    )
    st.markdown(content_html, unsafe_allow_html=True)
  else:
    with st.chat_message("assistant"):
      st.markdown(message["content"])

# 11. Entrada de chat principal
if prompt := st.chat_input("Preguntar a ORIA..."):
  user_text = prompt

  if (
      not st.session_state.current_chat_id
      or st.session_state.current_chat_id not in st.session_state.chats
  ):
    nuevo_id = str(uuid.uuid4())
    titulo = prompt.strip()
    if len(titulo) > 26:
      titulo = titulo[:26] + "..."
    titulo = titulo.capitalize()

    st.session_state.chats[nuevo_id] = {"title": titulo, "messages": []}
    st.session_state.current_chat_id = nuevo_id

  st.session_state.chats[st.session_state.current_chat_id]["messages"].append(
      {"role": "user", "content": user_text}
  )
  guardar_chats(st.session_state.chats)

  content_html = (
      '<div class="user-bubble-container"><div'
      f' class="user-bubble">{user_text}</div></div>'
  )
  st.markdown(content_html, unsafe_allow_html=True)

  with st.chat_message("assistant"):
    respuesta_texto = st.write_stream(
        obtener_respuesta_ia_stream(
            user_text,
            st.session_state.chats[st.session_state.current_chat_id][
                "messages"
            ],
        )
    )

  st.session_state.chats[st.session_state.current_chat_id]["messages"].append(
      {"role": "assistant", "content": respuesta_texto}
  )
  guardar_chats(st.session_state.chats)
  st.rerun()
