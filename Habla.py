import streamlit as st
import sqlite3
import os
import re
import time
import pandas as pd
import streamlit.components.v1 as components
import json
import threading
from streamlit_gsheets import GSheetsConnection
from google import genai
from google.genai import types
from pydantic import BaseModel

# ==============================================================================
# 1. CONFIGURACIÓN & RUTAS DE ALMACENAMIENTO
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "spanish_adaptive_coach.db")
st.set_page_config(page_title="🗣️ Tutor de Español", layout="wide", page_icon="☕")

# ==============================================================================
# 2. BASE DE DATOS LOCAL (DIAGNÓSTICO Y REGISTRO PEDAGÓGICO)
# ==============================================================================
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn

def init_db():
    conn = get_db_connection()
    with conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS diagnostics 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                       student_name TEXT, 
                       student_input TEXT, 
                       ai_feedback TEXT, 
                       grammar_topic TEXT, 
                       student_strengths TEXT, 
                       student_errors TEXT, 
                       mastery_status TEXT, 
                       used_pdf TEXT, 
                       timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    conn.close()

init_db()

def log_diagnostic_result(name, inp, feedback, topic, strengths, errors, mastery_status, used_pdf):
    try:
        conn = get_db_connection()
        with conn:
            cursor = conn.cursor()
            cursor.execute("""INSERT INTO diagnostics 
                          (student_name, student_input, ai_feedback, grammar_topic, student_strengths, student_errors, mastery_status, used_pdf) 
                          VALUES (?,?,?,?,?,?,?,?)""",
                          (name, inp, feedback, topic, strengths, errors, mastery_status, used_pdf))
            inserted_id = cursor.lastrowid
        conn.close()
        return inserted_id
    except sqlite3.OperationalError:
        return None

def get_student_unmastered_weaknesses(name, limit=3):
    try:
        conn = get_db_connection()
        query = """
            SELECT grammar_topic, student_errors 
            FROM diagnostics 
            WHERE student_name = ? AND grammar_topic != 'Analizando...' AND mastery_status != 'mastered'
            ORDER BY id DESC LIMIT ?
        """
        df = pd.read_sql_query(query, conn, params=(name, limit))
        conn.close()
        if df.empty:
            return "No hay dificultades persistentes registradas. Mantener la conversación fluida y natural."
        
        weakness_lines = []
        for _, row in df.iterrows():
            weakness_lines.append(f"- Objetivo implícito: {row['grammar_topic']} | Desliz para reformular: {row['student_errors']}")
        return "\n".join(weakness_lines)
    except Exception:
        return "Perfil del estudiante no disponible en este momento."

class PedagogicalDiagnosis(BaseModel):
    grammar_topic: str
    student_strengths: str
    student_errors: str
    mastery_status: str

# ==============================================================================
# 3. GESTIÓN DE ROSTER Y LECTURA SEGURA DE GOOGLE SHEETS
# ==============================================================================
def get_authorized_students() -> dict:
    try:
        gsheet_conn = st.connection("gsheets", type=GSheetsConnection)
        df = gsheet_conn.read(ttl=0)
    except Exception:
        df = None

    if df is None or df.empty:
        try:
            raw_url = None
            try:
                raw_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
            except Exception:
                raw_url = None
                
            if raw_url and "/d/" in raw_url:
                sheet_id = raw_url.split("/d/")[1].split("/")[0]
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                df = pd.read_csv(csv_url)
            else:
                return {"Edy": "peru2026"}
        except Exception:
            return {"Edy": "peru2026"}

    try:
        df.columns = [str(c).strip().lower() for c in df.columns]
        name_col = next((c for c in df.columns if any(k in c for k in ["name", "nombre", "student"])), df.columns[0])
        pin_col = next((c for c in df.columns if any(k in c for k in ["pin", "pass", "clave", "code"])), df.columns[1])

        df = df.dropna(subset=[name_col, pin_col])
        df[name_col] = df[name_col].astype(str).str.strip()
        df[pin_col] = df[pin_col].astype(str).str.strip()
        return dict(zip(df[name_col], df[pin_col]))
    except Exception:
        return {"Edy": "peru2026"}

# ==============================================================================
# 4. MOTOR DE AUDIO NATIVO EN ESPAÑOL (WEB SPEECH API)
# ==============================================================================
def clean_for_speech(text):
    clean = re.sub(r'[\_\-\(\)\[\]\{\}\*\#\«\»\<\>\/\\]', ' ', text)
    clean = clean.replace('❌', '').replace('✅', '').replace('🎯', '').replace('💡', '').replace('🔍', '').replace('💬', '').replace('📖', '')
    clean = re.sub(r'^\s*\d+\.\s*', '', clean, flags=re.MULTILINE)
    clean = re.sub(r'(?:https?|ftp)://\S+', '', clean)
    return re.sub(r'\s+', ' ', clean).strip()

def render_communicative_audio(full_response_text, auto_play=True, key_suffix=""):
    paragraphs = [p.strip() for p in full_response_text.split("\n") if p.strip()]
    cleaned_chunks = [clean_for_speech(p) for p in paragraphs if clean_for_speech(p)]
    payload = json.dumps(cleaned_chunks)
    auto_flag = "true" if auto_play else "false"

    html_code = f"""
    <div style="margin-top: 8px; margin-bottom: 6px;">
        <button id="btn_{key_suffix}" style="
            background-color: #0f172a;
            color: #38bdf8;
            border: 1px solid #0284c7;
            border-radius: 6px;
            padding: 6px 14px;
            font-size: 13px;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-weight: 500;
        ">
            🔊 Escuchar audio del tutor
        </button>
    </div>
    <script>
    (function() {{
        var chunks = {payload};
        var shouldAutoPlay = {auto_flag};
        var btn = document.getElementById("btn_{key_suffix}");

        function playFullSequence() {{
            var synth = (window.top && window.top.speechSynthesis) || window.speechSynthesis;
            if (!synth) return;
            synth.cancel();

            var voices = synth.getVoices();
            // Buscar una voz nativa en español
            var esVoice = voices.find(function(v) {{ 
                return v.lang && (v.lang.toLowerCase().startsWith("es")); 
            }}) || voices[0];

            if (!chunks || chunks.length === 0) return;

            var utterances = [];
            for (var i = 0; i < chunks.length; i++) {{
                var text = chunks[i];
                var utt = new SpeechSynthesisUtterance(text);
                utt.voice = esVoice;
                utt.lang = "es-ES";
                utt.rate = 0.95;
                utterances.push(utt);
            }}

            for (var j = 0; j < utterances.length - 1; j++) {{
                (function(index) {{
                    utterances[index].onend = function() {{ synth.speak(utterances[index + 1]); }};
                }})(j);
            }}

            synth.speak(utterances[0]);
        }}

        if (btn) {{
            btn.onclick = function() {{ playFullSequence(); }};
        }}

        if (shouldAutoPlay) {{
            var synth = (window.top && window.top.speechSynthesis) || window.speechSynthesis;
            if (synth && synth.getVoices().length > 0) {{
                playFullSequence();
            }} else if (synth) {{
                synth.onvoiceschanged = function() {{
                    playFullSequence();
                    synth.onvoiceschanged = null;
                }};
            }}
        }}
    }})();
    </script>
    """
    components.html(html_code, height=45)

# ==============================================================================
# 5. CONTROLADOR RESILIENTE DE LLAMADAS GEMINI
# ==============================================================================
def stream_with_retry(client_obj, model_name, contents_list, cfg, retries=3, delay=2.5):
    for attempt in range(retries):
        try:
            return client_obj.models.generate_content_stream(
                model=model_name,
                contents=contents_list,
                config=cfg
            )
        except Exception as e:
            err_msg = str(e)
            if ("503" in err_msg or "UNAVAILABLE" in err_msg) and attempt < retries - 1:
                time.sleep(delay * (attempt + 1))
                continue
            raise e

def generate_with_retry(client_obj, model_name, contents_val, cfg, retries=3, delay=2.5):
    for attempt in range(retries):
        try:
            return client_obj.models.generate_content(
                model=model_name,
                contents=contents_val,
                config=cfg
            )
        except Exception as e:
            err_msg = str(e)
            if ("503" in err_msg or "UNAVAILABLE" in err_msg) and attempt < retries - 1:
                time.sleep(delay * (attempt + 1))
                continue
            raise e

# ==============================================================================
# 6. BARRA LATERAL (AUTENTICACIÓN, CONFIGURACIÓN Y ANALÍTICA)
# ==============================================================================
with st.sidebar:
    st.title("🗣️ Tutor de Español")
    
    gemini_api_key = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            gemini_api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        gemini_api_key = None

    if not gemini_api_key:
        gemini_api_key = os.environ.get("GEMINI_API_KEY")

    if not gemini_api_key:
        api_key_input = st.text_input("🔑 Clave API de Gemini:", type="password")
        gemini_api_key = api_key_input

    st.divider()
    st.subheader("🔐 Acceso de Estudiante")
    students_dict = get_authorized_students()
    roster_names = ["-- Selecciona tu nombre --"] + sorted(list(students_dict.keys()))
    
    selected_student = st.selectbox("Nombre:", roster_names)
    entered_pin = st.text_input("PIN / Código:", type="password")
    
    authenticated = False
    if selected_student != "-- Selecciona tu nombre --" and entered_pin:
        if str(students_dict.get(selected_student, "")).strip() == entered_pin.strip():
            authenticated = True
            st.success(f"¡Bienvenido/a, {selected_student}!")
        else:
            st.error("PIN incorrecto.")

    st.divider()
    st.subheader("📚 Material de Apoyo (Opcional)")
    uploaded_file = st.file_uploader("Subir escenarios / lecturas (PDF, TXT, MD)", type=["pdf", "txt", "md"])
    pdf_bytes = None
    text_content = ""
    if uploaded_file:
        if uploaded_file.name.lower().endswith(".pdf"):
            pdf_bytes = uploaded_file.getvalue()
        else:
            text_content = uploaded_file.getvalue().decode("utf-8", errors="ignore")
        st.sidebar.success(f"✅ Material cargado: {uploaded_file.name}")

    st.divider()
    if st.button("🔄 Reiniciar conversación"):
        st.session_state.messages = []
        st.session_state.last_audio_bytes = None
        st.rerun()

    instructor_key = st.text_input("Acceso de Instructor:", type="password")
    if instructor_key == "peru2026":
        st.subheader("👨‍🏫 Analítica Docente")
        with st.expander("📊 Registro de Dificultades Orales", expanded=True):
            conn = get_db_connection()
            df = pd.read_sql_query("""
                SELECT student_name, grammar_topic, student_errors, mastery_status, student_input, timestamp 
                FROM diagnostics 
                ORDER BY timestamp DESC
            """, conn)
            conn.close()
            st.dataframe(df, use_container_width=True)

# ==============================================================================
# 7. SESIÓN PRINCIPAL & SALUDO INICIAL
# ==============================================================================
st.title("🗣️ Tutor de Español Conversacional")

if not gemini_api_key:
    st.warning("👈 Por favor ingresa tu GEMINI_API_KEY en la barra lateral para comenzar.")
    st.stop()

if not authenticated:
    st.warning("👈 Por favor selecciona tu nombre e ingresa tu PIN en la barra lateral.")
    st.stop()

@st.cache_resource
def get_gemini_client(key):
    return genai.Client(api_key=key)

client = get_gemini_client(gemini_api_key)
student_first_name = selected_student.split()[0]

INITIAL_GREETING = (
    f"¡Hola {student_first_name}! ¡Qué gusto saludarte! ☕\n\n"
    "Vamos a practicar conversación real en español en un ambiente tranquilo y divertido. "
    "**Puedes hablar usando el micrófono abajo o escribir tu respuesta en el chat.** "
    "Si no sabes cómo decir algo, puedes preguntarlo en inglés y yo te muestro cómo decirlo en español.\n\n"
    "Para empezar, **¿de qué te gustaría hablar hoy?**\n\n"
    "1. **Tu rutina diaria** (a qué hora te levantas, tus clases, qué haces cada día)\n"
    "2. **Un viaje o vacaciones memorables** (un lugar divertido que visitaste y qué hiciste)\n"
    "3. **Tus planes para el futuro** (lugares que quieres conocer o tus metas)\n\n"
    "👉 **Tu turno:** Elige el **1, 2 o 3** (o dime en tus propias palabras) para empezar."
)

if 'messages' not in st.session_state or st.session_state.get("active_user") != selected_student or not st.session_state.get("messages"):
    st.session_state.active_user = selected_student
    st.session_state.messages = [{
        "role": "assistant",
        "content": INITIAL_GREETING,
        "auto_play": True
    }]
    st.session_state.last_audio_bytes = None

for idx, m in enumerate(st.session_state.messages):
    with st.chat_message(m["role"]):
        st.markdown(m["content"])
        if m.get("audio_data"):
            st.audio(m["audio_data"], format="audio/wav")
        if m["role"] == "assistant" and m.get("content"):
            render_communicative_audio(
                m["content"], 
                auto_play=m.get("auto_play", False), 
                key_suffix=f"hist_{idx}"
            )
            m["auto_play"] = False

# ==============================================================================
# 8. ENTRADA DE USUARIO (MICRÓFONO O TEXTO)
# ==============================================================================
st.write("---")
col_audio, col_info = st.columns([1, 2])
with col_audio:
    recorded_audio = st.audio_input("🎙️ Habla aquí (Grabar y enviar):")
with col_info:
    st.caption("💡 **Consejo:** Presiona el botón rojo para hablar en español (o inglés si tienes dudas). Al terminar, presiona detener para enviar.")

text_input = st.chat_input("O escribe tu mensaje aquí si prefieres teclear...")

incoming_text = None
incoming_audio_bytes = None

if recorded_audio is not None:
    audio_bytes = recorded_audio.getvalue()
    if audio_bytes != st.session_state.get("last_audio_bytes"):
        st.session_state.last_audio_bytes = audio_bytes
        incoming_audio_bytes = audio_bytes
        incoming_text = "🎙️ [Mensaje de voz]"

if text_input:
    incoming_text = text_input

# ==============================================================================
# 9. MOTOR PEDAGÓGICO: COMUNICATIVO & FOCUS ON FORM (ESPAÑOL)
# ==============================================================================
if incoming_text:
    user_entry = {"role": "user", "content": incoming_text}
    if incoming_audio_bytes:
        user_entry["audio_data"] = incoming_audio_bytes
    st.session_state.messages.append(user_entry)

    with st.chat_message("user"):
        st.markdown(incoming_text)
        if incoming_audio_bytes:
            st.audio(incoming_audio_bytes, format="audio/wav")

    learner_weaknesses = get_student_unmastered_weaknesses(selected_student)

    system_instruction = (
        f"Eres un compañero de conversación amigable, cálido y paciente que ayuda a un estudiante adulto ({selected_student}, llamado '{student_first_name}') a desarrollar fluidez oral en ESPAÑOL.\n\n"
        f"OBJETIVOS DE APRENDIZAJE PREVIOS (RECICLAR DE FORMA IMPLÍCITA):\n"
        f"{learner_weaknesses}\n\n"
        "REGLAS PEDAGÓGICAS Y COMUNICATIVAS ESTRICTAS (NIVEL MCER A2–B1):\n"
        "1. EL TUTOR RESPONDE 100% EN ESPAÑOL:\n"
        "   - Todas tus respuestas deben ser en un español claro, natural e internacional.\n"
        "   - Evita modismos locales excesivamente oscuros o lunfardos difíciles.\n\n"
        "2. APOYO INMEDIATO SI EL ALUMNO PREGUNTA O HABLA EN INGLÉS:\n"
        "   - El estudiante tiene total libertad para escribir o hablar en inglés si no recuerda una palabra o si pregunta 'How do I say...?'.\n"
        "   - Si el estudiante usa inglés o tiene dudas, muéstrale de inmediato la forma natural en español dentro del contexto conversacional.\n"
        "     Ejemplo alumno: 'I don't know how to say I woke up late.'\n"
        "     Ejemplo tutor: 'En español puedes decir: \"¡Hoy me desperté tarde!\" ¿A qué hora sonó tu alarma?'\n\n"
        "3. ENFOQUE EN LA FORMA (FOCUS ON FORM - RECASTING NATURAL):\n"
        "   - NUNCA des explicaciones gramaticales teóricas, reglas explícitas ni lecciones como 'Tip gramatical:'.\n"
        "   - Si el estudiante comete un error (por ejemplo: 'Ayer yo ir al parque'), reformula la idea de forma natural en tu respuesta sin corregir directamente:\n"
        "     '¡Ah, fuiste al parque ayer! ¡Qué bien! ¿Con quién fuiste?'\n\n"
        "4. INTERVENCIONES BREVES Y PREGUNTAS ABIERTAS:\n"
        "   - Mantén tus respuestas breves: estrictamente 2 o 3 oraciones sencillas.\n"
        "   - Termina siempre con UNA pregunta sencilla y abierta para que el estudiante siga practicando de forma oral."
    )

    with st.chat_message("assistant"):
        try:
            contents = []
            for msg in st.session_state.messages[-4:]:
                content_text = msg.get("content", "")
                if content_text and not content_text.startswith("🚨"):
                    role = "user" if msg["role"] == "user" else "model"
                    parts = []
                    if msg.get("audio_data"):
                        parts.append(types.Part.from_bytes(data=msg["audio_data"], mime_type="audio/wav"))
                        parts.append(types.Part.from_text(text="[Escucha el mensaje de voz del estudiante. Si habló en inglés o tuvo dudas, dale la frase equivalente en español de inmediato, reformula cualquier error de forma natural en español y continúa la conversación.]"))
                    else:
                        parts.append(types.Part.from_text(text=content_text))
                    contents.append(types.Content(role=role, parts=parts))

            if pdf_bytes:
                contents[-1].parts.insert(0, types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"))
            elif text_content:
                contents[-1].parts.insert(0, types.Part.from_text(text=f"[MATERIAL CONTEXTUAL]:\n{text_content[:8000]}\n"))

            fast_config = types.GenerateContentConfig(
                system_instruction=system_instruction,
                max_output_tokens=400,
                temperature=0.6
            )

            def stream_response():
                response = stream_with_retry(
                    client_obj=client,
                    model_name='gemini-3.8-flash',
                    contents_list=contents,
                    cfg=fast_config,
                    retries=3,
                    delay=2.5
                )
                for chunk in response:
                    if chunk.text:
                        yield chunk.text.replace("<", "«").replace(">", "»")

            student_response = st.write_stream(stream_response)

            record_id = log_diagnostic_result(
                selected_student, incoming_text, student_response, 
                "Analizando...", "Analizando...", "Analizando...", "emerging",
                "YES" if uploaded_file else "NO"
            )

            # Diagnóstico pedagógico invisible en segundo plano (para analítica docente)
            def run_background_analysis(rec_id, query_content, resp, student_name):
                if not rec_id:
                    return
                try:
                    analysis_prompt = (
                        f"Analiza esta interacción oral de español como lengua extranjera exclusivamente para analítica del profesor.\n"
                        f"Entrada del estudiante: '{query_content}'\n"
                        f"Respuesta del tutor: '{resp}'\n\n"
                        "Genera un JSON con los campos:\n"
                        "- grammar_topic: Estructura gramatical meta (ej. Pretérito Indefinido, Gustar, Ser/Estar).\n"
                        "- student_strengths: Fortalezas orales o léxico utilizado.\n"
                        "- student_errors: Desliz o error de transferencia (o 'Ninguno' si fue correcto).\n"
                        "- mastery_status: 'needs_practice', 'emerging', o 'mastered'."
                    )
                    analysis_res = generate_with_retry(
                        client_obj=client,
                        model_name='gemini-3.8-flash',
                        contents_val=analysis_prompt,
                        cfg=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=PedagogicalDiagnosis,
                            max_output_tokens=250
                        ),
                        retries=3,
                        delay=2.5
                    )
                    diag_data = json.loads(analysis_res.text)
                    conn = get_db_connection()
                    with conn:
                        conn.execute("""
                            UPDATE diagnostics 
                            SET grammar_topic = ?, student_strengths = ?, student_errors = ?, mastery_status = ?
                            WHERE id = ?
                        """, (
                            diag_data.get("grammar_topic", "Fluidez Oral"), 
                            diag_data.get("student_strengths", "Interacción Espontánea"), 
                            diag_data.get("student_errors", "Desliz leve"), 
                            diag_data.get("mastery_status", "needs_practice"),
                            rec_id
                        ))
                    conn.close()
                except Exception:
                    pass

            threading.Thread(
                target=run_background_analysis,
                args=(record_id, incoming_text, student_response, selected_student),
                daemon=True
            ).start()

            new_idx = len(st.session_state.messages)
            render_communicative_audio(student_response, auto_play=True, key_suffix=f"dyn_{new_idx}")

            st.session_state.messages.append({
                "role": "assistant",
                "content": student_response,
                "auto_play": False
            })

        except Exception as e:
            st.error(f"🚨 Error de conexión: {e}")
