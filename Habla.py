import copy
import hmac
import json
import os
import random
import re
import sqlite3
import time
import urllib.parse
from io import BytesIO

import gspread
import openpyxl
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from gtts import gTTS
from streamlit_gsheets import GSheetsConnection

# ==============================================================================
# 1. CONFIGURACIÓN GENERAL (SPAN 107 - SPEAKING-ONLY EDITION - Habla.py)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "edgarvich_research_span107_speaking.db")

EXCEL_CANDIDATES = [
    os.path.join(BASE_DIR, "Hablaspan107_database.xlsx"),
    os.path.join(BASE_DIR, "span107.xlsx"),
    os.path.join(BASE_DIR, "estudiantes.xlsx"),
]
EXCEL_PATH = next((p for p in EXCEL_CANDIDATES if os.path.exists(p)), EXCEL_CANDIDATES[0])

st.set_page_config(
    page_title="🎙️ Coach Edgarvich - SPAN 107 Speaking Lab (A1)",
    layout="wide",
    page_icon="🇪🇸",
)

DEFAULT_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1u-bZBP3iJ2TAfvpjU509iZvLL6xqeJytWzDnlIwKuyY/edit"
ROSTER_TAB = "rosterspan107"
LEDGER_TAB = "speech_logs"
LEDGER_COLUMNS = [
    "timestamp",
    "student",
    "level",
    "topic",
    "task",
    "transcription",
    "communicative_success",
    "interlanguage_gap",
    "negotiation_move",
    "recast_target",
    "feedback",
]
MAX_ATTEMPTS = 2
LEDGER_BATCH_SIZE = 6
LEDGER_MAX_AGE_SECONDS = 120
SCENARIO_VARIANTS = 4
SCENARIO_CACHE_TTL = 12 * 3600
DEFAULT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-3.6-flash"
TEACHER_PASSWORD = "peru2026"

SPEAKING_TUTOR_RULES = """
You are Coach Edgarvich, an encouraging conversational Spanish speaking coach for English-speaking college students in Beginner Spanish (SPAN 107 - CEFR Level A1).
This is a SPOKEN interaction (Listening + Speaking only). The student speaks into a microphone and listens to your voice in Spanish.

LEVEL A1 SPOKEN SPANISH:
- Use clear, natural, everyday spoken Spanish: present tense (ser, estar, tener, ir, querer, necesitar, gustar, poder, hacer, comer, vivir), common campus/daily life vocabulary, and short rhythmic sentences (5-8 words) easy to process by ear.

SLA PEDAGOGICAL PRINCIPLES - FOLLOW ALL 5 RULES IN EVERY SPOKEN REPLY:
1. MANDATORY ORAL RECAST (Focus on Form / Lyster & Ranta): In your FIRST sentence, naturally rephrase what the student just said aloud in correct everyday spoken Spanish (fixing verb conjugation, ser/estar, gender/number agreement, or articles) without judging or grading. If it was accurate, echo it warmly in Spanish.
2. ZERO CORRECTIVE METALANGUAGE (Krashen Low Affective Filter): Never say "Incorrecto", "Mal", "Error de gramática", "Debes decir", and never give grammar rules. Keep spoken anxiety low.
3. NO DOUBLE QUESTIONS IN FEEDBACK: When validating an answer to a numbered speaking turn (Turns 1-5), ONLY give the warm oral recast/validation and XP confirmation in Spanish. NEVER add a new unnumbered question at the end of your feedback bubble, because the UI automatically presents the next speaking prompt below. (Only ask an open follow-up question in Spanish during the final 3-turn Personal Speaking Conversation).
4. NEGOTIATION OF MEANING & INFORMATION EXCHANGE (Long's Interaction Hypothesis - Turns 4 & 5): Act as a realistic spoken Spanish interlocutor (mesero/a, compañero/a de clase, recepcionista, vendedor/a) introducing a live spoken constraint or counter-offer in Spanish so the student must negotiate meaning orally.
5. ORAL BREVITY: Maximum 2 short sentences in Spanish so the student can easily understand your audio reply without cognitive overload.

LANGUAGE POLICY (TARGET-LANGUAGE IMMERSION & TOLERANCE OF AMBIGUITY):
- All audio prompts, spoken recasts, and hints MUST be in SPANISH.
- AVOID THE SLIPPERY SLOPE TO ENGLISH: If the student speaks English or says "I don't understand" / "No entiendo", DO NOT translate into English right away. First, paraphrase in simpler A1 Spanish using cognates (e.g., "problema", "importante", "diferente", "necesitar"), synonyms, or familiar context clues, and invite them to speak in Spanish. Only give a brief 1-sentence English clarification if they fail twice.
"""


def _secret(name, default=None):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


def _clean_sheet_url(value) -> str:
    value = str(value or "").strip().strip('"').strip("'")
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", value)
    if m:
        return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/edit"
    if re.fullmatch(r"[a-zA-Z0-9_-]{25,}", value):
        return f"https://docs.google.com/spreadsheets/d/{value}/edit"
    return ""


def get_spreadsheet_url() -> str:
    candidates = []
    try:
        candidates.append(st.secrets["connections"]["gsheets"]["spreadsheet"])
    except Exception:
        pass
    candidates.append(_secret("SPREADSHEET_URL"))
    candidates.append(DEFAULT_SPREADSHEET_URL)
    for c in candidates:
        url = _clean_sheet_url(c)
        if url:
            return url
    return _clean_sheet_url(DEFAULT_SPREADSHEET_URL)


SPREADSHEET_URL = get_spreadsheet_url()


def get_gsheets_conn():
    return st.connection("gsheets", type=GSheetsConnection)


MODEL_NAME = _secret("MODEL_NAME", DEFAULT_MODEL)
INSTRUCTOR_KEY = _secret("INSTRUCTOR_KEY", TEACHER_PASSWORD)
THINKING_LEVEL = str(_secret("THINKING_LEVEL", "") or "").strip().lower()


def md_safe(text) -> str:
    return str(text or "").replace("$", "\\$")


# ==============================================================================
# 2. ESTADO DE SESIÓN
# ==============================================================================
@st.cache_resource
def get_latency_log():
    return []


DEFAULTS = {
    "cefr_level": "A1",
    "phase": "onboarding",
    "selected_topic": "",
    "topic_label": "",
    "audio_scenario": "",
    "turns_list": [],
    "current_turn_idx": 0,
    "correct_turns_count": 0,
    "attempts": 0,
    "personal_turn": 0,
    "celebrated": False,
    "xp": 0,
    "streak": 0,
    "chat_messages": [],
    "last_processed_audio": None,
    "active_user": None,
    "ledger_error": None,
    "active_model": MODEL_NAME,
    "thinking_ok": bool(THINKING_LEVEL),
    "ledger_buffer": [],
    "ledger_last_flush": 0.0,
    "voice_on": True,
    "slow_speech": False,
    "show_transcript": False,
    "personal_question": "",
    "show_custom_topic": False,
}
for _k, _v in DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = copy.deepcopy(_v)


def reset_speaking_mission(full: bool = False):
    s = st.session_state
    flush_ledger(force=True)
    s.phase = "onboarding"
    s.selected_topic = ""
    s.topic_label = ""
    s.audio_scenario = ""
    s.turns_list = []
    s.current_turn_idx = 0
    s.correct_turns_count = 0
    s.attempts = 0
    s.personal_turn = 0
    s.celebrated = False
    s.chat_messages = []
    s.last_processed_audio = None
    s.personal_question = ""
    s.show_transcript = False
    s.show_custom_topic = False
    if full:
        s.xp = 0
        s.streak = 0


# ==============================================================================
# 3. PERSISTENCIA (SQLite + Excel + Google Sheets vía gspread con caché)
# ==============================================================================
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
    except Exception:
        pass
    return conn


def init_db():
    conn = get_db_connection()
    with conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS diagnostics_sla
               (id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                student TEXT,
                level TEXT,
                topic TEXT,
                task TEXT,
                transcription TEXT,
                communicative_success TEXT,
                interlanguage_gap TEXT,
                negotiation_move TEXT,
                recast_target TEXT,
                feedback TEXT)"""
        )
    conn.close()


init_db()


def append_to_excel_ledger(row_dict: dict) -> str | None:
    if not os.path.exists(EXCEL_PATH):
        return None
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb[LEDGER_TAB] if LEDGER_TAB in wb.sheetnames else wb.create_sheet(title=LEDGER_TAB)
        if ws.max_row == 0:
            ws.append(LEDGER_COLUMNS)
        ordered_vals = [row_dict.get(col, "") for col in LEDGER_COLUMNS]
        ws.append(ordered_vals)
        wb.save(EXCEL_PATH)
        return None
    except PermissionError:
        return "⚠️ Close Microsoft Excel so the app can save data."
    except Exception as e:
        return str(e)


def sheets_write_status():
    try:
        conns = st.secrets["connections"]
    except Exception as e:
        return False, f"Missing [connections] in secrets ({type(e).__name__})."
    try:
        gs = conns["gsheets"]
    except Exception:
        return False, f"Missing [connections.gsheets]. Found: {list(conns.keys())}"
    keys = list(gs.keys())
    if str(gs.get("type", "")) != "service_account":
        return False, f'Missing type = "service_account" inside [connections.gsheets]. Found: {keys}'
    if not _clean_sheet_url(gs.get("spreadsheet")) and not _clean_sheet_url(DEFAULT_SPREADSHEET_URL):
        return False, "Missing or invalid spreadsheet URL in [connections.gsheets]."
    missing = [k for k in ("project_id", "private_key", "client_email") if not gs.get(k)]
    if missing:
        return False, f"Missing service account fields: {missing}"
    return True, "ok"


def sheets_write_enabled() -> bool:
    return sheets_write_status()[0]


@st.cache_resource(ttl=1800)
def _get_gspread_sheet():
    gs = dict(st.secrets["connections"]["gsheets"])
    target_url = _clean_sheet_url(gs.get("spreadsheet")) or get_spreadsheet_url()
    if not target_url:
        raise ValueError("Missing spreadsheet URL in [connections.gsheets].")
    creds_dict = {k: v for k, v in gs.items() if k != "spreadsheet"}
    gc = gspread.service_account_from_dict(creds_dict)
    return gc.open_by_url(target_url)


def append_to_ledger(row):
    rows = row if isinstance(row, list) else [row]
    ok, reason = sheets_write_status()
    if not ok:
        return f"Write disabled: {reason}"
    try:
        sh = _get_gspread_sheet()
        try:
            ws = sh.worksheet(LEDGER_TAB)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=LEDGER_TAB, rows=1000, cols=len(LEDGER_COLUMNS))
            ws.append_row(LEDGER_COLUMNS)

        if not ws.row_values(1):
            ws.append_row(LEDGER_COLUMNS)

        formatted_rows = [[str(r.get(col, "")) for col in LEDGER_COLUMNS] for r in rows]
        ws.append_rows(formatted_rows, value_input_option="USER_ENTERED")

        return None
    except Exception as e:
        _get_gspread_sheet.clear()
        return str(e)


def log_diagnostic(
    name: str,
    level: str,
    topic: str,
    task: str,
    spoken_transcript: str,
    comm_success: str,
    interlanguage_gap: str = "none",
    negotiation_move: str = "DIRECT_ANSWER",
    recast_target: str = "",
    feedback: str = "",
):
    ts = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    row_data = {
        "timestamp": ts,
        "student": name,
        "level": level,
        "topic": topic or "General",
        "task": task,
        "transcription": spoken_transcript,
        "communicative_success": comm_success,
        "interlanguage_gap": interlanguage_gap or "none",
        "negotiation_move": negotiation_move or "DIRECT_ANSWER",
        "recast_target": (recast_target or "")[:200],
        "feedback": (feedback or "")[:250],
    }

    try:
        conn = get_db_connection()
        with conn:
            conn.execute(
                """INSERT INTO diagnostics_sla
                   (timestamp, student, level, topic, task, transcription,
                    communicative_success, interlanguage_gap, negotiation_move,
                    recast_target, feedback)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                tuple(row_data[c] for c in LEDGER_COLUMNS),
            )
        conn.close()
    except Exception as e:
        print(f"SQLite Error: {e}")

    append_to_excel_ledger(row_data)
    st.session_state.ledger_buffer.append(row_data)
    flush_ledger()


def flush_ledger(force: bool = False):
    s = st.session_state
    buf = s.ledger_buffer
    if not buf:
        return
    now = time.time()
    if not s.ledger_last_flush:
        s.ledger_last_flush = now
    due = force or len(buf) >= LEDGER_BATCH_SIZE or (now - s.ledger_last_flush) >= LEDGER_MAX_AGE_SECONDS
    if due and not force and s.ledger_error and (now - s.ledger_last_flush) < 30:
        due = False
    if not due:
        return
    err = append_to_ledger(list(buf))
    s.ledger_error = err
    s.ledger_last_flush = now
    s.ledger_buffer = [] if err is None else buf[-60:]


@st.cache_data(ttl=30, show_spinner=False)
def read_ledger_sheet(url: str):
    try:
        sh = _get_gspread_sheet()
        ws = sh.worksheet(LEDGER_TAB)
        records = ws.get_all_records()
        return pd.DataFrame(records)
    except Exception:
        return get_gsheets_conn().read(spreadsheet=url, worksheet=LEDGER_TAB, ttl=0)


def load_ledger_df():
    if os.path.exists(EXCEL_PATH):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=LEDGER_TAB, engine="openpyxl")
            if not df.dropna(how="all").empty:
                return df.dropna(how="all"), f"Local Excel ({os.path.basename(EXCEL_PATH)})"
        except Exception:
            pass

    target_url = get_spreadsheet_url()
    try:
        df = read_ledger_sheet(target_url)
        if df is not None and not df.dropna(how="all").empty:
            return df.dropna(how="all"), f"Google Sheets ({LEDGER_TAB})"
    except Exception:
        pass

    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM diagnostics_sla ORDER BY timestamp DESC", conn)
    conn.close()
    return df, "Local SQLite"


# ==============================================================================
# 4. ROSTER (LECTURA DESDE PESTAÑA 'rosterspan107')
# ==============================================================================
def _normalize_roster(df: pd.DataFrame) -> dict:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    name_c = next((c for c in df.columns if any(k in c for k in ["student", "name", "nombre"])), None)
    pin_c = next((c for c in df.columns if c != name_c and any(k in c for k in ["pin", "pass", "clave", "code"])), None)
    if name_c is None or pin_c is None:
        raise ValueError(f"Could not find 'student' and 'pin' columns. Columns: {list(df.columns)}")
    df = df.dropna(subset=[name_c, pin_c])
    names = df[name_c].astype(str).str.strip()
    pins = df[pin_c].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    roster = {n: p for n, p in zip(names, pins) if n and n.lower() != "nan" and p and p.lower() != "nan"}
    if not roster:
        raise ValueError(f"Tab '{ROSTER_TAB}' is empty or has no valid records.")
    return roster


def _public_csv_url(sheet_url: str, tab: str) -> str:
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", sheet_url)
    if not m:
        raise ValueError("Could not extract ID from URL.")
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(tab)}"


@st.cache_data(ttl=600, show_spinner=False)
def load_roster(sheet_url: str):
    for path in EXCEL_CANDIDATES:
        if os.path.exists(path):
            try:
                df = pd.read_excel(path, sheet_name=ROSTER_TAB, engine="openpyxl")
                return _normalize_roster(df), f"Local Excel ({os.path.basename(path)} -> {ROSTER_TAB})"
            except Exception:
                pass

    target_url = _clean_sheet_url(sheet_url) or get_spreadsheet_url()
    if target_url:
        try:
            sh = _get_gspread_sheet()
            ws = sh.worksheet(ROSTER_TAB)
            df = pd.DataFrame(ws.get_all_records())
            return _normalize_roster(df), f"Google Sheets ({ROSTER_TAB})"
        except Exception:
            pass
        try:
            conn = get_gsheets_conn()
            df = conn.read(spreadsheet=target_url, worksheet=ROSTER_TAB, ttl=0)
            return _normalize_roster(df), f"Google Sheets connection ({ROSTER_TAB})"
        except Exception:
            pass
        try:
            df = pd.read_csv(_public_csv_url(target_url, ROSTER_TAB), dtype=str)
            return _normalize_roster(df), f"Google Sheets public CSV ({ROSTER_TAB})"
        except Exception:
            pass

    return {"Edgar": "6421", "Percy": "1234"}, "Local backup roster"


# ==============================================================================
# 5. MOTOR DE AUDIO EN ESPAÑOL (COMPREHENSIBLE INPUT - NORMAL & SLOW SPEED)
# ==============================================================================
def clean_for_tts(text: str) -> str:
    text = re.sub(r"👉?\s*Model[oa]?\s*[A-B]:?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[^\w\s'’?!¿¡.,;:-áéíóúüñÁÉÍÓÚÜÑ]", " ", text)
    text = text.replace("_", " ")
    return re.sub(r"\s+", " ", text).strip()


@st.cache_data(show_spinner=False)
def generate_spanish_audio_bytes(text: str, slow: bool = False):
    cleaned = clean_for_tts(text)
    if not cleaned:
        return None
    try:
        tts = gTTS(text=cleaned, lang="es", tld="com", slow=slow)
        fp = BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        return fp.read()
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def generate_bilingual_welcome_audio(student_name_str: str):
    es_part = f"¡Hola, {student_name_str}! Bienvenido al laboratorio oral de español. ¡Hoy solo escuchamos y hablamos!"
    en_part = (
        "Here is how your Spanish speaking session works: "
        "First, pick a real-life topic and listen carefully to the Spanish audio scene. "
        "Second, tap the microphone and answer out loud in Spanish. "
        "Turns 1 to 3 check your listening comprehension; Turns 4 and 5 are a live roleplay where you negotiate a solution in Spanish. "
        "If the Spanish audio feels fast, turn on Slower Spanish Rate in the left sidebar. ¡Vamos!"
    )
    try:
        fp_es = BytesIO()
        gTTS(text=es_part, lang="es", tld="com", slow=False).write_to_fp(fp_es)
        fp_es.seek(0)

        fp_en = BytesIO()
        gTTS(text=en_part, lang="en", tld="com", slow=False).write_to_fp(fp_en)
        fp_en.seek(0)

        return fp_es.read() + fp_en.read()
    except Exception:
        return generate_spanish_audio_bytes(es_part)


# ==============================================================================
# 6. BARRA LATERAL
# ==============================================================================
with st.sidebar:
    st.title("🇪🇸 Coach Edgarvich")
    st.caption("SPAN 107 - Pure Speaking & Listening Lab (A1)")

    api_key = _secret("GEMINI_API_KEY")
    if not api_key:
        api_key = st.text_input("🔑 Gemini API Key:", type="password")

    st.divider()
    active_sheet_url = get_spreadsheet_url()
    try:
        roster, roster_source = load_roster(active_sheet_url)
        roster_error = None
    except Exception as e:
        roster, roster_source, roster_error = {}, None, str(e)

    if roster_error:
        st.error(f"❌ Error loading '{ROSTER_TAB}' roster:\n{roster_error}")

    st.caption(f"📂 Roster: {roster_source}")
    if st.button("🔄 Reload Student Roster"):
        load_roster.clear()
        st.rerun()

    names = ["-- Select your name --"] + sorted(roster.keys())
    student_name = st.selectbox("Student Login:", names)
    entered_pin = st.text_input("PIN:", type="password")

    authenticated = False
    if student_name != "-- Select your name --" and entered_pin:
        expected = str(roster.get(student_name, "")).strip()
        if expected and hmac.compare_digest(expected, entered_pin.strip()):
            authenticated = True
            st.success(f"¡Bienvenido/a, {student_name}!")
        else:
            st.error("Invalid PIN.")

    if authenticated:
        st.divider()
        c1, c2 = st.columns(2)
        c1.metric("⭐ XP", st.session_state.xp)
        c2.metric("🔥 Streak", st.session_state.streak)

        st.markdown(f"**CEFR Level:** `{st.session_state.cefr_level}`")
        st.markdown(f"**Current Phase:** `{st.session_state.phase.replace('_', ' ').title()}`")
        if st.button("🔄 Choose Another Topic"):
            reset_speaking_mission()
            st.rerun()

    st.divider()
    st.session_state.voice_on = st.toggle(
        "🔊 Auto-play Coach Spanish Voice",
        value=st.session_state.voice_on,
        help="Automatically plays the coach's spoken Spanish prompts.",
    )
    st.session_state.slow_speech = st.toggle(
        "🐢 Slower Spanish Rate (i+1 Scaffolding)",
        value=st.session_state.slow_speech,
        help="Slows down the spoken Spanish audio to support listening comprehension.",
    )

    # ---------------- Panel del Profesor (Teacher Access: peru2026) ----------------
    st.divider()
    instructor_input = st.text_input("Teacher Access:", type="password")
    effective_admin_key = str(INSTRUCTOR_KEY or TEACHER_PASSWORD)
    if instructor_input and hmac.compare_digest(instructor_input, effective_admin_key):
        st.subheader("👨‍🏫 Instructor SLA Control Panel")

        with st.expander("🔌 Connection Status", expanded=True):
            st.write(f"**Spreadsheet:** {active_sheet_url or '❌ empty'}")
            st.write(f"**Roster Worksheet:** `{ROSTER_TAB}`")
            st.write(f"**Logs Worksheet:** `{LEDGER_TAB}`")
            st.write(f"**Roster Source:** {roster_source or '❌ failed'} ({len(roster)} students)")
            st.write(f"**Gemini model:** `{st.session_state.active_model}`")
            _lat = get_latency_log()
            if _lat:
                _secs = sorted(x["seconds"] for x in _lat)
                st.write(
                    f"**Gemini latency (last {len(_lat)} calls):** avg {sum(_secs) / len(_secs):.1f}s · "
                    f"median {_secs[len(_secs) // 2]:.1f}s · max {_secs[-1]:.1f}s"
                )
            st.write(
                "**Sheets Writing:** "
                + ("✅ Service Account connected" if sheets_write_enabled() else "⚠️ Local SQLite/Excel only")
            )
            _ok, _reason = sheets_write_status()
            if not _ok:
                st.error(f"Why writing is disabled: {_reason}")
            if st.session_state.ledger_error:
                st.warning(f"Last Sheets write warning: {st.session_state.ledger_error}")
            if st.button("🧪 Write a test Spanish SLA row"):
                err = append_to_ledger({
                    "timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "student": "TEST_SPEAKER_ES",
                    "level": "A1",
                    "topic": "Café y horario",
                    "task": "Speaking Test Verification",
                    "transcription": "Yo quiero un café por favor",
                    "communicative_success": "ACHIEVED",
                    "interlanguage_gap": "none",
                    "negotiation_move": "DIRECT_ANSWER",
                    "recast_target": "Quiero un café, por favor.",
                    "feedback": "Conexión SLA verificada correctamente",
                })
                if err is None:
                    st.success(f"✅ Row written successfully to '{LEDGER_TAB}'.")
                else:
                    st.error(err)

        with st.expander("📊 Student SLA Speaking Ledger", expanded=True):
            if st.button("🔄 Refresh Analytics"):
                read_ledger_sheet.clear()
                st.rerun()
            df_diag, ledger_source = load_ledger_df()
            st.caption(f"Data source: {ledger_source}")
            if not df_diag.empty:
                students = ["All"] + sorted(df_diag["student"].dropna().astype(str).unique())
                chosen = st.selectbox("Filter by student:", students)
                df_show = df_diag if chosen == "All" else df_diag[df_diag["student"] == chosen]
                if "communicative_success" in df_show.columns:
                    summary = df_show.groupby(["student", "communicative_success"]).size().unstack(fill_value=0)
                    st.dataframe(summary)
                if "interlanguage_gap" in df_show.columns:
                    st.caption("Interlanguage Patterns (Selinker / Corder):")
                    gaps = df_show[df_show["interlanguage_gap"] != "none"]["interlanguage_gap"].value_counts()
                    if not gaps.empty:
                        st.bar_chart(gaps)
                st.dataframe(df_show)
                st.download_button(
                    "📥 Download SLA CSV", df_show.to_csv(index=False),
                    "spanish_speaking_sla_diagnostics.csv", "text/csv",
                )
            else:
                st.info(f"No speaking records found in '{LEDGER_TAB}' yet.")

# ==============================================================================
# 7. CLIENTE GEMINI Y CONTROL DE ACCESO (CON FALLBACK AUTOMÁTICO ANTE 503)
# ==============================================================================
if not api_key or not authenticated:
    st.info("👈 Please select your name and enter your PIN in the sidebar to start speaking Spanish.")
    st.stop()

if st.session_state.active_user != student_name:
    reset_speaking_mission(full=True)
    st.session_state.active_user = student_name


@st.cache_resource
def get_client(key):
    return genai.Client(api_key=key)


client = get_client(api_key)
student_first_name = student_name.split()[0]


def gen(contents, json_mode=False, temperature=0.2, retries=4, system=None, label="turn"):
    def build_cfg(with_thinking):
        kwargs = {"temperature": temperature}
        if json_mode:
            kwargs["response_mime_type"] = "application/json"
        if system:
            kwargs["system_instruction"] = system
        if with_thinking and THINKING_LEVEL:
            try:
                kwargs["thinking_config"] = types.ThinkingConfig(thinking_level=THINKING_LEVEL)
            except Exception:
                pass
        return types.GenerateContentConfig(**kwargs)

    use_thinking = bool(THINKING_LEVEL) and st.session_state.thinking_ok
    cfg = build_cfg(use_thinking)
    model = st.session_state.active_model or MODEL_NAME
    last_err = None
    for attempt in range(retries):
        started = time.perf_counter()
        try:
            resp = client.models.generate_content(model=model, contents=contents, config=cfg)
            log = get_latency_log()
            log.append({"label": label, "model": model, "seconds": round(time.perf_counter() - started, 2)})
            del log[:-60]
            return resp
        except Exception as e:
            last_err = e
            msg = str(e)
            if use_thinking and "thinking" in msg.lower() and ("400" in msg or "INVALID_ARGUMENT" in msg):
                use_thinking = False
                st.session_state.thinking_ok = False
                cfg = build_cfg(False)
                continue
            if ("404" in msg or "NOT_FOUND" in msg) and model != FALLBACK_MODEL:
                model = FALLBACK_MODEL
                st.session_state.active_model = FALLBACK_MODEL
                continue
            if any(k in msg for k in ("503", "UNAVAILABLE", "429")) and attempt < retries - 1:
                if model != FALLBACK_MODEL:
                    model = FALLBACK_MODEL
                time.sleep(1.0 * (attempt + 1))
                continue
            raise
    raise last_err


def parse_json(text: str):
    cleaned = (text or "").strip()
    match = re.search(r"(\{.*\}|\[.*\])", cleaned, re.DOTALL)
    if match:
        cleaned = match.group(1)
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.IGNORECASE).strip()
    data = json.loads(cleaned)
    if isinstance(data, list) and data:
        data = data[0]
    return data


FRIENDLY_ERROR = "⚠️ Technical hiccup connecting with your coach. Please tap the microphone and try again."

# ==============================================================================
# 8. GENERADOR DE ESCENARIOS ORALES EN ESPAÑOL (LISTENING + SPEAKING 3+2)
# ==============================================================================
def render_spoken_history(messages):
    """Muestra el historial de transcripciones y recasts previos en un acordeón discreto."""
    if len(messages) > 1:
        with st.expander(f"📜 Previous spoken turns ({len(messages)-1})", expanded=False):
            for old in messages[:-1]:
                with st.chat_message(old["role"]):
                    st.markdown(md_safe(old["content"]))


def render_oral_survival_phrases():
    """Frases orales de emergencia en español para negociar significado hablando."""
    with st.expander("🆘 Need help? Say one of these phrases in Spanish into your mic", expanded=False):
        st.markdown(
            "- 🎙️ *\"¿Puede repetir, por favor?\"* (Can you repeat, please?)  \n"
            "- 🎙️ *\"Más despacio, por favor.\"* (More slowly, please.)  \n"
            "- 🎙️ *\"¿Qué significa [palabra]?\"* (What does [word] mean?)  \n"
            "- 🎙️ *\"No entiendo. ¿Puede explicar con palabras simples?\"* (I don't understand. Can you explain in simple words?)"
        )


def format_speaking_prompt(t_item: dict, t_idx: int, total: int, prefix: str = "") -> str:
    t_label = "Listening Check" if t_item.get("type") == "listening_check" else "Live Roleplay Negotiation"
    header = f"{prefix}\n\n---\n" if prefix else ""
    body = f"**Turno Oral {t_idx + 1} de {total} ({t_label}):** {t_item['prompt']}"
    if t_idx < 3:
        return (
            f"{header}{body}\n\n"
            f"🗣️ Modelo A: \"{t_item.get('model_a', 'Yo creo que...')}\"\n"
            f"🗣️ Modelo B: \"{t_item.get('model_b', 'En esta situación...')}\""
        )
    return f"{header}{body}"


def generate_speaking_mission(topic: str, variant: int = 1):
    prompt = f"""
    You are an expert Spanish as a Second Language (ELE / SSL) instructor designing a SPOKEN (Listening + Speaking only) communicative mission grounded in Second Language Acquisition (Krashen Comprehensible Input, Long Interaction Hypothesis, Swain Output Hypothesis).
    Target learners: English-speaking college students in Beginner Spanish (SPAN 107 - CEFR Level A1).
    Topic: {topic}
    Variation #{variant}: create an authentic spoken everyday situation in a Spanish-speaking campus or city setting.

    SLA AUDIO & SPEAKING DESIGN RULES (ALL CONTENT IN SPANISH):
    1. SPOKEN AUDIO SCENE (45-65 words, natural spoken A1 Spanish):
       - Write as a short spoken monologue or voice message in SPANISH that the student will LISTEN to (no reading required).
       - Use clear, high-frequency A1 spoken Spanish (Presente de Indicativo, clear times, prices, places, and cognates).
       - Highlight 4-5 **key spoken clues** (times, places, constraints) with **double asterisks** (only shown if the student opens the optional transcript).
       - Present a realistic everyday conflict or scheduling/budget dilemma without giving away the solution.

    2. 5 SPOKEN INTERACTION TURNS IN SPANISH (3 + 2 STRUCTURE):
       - Turn 1 (Listening Interpretive Check - Verdadero/Falso + Spoken Reason in Spanish): A spoken statement in Spanish about the situation asking the student to say aloud whether it is "¿Verdadero o falso? ¿Por qué?".
       - Turn 2 (Conflict Comprehension): A spoken question in Spanish about why the character has a problem or a choice between two details heard in the audio.
       - Turn 3 (Constraint Check): A spoken question in Spanish about the obstacle (hora, dinero, lugar) preventing an easy solution.
       - Turn 4 (Live Roleplay - Information Exchange): Coach Edgarvich speaks IN CHARACTER in Spanish (e.g., mesero/a, compañero/a, recepcionista, vendedor/a) introducing a NEW spoken constraint or counter-offer not in the initial audio, asking the student to make an oral decision in Spanish.
       - Turn 5 (Spoken Negotiation / Request): The student must speak a polite request, excuse, or negotiation in Spanish to resolve the dilemma with that person.

       For each turn (1-5):
       - prompt: clear spoken prompt in A1 Spanish.
       - expected_core: target communicative meaning in A1 Spanish.
       - context_hint: clue from the spoken scene in Spanish.
       - model_a & model_b: short, natural spoken starters in A1 Spanish (5-7 words) suitable for oral shadowing.

    3. PERSONAL SPEAKING QUESTION: 1 open spoken question in A1 Spanish connecting the theme to the student's real life, plus 2 oral starters in Spanish.

    Return strictly a valid JSON object:
    {{
      "audio_scenario": "Escena corta hablada en español con **pistas clave**...",
      "personal_question": "¿Qué haces tú normalmente cuando tienes un problema así en la universidad?",
      "personal_starters": ["Normalmente, yo...", "Cuando pasa eso, prefiero..."],
      "turns": [
        {{"turn_num": 1, "type": "listening_check", "prompt": "...", "expected_core": "...", "context_hint": "...", "model_a": "...", "model_b": "..."}},
        {{"turn_num": 2, "type": "listening_check", "prompt": "...", "expected_core": "...", "context_hint": "...", "model_a": "...", "model_b": "..."}},
        {{"turn_num": 3, "type": "listening_check", "prompt": "...", "expected_core": "...", "context_hint": "...", "model_a": "...", "model_b": "..."}},
        {{"turn_num": 4, "type": "roleplay_decision", "prompt": "...", "expected_core": "...", "context_hint": "...", "model_a": "...", "model_b": "..."}},
        {{"turn_num": 5, "type": "roleplay_negotiation", "prompt": "...", "expected_core": "...", "context_hint": "...", "model_a": "...", "model_b": "..."}}
      ]
    }}
    """
    resp = gen([types.Part.from_text(text=prompt)], json_mode=True, temperature=0.25, label="speaking_mission_es")
    package = parse_json(resp.text)
    if not package.get("audio_scenario") or not package.get("turns"):
        raise ValueError("Model response missed audio scenario or speaking turns.")
    return package


@st.cache_data(ttl=SCENARIO_CACHE_TTL, show_spinner=False)
def cached_speaking_mission(topic: str, variant: int):
    return generate_speaking_mission(topic, variant)


def build_personal_opener(package) -> str:
    q = str(package.get("personal_question") or "").strip()
    if not q:
        return ""
    starters = [str(x).strip() for x in (package.get("personal_starters") or []) if str(x).strip()][:2]
    a = starters[0] if starters else "Normalmente, yo..."
    b = starters[1] if len(starters) > 1 else "En mi opinión..."
    st.session_state.personal_starters_cache = (a, b)
    return f"¡Excelente trabajo! Ahora hablemos de tu vida real, {student_first_name}. {q}"


def start_speaking_session(t_item: dict, total: int):
    st.session_state.current_turn_idx = 0
    st.session_state.correct_turns_count = 0
    st.session_state.attempts = 0
    st.session_state.phase = "speaking_practice"

    first_msg = format_speaking_prompt(t_item, 0, total)
    st.session_state.chat_messages = [{
        "role": "assistant",
        "content": first_msg,
        "auto_play": True,
    }]


def advance_speaking_turn(prefix: str):
    s = st.session_state
    s.current_turn_idx += 1
    s.attempts = 0
    total = len(s.turns_list)
    if s.current_turn_idx < total:
        nxt_item = s.turns_list[s.current_turn_idx]
        msg = format_speaking_prompt(nxt_item, s.current_turn_idx, total, prefix=prefix)
    elif s.correct_turns_count >= max(total - 1, 1):
        s.phase = "personal_conversation"
        s.personal_turn = 0
        s.celebrated = False
        msg = (
            f"{prefix}\n\n🎉 **¡Misión oral completada! Puntuación: {s.correct_turns_count}/{total}.**\n\n"
            "¡Ahora conversemos libremente en español sobre tu propia vida!"
        )
    else:
        s.phase = "review"
        msg = f"{prefix}\n\n📊 **Puntuación oral final: {s.correct_turns_count}/{total}.**"
    s.chat_messages.append({"role": "assistant", "content": msg, "auto_play": True})
    if s.phase != "speaking_practice":
        flush_ledger(force=True)


# ==============================================================================
# 9. INTERFAZ PRINCIPAL (100 % LISTENING & SPEAKING EN ESPAÑOL)
# ==============================================================================

# ---------------- FASE 1: SELECCIÓN DE TEMA DE CONVERSACIÓN ----------------
if st.session_state.phase == "onboarding":
    st.markdown("<div style='padding-top: 12px;'></div>", unsafe_allow_html=True)
    st.markdown(f"### 🇪🇸 ¡Hola {student_first_name}! Welcome to your Spanish Speaking Lab.")

    with st.container(border=True):
        st.markdown(
            "**🇪🇸 Coach Edgarvich:** *¡Hoy no escribimos! Solo escucha con atención y habla en español por tu micrófono.*  \n"
            "**🇺🇸 How this Speaking Lab works (SLA Communicative Approach):**  \n"
            "1. **Active Listening (Comprehensible Input):** Pick a topic and listen to the Spanish scene (turn on **🐢 Slower Spanish Rate** in the sidebar if needed).  \n"
            "2. **Speak Without Fear (Low Affective Filter):** Tap the microphone and speak in Spanish. Your coach will transcribe your speech, model natural Spanish (*Recast*), and guide you.  \n"
            "3. **3+2 Interaction:** Turns 1–3 check listening comprehension; Turns 4–5 put you in a live Spanish roleplay to solve a practical problem."
        )
        if st.session_state.voice_on:
            welcome_audio = generate_bilingual_welcome_audio(student_first_name)
            if welcome_audio:
                st.audio(welcome_audio, format="audio/mp3", autoplay=True)

    st.write("Choose a Spanish speaking topic to begin:")

    col1, col2 = st.columns(2)
    chosen_topic = None
    chosen_label = ""

    with col1:
        if st.button("☕ Pedir café antes de clase (Coffee & Morning Rush)", use_container_width=True):
            chosen_label = "Café y horario de clase"
            chosen_topic = "Ordering coffee and breakfast at a busy campus cafe in Spanish before an 8:00 AM class with very little time"
        if st.button("🎓 Grupo de estudio y horarios (Campus Study Group)", use_container_width=True):
            chosen_label = "Grupo de estudio"
            chosen_topic = "Coordinating a Spanish class group presentation on campus when two classmates have conflicting schedules"
        if st.button("🏥 En la farmacia (Pharmacy & Feeling Sick)", use_container_width=True):
            chosen_label = "En la farmacia"
            chosen_topic = "Explaining simple symptoms in Spanish at a pharmacy or campus clinic and asking for medicine"
    with col2:
        if st.button("🛒 Mercado y presupuesto (Market & Budget)", use_container_width=True):
            chosen_label = "Mercado y presupuesto"
            chosen_topic = "Shopping for fresh food at a local market in Spanish with a 15-dollar/euro budget when an item is sold out"
        if st.button("✈️ Hotel y reservación (Hotel Check-in)", use_container_width=True):
            chosen_label = "Hotel y reservación"
            chosen_topic = "Checking into a hotel late at night in Spanish when the reservation has the wrong room type"
        if st.button("🍽️ Restaurante y pedido (Restaurant Order)", use_container_width=True):
            chosen_label = "Restaurante y pedido"
            chosen_topic = "Ordering lunch at a restaurant in Spanish and politely fixing a mistake in the order"

    st.markdown("---")
    if st.button("🎙️ Propose my own Spanish speaking topic", type="secondary", use_container_width=True):
        st.session_state.show_custom_topic = not st.session_state.show_custom_topic

    if st.session_state.show_custom_topic:
        st.info("💡 Enter any everyday situation you want to practice speaking in Spanish (e.g., 'Taking a taxi in Madrid', 'Meeting a new roommate').")
        custom_input = st.text_input("Topic keyword:", placeholder="e.g., Meeting a new roommate on campus")
        if st.button("🚀 Start Spanish Speaking Session", type="primary"):
            if custom_input.strip():
                chosen_label = custom_input.strip()[:40]
                chosen_topic = f"A1 spoken Spanish interaction: {custom_input.strip()}"
            else:
                st.warning("Please enter a short topic name.")

    if chosen_topic:
        try:
            with st.spinner("Preparando tu escenario de audio y turnos orales en español..."):
                package = cached_speaking_mission(chosen_topic, random.randint(1, SCENARIO_VARIANTS))
            st.session_state.selected_topic = chosen_topic
            st.session_state.topic_label = chosen_label
            st.session_state.audio_scenario = package["audio_scenario"]
            st.session_state.turns_list = package["turns"]
            st.session_state.personal_question = build_personal_opener(package)
            start_speaking_session(package["turns"][0], len(package["turns"]))
            st.rerun()
        except Exception as e:
            st.error(f"{FRIENDLY_ERROR}\n\nDetail: {e}")

# ---------------- FASE 2: PRÁCTICA ORAL GUIADA (3+2 LISTENING & SPEAKING) ----------------
elif st.session_state.phase == "speaking_practice":
    total_t = len(st.session_state.turns_list)
    t_idx = st.session_state.current_turn_idx
    turn_obj = st.session_state.turns_list[t_idx]
    is_roleplay = t_idx >= 3

    phase_title = "Live Roleplay Negotiation (Spanish)" if is_roleplay else "Listening & Speaking Check (Spanish)"
    st.progress(min(t_idx / total_t, 1.0), text=f"{phase_title}: Turn {min(t_idx + 1, total_t)} of {total_t}")

    col_listen, col_speak = st.columns(2, gap="large")

    # ---------------- COLUMNA IZQUIERDA: AUDIO INPUT ----------------
    with col_listen:
        st.markdown("#### 🎧 Step 1: Listen to the Spanish Situation")
        with st.container(border=True):
            st.info("👂 **Listening Focus:** Listen for who is speaking, what they need, and the problem (hora, precio o lugar).")
            scenario_audio = generate_spanish_audio_bytes(
                st.session_state.audio_scenario,
                slow=st.session_state.slow_speech,
            )
            if scenario_audio:
                speed_tag = "🐢 Slow Spanish" if st.session_state.slow_speech else "▶️ Normal Spanish"
                st.caption(f"🔊 Play / Replay Scene Audio ({speed_tag}):")
                st.audio(scenario_audio, format="audio/mp3")

            st.markdown("---")
            with st.expander("👁️ Need Listening Support? Show Spanish Transcript", expanded=False):
                st.caption("Try listening twice before opening the transcript:")
                st.markdown(md_safe(st.session_state.audio_scenario))

    # ---------------- COLUMNA DERECHA: PRODUCCIÓN ORAL ----------------
    with col_speak:
        st.markdown("#### 🎙️ Step 2: Speak in Spanish with Coach Edgarvich")
        render_spoken_history(st.session_state.chat_messages)

        latest_msg = st.session_state.chat_messages[-1]
        with st.chat_message(latest_msg["role"]):
            st.markdown(md_safe(latest_msg["content"]))

        if is_roleplay:
            with st.expander("🗣️ Need a Spanish spoken starter? (Optional Shadowing)", expanded=False):
                st.markdown(
                    f"👉 **Modelo A:** \"{turn_obj.get('model_a', '')}\"  \n"
                    f"👉 **Modelo B:** \"{turn_obj.get('model_b', '')}\""
                )

        # ÚNICA ENTRADA DEL ESTUDIANTE: MICRÓFONO
        st.markdown("**🎙️ Tap the microphone to speak your answer in Spanish:**")
        rec_audio = st.audio_input("Record your Spanish voice:", key=f"speak_mic_turn_{t_idx}")

        render_oral_survival_phrases()

        # Reproductor de voz del coach
        st.caption("🔊 Listen to Coach Edgarvich:")
        prompt_audio = generate_spanish_audio_bytes(
            latest_msg["content"],
            slow=st.session_state.slow_speech,
        )
        if prompt_audio:
            st.audio(prompt_audio, format="audio/mp3", autoplay=latest_msg.get("auto_play", False))
        latest_msg["auto_play"] = False

        if rec_audio is not None:
            raw_bytes = rec_audio.getvalue()
            if raw_bytes != st.session_state.last_processed_audio:
                st.session_state.last_processed_audio = raw_bytes

                eval_instructions = (
                    "LIVE SPANISH ROLEPLAY EVALUATION (Turns 4 & 5 - Long's Interaction Hypothesis): "
                    "Listen carefully to the student's spoken Spanish audio. is_correct=true if the student communicates a realistic choice, polite request, or negotiation in Spanish that addresses the roleplay constraint, even with beginner A1 pronunciation or minor grammar errors. "
                    "Provide a warm spoken recast in 1 short sentence in Spanish. NEVER ask another question in feedback_text."
                    if is_roleplay else
                    "SPANISH LISTENING COMPREHENSION EVALUATION (Turns 1-3): "
                    "Listen carefully to the student's spoken Spanish audio. is_correct=true if their spoken meaning accurately reflects the Spanish audio scenario. "
                    "Provide a warm spoken recast in 1 short sentence in Spanish without metalanguage. NEVER ask another question in feedback_text."
                )

                eval_prompt = f"""
                Spanish Audio Scenario: "{st.session_state.audio_scenario}"
                Speaking Turn: {t_idx + 1} of {total_t}
                Coach Spanish Prompt: "{turn_obj['prompt']}"
                Expected Communicative Goal: "{turn_obj['expected_core']}"
                Attempt Number: {st.session_state.attempts + 1}

                {eval_instructions}

                TASK & SLA DIAGNOSTIC CLASSIFICATION:
                1. student_transcript: Exact transcription of what the student said in the audio (if silent/unintelligible, write "[audio poco claro]").
                2. intent: "INQUIRY" (if student asks for repetition, meaning, speaks English, or audio is unclear) or "ANSWER"
                3. is_correct: true/false based on spoken communicative meaning
                4. communicative_success: "ACHIEVED" (meaning clear & goal met), "PARTIAL" (partially understood or incomplete), or "BREAKDOWN" (unclear/off-topic/only English)
                5. interlanguage_gap: classify primary A1 Spanish interlanguage issue using ONE tag: "none", "ser_vs_estar", "verb_conjugation", "gender_number_agreement", "preposition_article", "vocabulary_gap", "pronunciation_intelligibility", or "l1_english"
                6. negotiation_move: classify learner's interactional move using ONE tag: "DIRECT_ANSWER", "ROLEPLAY_NEGOTIATION", "CLARIFICATION_REQUEST", or "L1_FALLBACK"
                7. recast: Natural rewrite in 1 short spoken A1 Spanish sentence modeling accurate form
                8. feedback_text:
                   - If intent is "ANSWER": give ONLY the warm spoken recast/validation in Spanish (1-2 short sentences). NEVER ask a new question here.
                   - If intent is "INQUIRY": give a simple spoken hint in A1 Spanish (cognates, synonyms, context clues) without translating into English unless Attempt Number >= 2.

                Return strictly JSON:
                {{
                  "student_transcript": "...",
                  "intent": "INQUIRY" or "ANSWER",
                  "is_correct": true or false,
                  "communicative_success": "ACHIEVED" or "PARTIAL" or "BREAKDOWN",
                  "interlanguage_gap": "...",
                  "negotiation_move": "...",
                  "recast": "...",
                  "feedback_text": "...",
                  "model_a": "...",
                  "model_b": "..."
                }}
                """
                parts = [
                    types.Part.from_bytes(data=raw_bytes, mime_type="audio/wav"),
                    types.Part.from_text(text=eval_prompt),
                ]

                try:
                    with st.spinner("🎧 Coach Edgarvich está escuchando tu español..."):
                        parsed = parse_json(
                            gen(parts, json_mode=True, temperature=0.1, system=SPEAKING_TUTOR_RULES, label="speak_eval_es").text
                        )
                except Exception as e:
                    st.session_state.chat_messages.append(
                        {"role": "assistant", "content": f"{FRIENDLY_ERROR}\n\n`{e}`", "auto_play": False}
                    )
                    st.rerun()

                transcript = (parsed.get("student_transcript") or "[audio hablado]").strip()
                st.session_state.chat_messages.append({"role": "user", "content": f"🎙️ *\"{transcript}\"* "})

                intent = parsed.get("intent", "ANSWER")
                fb_text = parsed.get("feedback_text", "")
                recast = (parsed.get("recast") or "").strip()
                is_correct = bool(parsed.get("is_correct", False))
                comm_success = parsed.get("communicative_success") or ("ACHIEVED" if is_correct else "PARTIAL")
                gap = parsed.get("interlanguage_gap") or "none"
                move = parsed.get("negotiation_move") or ("ROLEPLAY_NEGOTIATION" if is_roleplay else "DIRECT_ANSWER")

                topic_tag = st.session_state.topic_label or "Español A1"
                task_tag = f"Roleplay Turno {t_idx+1}" if is_roleplay else f"Comprensión Oral T{t_idx+1}"

                if intent == "INQUIRY":
                    st.session_state.chat_messages.append({
                        "role": "assistant",
                        "auto_play": True,
                        "content": (
                            f"💡 **Pista en español:** {fb_text}\n\n---\n"
                            f"👉 **Intenta decir:** {turn_obj['prompt']}\n\n"
                            f"🗣️ **Modelo A:** \"{turn_obj.get('model_a', '')}\"\n"
                            f"🗣️ **Modelo B:** \"{turn_obj.get('model_b', '')}\""
                        ),
                    })
                    log_diagnostic(
                        student_name, "A1", topic_tag, f"{task_tag} (Aclaración)",
                        transcript, "CLARIFICATION", gap, "CLARIFICATION_REQUEST", recast, fb_text
                    )
                elif is_correct:
                    points = 10 if st.session_state.attempts == 0 else 5
                    st.session_state.xp += points
                    st.session_state.streak += 1
                    st.session_state.correct_turns_count += 1
                    log_diagnostic(
                        student_name, "A1", topic_tag, task_tag,
                        transcript, comm_success, gap, move, recast, fb_text
                    )
                    advance_speaking_turn(f"{fb_text}\n\n⭐ **+{points} XP**")
                else:
                    st.session_state.attempts += 1
                    st.session_state.streak = 0
                    log_diagnostic(
                        student_name, "A1", topic_tag, task_tag,
                        transcript, comm_success, gap, move, recast, fb_text
                    )
                    if st.session_state.attempts >= MAX_ATTEMPTS:
                        sol_label = "Una forma natural de decirlo:" if is_roleplay else "Respuesta esperada:"
                        reveal = (
                            f"{fb_text}\n\n✅ **{sol_label}** \"{turn_obj['expected_core']}\"\n"
                            f"🎧 *Pista del audio:* {turn_obj['context_hint']}\n\n_¡Pasemos al siguiente turno!_"
                        )
                        log_diagnostic(
                            student_name, "A1", topic_tag, task_tag,
                            transcript, "REVEALED", gap, move, turn_obj["expected_core"], "Modelo oral revelado"
                        )
                        advance_speaking_turn(reveal)
                    else:
                        m_a = parsed.get("model_a") or f"Modelo A: {turn_obj.get('model_a', '')}"
                        m_b = parsed.get("model_b") or f"Modelo B: {turn_obj.get('model_b', '')}"
                        st.session_state.chat_messages.append({
                            "role": "assistant",
                            "auto_play": True,
                            "content": (
                                f"{fb_text}\n\n🗣️ **{m_a}**\n🗣️ **{m_b}**\n\n"
                                f"_Intento {st.session_state.attempts} de {MAX_ATTEMPTS} — ¡presiona el micrófono y dilo en español!_"
                            ),
                        })
                st.rerun()

# ---------------- FASE 3: REPASO ORAL ----------------
elif st.session_state.phase == "review":
    st.markdown("### 📊 Resumen de Práctica Oral (Español A1)")
    with st.expander("🎧 Escuchar de nuevo el escenario y ver transcripción", expanded=False):
        aud = generate_spanish_audio_bytes(st.session_state.audio_scenario, slow=st.session_state.slow_speech)
        if aud:
            st.audio(aud, format="audio/mp3")
        st.markdown(md_safe(st.session_state.audio_scenario))
    render_spoken_history(st.session_state.chat_messages)
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("🔁 Practicar este tema otra vez", use_container_width=True):
            start_speaking_session(st.session_state.turns_list[0], len(st.session_state.turns_list))
            st.rerun()
    with col_b:
        if st.button("🎙️ Elegir un tema nuevo", use_container_width=True):
            reset_speaking_mission()
            st.rerun()

# ---------------- FASE 4: CONVERSACIÓN PERSONAL LIBRE EN ESPAÑOL (VOZ) ----------------
elif st.session_state.phase == "personal_conversation":
    if st.session_state.personal_turn == 0 and st.session_state.personal_question:
        st.session_state.chat_messages.append(
            {"role": "assistant", "content": st.session_state.personal_question, "auto_play": True}
        )
        st.session_state.personal_turn = 1
        st.rerun()

    turn = st.session_state.personal_turn
    finished = turn > 3
    st.markdown(f"### 🗣️ Conversación Libre en Español (Turno {min(turn, 3)} de 3)")

    render_spoken_history(st.session_state.chat_messages)

    if finished:
        if not st.session_state.celebrated:
            flush_ledger(force=True)
            st.balloons()
            st.session_state.celebrated = True
        st.success("🏆 ¡FELICITACIONES! Completaste toda tu sesión oral en español.")
        st.metric("⭐ Total Speaking XP", st.session_state.xp)

        with st.container(border=True):
            st.markdown("#### 🧠 Quick Self-Reflection")
            reflection = st.radio(
                "How did your Spanish speaking practice feel today?",
                [
                    "💪 I understood the spoken Spanish without needing the transcript",
                    "🌱 Listening was challenging, but I solved the roleplay by speaking Spanish",
                    "🎯 I want to practice speaking on this topic one more time",
                ],
                key="speaking_reflection_radio",
            )
            if st.button("💾 Save Reflection & Pick a New Topic", type="primary"):
                log_diagnostic(
                    student_name, "A1", st.session_state.topic_label or "Español A1",
                    "Autoevaluación Oral", reflection, "REFLECTED", "none", "SELF_REFLECTION", "", "Sesión completada"
                )
                reset_speaking_mission()
                st.rerun()
        st.stop()

    latest_p_msg = st.session_state.chat_messages[-1]
    with st.chat_message(latest_p_msg["role"]):
        st.markdown(md_safe(latest_p_msg["content"]))

    starters_cache = st.session_state.get("personal_starters_cache")
    if turn == 1 and starters_cache:
        with st.expander("🗣️ ¿Necesitas un inicio de frase en español?", expanded=False):
            st.markdown(
                f"👉 **Modelo A:** \"{starters_cache[0]}\"  \n"
                f"👉 **Modelo B:** \"{starters_cache[1]}\""
            )

    # ÚNICA ENTRADA: MICRÓFONO
    st.markdown("**🎙️ Presiona el micrófono y responde en español:**")
    p_rec = st.audio_input("Habla en español:", key=f"speak_mic_personal_{turn}")

    render_oral_survival_phrases()

    st.caption("🔊 Escucha a Coach Edgarvich:")
    paud = generate_spanish_audio_bytes(latest_p_msg["content"], slow=st.session_state.slow_speech)
    if paud:
        st.audio(paud, format="audio/mp3", autoplay=latest_p_msg.get("auto_play", False))
    latest_p_msg["auto_play"] = False

    if p_rec is not None:
        raw = p_rec.getvalue()
        if raw != st.session_state.last_processed_audio:
            st.session_state.last_processed_audio = raw

            p_prompt = f"""
            Turn {turn} of 3. Free spoken personal conversation in A1 Spanish.
            Listen to the student's Spanish audio.
            1. Transcribe accurately what the student said in student_transcript.
            2. Classify interlanguage_gap ("none", "ser_vs_estar", "verb_conjugation", "gender_number_agreement", "preposition_article", "vocabulary_gap", "l1_english").
            3. Provide recast: natural A1 Spanish rewrite of what the student said.
            4. In your reply (message), give a natural spoken recast in Spanish in your FIRST sentence, then {"ask ONE open follow-up question in simple A1 Spanish" if turn < 3 else "give a warm spoken closing compliment in Spanish"}.
            Return strictly JSON:
            {{
              "student_transcript": "...",
              "interlanguage_gap": "...",
              "recast": "...",
              "message": "..."
            }}
            """
            parts = [
                types.Part.from_bytes(data=raw, mime_type="audio/wav"),
                types.Part.from_text(text=p_prompt),
            ]

            with st.spinner("🎧 Coach Edgarvich está escuchando..."):
                parsed = parse_json(
                    gen(parts, json_mode=True, temperature=0.3, system=SPEAKING_TUTOR_RULES, label="speak_personal_es").text
                )

            transcript = (parsed.get("student_transcript") or "[audio hablado]").strip()
            msg = parsed.get("message", "")
            recast = (parsed.get("recast") or "").strip()
            gap = parsed.get("interlanguage_gap") or "none"

            st.session_state.chat_messages.append({"role": "user", "content": f"🎙️ *\"{transcript}\"* "})
            st.session_state.xp += 5
            st.session_state.chat_messages.append({"role": "assistant", "content": f"{msg}\n\n⭐ **+5 XP**", "auto_play": True})
            log_diagnostic(
                student_name, "A1", st.session_state.topic_label or "Español A1",
                f"Conversación Libre T{turn}", transcript, "PRODUCED", gap, "FREE_PRODUCTION", recast, msg
            )
            st.session_state.personal_turn += 1
            st.rerun()