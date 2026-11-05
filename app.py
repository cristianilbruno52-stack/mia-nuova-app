import io
import os
import re
import json
import html
import base64
import difflib
from datetime import datetime, timezone
from xml.sax.saxutils import escape

import streamlit as st
import streamlit.components.v1 as components
from google import genai
from pypdf import PdfReader
from docx import Document
from ebooklib import epub, ITEM_DOCUMENT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

# ------------------------------------------------------------
# CONFIGURAZIONE
# ------------------------------------------------------------
st.set_page_config(
    page_title="LIBRO MASTER AI",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
:root {
  --navy: #0a1f44;
  --navy-2: #14356b;
  --gold: #c9a227;
  --gold-light: #e6c766;
  --paper: #f4f6fb;
}
.stApp { background: var(--paper); }

/* ---------- SIDEBAR navy ---------- */
section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, var(--navy) 0%, var(--navy-2) 100%);
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span {
  color: #e8edf7 !important;
}
section[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"],
section[data-testid="stSidebar"] .stTextInput input {
  background: rgba(255,255,255,.08) !important;
  color: #e8edf7 !important;
  border: 1px solid var(--gold) !important;
}

/* ---------- HERO ---------- */
.hero {
  display: flex; align-items: center; justify-content: space-between; gap: 1.2rem;
  padding: 1.6rem 2rem; border-radius: 20px; color: white;
  background: linear-gradient(120deg, var(--navy) 0%, var(--navy-2) 60%, #1c4a8a 100%);
  border: 1px solid var(--gold);
  box-shadow: 0 12px 32px rgba(10,31,68,.28);
  margin-bottom: 1.2rem;
}
.hero h1 { color: var(--gold-light); margin: 0; font-size: 2.1rem; letter-spacing: .5px; }
.hero p { color: #cdd8ee; margin: .4rem 0 0 0; font-size: 1.02rem; }
.hero .hero-art { flex: 0 0 auto; }

/* ---------- BOTTONI ---------- */
.stButton > button {
  border-radius: 10px; border: 1px solid var(--gold);
  background: #ffffff; color: var(--navy); font-weight: 600;
}
.stButton > button:hover { background: var(--gold-light); color: var(--navy); border-color: var(--gold); }
.stButton > button[kind="primary"] {
  background: linear-gradient(120deg, var(--gold), var(--gold-light)) !important;
  color: var(--navy) !important; border: none !important;
  box-shadow: 0 4px 14px rgba(201,162,39,.4);
}
.stButton > button[kind="primary"]:hover { filter: brightness(1.05); }

/* ---------- METRICHE ---------- */
div[data-testid="stMetric"] {
  background: #ffffff; border: 1px solid #d8c48a; border-left: 5px solid var(--gold);
  border-radius: 12px; padding: 12px;
  box-shadow: 0 3px 10px rgba(10,31,68,.06);
}
div[data-testid="stMetric"] label { color: var(--navy-2) !important; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: var(--navy) !important; }

/* ---------- TAB ---------- */
div[data-testid="stTabs"] button[role="tab"] {
  color: var(--navy-2); font-weight: 600; border-radius: 10px 10px 0 0;
}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
  color: var(--navy); border-bottom: 3px solid var(--gold) !important;
}

/* ---------- TEXTAREA / INPUT ---------- */
div[data-testid="stTextArea"] textarea,
.stTextInput input {
  background: #ffffff !important;
  border: 1px solid #c9cfe0 !important;
  border-radius: 12px !important;
  color: var(--navy) !important;
}
div[data-testid="stTextArea"] textarea:focus,
.stTextInput input:focus { border-color: var(--gold) !important; }

/* ---------- EXPANDER ---------- */
div[data-testid="stExpander"] {
  background: #ffffff; border: 1px solid #d8c48a; border-radius: 12px;
}
div[data-testid="stExpander"] summary { color: var(--navy); font-weight: 600; }

/* ---------- DIVIDER / CAPTION ---------- */
hr { border-color: #d8c48a !important; }
.stCaption, .stMarkdown p { color: #4a5a7a; }
</style>
""", unsafe_allow_html=True)

EDITOR_KEY = "book_editor"
VERSIONS_DIR = "versions"
ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

# Tracce disponibili in ordine di preferenza: (nome file, MIME)
MUSIC_TRACKS = [
    ("suno_track.m4a", "audio/mp4", "Accoglienza Dinamica · Suno"),
    ("suno_track.mp3", "audio/mpeg", "Accoglienza Dinamica · Suno"),
    ("library_ambience.wav", "audio/wav", "Atmosfera da libreria"),
]

def get_music_file():
    """Restituisce (percorso, mime, etichetta) della prima traccia trovata."""
    for name, mime, label in MUSIC_TRACKS:
        path = os.path.join(ASSETS_DIR, name)
        if os.path.isfile(path):
            return path, mime, label
    return None, None, None

# ------------------------------------------------------------
# STATO
# ------------------------------------------------------------
def init_state():
    defaults = {
        "original_text": "",
        "working_text": "",
        "versions": [],
        "metadata": {},
        "filename": "libro",
        "last_upload": None,
        "google_api_key": "",
        EDITOR_KEY: "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_state()

def sync_editor():
    """Aggiorna la text_area dopo modifiche programmatiche del testo."""
    st.session_state[EDITOR_KEY] = st.session_state.working_text

# ------------------------------------------------------------
# PLAYER AUDIO — musica di sottofondo "atmosfera da libreria"
# ------------------------------------------------------------
def render_music_player():
    """
    Player audio con autoplay, loop e cursore del volume.
    L'audio è incorporato come data URI (funziona offline).
    Nota: i browser possono bloccare l'autoplay prima del primo clic
    dell'utente; in quel caso basta premere il pulsante play.
    """
    music_path, music_mime, music_label = get_music_file()
    if not music_path:
        st.info("Nessuna traccia trovata in `assets/`: aggiungi `suno_track.m4a` "
                "oppure esegui `python make_ambience.py`.")
        return
    with open(music_path, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode("ascii")
    components.html(f"""
<div style="background: linear-gradient(120deg,#0a1f44,#14356b);
            border: 1px solid #c9a227; border-radius: 14px; padding: 10px 14px;
            color: #e8edf7; font-family: sans-serif; font-size: 13px;
            box-shadow: 0 6px 18px rgba(10,31,68,.35);">
  <div style="display:flex; align-items:center; gap:10px;">
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" style="flex:0 0 auto;">
      <path d="M9 18V6l10-2v12" stroke="#e6c766" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="7" cy="18" r="2.6" fill="#c9a227"/>
      <circle cx="17" cy="16" r="2.6" fill="#c9a227"/>
    </svg>
    <span style="white-space:nowrap; color:#e6c766; font-weight:600;">{music_label}</span>
    <button id="mp-play" onclick="mpToggle()"
            style="background: linear-gradient(120deg,#c9a227,#e6c766); color:#0a1f44;
                   border:none; border-radius:8px;
                   height:30px; padding:0 12px; cursor:pointer; font-size:13px;
                   font-weight:bold; white-space:nowrap;">▶ Play</button>
    <span>🔈</span>
    <input id="mp-vol" type="range" min="0" max="100" value="30"
           oninput="mpSetVol(this.value)"
           style="flex:1; accent-color:#c9a227; cursor:pointer;">
    <span id="mp-vol-label" style="width:34px; text-align:right; color:#e6c766;">30%</span>
  </div>
  <audio id="mp-audio" src="data:{music_mime};base64,{b64}" loop autoplay></audio>
</div>
<script>
function mpSetBtn(playing) {{
  document.getElementById('mp-play').textContent = playing ? '⏸ Pausa' : '▶ Play';
}}
function mpToggle() {{
  var a = document.getElementById('mp-audio');
  if (a.paused) {{
    var p = a.play();
    if (p !== undefined) {{ p.then(function() {{ mpSetBtn(true); }}).catch(function() {{ mpSetBtn(false); }}); }}
    else {{ mpSetBtn(true); }}
  }} else {{ a.pause(); mpSetBtn(false); }}
}}
function mpSetVol(v) {{
  document.getElementById('mp-audio').volume = v / 100;
  document.getElementById('mp-vol-label').textContent = v + '%';
}}
var a = document.getElementById('mp-audio');
a.volume = 0.3;
a.addEventListener('play', function() {{ mpSetBtn(true); }});
a.addEventListener('pause', function() {{ mpSetBtn(false); }});
var p = a.play();
if (p !== undefined) {{ p.catch(function() {{ mpSetBtn(false); }}); }}
</script>
""", height=64)

# ------------------------------------------------------------
# CATTURA ERRORI
# ------------------------------------------------------------
def safe_call(func, *args, **kwargs):
    """Esegue una funzione restituendo (risultato, errore)."""
    try:
        return func(*args, **kwargs), None
    except Exception as error:  # noqa: BLE001
        return None, error

# ------------------------------------------------------------
# UTILITÀ FILE
# ------------------------------------------------------------
def extract_epub(uploaded_bytes: bytes) -> str:
    book = epub.read_epub(io.BytesIO(uploaded_bytes))
    sections = []
    for item in book.get_items_of_type(ITEM_DOCUMENT):
        raw = item.get_content().decode("utf-8", errors="ignore")
        raw = re.sub(r"<(script|style)\b.*?</\1\s*>", " ", raw, flags=re.I | re.S)
        raw = re.sub(r"</(p|div|h[1-6]|li|br|section)\s*>", "\n", raw, flags=re.I)
        raw = re.sub(r"<[^>]+>", " ", raw)
        clean = html.unescape(raw)
        clean = re.sub(r"[ \t]+", " ", clean)
        clean = re.sub(r"\n\s*\n+", "\n\n", clean)
        if clean.strip():
            sections.append(clean.strip())
    return "\n\n".join(sections)

def extract_document(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith((".txt", ".md")):
        return data.decode("utf-8-sig", errors="replace")

    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        return "\n\n".join((page.extract_text() or "") for page in reader.pages)

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text for cell in row.cells))
        return "\n\n".join(parts)

    if name.endswith(".epub"):
        return extract_epub(data)

    raise ValueError("Formato non supportato. Carica PDF, EPUB, DOCX, TXT o MD.")

def make_docx(text: str) -> bytes:
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def make_pdf(text: str) -> bytes:
    output = io.BytesIO()
    pdf = SimpleDocTemplate(
        output, pagesize=A4,
        rightMargin=20 * mm, leftMargin=20 * mm,
        topMargin=20 * mm, bottomMargin=20 * mm,
    )
    styles = getSampleStyleSheet()
    story = []
    for line in text.splitlines():
        if line.strip():
            story.append(Paragraph(escape(line), styles["BodyText"]))
        story.append(Spacer(1, 5))
    pdf.build(story)
    return output.getvalue()

def make_epub(text: str, title: str, author: str) -> bytes:
    book = epub.EpubBook()
    book.set_identifier(f"libromaster-{datetime.now(timezone.utc).timestamp()}")
    book.set_title(title or "Libro revisionato")
    book.set_language("it")
    book.add_author(author or "Autore")
    paragraphs = [
        f"<p>{html.escape(line)}</p>"
        for line in text.splitlines()
        if line.strip()
    ]
    content = (
        "<html><head><title></title></head><body>"
        + "\n".join(paragraphs)
        + "</body></html>"
    )
    chapter = epub.EpubHtml(
        title=title or "Testo", file_name="chapter.xhtml", lang="it"
    )
    chapter.content = content
    book.add_item(chapter)
    book.toc = (chapter,)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav", chapter]
    output = io.BytesIO()
    epub.write_epub(output, book)
    return output.getvalue()

# ------------------------------------------------------------
# AI (GOOGLE GEMINI)
# ------------------------------------------------------------
def get_api_key() -> str:
    """Chiave Gemini da sessione, Secrets o variabili d'ambiente."""
    session_key = st.session_state.get("google_api_key", "")
    if session_key:
        return session_key
    try:
        return (
            st.secrets.get("GOOGLE_API_KEY", "")
            or st.secrets.get("GEMINI_API_KEY", "")
            or os.getenv("GOOGLE_API_KEY", "")
            or os.getenv("GEMINI_API_KEY", "")
        )
    except Exception:
        return os.getenv("GOOGLE_API_KEY", "") or os.getenv("GEMINI_API_KEY", "")

def split_text(text: str, max_chars: int = 9000):
    """Divide il testo in blocchi senza alterare gli a capo originali."""
    paragraphs = text.split("\n")
    chunks, current = [], ""
    for paragraph in paragraphs:
        candidate = current + ("\n" if current else "") + paragraph
        if len(candidate) > max_chars and current:
            chunks.append(current)
            current = paragraph
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks or [""]

def ask_ai(prompt: str, system_prompt: str, model: str, temperature: float = 0.2) -> str:
    """
    Chiama un modello Gemini tramite il nuovo SDK google-genai.
    Il system prompt viene passato come system_instruction nella config.
    """
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError(
            "Chiave API non configurata. Inseriscila nella sidebar, oppure "
            "imposta GOOGLE_API_KEY nei Secrets di Streamlit o come variabile d'ambiente."
        )
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=genai.types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=temperature,
        ),
    )
    if not response.candidates:
        raise RuntimeError(
            "Risposta vuota o bloccata dal modello (nessun candidato disponibile)."
        )
    return response.text.strip()

def parse_json_response(raw: str) -> dict:
    """Estrae il JSON anche se l'AI lo racchiude in blocchi markdown."""
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", raw.strip(), flags=re.I)
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Nessun oggetto JSON valido trovato nella risposta.")
    return json.loads(cleaned[start:end + 1])

# ------------------------------------------------------------
# VERSIONI (con persistenza su disco, best-effort)
# ------------------------------------------------------------
def save_version(label: str):
    entry = {
        "label": label,
        "text": st.session_state.working_text,
        "created_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    st.session_state.versions.append(entry)
    try:
        os.makedirs(VERSIONS_DIR, exist_ok=True)
        fname = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + re.sub(
            r"[^\w\-]+", "_", label
        )[:40] + ".json"
        with open(os.path.join(VERSIONS_DIR, fname), "w", encoding="utf-8") as fh:
            json.dump(entry, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass  # la cronologia resta comunque in sessione

def load_versions_from_disk():
    """Ricarica la cronologia salvata, se presente e non già caricata."""
    if st.session_state.versions or not os.path.isdir(VERSIONS_DIR):
        return
    entries = []
    try:
        for fname in sorted(os.listdir(VERSIONS_DIR)):
            if fname.endswith(".json"):
                with open(os.path.join(VERSIONS_DIR, fname), encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict) and "text" in data:
                    entries.append({
                        "label": data.get("label", fname),
                        "text": data.get("text", ""),
                        "created_at": data.get("created_at", ""),
                    })
    except Exception:
        entries = []
    st.session_state.versions = entries

load_versions_from_disk()

def run_revision(text: str, task: str, language: str, intensity: str,
                 preserve_style: bool, model: str):
    pieces = split_text(text)
    output = []
    progress = st.progress(0, text="Preparazione revisione...")
    errors = []
    for index, piece in enumerate(pieces):
        system = (
            f"Sei un revisore editoriale professionista. Scrivi in {language}. "
            f"Intervento richiesto: {task}. Intensità: {intensity}. "
            + (
                "Conserva rigorosamente voce, tono, significato, fatti, trama, nomi propri "
                "e ordine degli eventi dell'autore. "
                if preserve_style else
                "Puoi riformulare per migliorare la chiarezza, senza inventare fatti o cambiare il significato. "
            )
            + "Correggi gli errori grammaticali, ortografici, sintattici e di punteggiatura pertinenti. "
            "Non riassumere, non eliminare parti, non aggiungere spiegazioni o commenti. "
            "Mantieni titoli, paragrafi, elenchi e interruzioni di riga. "
            "Restituisci esclusivamente il testo revisionato."
        )
        try:
            output.append(ask_ai(piece, system, model))
        except Exception as error:  # noqa: BLE001
            errors.append(f"Blocco {index + 1}: {error}")
            output.append(piece)  # conserva il testo originale del blocco fallito
        progress.progress(
            (index + 1) / len(pieces),
            text=f"Revisione blocco {index + 1}/{len(pieces)}",
        )
    progress.empty()
    if errors:
        st.warning(
            "Revisione parziale: alcuni blocchi non sono stati elaborati. "
            + "; ".join(errors)
        )
    return "\n".join(output)

# ------------------------------------------------------------
# INTERFACCIA
# ------------------------------------------------------------
st.markdown("""
<div class="hero">
  <div>
    <h1>📚 LIBRO MASTER AI</h1>
    <p>La tua redazione editoriale intelligente: correggi, migliora, organizza e prepara i tuoi libri.</p>
  </div>
  <div class="hero-art">
    <svg width="150" height="110" viewBox="0 0 150 110" fill="none" xmlns="http://www.w3.org/2000/svg">
      <!-- libro aperto -->
      <path d="M75 26 C58 14 32 12 14 19 L14 88 C32 81 58 83 75 95 C92 83 118 81 136 88 L136 19 C118 12 92 14 75 26 Z"
            fill="#12305e" stroke="#c9a227" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M75 26 L75 95" stroke="#c9a227" stroke-width="2.5"/>
      <!-- righe pagina sinistra -->
      <path d="M24 34 C38 30 56 31 68 38" stroke="#e6c766" stroke-width="2" stroke-linecap="round"/>
      <path d="M24 46 C38 42 56 43 68 50" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".85"/>
      <path d="M24 58 C38 54 56 55 68 62" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".7"/>
      <!-- righe pagina destra -->
      <path d="M82 38 C94 31 112 30 126 34" stroke="#e6c766" stroke-width="2" stroke-linecap="round"/>
      <path d="M82 50 C94 43 112 42 126 46" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".85"/>
      <path d="M82 62 C94 55 112 54 126 58" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".7"/>
      <!-- segnalibro dorato -->
      <path d="M118 14 L118 40 L112 35 L106 40 L106 16 Z" fill="#c9a227"/>
      <!-- stelle/sparkles -->
      <path d="M30 8 l2.2 4.6 4.6 2.2 -4.6 2.2 -2.2 4.6 -2.2 -4.6 -4.6 -2.2 4.6 -2.2 Z" fill="#e6c766"/>
      <path d="M132 60 l1.8 3.8 3.8 1.8 -3.8 1.8 -1.8 3.8 -1.8 -3.8 -3.8 -1.8 3.8 -1.8 Z" fill="#c9a227" opacity=".9"/>
    </svg>
  </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    music_on = st.toggle("🎵 Musica di sottofondo", value=True,
                         help="Avvia automaticamente la traccia Suno 'Accoglienza Dinamica' "
                              "(rilassante, atmosfera da libreria). "
                              "Se il browser blocca l'autoplay, premi ▶ Play nel player.")
    if music_on:
        render_music_player()
        st.caption("Il cursore nel player regola il volume in tempo reale.")
    st.divider()
    st.header("Impostazioni")
    model = st.selectbox(
        "Modello AI (Gemini)",
        ["gemini-2.5-flash-lite", "gemini-2.5-flash", "gemini-2.5-pro"],
        index=1,
    )
    language = st.selectbox("Lingua", ["italiano", "inglese", "francese", "spagnolo", "tedesco"])
    intensity = st.select_slider(
        "Intensità revisione", ["Leggera", "Standard", "Approfondita"], value="Standard"
    )
    preserve_style = st.toggle("Mantieni stile e voce dell'autore", value=True)

    st.divider()
    st.subheader("Chiave API Google (Gemini)")
    if get_api_key():
        st.success("Chiave AI configurata")
    else:
        st.warning("Chiave AI non configurata")
    key_input = st.text_input(
        "GOOGLE_API_KEY (lascia vuoto per usare Secrets/env)",
        type="password",
        help="Ottienila su https://aistudio.google.com/apikey",
    )
    if key_input:
        st.session_state.google_api_key = key_input
        st.success("Chiave salvata per questa sessione")

tab_editor, tab_seo, tab_history, tab_export = st.tabs([
    "✍️ Editor e correzione", "🔎 Titoli e SEO", "🕘 Versioni", "⬇️ Esporta"
])

with tab_editor:
    st.subheader("Importa il tuo libro")
    uploaded = st.file_uploader(
        "Carica un documento",
        type=["pdf", "epub", "docx", "txt", "md"],
        help="PDF con testo selezionabile, EPUB, DOCX, TXT o Markdown.",
    )
    import_col, restore_col = st.columns(2)
    with import_col:
        if st.button("Importa file", type="primary", disabled=uploaded is None):
            extracted, error = safe_call(extract_document, uploaded)
            if error:
                st.error(f"Errore durante l'importazione: {error}")
            elif not extracted.strip():
                st.error("Non è stato trovato testo. Se il PDF è una scansione, serve OCR.")
            else:
                st.session_state.original_text = extracted
                st.session_state.working_text = extracted
                st.session_state.filename = os.path.splitext(uploaded.name)[0]
                st.session_state.last_upload = uploaded.name
                save_version("Importazione originale")
                sync_editor()
                st.success(f"Importato: {uploaded.name} ({len(extracted):,} caratteri)")
    with restore_col:
        if st.button("Ripristina originale", disabled=not bool(st.session_state.original_text)):
            st.session_state.working_text = st.session_state.original_text
            save_version("Ripristino originale")
            sync_editor()
            st.success("Testo originale ripristinato.")

    st.divider()
    st.subheader("Correzione intelligente")
    task = st.selectbox("Scegli il tipo di revisione", [
        "Correzione completa",
        "Solo grammatica e ortografia",
        "Punteggiatura e sintassi",
        "Chiarezza e leggibilità",
        "Uniformità di nomi, date e terminologia",
        "Revisione editoriale approfondita",
    ])
    if st.button("✨ Avvia correzione AI", type="primary",
                 disabled=not bool(st.session_state.working_text)):
        if not get_api_key():
            st.error("Configura prima la chiave API di Google nella sidebar.")
        else:
            save_version("Backup prima della revisione")
            with st.spinner("L'AI sta revisionando il testo..."):
                revised, error = safe_call(
                    run_revision,
                    st.session_state.working_text, task, language,
                    intensity, preserve_style, model,
                )
            if error:
                st.error(f"Revisione non completata: {error}")
            else:
                st.session_state.working_text = revised
                save_version(f"Revisione AI: {task}")
                sync_editor()
                st.success(
                    "Revisione terminata. Controlla sempre il risultato prima della pubblicazione."
                )

    st.divider()
    st.subheader("Testo completo")
    edited = st.text_area(
        "Puoi modificare il testo direttamente qui",
        height=480,
        key=EDITOR_KEY,
        placeholder="Incolla qui il testo del tuo libro oppure importa un file...",
    )
    if edited != st.session_state.working_text:
        st.session_state.working_text = edited

    words = len(re.findall(r"\b[\w'-]+\b", st.session_state.working_text, flags=re.UNICODE))
    chapters = len(re.findall(
        r"(?im)^\s*(capitolo|chapter|parte|prologo|epilogo)\b",
        st.session_state.working_text
    ))
    reading_minutes = max(1, round(words / 200))  # ~200 parole/minuto
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Caratteri", f"{len(st.session_state.working_text):,}")
    m2.metric("Parole", f"{words:,}")
    m3.metric("Capitoli rilevati", chapters)
    m4.metric("Lettura stimata", f"~{reading_minutes} min")

    if st.session_state.original_text and st.session_state.original_text != st.session_state.working_text:
        with st.expander("🔍 Confronta con il testo originale (diff)"):
            diff_lines = list(difflib.unified_diff(
                st.session_state.original_text.splitlines(),
                st.session_state.working_text.splitlines(),
                fromfile="originale",
                tofile="revisionato",
                lineterm="",
            ))
            st.code("\n".join(diff_lines[:400]) +
                    ("\n… (diff troncato)" if len(diff_lines) > 400 else ""),
                    language="diff")

with tab_seo:
    st.subheader("Metadati editoriali e SEO")
    metadata = st.session_state.metadata
    title = st.text_input("Titolo", value=metadata.get("titolo_input", ""))
    subtitle = st.text_input("Sottotitolo", value=metadata.get("sottotitolo_input", ""))
    author = st.text_input("Autore", value=metadata.get("autore_input", ""))
    genre = st.text_input("Genere / categoria", value=metadata.get("genere_input", ""))
    audience = st.text_input("Pubblico di riferimento", value=metadata.get("pubblico_input", ""))

    if st.button("Genera descrizione, keyword e tag SEO", type="primary",
                 disabled=not bool(st.session_state.working_text)):
        if not get_api_key():
            st.error("Configura prima la chiave API di Google nella sidebar.")
        else:
            sample = st.session_state.working_text[:14000]
            prompt = f"""
Analizza il testo del libro e crea metadati editoriali accurati.
Titolo attuale: {title}
Sottotitolo: {subtitle}
Autore: {author}
Genere: {genre}
Pubblico: {audience}

TESTO:
{sample}

Restituisci ESCLUSIVAMENTE un oggetto JSON valido in italiano con queste chiavi:
titolo_suggerito, sottotitolo_suggerito, descrizione_breve,
descrizione_lunga, meta_title (massimo 60 caratteri),
meta_description (massimo 160 caratteri),
keyword_principali (array di 7 stringhe), tag (array di 12 stringhe),
categorie_suggerite (array di 3 stringhe).
Non inventare contenuti, premi, recensioni o caratteristiche non presenti nel testo.
"""
            with st.spinner("Generazione metadati in corso..."):
                raw, error = safe_call(
                    ask_ai, prompt,
                    "Sei un consulente editoriale ed esperto SEO. Rispondi solo con JSON valido, senza markdown.",
                    model,
                )
            if error:
                st.error(f"Non è stato possibile generare i metadati: {error}")
            else:
                parsed, error = safe_call(parse_json_response, raw)
                if error:
                    st.error(f"La risposta dell'AI non è un JSON valido: {error}")
                else:
                    parsed.update({
                        "titolo_input": title,
                        "sottotitolo_input": subtitle,
                        "autore_input": author,
                        "genere_input": genre,
                        "pubblico_input": audience,
                    })
                    st.session_state.metadata = parsed
                    st.success("Metadati generati.")

    if st.session_state.metadata:
        st.json(st.session_state.metadata)
        st.download_button(
            "Scarica metadati JSON",
            data=json.dumps(st.session_state.metadata, ensure_ascii=False, indent=2),
            file_name="metadati_libro_seo.json",
            mime="application/json",
        )

with tab_history:
    st.subheader("Cronologia e ripristino")
    st.caption(
        "Le versioni sono salvate anche su disco (cartella `versions/`) e ricaricate "
        "al riavvio dell'app, quando possibile."
    )
    if not st.session_state.versions:
        st.info("La cronologia si popola quando importi un file o avvii una revisione.")
    else:
        for reverse_index, version in enumerate(reversed(st.session_state.versions)):
            version_index = len(st.session_state.versions) - 1 - reverse_index
            with st.expander(f"{version['label']} — {version['created_at']}"):
                st.caption(f"{len(version['text']):,} caratteri")
                st.text(version["text"][:1500] + ("…" if len(version["text"]) > 1500 else ""))
                if st.button("Ripristina questa versione", key=f"restore_version_{version_index}"):
                    st.session_state.working_text = version["text"]
                    save_version(f"Ripristino: {version['label']}")
                    sync_editor()
                    st.success("Versione ripristinata. Torna alla scheda Editor.")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Salva snapshot manuale", disabled=not bool(st.session_state.working_text)):
                save_version("Snapshot manuale")
                st.success("Snapshot salvato.")
        with col_b:
            st.download_button(
                "Scarica tutta la cronologia (JSON)",
                data=json.dumps(st.session_state.versions, ensure_ascii=False, indent=2),
                file_name="cronologia_versioni.json",
                mime="application/json",
            )

with tab_export:
    st.subheader("Esporta il tuo libro")
    text_to_export = st.session_state.working_text
    base_name = st.session_state.filename or "libro_revisionato"
    if not text_to_export.strip():
        st.info("Importa o incolla un testo per attivare l'esportazione.")
    else:
        st.download_button(
            "Scarica TXT",
            data=text_to_export.encode("utf-8"),
            file_name=f"{base_name}_revisionato.txt",
            mime="text/plain",
        )
        st.download_button(
            "Scarica Markdown",
            data=text_to_export.encode("utf-8"),
            file_name=f"{base_name}_revisionato.md",
            mime="text/markdown",
        )
        docx_bytes, docx_err = safe_call(make_docx, text_to_export)
        pdf_bytes, pdf_err = safe_call(make_pdf, text_to_export)
        st.download_button(
            "Scarica Word (DOCX)",
            data=docx_bytes or b"",
            file_name=f"{base_name}_revisionato.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            disabled=docx_err is not None,
        )
        st.download_button(
            "Scarica PDF",
            data=pdf_bytes or b"",
            file_name=f"{base_name}_revisionato.pdf",
            mime="application/pdf",
            disabled=pdf_err is not None,
        )
        if docx_err:
            st.error(f"Errore DOCX: {docx_err}")
        if pdf_err:
            st.error(f"Errore PDF: {pdf_err}")
        export_title = st.text_input(
            "Titolo per EPUB",
            value=st.session_state.metadata.get("titolo_suggerito")
                  or st.session_state.metadata.get("titolo_input")
                  or base_name,
        )
        export_author = st.text_input(
            "Autore per EPUB",
            value=st.session_state.metadata.get("autore_input", ""),
        )
        epub_bytes, epub_err = safe_call(make_epub, text_to_export, export_title, export_author)
        st.download_button(
            "Scarica EPUB",
            data=epub_bytes or b"",
            file_name=f"{base_name}_revisionato.epub",
            mime="application/epub+zip",
            disabled=epub_err is not None,
        )
        if epub_err:
            st.error(f"Errore EPUB: {epub_err}")

st.divider()
st.caption(
    "LIBRO MASTER AI — I PDF scannerizzati richiedono OCR. "
    "L'importazione e l'esportazione ricostruiscono il testo e non conservano necessariamente "
    "impaginazione, immagini, note o grafica originale. Controlla sempre il testo revisionato."
)
