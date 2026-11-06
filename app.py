# -*- coding: utf-8 -*-
"""
LIBRO MASTER AI  —  versione potenziata
=======================================
Principio guida: l'app NON deve mai "indovinare" in silenzio.
 - Pagine reali (PDF) distinte da pagine stimate (parole / parole-per-pagina).
 - Capitoli rilevati dal testo, con pagina di inizio reale nei PDF e controllo
   della numerazione (buchi, doppioni).
 - Ogni file importato viene verificato (impronta SHA-256, pagine vuote/scansioni,
   codifica, immagini perse, coerenza indice/capitoli).
 - Ogni revisione AI è validata blocco per blocco (lunghezza, numeri, commenti
   dell'AI, risposte troncate) e il risultato è registrato in un log.
 - Pulsante "Ripulisci tutto da zero" nella sidebar.
"""
import io
import os
import re
import json
import html
import time
import base64
import shutil
import hashlib
import zipfile
import difflib
import copy
from bisect import bisect_right
from collections import Counter
from datetime import datetime, timezone
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from google.genai import types
from pypdf import PdfReader
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph as DocxParagraph
from ebooklib import epub, ITEM_DOCUMENT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak

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
.stButton > button, .stDownloadButton > button {
  border-radius: 10px; border: 1px solid var(--gold);
  background: #ffffff; color: var(--navy); font-weight: 600;
}
.stButton > button:hover, .stDownloadButton > button:hover {
  background: var(--gold-light); color: var(--navy); border-color: var(--gold);
}
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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERSIONS_DIR = os.path.join(BASE_DIR, "versions")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
MAX_VERSIONS = 40
DEFAULT_WORDS_PER_PAGE = 250

LANG_CODES = {"italiano": "it", "inglese": "en", "francese": "fr",
              "spagnolo": "es", "tedesco": "de"}

# Tracce disponibili in ordine di preferenza: (nome file, MIME, etichetta)
MUSIC_TRACKS = [
    ("suno_track.m4a", "audio/mp4", "Accoglienza Dinamica · Suno"),
    ("suno_track.mp3", "audio/mpeg", "Accoglienza Dinamica · Suno"),
    ("library_ambience.wav", "audio/wav", "Atmosfera da libreria"),
]

WORD_RE = re.compile(r"\w+(?:['’-]\w+)*", re.UNICODE)
CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
INVISIBLE_RE = re.compile(r"[\u00ad\u200b\u200c\u200d\u2060\ufeff]")
XML_BAD_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
PAGE_NUM_LINE = re.compile(
    r"[ \t]*(?:(?:pag(?:ina|\.)?|page|p\.)[ \t]*)?[-–—]?[ \t]*\d{1,4}[ \t]*[-–—]?[ \t]*"
    r"(?:(?:di|of|/)[ \t]*\d{1,4})?[ \t]*", re.I)

_NUM = r"(?:\d{1,3}|[ivxlcdm]{1,8}|[a-zàèéìòù]{3,12})"
CHAPTER_RE = re.compile(
    r"^\s*(?:"
    rf"(?:capitolo|cap\.|chapter|parte|part|libro|book|sezione|section)\s+{_NUM}(?:\s*[:.\-–—]\s*.{{0,60}})?"
    r"|(?:prologo|epilogo|prefazione|introduzione|premessa|conclusioni|conclusione|postfazione|"
    r"appendice|ringraziamenti|prologue|epilogue|preface|introduction|conclusion|afterword|"
    r"appendix|acknowledg(?:e)?ments)(?:\s*[:.\-–—]?\s*.{0,50})?"
    r"|\#{1,2}\s+\S.{0,80}"
    r")\s*$",
    re.I,
)
ROMAN_RE = re.compile(r"^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")

STOPWORDS = set("""
come anche della delle dello degli dalla dalle nella nelle negli sono essere aveva erano loro questo
questa quello quella quando dove cosa tutto tutti tutta ancora molto dopo prima perché però così solo
senza sulla suoi hanno stato stata fatto fare ogni sempre mentre dentro verso anni quel quei quelle
quegli nelle nello alla alle allo agli sul sulle sulla suo sua sue mio mia miei mie tuo tua tuoi
che cui chi era ero stavo stava fosse poi più già dalla dallo dagli lei lui noi voi essi
with that this have from they will would there their what about which when your were been
""".split())

# ------------------------------------------------------------
# STATO
# ------------------------------------------------------------
DEFAULTS = {
    "original_text": "",
    "versions": [],
    "versions_loaded": False,
    "metadata": {},
    "filename": "libro",
    "last_upload": None,
    "google_api_key": "",
    EDITOR_KEY: "",
    "pending_text": None,
    "flash": [],
    "doc_info": {},
    "revision_log": [],
    "grammar": None,
    "uploader_gen": 0,
    "seo_title": "",
    "seo_subtitle": "",
    "seo_author": "",
    "seo_genre": "",
    "seo_audience": "",
}


def init_state():
    for key, value in DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = copy.deepcopy(value)


init_state()

# Un testo impostato da codice va applicato PRIMA che la text_area venga creata
# (Streamlit vieta di modificare lo stato di un widget già istanziato).
if st.session_state.pending_text is not None:
    st.session_state[EDITOR_KEY] = st.session_state.pending_text
    st.session_state.pending_text = None


def get_text() -> str:
    return st.session_state.get(EDITOR_KEY, "") or ""


def flash(msg: str, kind: str = "success"):
    st.session_state.flash.append((msg, kind))


def show_flashes():
    icons = {"success": "✅", "warning": "⚠️", "error": "❌", "info": "ℹ️"}
    for msg, kind in st.session_state.flash:
        st.toast(msg, icon=icons.get(kind, "ℹ️"))
    st.session_state.flash = []


def apply_text(new_text: str, label: str, notify: str = "", kind: str = "success"):
    """Salva una versione, imposta il testo nell'editor e ricarica la pagina."""
    save_version(label, new_text)
    st.session_state.pending_text = new_text
    if notify:
        flash(notify, kind)
    st.rerun()


def safe_call(func, *args, **kwargs):
    """Esegue una funzione restituendo (risultato, errore)."""
    try:
        return func(*args, **kwargs), None
    except Exception as error:  # noqa: BLE001
        return None, error


def count_words(text: str) -> int:
    return len(WORD_RE.findall(text))


# ------------------------------------------------------------
# PLAYER AUDIO — musica di sottofondo "atmosfera da libreria"
# ------------------------------------------------------------
def get_music_file():
    for name, mime, label in MUSIC_TRACKS:
        path = os.path.join(ASSETS_DIR, name)
        if os.path.isfile(path):
            return path, mime, label
    return None, None, None


def render_music_player():
    """Player con autoplay, loop e volume. L'audio è un data URI (funziona offline)."""
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
# VERSIONI (persistenza su disco, best-effort)
# ------------------------------------------------------------
def save_version(label: str, text: str):
    entry = {
        "label": label,
        "text": text,
        "created_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    st.session_state.versions.append(entry)
    if len(st.session_state.versions) > MAX_VERSIONS:
        del st.session_state.versions[:-MAX_VERSIONS]
    try:
        os.makedirs(VERSIONS_DIR, exist_ok=True)
        fname = datetime.now().strftime("%Y%m%d_%H%M%S_%f") + "_" + re.sub(
            r"[^\w\-]+", "_", label)[:40] + ".json"
        with open(os.path.join(VERSIONS_DIR, fname), "w", encoding="utf-8") as fh:
            json.dump(entry, fh, ensure_ascii=False)
        files = sorted(f for f in os.listdir(VERSIONS_DIR) if f.endswith(".json"))
        for old in files[:-MAX_VERSIONS]:
            os.remove(os.path.join(VERSIONS_DIR, old))
    except Exception:  # noqa: BLE001
        pass  # la cronologia resta comunque in sessione


def load_versions_from_disk():
    if st.session_state.versions_loaded:
        return
    st.session_state.versions_loaded = True
    if st.session_state.versions or not os.path.isdir(VERSIONS_DIR):
        return
    entries = []
    try:
        for fname in sorted(os.listdir(VERSIONS_DIR))[-MAX_VERSIONS:]:
            if fname.endswith(".json"):
                with open(os.path.join(VERSIONS_DIR, fname), encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict) and "text" in data:
                    entries.append({
                        "label": data.get("label", fname),
                        "text": data.get("text", ""),
                        "created_at": data.get("created_at", ""),
                    })
    except Exception:  # noqa: BLE001
        entries = []
    st.session_state.versions = entries


load_versions_from_disk()


# ------------------------------------------------------------
# RESET TOTALE
# ------------------------------------------------------------
def reset_everything(delete_disk: bool, keep_key: bool):
    """Riporta l'app allo stato iniziale (eseguita come callback del pulsante)."""
    key = st.session_state.get("google_api_key", "") if keep_key else ""
    gen = st.session_state.get("uploader_gen", 0) + 1
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    if delete_disk:
        shutil.rmtree(VERSIONS_DIR, ignore_errors=True)
    try:
        st.cache_data.clear()
    except Exception:  # noqa: BLE001
        pass
    init_state()
    st.session_state.google_api_key = key
    st.session_state.uploader_gen = gen          # svuota anche il file_uploader
    st.session_state.versions_loaded = True      # non ricaricare la cronologia dal disco
    flash("Pagina ripulita: tutto è ripartito da zero.", "success")


# ------------------------------------------------------------
# ESTRAZIONE DOCUMENTI (con verifica)
# ------------------------------------------------------------
def decode_text(data: bytes):
    try:
        return data.decode("utf-8-sig"), "UTF-8"
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace"), "Windows-1252 (fallback, verifica gli accenti)"


def extract_pdf(data: bytes):
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        try:
            ok = reader.decrypt("")
        except Exception:  # noqa: BLE001
            ok = 0
        if not ok:
            raise ValueError("Il PDF è protetto da password: impossibile leggerlo.")
    parts, offsets, table, empty = [], [], [], []
    pos = 0
    for number, page in enumerate(reader.pages, start=1):
        err = ""
        try:
            page_text = page.extract_text() or ""
        except Exception as exc:  # noqa: BLE001
            page_text, err = "", str(exc)[:80]
        offsets.append(pos)
        parts.append(page_text)
        pos += len(page_text) + 2          # separatore "\n\n"
        words = count_words(page_text)
        if words < 5:
            status = "vuota / immagine"
            empty.append(number)
        elif words < 30:
            status = "poco testo"
        else:
            status = "ok"
        table.append({"pagina": number, "parole": words,
                      "caratteri": len(page_text.strip()), "stato": status,
                      "errore": err})
    meta = {}
    try:
        for k, v in (reader.metadata or {}).items():
            meta[str(k).lstrip("/")] = str(v)[:120]
    except Exception:  # noqa: BLE001
        meta = {}
    info = {
        "pagine_reali": len(reader.pages),
        "pagine_vuote": empty,
        "tabella_pagine": table,
        "offset_pagine": offsets,
        "metadati_pdf": meta,
    }
    return "\n\n".join(parts), info


def docx_declared_pages(data: bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            xml = z.read("docProps/app.xml").decode("utf-8", "ignore")
        m = re.search(r"<Pages>(\d+)</Pages>", xml)
        return int(m.group(1)) if m else None
    except Exception:  # noqa: BLE001
        return None


def extract_docx(data: bytes):
    doc = Document(io.BytesIO(data))
    parts, headings = [], []
    for child in doc.element.body.iterchildren():     # mantiene l'ordine reale
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            para = DocxParagraph(child, doc)
            if para.text.strip():
                parts.append(para.text)
                try:
                    style = para.style.name if para.style is not None else ""
                except Exception:  # noqa: BLE001
                    style = ""
                if re.fullmatch(r"(?i)(heading 1|titolo 1)", (style or "").strip()):
                    headings.append(para.text.strip())
        elif tag == "tbl":
            table = Table(child, doc)
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    info = {
        "tabelle": len(doc.tables),
        "immagini": len(doc.inline_shapes),
        "titoli_heading1": headings,
        "pagine_dichiarate": docx_declared_pages(data),
    }
    return "\n\n".join(parts), info


def _flatten_toc(toc):
    out = []
    for entry in toc:
        if isinstance(entry, tuple):
            out.append(getattr(entry[0], "title", "") or "")
            out.extend(_flatten_toc(entry[1]))
        elif isinstance(entry, (list,)):
            out.extend(_flatten_toc(entry))
        else:
            out.append(getattr(entry, "title", "") or "")
    return out


def extract_epub(data: bytes):
    book = epub.read_epub(io.BytesIO(data))
    sections, docs = [], 0
    for item in book.get_items_of_type(ITEM_DOCUMENT):
        if isinstance(item, epub.EpubNav) or item.get_name().lower().endswith("nav.xhtml"):
            continue  # l'indice di navigazione duplicherebbe i titoli dei capitoli
        docs += 1
        raw = item.get_content().decode("utf-8", errors="ignore")
        raw = re.sub(r"<(script|style)\b.*?</\1\s*>", " ", raw, flags=re.I | re.S)
        raw = re.sub(r"</(p|div|h[1-6]|li|br|section)\s*>", "\n", raw, flags=re.I)
        raw = re.sub(r"<[^>]+>", " ", raw)
        clean = html.unescape(raw)
        clean = re.sub(r"[ \t]+", " ", clean)
        clean = re.sub(r"\n\s*\n+", "\n\n", clean)
        if clean.strip():
            sections.append(clean.strip())
    toc_titles = [t.strip() for t in _flatten_toc(book.toc) if t and t.strip()]
    info = {"file_documento": docs, "voci_indice": len(toc_titles),
            "titoli_indice": toc_titles[:300]}
    return "\n\n".join(sections), info


def extract_document(uploaded_file):
    """Restituisce (testo, info). `info` contiene i fatti verificabili sul file."""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    info = {
        "nome": uploaded_file.name,
        "dimensione_byte": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "importato_il": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    if name.endswith((".txt", ".md")):
        text, enc = decode_text(data)
        info.update(formato="Markdown" if name.endswith(".md") else "TXT", codifica=enc)
    elif name.endswith(".pdf"):
        text, extra = extract_pdf(data)
        info.update(formato="PDF", **extra)
    elif name.endswith(".docx"):
        text, extra = extract_docx(data)
        info.update(formato="DOCX", **extra)
    elif name.endswith(".epub"):
        text, extra = extract_epub(data)
        info.update(formato="EPUB", **extra)
    else:
        raise ValueError("Formato non supportato. Carica PDF, EPUB, DOCX, TXT o MD.")
    info["parole"] = count_words(text)
    info["caratteri"] = len(text)
    return text, info


# ------------------------------------------------------------
# CAPITOLI
# ------------------------------------------------------------
def roman_to_int(s: str):
    s = s.upper()
    if not s or not ROMAN_RE.match(s):
        return None
    vals = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    total = 0
    for i, ch in enumerate(s):
        v = vals[ch]
        total += -v if i + 1 < len(s) and vals[s[i + 1]] > v else v
    return total


def chapter_number(title: str):
    m = re.match(r"^\s*(?:capitolo|cap\.|chapter)\s+(\d{1,3}|[ivxlcdm]{1,8})\b", title, re.I)
    if not m:
        return None
    v = m.group(1)
    return int(v) if v.isdigit() else roman_to_int(v)


def detect_chapters(text: str, extra_pattern: str = ""):
    """Righe che sembrano intestazioni di capitolo, con offset e numero di riga."""
    extra = None
    if extra_pattern.strip():
        try:
            extra = re.compile(extra_pattern, re.I)
        except re.error:
            extra = None
    chapters, pos = [], 0
    for line_no, line in enumerate(text.split("\n")):
        s = line.strip()
        if s and len(s) <= 80 and not s.endswith((",", ";")) and (
                CHAPTER_RE.match(line) or (extra is not None and extra.search(s))):
            chapters.append({"title": s, "start": pos, "line": line_no})
        pos += len(line) + 1
    return chapters


def chapter_spans(text: str, chapters):
    """Intervalli di testo per capitolo (con eventuale testo iniziale)."""
    if not chapters:
        return []
    bounds = [c["start"] for c in chapters]
    spans = []
    if text[:bounds[0]].strip():
        spans.append({"title": "(testo prima del primo capitolo)", "start": 0,
                      "end": bounds[0], "front": True})
    for i, c in enumerate(chapters):
        end = bounds[i + 1] if i + 1 < len(bounds) else len(text)
        spans.append({"title": c["title"], "start": c["start"], "end": end,
                      "front": False, "line": c["line"]})
    return spans


def chapter_table(text: str, chapters, wpp: int, page_offsets=None):
    rows, n = [], 0
    for sp in chapter_spans(text, chapters):
        words = count_words(text[sp["start"]:sp["end"]])
        if not sp["front"]:
            n += 1
        row = {"#": "—" if sp["front"] else str(n), "Titolo": sp["title"], "Parole": words,
               "Pagine stimate": round(words / max(1, wpp), 1)}
        if page_offsets:
            row["Pagina inizio (PDF reale)"] = bisect_right(page_offsets, sp["start"])
            row["Pagina fine (PDF reale)"] = bisect_right(page_offsets, max(sp["start"], sp["end"] - 1))
        rows.append(row)
    return rows


# ------------------------------------------------------------
# DIAGNOSI LOCALE (nessuna AI: numeri verificabili)
# ------------------------------------------------------------
def repeated_lines(text: str, min_count: int = 6):
    counts = Counter()
    for line in text.split("\n"):
        s = line.strip()
        if 8 <= len(s) <= 80 and not s.endswith((".", "!", "?", "…", "»", "”", '"', "-", ",")) \
                and not CHAPTER_RE.match(s):
            counts[re.sub(r"\d+", "#", s.lower())] += 1
    return {k: v for k, v in counts.items() if v >= min_count}


@st.cache_data(show_spinner=False, max_entries=6)
def find_issues(text: str, pages: int = 0):
    min_rep = max(4, int(pages * 0.3)) if pages else 6
    lines = text.split("\n")
    return {
        "Spazi doppi": len(re.findall(r"[ \t]{2,}", text)),
        "Parole spezzate a fine riga (trattino)": len(re.findall(r"[a-zà-ÿ]-\n[ \t]*[a-zà-ÿ]", text)),
        "Righe con solo un numero (numeri di pagina?)": sum(1 for l in lines if l.strip() and PAGE_NUM_LINE.fullmatch(l)),
        "Righe ripetute (intestazioni/piè di pagina?)": len(repeated_lines(text, min_rep)),
        "Spazio prima della punteggiatura": len(re.findall(r"(?<=\w) +(?:[,;:!?]|\.(?!\.))", text)),
        "Caratteri anomali (�, controllo, invisibili)": text.count("\ufffd") + len(CONTROL_RE.findall(text)) + len(INVISIBLE_RE.findall(text)),
        "Parole ripetute consecutive (es. «che che»)": len(re.findall(r"\b(\w{3,})\s+\1\b", text, re.I)),
        "Tre o più righe vuote consecutive": len(re.findall(r"\n{4,}", text)),
        "Righe lunghissime (> 2000 caratteri)": sum(1 for l in lines if len(l) > 2000),
    }


def gulpease_label(g: float) -> str:
    if g >= 80:
        return "molto facile"
    if g >= 60:
        return "facile / medio"
    if g >= 40:
        return "difficile"
    return "molto difficile"


@st.cache_data(show_spinner=False, max_entries=6)
def text_stats(text: str):
    words = WORD_RE.findall(text)
    n_words = len(words)
    letters = sum(ch.isalpha() for ch in text)
    sentences = [s for s in re.split(r"(?<=[.!?…])\s+", text) if s.strip()]
    paragraphs = [p for p in re.split(r"\n\s*\n", text) if p.strip()]
    lower = [w.lower() for w in words]
    gulp = 89 + (300 * len(sentences) - 10 * letters) / max(1, n_words)
    top = Counter(w for w in lower if len(w) >= 4 and w not in STOPWORDS).most_common(15)
    longest = sorted(sentences, key=len, reverse=True)[:5]
    return {
        "parole": n_words,
        "frasi": len(sentences),
        "paragrafi": len(paragraphs),
        "parole_per_frase": round(n_words / max(1, len(sentences)), 1),
        "gulpease": round(gulp, 1),
        "varieta_lessicale": round(len(set(lower)) / max(1, n_words), 3),
        "top": top,
        "frasi_lunghe": [s[:400] for s in longest],
    }


def proper_nouns(text: str) -> Counter:
    c = Counter()
    for sent in re.split(r"(?<=[.!?…:])\s+|\n+", text):
        toks = re.findall(r"[A-Za-zÀ-ÿ]+", sent)
        for tok in toks[1:]:
            if len(tok) >= 3 and tok[0].isupper() and not tok.isupper():
                c[tok] += 1
    return c


def integrity_report(orig: str, cur: str):
    on = Counter(re.findall(r"\d+(?:[.,]\d+)*", orig))
    cn = Counter(re.findall(r"\d+(?:[.,]\d+)*", cur))
    op, cp = proper_nouns(orig), proper_nouns(cur)
    gone = [n for n, _ in op.most_common() if n not in cp]
    new = [n for n, _ in cp.most_common() if n not in op]
    ow, cw = count_words(orig), count_words(cur)
    return {
        "parole_originale": ow,
        "parole_attuale": cw,
        "variazione_parole_%": round((cw - ow) / max(1, ow) * 100, 1),
        "numeri_scomparsi": sorted((on - cn).elements())[:60],
        "numeri_nuovi": sorted((cn - on).elements())[:60],
        "nomi_scomparsi": gone[:40],
        "nomi_nuovi": new[:40],
    }


def build_checks(text: str, info: dict, chapters, wpp: int):
    """Elenco di (livello, messaggio) con soli fatti misurati sul file/testo."""
    if not text.strip():
        return [("err", "Nessun testo disponibile da verificare.")]
    checks = []
    add = lambda lvl, msg: checks.append((lvl, msg))  # noqa: E731
    words = count_words(text)
    fmt = info.get("formato", "Testo incollato")
    pages_for_headers = 0

    # ---- pagine ----
    if fmt == "PDF":
        total = info["pagine_reali"]
        pages_for_headers = total
        empty = info["pagine_vuote"]
        with_text = total - len(empty)
        add("ok", f"Il PDF contiene realmente **{total} pagine**; **{with_text}** hanno testo estraibile.")
        if empty:
            ratio = len(empty) / max(1, total)
            sample = ", ".join(map(str, empty[:15])) + ("…" if len(empty) > 15 else "")
            add("err" if ratio > 0.3 else "warn",
                f"{len(empty)} pagine ({ratio:.0%}) non hanno testo estraibile (pagine: {sample}). "
                "Possono essere immagini, pagine bianche o scansioni: in quel caso serve OCR "
                "e il testo importato è **incompleto**.")
        avg = words / max(1, with_text)
        add("info", f"Media misurata: {avg:.0f} parole per pagina con testo (la stima usa {wpp}). "
                    f"Pagine stimate dal testo: {words / wpp:.0f} contro {total} reali.")
    elif fmt == "DOCX":
        dp = info.get("pagine_dichiarate")
        if dp:
            add("info", f"Word dichiara {dp} pagine nel file (valore salvato da Word, può essere "
                        f"obsoleto). Stima dal testo: ~{words / wpp:.0f}.")
        else:
            add("info", f"Il DOCX non riporta il numero di pagine. Stima dal testo: ~{words / wpp:.0f}.")
        if info.get("immagini"):
            add("warn", f"{info['immagini']} immagini nel DOCX: non vengono importate (solo testo).")
        if info.get("tabelle"):
            add("info", f"{info['tabelle']} tabelle importate come righe di testo separate da « | ».")
    elif fmt == "EPUB":
        add("info", f"L'EPUB è a flusso e non ha pagine fisiche: pagine stimate ~{words / wpp:.0f} "
                    f"(a {wpp} parole/pagina). File di contenuto letti: {info.get('file_documento', '?')}.")
    elif fmt in ("TXT", "Markdown"):
        add("info", f"File di testo: nessuna pagina reale. Pagine stimate ~{words / wpp:.0f}.")
        if "fallback" in info.get("codifica", ""):
            add("warn", "Il file non era UTF-8: è stato letto come Windows-1252. Controlla accenti e apostrofi.")
    else:
        add("info", f"Testo nell'editor: pagine stimate ~{words / wpp:.0f} (a {wpp} parole/pagina).")

    if words < 2000:
        add("warn", f"Solo {words} parole (~{words / wpp:.0f} pagine): è breve per un libro. "
                    "Controlla di aver importato il file completo.")

    bad = text.count("\ufffd")
    if bad:
        add("warn", f"{bad} caratteri non decodificabili («�»): il testo contiene simboli corrotti.")

    # ---- capitoli ----
    if not chapters:
        add("warn", "Nessun capitolo riconosciuto automaticamente (cerco «Capitolo N», «Parte», «Prologo», "
                    "«Introduzione», titoli Markdown…). Se il libro ha un altro formato, usa il "
                    "*pattern personalizzato* nella sidebar.")
    else:
        add("ok", f"**{len(chapters)} intestazioni di capitolo/sezione** riconosciute nel testo.")
        nums = [n for n in (chapter_number(c["title"]) for c in chapters) if n is not None]
        if len(nums) >= 3:
            missing = sorted(set(range(min(nums), max(nums) + 1)) - set(nums))
            dups = sorted(n for n, k in Counter(nums).items() if k > 1)
            if missing:
                add("warn", f"Numerazione con buchi: mancano i capitoli {', '.join(map(str, missing[:15]))}"
                            f"{'…' if len(missing) > 15 else ''} (testo perso o titoli non riconosciuti).")
            if dups:
                add("warn", f"Numeri di capitolo ripetuti: {', '.join(map(str, dups[:15]))} "
                            "(forse intestazioni correnti in testa alle pagine).")
            if not missing and not dups:
                add("ok", f"Numerazione coerente: capitoli {min(nums)}–{max(nums)} senza buchi né doppioni.")
        titles = Counter(c["title"].lower() for c in chapters)
        repeated = [t for t, k in titles.items() if k > 1]
        if repeated:
            add("warn", f"Titoli identici ripetuti: {', '.join(repr(t) for t in repeated[:6])}.")
        spans = [sp for sp in chapter_spans(text, chapters) if not sp["front"]]
        sizes = [count_words(text[sp["start"]:sp["end"]]) for sp in spans]
        if len(sizes) >= 3 and max(sizes) > 0.5 * sum(sizes):
            add("warn", "Un solo capitolo contiene oltre metà del libro: probabilmente altre intestazioni non sono state riconosciute.")
        tiny = [sp["title"] for sp, w in zip(spans, sizes) if w < 50]
        if tiny:
            add("warn", f"Capitoli con meno di 50 parole: {', '.join(repr(t[:30]) for t in tiny[:6])}.")
        if fmt == "DOCX" and info.get("titoli_heading1") is not None:
            h = len(info["titoli_heading1"])
            if h != len(chapters):
                add("info", f"Il DOCX ha {h} titoli «Heading 1» ma il testo ne mostra {len(chapters)} "
                            "come intestazioni: i due conteggi non coincidono.")
        if fmt == "EPUB" and info.get("voci_indice") is not None:
            if info["voci_indice"] != len(chapters):
                add("info", f"L'indice dell'EPUB ha {info['voci_indice']} voci, il testo {len(chapters)} intestazioni riconosciute.")

    # ---- qualità del testo estratto ----
    iss = find_issues(text, pages_for_headers)
    if iss["Parole spezzate a fine riga (trattino)"] > 20:
        add("warn", f"{iss['Parole spezzate a fine riga (trattino)']} parole spezzate a fine riga: usa «Pulizia automatica».")
    if iss["Righe ripetute (intestazioni/piè di pagina?)"] > 0:
        add("warn", f"{iss['Righe ripetute (intestazioni/piè di pagina?)']} righe che si ripetono (intestazioni/piè di pagina): "
                    "vanno rimosse con «Pulizia automatica».")
    if iss["Righe con solo un numero (numeri di pagina?)"] > 5:
        add("info", f"{iss['Righe con solo un numero (numeri di pagina?)']} righe contengono solo un numero (probabili numeri di pagina).")
    return checks


def report_markdown(info, checks, rows_orig, stats, integ, wpp):
    icon = {"ok": "✅", "warn": "⚠️", "err": "❌", "info": "ℹ️"}
    L = [f"# Rapporto di verifica — {info.get('nome', 'testo incollato')}",
         f"_Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')}_", ""]
    if info:
        L += ["## File", f"- Formato: {info.get('formato')}",
              f"- Dimensione: {info.get('dimensione_byte', 0):,} byte",
              f"- SHA-256: `{info.get('sha256')}`"]
        if info.get("pagine_reali"):
            L.append(f"- Pagine reali: {info['pagine_reali']}")
        L.append("")
    L += ["## Esito dei controlli"]
    L += [f"- {icon.get(l, '•')} {m.replace('**', '')}" for l, m in checks]
    L += ["", f"## Numeri del testo (pagine stimate a {wpp} parole/pagina)",
          f"- Parole: {stats['parole']:,}", f"- Frasi: {stats['frasi']:,}",
          f"- Paragrafi: {stats['paragrafi']:,}",
          f"- Parole per frase: {stats['parole_per_frase']}",
          f"- Indice Gulpease (italiano): {stats['gulpease']} ({gulpease_label(stats['gulpease'])})", ""]
    if rows_orig:
        L += ["## Capitoli", "| " + " | ".join(rows_orig[0].keys()) + " |",
              "|" + "---|" * len(rows_orig[0])]
        L += ["| " + " | ".join(str(v) for v in r.values()) + " |" for r in rows_orig]
        L.append("")
    if integ:
        L += ["## Integrità rispetto all'originale",
              f"- Parole: {integ['parole_originale']:,} → {integ['parole_attuale']:,} "
              f"({integ['variazione_parole_%']:+}%)",
              f"- Numeri scomparsi: {', '.join(integ['numeri_scomparsi']) or 'nessuno'}",
              f"- Numeri nuovi: {', '.join(integ['numeri_nuovi']) or 'nessuno'}",
              f"- Nomi propri scomparsi: {', '.join(integ['nomi_scomparsi']) or 'nessuno'}",
              f"- Nomi propri nuovi: {', '.join(integ['nomi_nuovi']) or 'nessuno'}", ""]
    L += ["---", "_Limiti: la verifica misura il testo estratto dal file. Non può sapere se il "
          "PDF contiene immagini con testo, né giudicare la qualità letteraria._"]
    return "\n".join(L)


# ------------------------------------------------------------
# PULIZIA LOCALE
# ------------------------------------------------------------
CLEAN_OPTS = [
    ("ctrl", "Rimuovi caratteri invisibili e di controllo", True),
    ("pagenum", "Rimuovi le righe che contengono solo il numero di pagina", True),
    ("headers", "Rimuovi intestazioni / piè di pagina ripetuti", True),
    ("dehyph", "Ricongiungi le parole spezzate a fine riga (para-⏎mola)", True),
    ("reflow", "Ricompatta le righe spezzate a metà frase (tipico dei PDF)", False),
    ("spaces", "Elimina spazi doppi e spazi a fine riga", True),
    ("punct", "Togli lo spazio prima di , ; : ! ? . (non per il francese)", True),
    ("ellipsis", "Uniforma i puntini di sospensione (...) in …", False),
    ("blank", "Riduci le righe vuote multiple a una sola", True),
]


def reflow_lines(text: str):
    out, buf, joined = [], "", 0
    for line in text.split("\n"):
        s = line.strip()
        if not s:
            if buf:
                out.append(buf)
                buf = ""
            out.append("")
            continue
        if buf and not re.search(r"[.!?…:;»”\"')\]]$", buf) \
                and not CHAPTER_RE.match(buf) and s[0].islower():
            buf += " " + s
            joined += 1
        else:
            if buf:
                out.append(buf)
            buf = s
    if buf:
        out.append(buf)
    return "\n".join(out), joined


def clean_text(text: str, opts: dict, min_repeat: int = 6):
    rep = {}
    t = text.replace("\r\n", "\n").replace("\r", "\n")
    if opts.get("ctrl"):
        t, n1 = re.subn(r"(\w)\u00ad\n(\w)", r"\1\2", t)
        t, n2 = INVISIBLE_RE.subn("", t)
        t, n3 = CONTROL_RE.subn("", t)
        rep["Caratteri invisibili / di controllo rimossi"] = n1 + n2 + n3
    if opts.get("pagenum"):
        lines = t.split("\n")
        kept = [l for l in lines if not (l.strip() and PAGE_NUM_LINE.fullmatch(l))]
        rep["Righe «numero di pagina» rimosse"] = len(lines) - len(kept)
        t = "\n".join(kept)
    if opts.get("headers"):
        repeated = repeated_lines(t, min_repeat)
        lines = t.split("\n")
        kept = [l for l in lines
                if not (l.strip() and re.sub(r"\d+", "#", l.strip().lower()) in repeated)]
        rep["Intestazioni / piè di pagina rimossi"] = len(lines) - len(kept)
        t = "\n".join(kept)
    if opts.get("dehyph"):
        t, n = re.subn(r"([a-zà-ÿ])-\n[ \t]*([a-zà-ÿ])", r"\1\2", t)
        rep["Parole ricongiunte"] = n
    if opts.get("reflow"):
        t, n = reflow_lines(t)
        rep["Righe ricompattate"] = n
    if opts.get("spaces"):
        t, n1 = re.subn(r"[ \t]{2,}", " ", t)
        t, n2 = re.subn(r"[ \t]+\n", "\n", t)
        rep["Spazi superflui eliminati"] = n1 + n2
    if opts.get("punct"):
        t, n = re.subn(r"(?<=\w) +([,;:!?]|\.(?!\.))", r"\1", t)
        rep["Spazi prima della punteggiatura tolti"] = n
    if opts.get("ellipsis"):
        t, n = re.subn(r"\.{3,}", "…", t)
        rep["Puntini uniformati"] = n
    if opts.get("blank"):
        t, n = re.subn(r"\n{3,}", "\n\n", t)
        rep["Gruppi di righe vuote ridotti"] = n
    return t.strip(), rep


# ------------------------------------------------------------
# AI (GOOGLE GEMINI)
# ------------------------------------------------------------
def get_api_key() -> str:
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
    except Exception:  # noqa: BLE001
        return os.getenv("GOOGLE_API_KEY", "") or os.getenv("GEMINI_API_KEY", "")


@st.cache_resource(show_spinner=False)
def get_client(api_key: str):
    return genai.Client(api_key=api_key)


def split_text(text: str, max_chars: int = 9000):
    """Divide il testo in blocchi di righe intere (rimettendoli insieme con \\n si
    riottiene esattamente l'originale; una riga > max_chars viene spezzata su uno spazio)."""
    lines = []
    for line in text.split("\n"):
        while len(line) > max_chars:
            cut = line.rfind(" ", 0, max_chars)
            cut = cut if cut > max_chars // 2 else max_chars
            lines.append(line[:cut])
            line = line[cut:]
        lines.append(line)
    chunks, current = [], ""
    for line in lines:
        candidate = current + ("\n" if current else "") + line
        if len(candidate) > max_chars and current:
            chunks.append(current)
            current = line
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks or [""]


def ask_ai(prompt: str, system_prompt: str, model: str, temperature: float = 0.2) -> str:
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError(
            "Chiave API non configurata. Inseriscila nella sidebar, oppure imposta "
            "GOOGLE_API_KEY nei Secrets di Streamlit o come variabile d'ambiente.")
    client = get_client(api_key)
    response = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt, temperature=temperature),
            )
            break
        except Exception as exc:  # noqa: BLE001
            msg = str(exc).upper()
            transient = any(k in msg for k in (
                "429", "503", "UNAVAILABLE", "RESOURCE_EXHAUSTED", "TIMEOUT", "DEADLINE"))
            if attempt < 2 and transient:
                time.sleep(3 * (attempt + 1))
                continue
            raise
    candidates = getattr(response, "candidates", None)
    if not candidates:
        raise RuntimeError("Risposta vuota o bloccata dal modello (nessun candidato).")
    reason = str(getattr(candidates[0], "finish_reason", "") or "")
    if reason and "STOP" not in reason.upper():
        raise RuntimeError(f"Risposta incompleta o bloccata (motivo: {reason}).")
    out = (response.text or "").strip()
    if not out:
        raise RuntimeError("Il modello ha restituito un testo vuoto.")
    return out


def parse_json_response(raw: str) -> dict:
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", raw.strip(), flags=re.I)
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Nessun oggetto JSON valido trovato nella risposta.")
    data = json.loads(cleaned[start:end + 1])
    if not isinstance(data, dict):
        raise ValueError("Il JSON ricevuto non è un oggetto.")
    return data


def strip_fences(text: str) -> str:
    return re.sub(r"^\s*```[a-z]*\s*\n|\n\s*```\s*$", "", text.strip(), flags=re.I)


def validate_block(src: str, out: str, intensity: str):
    lo, hi = {"Leggera": (0.92, 1.08), "Standard": (0.85, 1.15),
              "Approfondita": (0.75, 1.30)}[intensity]
    notes = []
    ratio = len(out) / max(1, len(src))
    if not lo <= ratio <= hi:
        notes.append(f"lunghezza {ratio:.0%} dell'originale (atteso {lo:.0%}–{hi:.0%})")
    if Counter(re.findall(r"\d+", src)) != Counter(re.findall(r"\d+", out)):
        notes.append("numeri modificati")
    if re.match(r"\s*(ecco (il|qui|la|di seguito)|certo[,!]|sure[,!]|here is|here's)", out, re.I):
        notes.append("possibile commento introduttivo dell'AI")
    if "```" in out:
        notes.append("contiene blocchi markdown")
    return ratio, notes


def build_system_prompt(task, language, intensity, preserve_style):
    return (
        f"Sei un revisore editoriale professionista. Scrivi in {language}. "
        f"Intervento richiesto: {task}. Intensità: {intensity}. "
        + ("Conserva rigorosamente voce, tono, significato, fatti, trama, nomi propri "
           "e ordine degli eventi dell'autore. " if preserve_style else
           "Puoi riformulare per migliorare la chiarezza, senza inventare fatti o cambiare il significato. ")
        + "Correggi gli errori grammaticali, ortografici, sintattici e di punteggiatura pertinenti. "
          "Non modificare numeri, date, cifre e nomi propri. Non tradurre. "
          "Non riassumere, non eliminare parti, non aggiungere spiegazioni o commenti. "
          "Mantieni titoli, paragrafi, elenchi e interruzioni di riga. "
          "Restituisci esclusivamente il testo revisionato."
    )


def run_revision(segments, task, language, intensity, preserve_style, model, strict):
    """Revisiona una lista di segmenti. Restituisce (segmenti_nuovi, log)."""
    system = build_system_prompt(task, language, intensity, preserve_style)
    plan = [(si, piece) for si, seg in enumerate(segments) for piece in split_text(seg)]
    outs = [[] for _ in segments]
    log = []
    bar = st.progress(0, text="Preparazione revisione...")
    total = max(1, len(plan))
    for n, (si, piece) in enumerate(plan, start=1):
        if not piece.strip():
            outs[si].append(piece)
            bar.progress(n / total, text=f"Revisione blocco {n}/{total}")
            continue
        ratio, notes, status, final = None, [], "", piece
        try:
            raw = ask_ai(piece, system, model)
            lead = piece[:len(piece) - len(piece.lstrip())]
            trail = piece[len(piece.rstrip()):]
            candidate = lead + strip_fences(raw).strip() + trail
            ratio, notes = validate_block(piece, candidate, intensity)
            if notes and strict:
                status = "scartato (originale conservato)"
            else:
                status = "accettato con avviso" if notes else "accettato"
                final = candidate
        except Exception as exc:  # noqa: BLE001
            status = "errore (originale conservato)"
            notes = [str(exc)[:220]]
        outs[si].append(final)
        log.append({"blocco": n, "stato": status,
                    "lunghezza_%": None if ratio is None else round(ratio * 100),
                    "note": "; ".join(notes)})
        bar.progress(n / total, text=f"Revisione blocco {n}/{total}")
    bar.empty()
    return ["\n".join(o) for o in outs], log


def parse_json_list(raw: str) -> list:
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", raw.strip(), flags=re.I)
    start, end = cleaned.find("["), cleaned.rfind("]")
    if start == -1 or end == -1 or end < start:
        raise ValueError("Nessuna lista JSON nella risposta.")
    data = json.loads(cleaned[start:end + 1])
    if not isinstance(data, list):
        raise ValueError("Il JSON ricevuto non è una lista.")
    return [d for d in data if isinstance(d, dict)]


# ---- regole locali (italiano): funzionano senza AI e senza chiave API ----
WORD_FIXES = {
    "perchè": "perché", "poichè": "poiché", "affinchè": "affinché", "benchè": "benché",
    "finchè": "finché", "cosicchè": "cosicché", "giacchè": "giacché", "pultroppo": "purtroppo",
    "sopratutto": "soprattutto", "propio": "proprio", "daccordo": "d'accordo",
    "accellerare": "accelerare", "eccezzione": "eccezione", "qual'è": "qual è",
    "quì": "qui", "quà": "qua", "fà": "fa", "stà": "sta", "sù": "su", "nò": "no",
    "pò": "po'", "sè": "sé",
}
_PARTICIPLES = (r"(?:fatto|detto|visto|preso|avuto|dato|mangiato|scritto|letto|capito|"
                r"voluto|potuto|saputo|portato)")
_WORD_FIX_RE = re.compile(
    r"(?<![\w'’])(" + "|".join(re.escape(k) for k in sorted(WORD_FIXES, key=len, reverse=True)) + r")(?!\w)",
    re.I)
_BAD_ELISION_RE = re.compile(r"\b(un|l|dell|nell|all|sull|quest|quell|d|dov|com)['’] +(?=[a-zà-ÿ])", re.I)
_REPEAT_OK = {"bene", "piano", "così", "forte", "presto", "adagio", "poco", "giù", "sempre", "tanto", "pian"}


def _case_like(src: str, fix: str) -> str:
    if len(src) > 1 and src.isupper():
        return fix.upper()
    return fix[0].upper() + fix[1:] if src[:1].isupper() else fix


def local_grammar_check(text: str, limit: int = 3000):
    """Controlli deterministici: ogni voce cita testo realmente presente, con posizione esatta."""
    found = []

    def add(m, fix, tipo, spieg, auto=True):
        found.append({"pos": m.start(), "errore": m.group(0), "correzione": fix, "tipo": tipo,
                      "spiegazione": spieg, "auto": auto, "fonte": "regole"})

    for m in _WORD_FIX_RE.finditer(text):
        add(m, _case_like(m.group(1), WORD_FIXES[m.group(1).lower()]), "ortografia", "Forma errata (accento/apostrofo/grafia).")
    for m in re.finditer(r"\bun po(?![\w'’])", text, re.I):
        add(m, m.group(0) + "'", "ortografia", "«un po'» è un troncamento: richiede l'apostrofo.")
    for m in re.finditer(rf"\b(lui|lei|egli|ella|che|chi|non|mai|già)\s+a\s+({_PARTICIPLES})\b", text, re.I):
        add(m, f"{m.group(1)} ha {m.group(2)}", "grammatica", "Verbo avere: «ha», non la preposizione «a».")
    for m in re.finditer(rf"\b(io|non|che|già|mai)\s+o\s+({_PARTICIPLES})\b", text, re.I):
        add(m, f"{m.group(1)} ho {m.group(2)}", "grammatica", "Verbo avere: «ho», non la congiunzione «o».")
    for m in re.finditer(r"\b(lui|lei|egli|ella|non|questo|quello|tutto|ciò|che)\s+e\s+"
                         r"(stato|stata|stati|state|vero|vera|giusto|giusta|possibile|necessario|bello|bella)\b", text, re.I):
        add(m, f"{m.group(1)} è {m.group(2)}", "grammatica", "Verbo essere: «è», non la congiunzione «e».")
    for m in _BAD_ELISION_RE.finditer(text):
        add(m, m.group(0).rstrip(), "punteggiatura", "Dopo l'apostrofo di elisione non va lo spazio.")
    for m in re.finditer(r"[A-Za-zÀ-ÿ]{2,}[,;:][A-Za-zÀ-ÿ]{2,}", text):
        pu = re.search(r"[,;:]", m.group(0))
        add(m, m.group(0)[:pu.end()] + " " + m.group(0)[pu.end():], "punteggiatura", "Manca lo spazio dopo la punteggiatura.")
    for m in re.finditer(r"[a-zà-ÿ]{3,}\.[A-ZÀ-Ý][a-zà-ÿ]{2,}", text):
        if not re.search(r"www|http|\.(?:com|it|org|net)\b", m.group(0), re.I):
            d = m.group(0).index(".")
            add(m, m.group(0)[:d + 1] + " " + m.group(0)[d + 1:], "punteggiatura", "Manca lo spazio dopo il punto.")
    for m in re.finditer(r"[,;]{2,}", text):
        add(m, m.group(0)[0], "punteggiatura", "Punteggiatura doppia.")
    for m in re.finditer(r"(?<![.])\.\.(?![.])", text):
        add(m, ".", "punteggiatura", "Due punti consecutivi: punto singolo o puntini di sospensione (…).")
    for m in re.finditer(r"(?<=\w) +([,;:!?]|\.(?!\.))", text):
        add(m, m.group(1), "punteggiatura", "Spazio prima della punteggiatura (in italiano non va).")
    for m in re.finditer(r"\b(\w{3,})\s+\1\b", text, re.I):
        if m.group(1).lower() not in _REPEAT_OK:
            add(m, m.group(1), "grammatica", "Parola ripetuta di seguito (controlla che non sia voluto).")
    pos = 0
    for para in re.split(r"(\n\s*\n)", text):
        for o, c, nome in (("(", ")", "parentesi"), ("«", "»", "virgolette caporali")):
            if para.count(o) != para.count(c):
                exc = para.strip()[:60]
                if exc:
                    found.append({"pos": pos + para.index(exc), "errore": exc, "correzione": "—",
                                  "tipo": "punteggiatura",
                                  "spiegazione": f"{nome.capitalize()} non bilanciate in questo paragrafo: da verificare a mano.",
                                  "auto": False, "fonte": "regole"})
        pos += len(para)
    found.sort(key=lambda d: d["pos"])
    out, last_end = [], -1
    for it in found:
        if it["pos"] >= last_end or not it["auto"]:
            out.append(it)
            last_end = max(last_end, it["pos"] + len(it["errore"])) if it["auto"] else last_end
    return out[:limit]


def merge_items(ai_items, local_items):
    merged = [dict(i, auto=True, fonte="AI") for i in ai_items]
    for it in local_items:
        a, b = it["pos"], it["pos"] + len(it["errore"])
        if not any(a < o["pos"] + len(o["errore"]) and o["pos"] < b for o in merged):
            merged.append(it)
    merged.sort(key=lambda d: d["pos"])
    return merged



def check_grammar(text: str, model: str, language: str, block_chars: int = 6000):
    """Segnala errori SENZA modificare il testo. Ogni segnalazione deve citare
    una frase realmente presente nel blocco: le altre vengono scartate e contate."""
    system = (
        f"Sei un correttore di bozze esperto di {language}. Trova SOLO errori oggettivi di "
        "grammatica, ortografia, concordanza, punteggiatura e sintassi. Ignora scelte di stile, "
        "dialetti, forme volutamente colloquiali e nomi propri. Per ogni errore copia in 'errore' "
        "la porzione di testo ESATTA (max 80 caratteri, identica all'originale). Rispondi "
        "esclusivamente con una lista JSON di oggetti con chiavi: errore, correzione, tipo "
        "(grammatica|ortografia|concordanza|punteggiatura|sintassi), spiegazione (max 15 parole). "
        "Se non ci sono errori rispondi []. Nessun markdown, nessun commento."
    )
    chunks = split_text(text, block_chars)
    items, failed, discarded, seen, approx = [], [], 0, set(), 0
    bar = st.progress(0, text="Controllo grammaticale...")
    for n, chunk in enumerate(chunks, start=1):
        if chunk.strip():
            try:
                for it in parse_json_list(ask_ai(chunk, system, model, 0.0)):
                    err = str(it.get("errore", "")).strip()
                    fix = str(it.get("correzione", "")).strip()
                    if not err or not fix or err == fix or err not in chunk:
                        discarded += 1
                        continue
                    pos = text.find(err, max(0, approx - 300))
                    if pos == -1:
                        pos = text.find(err)
                    if pos == -1 or (pos, err) in seen:
                        discarded += 1
                        continue
                    seen.add((pos, err))
                    items.append({"pos": pos, "errore": err, "correzione": fix,
                                  "tipo": str(it.get("tipo", ""))[:30],
                                  "spiegazione": str(it.get("spiegazione", ""))[:160],
                                  "auto": True, "fonte": "AI"})
            except Exception as exc:  # noqa: BLE001
                failed.append(f"Blocco {n}: {str(exc)[:150]}")
        approx += len(chunk) + 1
        bar.progress(n / len(chunks), text=f"Controllo blocco {n}/{len(chunks)}")
    bar.empty()
    items.sort(key=lambda d: d["pos"])
    return {"sig": hashlib.sha1(text.encode("utf-8", "ignore")).hexdigest(),
            "items": items, "blocks": len(chunks), "failed": failed, "discarded": discarded}


# ------------------------------------------------------------
# ESPORTAZIONE
# ------------------------------------------------------------
def sanitize(s: str) -> str:
    return XML_BAD_RE.sub("", s)


def build_docx(text, title, author, extra):
    doc = Document()
    if title:
        doc.add_heading(sanitize(title), 0)
    heading_lines = {c["line"] for c in detect_chapters(text, extra)}
    for i, line in enumerate(text.split("\n")):
        s = sanitize(line).strip()
        if not s:
            continue
        if i in heading_lines:
            doc.add_heading(s.lstrip("# ").strip(), level=1)
        else:
            doc.add_paragraph(s)
    doc.core_properties.title = title or ""
    doc.core_properties.author = author or ""
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def build_pdf(text, title, author, extra):
    out = io.BytesIO()

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(A4[0] / 2, 10 * mm, str(doc_.page))
        canvas.restoreState()

    pdf = SimpleDocTemplate(out, pagesize=A4, title=title or "", author=author or "",
                            rightMargin=20 * mm, leftMargin=20 * mm,
                            topMargin=20 * mm, bottomMargin=20 * mm)
    styles = getSampleStyleSheet()
    story = []
    if title:
        story += [Paragraph(escape(sanitize(title)), styles["Title"]), Spacer(1, 12)]
    heading_lines = {c["line"] for c in detect_chapters(text, extra)}
    first_heading = True
    for i, line in enumerate(text.split("\n")):
        s = sanitize(line).strip()
        if not s:
            continue
        if i in heading_lines:
            if not first_heading or story:
                story.append(PageBreak())
            first_heading = False
            story.append(Paragraph(escape(s.lstrip("# ").strip()), styles["Heading1"]))
        else:
            story.append(Paragraph(escape(s), styles["BodyText"]))
            story.append(Spacer(1, 4))
    if not story:
        story.append(Paragraph(" ", styles["BodyText"]))
    pdf.build(story, onFirstPage=footer, onLaterPages=footer)
    return out.getvalue()


def build_epub(text, title, author, lang, extra):
    book = epub.EpubBook()
    book.set_identifier(f"libromaster-{datetime.now(timezone.utc).timestamp()}")
    book.set_title(title or "Libro revisionato")
    book.set_language(lang)
    book.add_author(author or "Autore")
    chapters = detect_chapters(text, extra)
    spans = chapter_spans(text, chapters) or [
        {"title": title or "Testo", "start": 0, "end": len(text), "front": True}]
    items = []
    for i, sp in enumerate(spans, start=1):
        seg = text[sp["start"]:sp["end"]]
        lines = [l for l in seg.split("\n") if l.strip()]
        body = []
        for j, l in enumerate(lines):
            clean = html.escape(sanitize(l.strip()))
            if j == 0 and not sp["front"]:
                body.append(f"<h1>{html.escape(sp['title'].lstrip('# ').strip())}</h1>")
            else:
                body.append(f"<p>{clean}</p>")
        name = "Inizio" if sp["front"] and chapters else sp["title"][:60]
        ch = epub.EpubHtml(title=name, file_name=f"cap_{i:03d}.xhtml", lang=lang)
        ch.content = "\n".join(body) or "<p> </p>"
        book.add_item(ch)
        items.append(ch)
    book.toc = tuple(items)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav"] + items
    out = io.BytesIO()
    epub.write_epub(out, book)
    return out.getvalue()


# ============================================================
#                          INTERFACCIA
# ============================================================
st.markdown("""
<div class="hero">
  <div>
    <h1>📚 LIBRO MASTER AI</h1>
    <p>Redazione editoriale intelligente: verifica, correggi, pulisci, organizza ed esporta i tuoi libri — con numeri reali, non promesse.</p>
  </div>
  <div class="hero-art">
    <svg width="150" height="110" viewBox="0 0 150 110" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M75 26 C58 14 32 12 14 19 L14 88 C32 81 58 83 75 95 C92 83 118 81 136 88 L136 19 C118 12 92 14 75 26 Z"
            fill="#12305e" stroke="#c9a227" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M75 26 L75 95" stroke="#c9a227" stroke-width="2.5"/>
      <path d="M24 34 C38 30 56 31 68 38" stroke="#e6c766" stroke-width="2" stroke-linecap="round"/>
      <path d="M24 46 C38 42 56 43 68 50" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".85"/>
      <path d="M24 58 C38 54 56 55 68 62" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".7"/>
      <path d="M82 38 C94 31 112 30 126 34" stroke="#e6c766" stroke-width="2" stroke-linecap="round"/>
      <path d="M82 50 C94 43 112 42 126 46" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".85"/>
      <path d="M82 62 C94 55 112 54 126 58" stroke="#e6c766" stroke-width="2" stroke-linecap="round" opacity=".7"/>
      <path d="M118 14 L118 40 L112 35 L106 40 L106 16 Z" fill="#c9a227"/>
      <path d="M30 8 l2.2 4.6 4.6 2.2 -4.6 2.2 -2.2 4.6 -2.2 -4.6 -4.6 -2.2 4.6 -2.2 Z" fill="#e6c766"/>
      <path d="M132 60 l1.8 3.8 3.8 1.8 -3.8 1.8 -1.8 3.8 -1.8 -3.8 -3.8 -1.8 3.8 -1.8 Z" fill="#c9a227" opacity=".9"/>
    </svg>
  </div>
</div>
""", unsafe_allow_html=True)

show_flashes()

# ---------------------------- SIDEBAR ----------------------------
with st.sidebar:
    music_on = st.toggle("🎵 Musica di sottofondo", value=True,
                         help="Avvia automaticamente la traccia Suno 'Accoglienza Dinamica'. "
                              "Se il browser blocca l'autoplay, premi ▶ Play nel player.")
    if music_on:
        render_music_player()
        st.caption("Il cursore nel player regola il volume in tempo reale.")
    st.divider()
    st.header("Impostazioni")
    model = st.selectbox(
        "Modello AI (Gemini)",
        ["gemini-2.5-flash-lite", "gemini-2.5-flash", "gemini-2.5-pro"], index=1)
    language = st.selectbox("Lingua", list(LANG_CODES.keys()))
    intensity = st.select_slider(
        "Intensità revisione", ["Leggera", "Standard", "Approfondita"], value="Standard")
    preserve_style = st.toggle("Mantieni stile e voce dell'autore", value=True)

    st.divider()
    st.subheader("Verità e controlli")
    wpp = st.number_input("Parole per pagina (per le pagine STIMATE)", min_value=100,
                          max_value=600, value=DEFAULT_WORDS_PER_PAGE, step=10,
                          help="Un libro tascabile ha circa 250–300 parole per pagina. "
                               "Serve solo per EPUB/DOCX/TXT; nei PDF le pagine reali sono contate dal file.")
    strict = st.toggle("Protezione anti-alterazione AI", value=True,
                       help="Scarta i blocchi in cui l'AI cambia numeri, accorcia/allunga troppo "
                            "o aggiunge commenti: viene conservato il testo originale.")
    custom_pat = st.text_input("Pattern capitoli personalizzato (regex, opzionale)",
                               placeholder=r"es.  ^ANNO \d+")
    if custom_pat:
        try:
            re.compile(custom_pat)
        except re.error as exc:
            st.error(f"Regex non valida: {exc}")
            custom_pat = ""

    st.divider()
    st.subheader("Chiave API Google (Gemini)")
    if get_api_key():
        st.success("Chiave AI configurata")
    else:
        st.warning("Chiave AI non configurata")
    key_input = st.text_input(
        "GOOGLE_API_KEY (lascia vuoto per usare Secrets/env)", type="password",
        help="Ottienila su https://aistudio.google.com/apikey")
    if key_input:
        st.session_state.google_api_key = key_input
        st.success("Chiave salvata per questa sessione")
    if st.button("Testa connessione AI", disabled=not get_api_key()):
        reply, err = safe_call(ask_ai, "Rispondi solo con la parola OK.",
                               "Rispondi in modo brevissimo.", model)
        if err:
            st.error(f"Connessione fallita: {err}")
        else:
            st.success(f"Connessione riuscita ({model}).")

    st.divider()
    st.subheader("🧽 Ripulisci tutto da zero")
    st.caption("Cancella testo, originale, cronologia, metadati, report e svuota il caricamento file.")
    wipe_disk = st.checkbox("Elimina anche le versioni salvate su disco", value=True)
    keep_key = st.checkbox("Mantieni la chiave API", value=True)
    confirm_reset = st.checkbox("Confermo: voglio ripartire da zero")
    st.button("🗑️ RIPULISCI TUTTO", disabled=not confirm_reset, on_click=reset_everything,
              args=(wipe_disk, keep_key), type="primary")

# Valori condivisi da tutti i tab
text_now = get_text()
chapters_now = detect_chapters(text_now, custom_pat)
info = st.session_state.doc_info
original = st.session_state.original_text

tab_editor, tab_verify, tab_tools, tab_ai, tab_gram, tab_seo, tab_history, tab_export = st.tabs([
    "✍️ Importa ed Editor", "🧪 Verifica e verità", "🧹 Pulizia e strumenti",
    "✨ Revisione AI", "📝 Errori grammaticali", "🔎 Titoli e SEO", "🕘 Versioni", "⬇️ Esporta",
])

# ============================ EDITOR ============================
with tab_editor:
    st.subheader("Importa il tuo libro")
    uploaded = st.file_uploader(
        "Carica un documento", type=["pdf", "epub", "docx", "txt", "md"],
        key=f"uploader_{st.session_state.uploader_gen}",
        help="PDF con testo selezionabile, EPUB, DOCX, TXT o Markdown. "
             "Ogni file viene verificato in fase di importazione.")
    import_col, restore_col = st.columns(2)
    with import_col:
        if st.button("Importa e verifica file", type="primary", disabled=uploaded is None):
            result, error = safe_call(extract_document, uploaded)
            if error:
                st.error(f"Errore durante l'importazione: {error}")
            elif not result[0].strip():
                st.error("Non è stato trovato testo. Se il PDF è una scansione serve l'OCR: "
                         "l'app non può inventare il testo che non c'è.")
            else:
                extracted, new_info = result
                st.session_state.original_text = extracted
                st.session_state.doc_info = new_info
                st.session_state.filename = os.path.splitext(uploaded.name)[0]
                st.session_state.last_upload = uploaded.name
                st.session_state.revision_log = []
                st.session_state.metadata = {}
                n_chap = len(detect_chapters(extracted, custom_pat))
                if new_info["formato"] == "PDF":
                    pages = (f"{new_info['pagine_reali']} pagine reali "
                             f"({new_info['pagine_reali'] - len(new_info['pagine_vuote'])} con testo)")
                else:
                    pages = f"~{new_info['parole'] / wpp:.0f} pagine stimate"
                kind = "warning" if new_info.get("pagine_vuote") else "success"
                apply_text(extracted, "Importazione originale",
                           f"Importato {uploaded.name}: {pages}, {n_chap} capitoli rilevati, "
                           f"{new_info['parole']:,} parole. Vedi la scheda «Verifica e verità».", kind)
    with restore_col:
        if st.button("Ripristina originale", disabled=not bool(original)):
            apply_text(original, "Ripristino originale", "Testo originale ripristinato.")

    st.divider()
    st.subheader("Testo completo")
    edited = st.text_area(
        "Puoi modificare il testo direttamente qui", height=480, key=EDITOR_KEY,
        placeholder="Incolla qui il testo del tuo libro oppure importa un file...")

    words = count_words(edited)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Caratteri", f"{len(edited):,}")
    m2.metric("Parole", f"{words:,}")
    m3.metric("Capitoli rilevati", len(chapters_now),
              help="Intestazioni riconosciute nel testo. Dettagli nella scheda «Verifica e verità».")
    if info.get("formato") == "PDF" and edited == original:
        m4.metric("Pagine reali (PDF)", info["pagine_reali"],
                  help="Contate direttamente nel file PDF importato.")
    else:
        m4.metric("Pagine stimate", f"~{words / wpp:.0f}",
                  help=f"Stima: parole / {wpp}. Non sono pagine reali.")
    m5.metric("Lettura stimata", f"~{max(1, round(words / 200))} min")

    if original and original != edited:
        with st.expander("🔍 Confronta con il testo originale (diff)"):
            diff_lines = list(difflib.unified_diff(
                original.splitlines(), edited.splitlines(),
                fromfile="originale", tofile="attuale", lineterm=""))
            st.caption(f"{len(diff_lines)} righe di differenza.")
            st.code("\n".join(diff_lines[:400]) +
                    ("\n… (diff troncato)" if len(diff_lines) > 400 else ""), language="diff")

# ========================= VERIFICA E VERITÀ =========================
with tab_verify:
    st.subheader("Verifica e verità sul tuo libro")
    st.caption("Tutto ciò che vedi qui è misurato sul file o sul testo, non generato dall'AI. "
               "Dove un valore è una stima, è scritto chiaramente.")
    if not text_now.strip():
        st.info("Importa un file o incolla un testo per attivare la verifica.")
    else:
        subject = original if original.strip() else text_now
        subject_info = info if original.strip() else {}
        subject_chapters = detect_chapters(subject, custom_pat)
        offsets = subject_info.get("offset_pagine") if subject_info.get("formato") == "PDF" else None

        st.markdown("#### 1 · Identità del file")
        if subject_info:
            c1, c2, c3 = st.columns(3)
            c1.metric("File", subject_info["nome"])
            c2.metric("Formato", subject_info["formato"])
            c3.metric("Dimensione", f"{subject_info['dimensione_byte'] / 1024:,.0f} KB")
            st.caption("Impronta SHA-256 (identifica in modo univoco questa versione del file):")
            st.code(subject_info["sha256"], language=None)
            if subject_info.get("metadati_pdf"):
                with st.expander("Metadati incorporati nel PDF"):
                    st.json(subject_info["metadati_pdf"])
        else:
            st.caption("Nessun file importato: la verifica riguarda il testo presente nell'editor.")

        st.markdown("#### 2 · Numeri reali")
        sw = count_words(subject)
        v1, v2, v3, v4, v5 = st.columns(5)
        if subject_info.get("formato") == "PDF":
            v1.metric("Pagine reali", subject_info["pagine_reali"], help="Contate nel file PDF.")
            v2.metric("Pagine con testo", subject_info["pagine_reali"] - len(subject_info["pagine_vuote"]))
        elif subject_info.get("formato") == "DOCX":
            v1.metric("Pagine dichiarate da Word", subject_info.get("pagine_dichiarate") or "n/d",
                      help="Valore salvato da Word nel file; può essere obsoleto.")
            v2.metric("Immagini non importate", subject_info.get("immagini", 0))
        else:
            v1.metric("Pagine reali", "n/d", help="Il formato non ha pagine fisiche.")
            v2.metric("Pagine con testo", "n/d")
        v3.metric("Capitoli rilevati", len(subject_chapters))
        v4.metric("Parole", f"{sw:,}")
        v5.metric(f"Pagine stimate", f"~{sw / wpp:.0f}", help=f"Parole / {wpp}: è una STIMA.")

        st.markdown("#### 3 · Esito dei controlli")
        checks = build_checks(subject, subject_info, subject_chapters, wpp)
        icon = {"ok": "✅", "warn": "⚠️", "err": "❌", "info": "ℹ️"}
        for level, msg in checks:
            st.markdown(f"{icon.get(level, '•')} {msg}")

        if subject_info.get("formato") == "PDF" and subject_info.get("tabella_pagine"):
            with st.expander("📄 Dettaglio pagina per pagina (PDF)"):
                df_pages = pd.DataFrame(subject_info["tabella_pagine"])
                st.bar_chart(df_pages.set_index("pagina")["parole"])
                st.caption("Parole estratte per pagina: i buchi indicano pagine vuote o immagini.")
                st.dataframe(df_pages, hide_index=True)

        st.markdown("#### 4 · Capitoli")
        rows_orig = chapter_table(subject, subject_chapters, wpp, offsets)
        if rows_orig:
            st.caption("Pagina inizio/fine = pagine reali del PDF originale." if offsets else
                       f"Pagine stimate a {wpp} parole/pagina.")
            st.dataframe(pd.DataFrame(rows_orig), hide_index=True)
        else:
            st.info("Nessun capitolo riconosciuto: prova il pattern personalizzato nella sidebar.")
        if original.strip() and text_now != original:
            with st.expander("Capitoli del testo ATTUALE (modificato)"):
                rows_cur = chapter_table(text_now, chapters_now, wpp)
                if rows_cur:
                    st.dataframe(pd.DataFrame(rows_cur), hide_index=True)
                else:
                    st.write("Nessun capitolo riconosciuto nel testo attuale.")

        st.markdown("#### 5 · Integrità rispetto all'originale")
        integ = None
        if original.strip() and text_now != original:
            integ = integrity_report(original, text_now)
            i1, i2, i3 = st.columns(3)
            i1.metric("Parole originale → attuale",
                      f"{integ['parole_originale']:,} → {integ['parole_attuale']:,}",
                      f"{integ['variazione_parole_%']:+}%")
            i2.metric("Numeri scomparsi", len(integ["numeri_scomparsi"]))
            i3.metric("Nomi propri scomparsi", len(integ["nomi_scomparsi"]))
            if integ["numeri_scomparsi"] or integ["numeri_nuovi"]:
                st.warning(f"Numeri scomparsi: {', '.join(integ['numeri_scomparsi'][:25]) or '—'}  |  "
                           f"Numeri nuovi: {', '.join(integ['numeri_nuovi'][:25]) or '—'}")
            else:
                st.success("Tutti i numeri dell'originale sono presenti e nessuno è stato aggiunto.")
            if integ["nomi_scomparsi"] or integ["nomi_nuovi"]:
                st.warning(f"Nomi propri scomparsi: {', '.join(integ['nomi_scomparsi'][:25]) or '—'}  |  "
                           f"Nomi propri nuovi: {', '.join(integ['nomi_nuovi'][:25]) or '—'}")
            else:
                st.success("Tutti i nomi propri (parole capitalizzate a metà frase) sono invariati.")
            st.caption("Controllo euristico: una correzione legittima può cambiare un nome scritto male. "
                       "Verifica sempre i casi segnalati.")
        else:
            st.caption("Il testo attuale coincide con l'originale (o non c'è un originale da confrontare).")

        st.markdown("#### 6 · Rapporto scaricabile")
        stats_subject = text_stats(subject)
        md = report_markdown(subject_info, checks, rows_orig, stats_subject, integ, wpp)
        d1, d2 = st.columns(2)
        d1.download_button("Scarica rapporto (Markdown)", data=md.encode("utf-8"),
                           file_name="rapporto_verifica.md", mime="text/markdown")
        d2.download_button(
            "Scarica rapporto (JSON)",
            data=json.dumps({"file": {k: v for k, v in subject_info.items()
                                      if k not in ("offset_pagine", "tabella_pagine")},
                             "controlli": [{"livello": l, "messaggio": m} for l, m in checks],
                             "capitoli": rows_orig, "integrita": integ},
                            ensure_ascii=False, indent=2, default=str),
            file_name="rapporto_verifica.json", mime="application/json")

        with st.expander("ℹ️ Cosa significa «verificato» (e cosa no)"):
            st.markdown(
                "- **Verificato** = misurato direttamente: pagine del PDF, pagine senza testo, "
                "parole, capitoli riconosciuti, numerazione, impronta SHA-256.\n"
                "- **Stimato** = parole ÷ parole-per-pagina (EPUB, DOCX, TXT): non sono pagine reali.\n"
                "- **Non verificabile** = testo dentro immagini, qualità letteraria, correttezza dei "
                "fatti narrati, copyright. Per le scansioni serve un OCR.\n"
                "- I capitoli sono riconosciuti dal *testo* (es. «Capitolo 3»): un titolo scritto in "
                "modo insolito può sfuggire (usa il pattern personalizzato).")

# ========================= PULIZIA E STRUMENTI =========================
with tab_tools:
    st.subheader("Diagnosi locale (senza AI)")
    if not text_now.strip():
        st.info("Importa o incolla un testo per usare gli strumenti.")
    else:
        pages_hint = info.get("pagine_reali", 0) if (info.get("formato") == "PDF" and text_now == original) else 0
        stats = text_stats(text_now)
        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric("Frasi", f"{stats['frasi']:,}")
        s2.metric("Paragrafi", f"{stats['paragrafi']:,}")
        s3.metric("Parole per frase", stats["parole_per_frase"])
        s4.metric("Leggibilità Gulpease", stats["gulpease"], gulpease_label(stats["gulpease"]),
                  help="Indice di leggibilità per l'italiano (0–100). Per altre lingue non è affidabile.")
        s5.metric("Varietà lessicale", stats["varieta_lessicale"],
                  help="Parole diverse / parole totali. Dipende dalla lunghezza del testo.")
        issues = find_issues(text_now, pages_hint)
        df_iss = pd.DataFrame([{"Problema": k, "Occorrenze": v} for k, v in issues.items()])
        col_a, col_b = st.columns([3, 2])
        with col_a:
            st.markdown("**Problemi tecnici rilevati**")
            st.dataframe(df_iss, hide_index=True)
        with col_b:
            st.markdown("**Parole più usate** (≥ 4 lettere, senza parole comuni)")
            if stats["top"]:
                st.bar_chart(pd.DataFrame(stats["top"], columns=["parola", "volte"]).set_index("parola"))
        with st.expander("Le 5 frasi più lunghe"):
            for s in stats["frasi_lunghe"]:
                st.write("• " + s + ("…" if len(s) >= 400 else ""))

        st.divider()
        st.subheader("Pulizia automatica")
        st.caption("Operazioni locali, istantanee e ripetibili. Prima di applicare viene salvato "
                   "un backup nella scheda Versioni.")
        opts = {}
        oc1, oc2 = st.columns(2)
        for i, (key, label, default) in enumerate(CLEAN_OPTS):
            with (oc1 if i % 2 == 0 else oc2):
                opts[key] = st.checkbox(label, value=default, key=f"clean_{key}")
        p1, p2 = st.columns(2)
        min_rep = max(4, int(pages_hint * 0.3)) if pages_hint else 6
        if p1.button("👁️ Anteprima (non modifica)"):
            cleaned, rep = clean_text(text_now, opts, min_rep)
            st.write(f"Caratteri: {len(text_now):,} → {len(cleaned):,}  ·  "
                     f"Parole: {count_words(text_now):,} → {count_words(cleaned):,}")
            st.dataframe(pd.DataFrame([{"Operazione": k, "Interventi": v} for k, v in rep.items()]),
                         hide_index=True)
        if p2.button("🧹 Applica pulizia", type="primary"):
            cleaned, rep = clean_text(text_now, opts, min_rep)
            total_ops = sum(rep.values())
            save_version("Backup prima della pulizia", text_now)
            summary = "; ".join(f"{k}: {v}" for k, v in rep.items() if v) or "nessun intervento necessario"
            apply_text(cleaned, "Pulizia automatica",
                       f"Pulizia completata ({total_ops} interventi): {summary}.")

        st.divider()
        st.subheader("Cerca e sostituisci")
        f1, f2 = st.columns(2)
        needle = f1.text_input("Cerca", key="find_text")
        repl = f2.text_input("Sostituisci con", key="repl_text")
        fo1, fo2 = st.columns(2)
        use_regex = fo1.checkbox("Espressione regolare", key="find_regex")
        case_sens = fo2.checkbox("Distingui maiuscole/minuscole", key="find_case")
        pattern = None
        if needle:
            try:
                pattern = re.compile(needle if use_regex else re.escape(needle),
                                     0 if case_sens else re.I)
            except re.error as exc:
                st.error(f"Espressione non valida: {exc}")
        if pattern is not None:
            matches = list(pattern.finditer(text_now))
            st.write(f"**{len(matches)}** occorrenze trovate.")
            for m in matches[:8]:
                a, b = max(0, m.start() - 40), min(len(text_now), m.end() + 40)
                st.code(text_now[a:b].replace("\n", " ⏎ "), language=None)
            if st.button("Sostituisci tutto", disabled=not matches):
                try:
                    new_text = pattern.sub((lambda _m: repl) if not use_regex else repl, text_now)
                except re.error as exc:
                    st.error(f"Sostituzione non valida: {exc}")
                else:
                    apply_text(new_text, f"Sostituzione: {needle[:25]}",
                               f"{len(matches)} occorrenze sostituite.")

# ============================ REVISIONE AI ============================
with tab_ai:
    st.subheader("Correzione intelligente (AI)")
    task = st.selectbox("Scegli il tipo di revisione", [
        "Correzione completa", "Solo grammatica e ortografia", "Punteggiatura e sintassi",
        "Chiarezza e leggibilità", "Uniformità di nomi, date e terminologia",
        "Revisione editoriale approfondita"])
    spans = chapter_spans(text_now, chapters_now)
    scope = st.radio("Ambito", ["Tutto il testo", "Solo i capitoli selezionati"],
                     horizontal=True, disabled=not spans,
                     help="Per revisionare solo alcuni capitoli servono capitoli rilevati.")
    selected_idx = []
    if spans and scope.startswith("Solo"):
        labels = [f"{i + 1}. {sp['title'][:60]}" for i, sp in enumerate(spans)]
        picked = st.multiselect("Capitoli da revisionare", labels)
        selected_idx = [labels.index(p) for p in picked]
    use_chapters = bool(spans) and scope.startswith("Solo")
    if use_chapters:
        seg_ranges = [(spans[i]["start"], spans[i]["end"]) for i in sorted(selected_idx)]
    else:
        seg_ranges = [(0, len(text_now))]
    segments_src = [text_now[a:b] for a, b in seg_ranges] if text_now.strip() else []
    n_calls = sum(1 for seg in segments_src for p in split_text(seg) if p.strip())
    chars_sent = sum(len(seg) for seg in segments_src)
    st.caption(f"Verranno inviati **{n_calls} blocchi** a Google Gemini ({chars_sent:,} caratteri). "
               "Il testo lascia questa app e viene elaborato dall'API di Google.")
    if n_calls > 40:
        st.warning("Sono molti blocchi: tempi e costi API possono essere elevati. "
                   "Valuta di revisionare un capitolo alla volta.")
    can_run = bool(segments_src) and n_calls > 0
    if st.button("✨ Avvia correzione AI", type="primary", disabled=not can_run):
        if not get_api_key():
            st.error("Configura prima la chiave API di Google nella sidebar.")
        else:
            save_version("Backup prima della revisione", text_now)
            with st.spinner("L'AI sta revisionando il testo..."):
                result, error = safe_call(run_revision, segments_src, task, language,
                                          intensity, preserve_style, model, strict)
            if error:
                st.error(f"Revisione non completata: {error}")
            else:
                new_segments, log = result
                new_text = text_now
                for (a, b), seg in sorted(zip(seg_ranges, new_segments), reverse=True):
                    new_text = new_text[:a] + seg + new_text[b:]
                st.session_state.revision_log = log
                ok = sum(1 for e in log if e["stato"].startswith("accettato"))
                bad = len(log) - ok
                if new_text == text_now:
                    flash("Nessuna modifica applicata: tutti i blocchi sono stati scartati o già corretti.",
                          "warning")
                    st.rerun()
                kind = "warning" if bad else "success"
                apply_text(new_text, f"Revisione AI: {task}",
                           f"Revisione terminata: {ok}/{len(log)} blocchi applicati, {bad} scartati o in errore. "
                           "Controlla il report.", kind)

    log = st.session_state.revision_log
    if log:
        st.markdown("#### Report dell'ultima revisione")
        ok_n = sum(1 for e in log if e["stato"] == "accettato")
        warn_n = sum(1 for e in log if e["stato"] == "accettato con avviso")
        rej_n = sum(1 for e in log if e["stato"].startswith("scartato"))
        err_n = sum(1 for e in log if e["stato"].startswith("errore"))
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Accettati", ok_n)
        r2.metric("Accettati con avviso", warn_n)
        r3.metric("Scartati (originale tenuto)", rej_n)
        r4.metric("Errori", err_n)
        if rej_n or err_n:
            st.warning("Alcuni blocchi NON sono stati revisionati: sono rimasti com'erano. "
                       "Puoi rilanciare la revisione per riprovarli.")
        st.dataframe(pd.DataFrame(log), hide_index=True)
    st.caption("Controlla sempre il risultato prima della pubblicazione: l'AI può sbagliare anche "
               "quando i controlli automatici non rilevano anomalie.")

# ======================== ERRORI GRAMMATICALI ========================
with tab_gram:
    st.subheader("Segnalazione errori grammaticali")
    st.caption("Il controllo NON modifica il testo: elenca gli errori con pagina, capitolo, "
               "correzione proposta e spiegazione. Poi scegli tu cosa applicare.")
    is_it = language == "italiano"
    gram_blocks = len([c for c in split_text(text_now, 6000) if c.strip()]) if text_now.strip() else 0
    sig_now = hashlib.sha1(text_now.encode("utf-8", "ignore")).hexdigest()
    b1, b2 = st.columns(2)
    with b1:
        st.markdown("**⚡ Controllo rapido** — regole locali, gratis, senza chiave API")
        st.caption("Accenti e apostrofi (perchè, qual'è, un pò), «ha/ho/è» scritti male, punteggiatura, "
                   "parole ripetute, parentesi non chiuse. Solo italiano.")
        if st.button("⚡ Controlla (senza AI)", disabled=not (text_now.strip() and is_it), key="run_local"):
            st.session_state.grammar = {"sig": sig_now, "items": local_grammar_check(text_now),
                                        "blocks": 0, "failed": [], "discarded": 0, "source": "regole locali"}
        if not is_it:
            st.caption("Le regole locali sono disponibili solo con lingua «italiano».")
    with b2:
        st.markdown("**📝 Controllo completo con AI** — regole locali + Gemini")
        st.caption(f"Invia **{gram_blocks} blocchi** a Google Gemini. Trova anche errori che le regole non vedono.")
        if st.button("📝 Controlla con AI", type="primary", disabled=gram_blocks == 0, key="run_grammar"):
            if not get_api_key():
                st.error("Configura prima la chiave API di Google nella sidebar.")
            else:
                res, err = safe_call(check_grammar, text_now, model, language)
                if err:
                    st.error(f"Controllo non completato: {err}")
                else:
                    res["items"] = merge_items(res["items"], local_grammar_check(text_now) if is_it else [])
                    res["source"] = "AI + regole locali" if is_it else "AI"
                    st.session_state.grammar = res
    st.caption("Nessun controllo automatico è completo: «nessun errore» non è una garanzia che il testo sia perfetto.")

    g = st.session_state.grammar
    if g:
        st.markdown(f"##### Risultato · fonte: {g.get('source', 'AI')}")
        if g["sig"] != sig_now:
            st.warning("Il testo è cambiato dopo il controllo: gli errori elencati potrebbero non "
                       "essere più aggiornati. Rilancia il controllo.")
        g1, g2, g3, g4 = st.columns(4)
        g1.metric("Errori segnalati", len(g["items"]))
        g2.metric("Blocchi AI controllati",
                  "—" if not g["blocks"] else f"{g['blocks'] - len(g['failed'])}/{g['blocks']}")
        g3.metric("Segnalazioni AI scartate", g["discarded"],
                  help="L'AI ha citato un testo che non esiste nel libro: scartata per non inventare errori.")
        g4.metric("Errori per 1000 parole", round(len(g["items"]) / max(1, count_words(text_now)) * 1000, 1))
        if g["failed"]:
            st.warning("Alcuni blocchi NON sono stati controllati: " + "; ".join(g["failed"][:5]))
        if not g["items"]:
            st.success("Nessun errore segnalato dai controlli eseguiti.")
        else:
            offsets = info.get("offset_pagine") if (info.get("formato") == "PDF" and text_now == original) else None
            rows = []
            for it in g["items"]:
                row = {"Applica": False}
                if offsets:
                    row["Pagina PDF"] = bisect_right(offsets, it["pos"])
                prev = [c["title"] for c in chapters_now if c["start"] <= it["pos"]]
                row["Capitolo"] = prev[-1][:40] if prev else "—"
                row.update({"Fonte": it.get("fonte", "AI"), "Tipo": it["tipo"], "Errore": it["errore"],
                            "Correzione": it["correzione"], "Spiegazione": it["spiegazione"]})
                rows.append(row)
            df_g = pd.DataFrame(rows)
            if not offsets and info.get("formato") == "PDF":
                st.caption("Il numero di pagina compare solo finché il testo coincide con il PDF originale importato.")
            edited_g = st.data_editor(df_g, hide_index=True, key="grammar_editor",
                                      disabled=[c for c in df_g.columns if c != "Applica"])
            st.download_button("Scarica elenco errori (CSV)",
                               data=df_g.drop(columns=["Applica"]).to_csv(index=False).encode("utf-8-sig"),
                               file_name="errori_grammaticali.csv", mime="text/csv")
            chosen = [i for i, v in enumerate(edited_g["Applica"].tolist()) if v]
            st.caption("Le voci con correzione «—» (parentesi non bilanciate) vanno sistemate a mano.")
            if st.button(f"Applica le {len(chosen)} correzioni selezionate", disabled=not chosen):
                if g["sig"] != sig_now:
                    st.error("Il testo è cambiato: rilancia il controllo prima di applicare.")
                else:
                    new_text, done, skipped = text_now, 0, 0
                    for i in sorted(chosen, key=lambda k: g["items"][k]["pos"], reverse=True):
                        it = g["items"][i]
                        a, b = it["pos"], it["pos"] + len(it["errore"])
                        if it.get("auto", True) and new_text[a:b] == it["errore"]:
                            new_text = new_text[:a] + it["correzione"] + new_text[b:]
                            done += 1
                        else:
                            skipped += 1
                    st.session_state.grammar = None
                    apply_text(new_text, "Correzioni grammaticali selezionate",
                               f"{done} correzioni applicate, {skipped} saltate (da fare a mano). "
                               "Rilancia il controllo per verificare.",
                               "success" if done else "warning")

# ============================== SEO ==============================
with tab_seo:
    st.subheader("Metadati editoriali e SEO")
    title = st.text_input("Titolo", key="seo_title")
    subtitle = st.text_input("Sottotitolo", key="seo_subtitle")
    author = st.text_input("Autore", key="seo_author")
    genre = st.text_input("Genere / categoria", key="seo_genre")
    audience = st.text_input("Pubblico di riferimento", key="seo_audience")

    if st.button("Genera descrizione, keyword e tag SEO", type="primary",
                 disabled=not bool(text_now)):
        if not get_api_key():
            st.error("Configura prima la chiave API di Google nella sidebar.")
        else:
            n = len(text_now)
            if n <= 14000:
                sample = text_now
            else:
                mid = n // 2
                sample = (text_now[:6000] + "\n[...]\n" + text_now[mid - 2000:mid + 2000]
                          + "\n[...]\n" + text_now[-4000:])
            prompt = f"""
Analizza il testo del libro e crea metadati editoriali accurati.
Titolo attuale: {title}
Sottotitolo: {subtitle}
Autore: {author}
Genere: {genre}
Pubblico: {audience}

TESTO (estratti: inizio, metà e fine del libro):
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
                    model)
            if error:
                st.error(f"Non è stato possibile generare i metadati: {error}")
            else:
                parsed, error = safe_call(parse_json_response, raw)
                if error:
                    st.error(f"La risposta dell'AI non è un JSON valido: {error}")
                else:
                    st.session_state.metadata = parsed
                    st.success("Metadati generati.")

    meta = st.session_state.metadata
    if meta:
        st.caption("L'AI ha letto solo estratti (inizio, metà e fine) se il libro supera 14.000 caratteri.")
        for key_name, limit in (("meta_title", 60), ("meta_description", 160)):
            val = str(meta.get(key_name, ""))
            if val:
                length = len(val)
                (st.success if length <= limit else st.warning)(
                    f"{key_name}: {length}/{limit} caratteri" +
                    ("" if length <= limit else " — SUPERA il limite, accorcialo a mano."))

        def apply_suggested():
            m = st.session_state.metadata
            if m.get("titolo_suggerito"):
                st.session_state.seo_title = str(m["titolo_suggerito"])
            if m.get("sottotitolo_suggerito"):
                st.session_state.seo_subtitle = str(m["sottotitolo_suggerito"])

        st.button("Usa titolo e sottotitolo suggeriti", on_click=apply_suggested)
        st.json(meta)
        st.download_button("Scarica metadati JSON",
                           data=json.dumps(meta, ensure_ascii=False, indent=2),
                           file_name="metadati_libro_seo.json", mime="application/json")

# ============================ VERSIONI ============================
with tab_history:
    st.subheader("Cronologia e ripristino")
    st.caption(f"Si conservano le ultime {MAX_VERSIONS} versioni, anche su disco "
               "(cartella `versions/`), e vengono ricaricate al riavvio dell'app.")
    if not st.session_state.versions:
        st.info("La cronologia si popola quando importi un file o avvii un'operazione.")
    else:
        for reverse_index, version in enumerate(reversed(st.session_state.versions)):
            version_index = len(st.session_state.versions) - 1 - reverse_index
            vw = count_words(version["text"])
            with st.expander(f"{version['label']} — {version['created_at']}"):
                st.caption(f"{len(version['text']):,} caratteri · {vw:,} parole")
                st.text(version["text"][:1500] + ("…" if len(version["text"]) > 1500 else ""))
                if st.button("Ripristina questa versione", key=f"restore_version_{version_index}"):
                    apply_text(version["text"], f"Ripristino: {version['label']}",
                               "Versione ripristinata.")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            if st.button("Salva snapshot manuale", disabled=not bool(text_now)):
                save_version("Snapshot manuale", text_now)
                st.success("Snapshot salvato.")
        with col_b:
            st.download_button(
                "Scarica tutta la cronologia (JSON)",
                data=json.dumps(st.session_state.versions, ensure_ascii=False, indent=2),
                file_name="cronologia_versioni.json", mime="application/json")
        with col_c:
            if st.button("Elimina la cronologia"):
                st.session_state.versions = []
                shutil.rmtree(VERSIONS_DIR, ignore_errors=True)
                st.rerun()

# ============================== ESPORTA ==============================
with tab_export:
    st.subheader("Esporta il tuo libro")
    base_name = st.session_state.filename or "libro_revisionato"
    if not text_now.strip():
        st.info("Importa o incolla un testo per attivare l'esportazione.")
    else:
        default_title = (st.session_state.seo_title
                         or str(st.session_state.metadata.get("titolo_suggerito", ""))
                         or base_name)
        ex1, ex2 = st.columns(2)
        export_title = ex1.text_input("Titolo", value=default_title)
        export_author = ex2.text_input("Autore", value=st.session_state.seo_author)
        st.caption(f"I {len(chapters_now)} capitoli rilevati diventano titoli (DOCX), "
                   "inizio pagina con numerazione (PDF) e indice navigabile (EPUB).")

        t1, t2 = st.columns(2)
        t1.download_button("Scarica TXT", data=text_now.encode("utf-8"),
                           file_name=f"{base_name}_revisionato.txt", mime="text/plain")
        t2.download_button("Scarica Markdown", data=text_now.encode("utf-8"),
                           file_name=f"{base_name}_revisionato.md", mime="text/markdown")

        sig = hashlib.sha1("|".join([text_now, export_title, export_author,
                                     custom_pat, language]).encode("utf-8", "ignore")).hexdigest()
        formats = [
            ("docx", "Word (DOCX)", lambda: build_docx(text_now, export_title, export_author, custom_pat),
             "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            ("pdf", "PDF", lambda: build_pdf(text_now, export_title, export_author, custom_pat),
             "application/pdf"),
            ("epub", "EPUB", lambda: build_epub(text_now, export_title, export_author,
                                                LANG_CODES[language], custom_pat),
             "application/epub+zip"),
        ]
        cols = st.columns(3)
        for col, (ext, label, builder, mime) in zip(cols, formats):
            with col:
                slot = f"export_{ext}"
                if st.button(f"Prepara {label}", key=f"prep_{ext}"):
                    with st.spinner(f"Creo {label}..."):
                        data_bytes, err = safe_call(builder)
                    st.session_state[slot] = {"sig": sig, "data": data_bytes, "err": str(err) if err else ""}
                entry = st.session_state.get(slot)
                if entry and entry["sig"] == sig:
                    if entry["err"]:
                        st.error(f"Errore {label}: {entry['err']}")
                    else:
                        st.download_button(f"⬇️ Scarica {label}", data=entry["data"],
                                           file_name=f"{base_name}_revisionato.{ext}", mime=mime,
                                           key=f"dl_{ext}")
                elif entry:
                    st.caption("Il testo è cambiato: preparalo di nuovo.")

st.divider()
st.caption(
    "LIBRO MASTER AI — I PDF scannerizzati richiedono OCR. L'importazione e l'esportazione "
    "ricostruiscono il testo e non conservano necessariamente impaginazione, immagini, note o "
    "grafica originale. Controlla sempre il testo revisionato."
)
