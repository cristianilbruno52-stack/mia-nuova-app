# -*- coding: utf-8 -*-
"""LIBRO MASTER AI — versione corretta.

- Modello Gemini configurabile, con gemini-3.8-flash come predefinito.
- Suddivisione del testo senza perdere caratteri e con offset globali esatti.
- Contesto precedente/successivo inviato all'AI, ma mai corretto.
- Caporali controllate sull'intero documento, non sulle singole pagine PDF.
- Un blocco AI fallito non viene mai conteggiato come revisionato.
"""
from __future__ import annotations

import copy
import base64
import hashlib
import html
import io
import json
import os
import re
import shutil
import time
import zipfile
from bisect import bisect_right
from collections import Counter
from datetime import datetime, timezone
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph as DocxParagraph
from ebooklib import ITEM_DOCUMENT, epub
from google import genai
from google.genai import types
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer

st.set_page_config(page_title="LIBRO MASTER AI", page_icon="📚", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERSIONS_DIR = os.path.join(BASE_DIR, "versions")
MAX_VERSIONS = 40
DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
LANG_CODES = {"italiano": "it", "inglese": "en", "francese": "fr", "spagnolo": "es", "tedesco": "de"}
WORD_RE = re.compile(r"\w+(?:['’-]\w+)*", re.UNICODE)
CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
INVISIBLE_RE = re.compile(r"[\u00ad\u200b\u200c\u200d\u2060\ufeff]")
XML_BAD_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
PAGE_NUM_LINE = re.compile(r"[ \t]*(?:(?:pag(?:ina|\.)?|page|p\.)[ \t]*)?[-–—]?[ \t]*\d{1,4}[ \t]*[-–—]?[ \t]*(?:(?:di|of|/)[ \t]*\d{1,4})?[ \t]*", re.I)
_NUM = r"(?:\d{1,3}|[ivxlcdm]{1,8}|[a-zàèéìòù]{3,12})"
CHAPTER_RE = re.compile(
    r"^\s*(?:"
    rf"(?:capitolo|cap\.|chapter|parte|part|libro|book|sezione|section)\s+{_NUM}(?:\s*[:.\-–—]\s*.{{0,80}})?"
    r"|(?:prologo|epilogo|prefazione|introduzione|premessa|conclusioni|conclusione|postfazione|appendice|ringraziamenti|prologue|epilogue|preface|introduction|conclusion|afterword|appendix|acknowledg(?:e)?ments)(?:\s*[:.\-–—]?\s*.{0,50})?"
    r"|\#{1,2}\s+\S.{0,80}"
    r")\s*$", re.I)
ROMAN_RE = re.compile(r"^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$")

DEFAULTS = {
    "original_text": "", "versions": [], "metadata": {}, "filename": "libro",
    "google_api_key": "", "book_editor": "", "pending_text": None,
    "flash": [], "doc_info": {}, "revision_log": [], "grammar": None,
    "uploader_gen": 0, "seo_title": "", "seo_author": "",
}


def init_state():
    for key, value in DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = copy.deepcopy(value)


init_state()
if st.session_state.pending_text is not None:
    st.session_state.book_editor = st.session_state.pending_text
    st.session_state.pending_text = None


def get_text():
    return st.session_state.get("book_editor", "") or ""


def flash(msg, kind="success"):
    st.session_state.flash.append((msg, kind))


def show_flashes():
    icons = {"success": "✅", "warning": "⚠️", "error": "❌", "info": "ℹ️"}
    for msg, kind in st.session_state.flash:
        st.toast(msg, icon=icons.get(kind, "ℹ️"))
    st.session_state.flash = []


def count_words(text):
    return len(WORD_RE.findall(text))


def safe_call(func, *args, **kwargs):
    try:
        return func(*args, **kwargs), None
    except Exception as exc:  # noqa: BLE001
        return None, exc


def save_version(label, text):
    entry = {"label": label, "text": text, "created_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
    st.session_state.versions.append(entry)
    st.session_state.versions = st.session_state.versions[-MAX_VERSIONS:]
    try:
        os.makedirs(VERSIONS_DIR, exist_ok=True)
        name = datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".json"
        with open(os.path.join(VERSIONS_DIR, name), "w", encoding="utf-8") as fh:
            json.dump(entry, fh, ensure_ascii=False)
        files = sorted(f for f in os.listdir(VERSIONS_DIR) if f.endswith(".json"))
        for old in files[:-MAX_VERSIONS]:
            os.remove(os.path.join(VERSIONS_DIR, old))
    except OSError:
        pass


def apply_text(new_text, label, notify="", kind="success"):
    save_version(label, new_text)
    st.session_state.pending_text = new_text
    st.session_state.grammar = None
    if notify:
        flash(notify, kind)
    st.rerun()


# ---------------------------------------------------------------------
# IMPORTAZIONE
# ---------------------------------------------------------------------
def decode_text(data):
    try:
        return data.decode("utf-8-sig"), "UTF-8"
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace"), "Windows-1252 (verifica gli accenti)"


def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted and not reader.decrypt(""):
        raise ValueError("Il PDF è protetto da password.")
    parts, offsets, table, empty = [], [], [], []
    pos = 0
    for number, page in enumerate(reader.pages, 1):
        try:
            page_text = page.extract_text() or ""
            err = ""
        except Exception as exc:  # noqa: BLE001
            page_text, err = "", str(exc)[:100]
        offsets.append(pos)
        parts.append(page_text)
        pos += len(page_text) + 2
        words = count_words(page_text)
        if words < 5:
            empty.append(number)
        table.append({"pagina": number, "parole": words, "caratteri": len(page_text.strip()),
                      "stato": "vuota / immagine" if words < 5 else "ok", "errore": err})
    return "\n\n".join(parts), {
        "pagine_reali": len(reader.pages), "pagine_vuote": empty,
        "tabella_pagine": table, "offset_pagine": offsets,
    }


def docx_declared_pages(data):
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            xml = z.read("docProps/app.xml").decode("utf-8", "ignore")
        match = re.search(r"<Pages>(\d+)</Pages>", xml)
        return int(match.group(1)) if match else None
    except Exception:  # noqa: BLE001
        return None


def extract_docx(data):
    doc = Document(io.BytesIO(data))
    parts = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            para = DocxParagraph(child, doc)
            if para.text.strip():
                parts.append(para.text)
        elif tag == "tbl":
            for row in Table(child, doc).rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n\n".join(parts), {"tabelle": len(doc.tables), "immagini": len(doc.inline_shapes),
                                "pagine_dichiarate": docx_declared_pages(data)}


def extract_epub(data):
    book = epub.read_epub(io.BytesIO(data))
    sections = []
    for item in book.get_items_of_type(ITEM_DOCUMENT):
        if isinstance(item, epub.EpubNav) or item.get_name().lower().endswith("nav.xhtml"):
            continue
        raw = item.get_content().decode("utf-8", errors="ignore")
        raw = re.sub(r"<(script|style)\b.*?</\1\s*>", " ", raw, flags=re.I | re.S)
        raw = re.sub(r"</(p|div|h[1-6]|li|br|section)\s*>", "\n", raw, flags=re.I)
        clean = html.unescape(re.sub(r"<[^>]+>", " ", raw))
        clean = re.sub(r"[ \t]+", " ", clean)
        clean = re.sub(r"\n\s*\n+", "\n\n", clean)
        if clean.strip():
            sections.append(clean.strip())
    return "\n\n".join(sections), {"file_documento": len(sections)}


def extract_document(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    info = {"nome": uploaded_file.name, "dimensione_byte": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
            "importato_il": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
    if name.endswith((".txt", ".md")):
        text, encoding = decode_text(data)
        info.update(formato="Markdown" if name.endswith(".md") else "TXT", codifica=encoding)
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
        raise ValueError("Formato non supportato.")
    info.update(parole=count_words(text), caratteri=len(text))
    return text, info


# ---------------------------------------------------------------------
# CAPITOLI E STATISTICHE
# ---------------------------------------------------------------------
def roman_to_int(value):
    value = value.upper()
    if not value or not ROMAN_RE.match(value):
        return None
    vals = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    total = 0
    for i, char in enumerate(value):
        total += -vals[char] if i + 1 < len(value) and vals[value[i + 1]] > vals[char] else vals[char]
    return total


def chapter_number(title):
    match = re.match(r"^\s*(?:capitolo|cap\.|chapter)\s+(\d{1,3}|[ivxlcdm]{1,8})\b", title, re.I)
    if not match:
        return None
    value = match.group(1)
    return int(value) if value.isdigit() else roman_to_int(value)


def detect_chapters(text, extra_pattern=""):
    extra = None
    if extra_pattern.strip():
        try:
            extra = re.compile(extra_pattern, re.I)
        except re.error:
            pass
    chapters, pos = [], 0
    for line_no, line in enumerate(text.split("\n")):
        stripped = line.strip()
        if stripped and len(stripped) <= 100 and (CHAPTER_RE.match(line) or (extra and extra.search(stripped))):
            chapters.append({"title": stripped, "start": pos, "line": line_no})
        pos += len(line) + 1
    return chapters


def chapter_spans(text, chapters):
    if not chapters:
        return []
    spans = []
    if text[:chapters[0]["start"]].strip():
        spans.append({"title": "(testo prima del primo capitolo)", "start": 0,
                      "end": chapters[0]["start"], "front": True})
    for i, chapter in enumerate(chapters):
        end = chapters[i + 1]["start"] if i + 1 < len(chapters) else len(text)
        spans.append({"title": chapter["title"], "start": chapter["start"], "end": end,
                      "front": False, "line": chapter["line"]})
    return spans


# ---------------------------------------------------------------------
# PULIZIA LOCALE
# ---------------------------------------------------------------------
def repeated_lines(text, min_count=6):
    counts = Counter()
    for line in text.split("\n"):
        stripped = line.strip()
        if 8 <= len(stripped) <= 80 and not stripped.endswith((".", "!", "?", "…", "»", "”", '"', "-", ",")) and not CHAPTER_RE.match(stripped):
            counts[re.sub(r"\d+", "#", stripped.lower())] += 1
    return {key: val for key, val in counts.items() if val >= min_count}


def clean_text(text, options, min_repeat=6):
    report = {}
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if options.get("ctrl"):
        text, n1 = re.subn(r"(\w)\u00ad\n(\w)", r"\1\2", text)
        text, n2 = INVISIBLE_RE.subn("", text)
        text, n3 = CONTROL_RE.subn("", text)
        report["Caratteri invisibili rimossi"] = n1 + n2 + n3
    if options.get("pagenum"):
        lines = text.split("\n")
        kept = [line for line in lines if not (line.strip() and PAGE_NUM_LINE.fullmatch(line))]
        report["Numeri di pagina rimossi"] = len(lines) - len(kept)
        text = "\n".join(kept)
    if options.get("headers"):
        repeated = repeated_lines(text, min_repeat)
        lines = text.split("\n")
        kept = [line for line in lines if not (line.strip() and re.sub(r"\d+", "#", line.strip().lower()) in repeated)]
        report["Intestazioni ripetute rimosse"] = len(lines) - len(kept)
        text = "\n".join(kept)
    if options.get("dehyph"):
        text, count = re.subn(r"([a-zà-ÿ])-\n[ \t]*([a-zà-ÿ])", r"\1\2", text)
        report["Parole ricongiunte"] = count
    if options.get("spaces"):
        text, n1 = re.subn(r"[ \t]{2,}", " ", text)
        text, n2 = re.subn(r"[ \t]+\n", "\n", text)
        report["Spazi superflui eliminati"] = n1 + n2
    if options.get("punct"):
        text, count = re.subn(r"(?<=\w) +([,;:!?]|\.(?!\.))", r"\1", text)
        report["Spazi prima della punteggiatura tolti"] = count
    if options.get("blank"):
        text, count = re.subn(r"\n{3,}", "\n\n", text)
        report["Righe vuote ridotte"] = count
    return text.strip(), report


# ---------------------------------------------------------------------
# BLOCCHI: offset esatti e ricomposizione senza alterazioni
# ---------------------------------------------------------------------
def split_text_spans(text, max_chars=6000):
    """Restituisce coppie (inizio, fine). ''.join dei blocchi = testo originale."""
    if max_chars < 500:
        raise ValueError("max_chars troppo piccolo")
    spans, start = [], 0
    while start < len(text):
        if len(text) - start <= max_chars:
            spans.append((start, len(text)))
            break
        limit = start + max_chars
        floor = start + max_chars // 2
        cut = text.rfind("\n\n", floor, limit)
        if cut >= floor:
            cut += 2
        else:
            cut = text.rfind("\n", floor, limit)
            if cut >= floor:
                cut += 1
            else:
                sentence = max(text.rfind(". ", floor, limit), text.rfind("! ", floor, limit), text.rfind("? ", floor, limit))
                space = text.rfind(" ", floor, limit)
                cut = sentence + 2 if sentence >= floor else (space + 1 if space >= floor else limit)
        if cut <= start:
            cut = limit
        spans.append((start, cut))
        start = cut
    return spans or [(0, 0)]


def split_text(text, max_chars=9000):
    return [text[a:b] for a, b in split_text_spans(text, max_chars)]


def make_chunks(text, max_chars=6000, context_chars=1400):
    chunks = []
    for number, (start, end) in enumerate(split_text_spans(text, max_chars), 1):
        chunks.append({
            "number": number, "start": start, "end": end, "text": text[start:end],
            "previous": text[max(0, start - context_chars):start],
            "next": text[end:min(len(text), end + context_chars)],
        })
    return chunks


# ---------------------------------------------------------------------
# GEMINI
# ---------------------------------------------------------------------
class QuotaExhaustedError(RuntimeError):
    """Quota API esaurita: riprovare subito aumenterebbe soltanto gli errori."""


def get_api_key():
    if st.session_state.get("google_api_key"):
        return st.session_state.google_api_key
    try:
        return st.secrets.get("GOOGLE_API_KEY", "") or st.secrets.get("GEMINI_API_KEY", "") or os.getenv("GOOGLE_API_KEY", "") or os.getenv("GEMINI_API_KEY", "")
    except Exception:  # noqa: BLE001
        return os.getenv("GOOGLE_API_KEY", "") or os.getenv("GEMINI_API_KEY", "")


@st.cache_resource(show_spinner=False)
def get_client(api_key):
    return genai.Client(api_key=api_key)


def ask_ai(prompt, system_prompt, model, temperature=0.0, json_mode=False):
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError("Chiave API non configurata.")
    client = get_client(api_key)
    response = None
    for attempt in range(3):
        try:
            config_kwargs = {"system_instruction": system_prompt, "temperature": temperature}
            if json_mode:
                config_kwargs["response_mime_type"] = "application/json"
            response = client.models.generate_content(
                model=model, contents=prompt,
                config=types.GenerateContentConfig(**config_kwargs))
            break
        except Exception as exc:  # noqa: BLE001
            message = str(exc)
            upper = message.upper()
            if "429" in upper or "RESOURCE_EXHAUSTED" in upper or "QUOTA" in upper:
                # Un 429 di quota non viene ritentato: piano/billing oppure reset
                # della finestra temporale sono azioni esterne all'applicazione.
                raise QuotaExhaustedError(
                    "Quota Gemini esaurita. Controlla piano e fatturazione oppure attendi il reset "
                    "della quota. Le richieste rimanenti sono state fermate per evitare sprechi.") from exc
            transient = any(key in upper for key in ("503", "UNAVAILABLE", "TIMEOUT", "DEADLINE"))
            if attempt < 2 and transient:
                time.sleep(3 * (attempt + 1))
                continue
            if "404" in upper or "NOT_FOUND" in upper or "NO LONGER AVAILABLE" in upper:
                raise RuntimeError(f"Modello non disponibile ({model}). Aggiorna il nome del modello nella sidebar. Dettaglio: {message[:220]}") from exc
            raise
    candidates = getattr(response, "candidates", None) if response else None
    if not candidates:
        raise RuntimeError("Risposta vuota o bloccata dal modello.")
    reason = str(getattr(candidates[0], "finish_reason", "") or "")
    if reason and "STOP" not in reason.upper():
        raise RuntimeError(f"Risposta incompleta o bloccata: {reason}.")
    output = (response.text or "").strip()
    if not output:
        raise RuntimeError("Il modello ha restituito un testo vuoto.")
    return output


def strip_fences(text):
    return re.sub(r"^\s*```[a-z]*\s*\n|\n\s*```\s*$", "", text.strip(), flags=re.I)


def parse_json_list(raw):
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", raw.strip(), flags=re.I)
    start, end = cleaned.find("["), cleaned.rfind("]")
    if start == -1 or end == -1 or end < start:
        raise ValueError("Nessuna lista JSON nella risposta.")
    data = json.loads(cleaned[start:end + 1])
    if not isinstance(data, list):
        raise ValueError("Il JSON ricevuto non è una lista.")
    return [item for item in data if isinstance(item, dict)]


def parse_json_object(raw):
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", raw.strip(), flags=re.I)
    start, end = cleaned.find("{"), cleaned.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Nessun oggetto JSON nella risposta.")
    data = json.loads(cleaned[start:end + 1])
    if not isinstance(data, dict):
        raise ValueError("Il JSON ricevuto non è un oggetto.")
    return data


def grammar_system_prompt(language):
    return f"""Sei un correttore di bozze professionista di {language}. Segnala SOLO errori oggettivi e certi di ortografia, grammatica, concordanza, sintassi e punteggiatura. Non proporre interventi stilistici e non riscrivere frasi corrette.
Rispetta nomi propri, dialoghi, forme colloquiali intenzionali e voce dell'autore. Non cambiare le caporali con trattini.
Per ogni errore copia nel campo 'errore' una porzione ESATTA e contigua del TESTO DA CONTROLLARE (massimo 100 caratteri), e nel campo 'correzione' il solo frammento corretto. Il frammento deve essere sufficientemente lungo da individuare la posizione senza ambiguità.
Ignora intestazioni, numeri di pagina, piè di pagina e artefatti da estrazione PDF. Non segnalare frasi incomplete dovute soltanto al taglio del blocco. Prima di segnalare delimitatori non bilanciati usa anche il contesto: un dialogo può continuare oltre il blocco.
Rispondi ESCLUSIVAMENTE con una lista JSON. Ogni oggetto ha chiavi: errore, correzione, tipo (grammatica|ortografia|concordanza|punteggiatura|sintassi), spiegazione (massimo 20 parole). Se non ci sono errori certi rispondi []. Nessun Markdown e nessun commento."""


def grammar_user_prompt(chunk, total):
    return f"""BLOCCO {chunk['number']} DI {total}

CONTESTO PRECEDENTE — NON ANALIZZARE E NON CORREGGERE:
<<<CONTESTO_PRECEDENTE
{chunk['previous']}
CONTESTO_PRECEDENTE>>>

TESTO DA CONTROLLARE:
<<<TESTO_DA_CONTROLLARE
{chunk['text']}
TESTO_DA_CONTROLLARE>>>

CONTESTO SUCCESSIVO — NON ANALIZZARE E NON CORREGGERE:
<<<CONTESTO_SUCCESSIVO
{chunk['next']}
CONTESTO_SUCCESSIVO>>>"""


def verification_system_prompt(language):
    return f"""Sei un secondo revisore indipendente di {language}. Il tuo compito NON è trovare nuovi errori, ma VALIDARE le correzioni proposte da un altro revisore.
Una proposta è valida soltanto se: (1) il frammento originale contiene davvero un errore oggettivo; (2) la correzione elimina esattamente quell'errore; (3) non cambia stile, significato, nomi, fatti o intenzioni narrative; (4) è grammaticalmente corretta nel contesto completo.
Respingi proposte stilistiche, opinabili, inventate, dovute a tagli di pagina o blocco, oppure applicate a dialoghi e forme colloquiali intenzionali. Se hai un dubbio ragionevole, valido=false.
Rispondi ESCLUSIVAMENTE con una lista JSON di oggetti con chiavi: indice (numero intero), valido (booleano), motivo (massimo 15 parole). Includi tutti gli indici ricevuti. Nessun Markdown."""


def verify_candidates(chunk, candidates, model, language):
    payload = [{"indice": index, "errore": item["errore"], "correzione": item["correzione"],
                "spiegazione": item["spiegazione"]} for index, item in enumerate(candidates)]
    prompt = f"""TESTO ORIGINALE COMPLETO DEL BLOCCO:
<<<BLOCCO
{chunk['text']}
BLOCCO>>>

CONTESTO PRECEDENTE:
{chunk['previous']}

CONTESTO SUCCESSIVO:
{chunk['next']}

PROPOSTE DA VALIDARE:
{json.dumps(payload, ensure_ascii=False)}"""
    raw = ask_ai(prompt, verification_system_prompt(language), model, 0.0, json_mode=True)
    verdicts = {}
    for item in parse_json_list(raw):
        try:
            index = int(item.get("indice"))
        except (TypeError, ValueError):
            continue
        if 0 <= index < len(candidates):
            verdicts[index] = (item.get("valido") is True, str(item.get("motivo", ""))[:160])
    return verdicts


def check_grammar_ai(text, model, language, block_chars=6000):
    """Doppio passaggio: individuazione + validazione indipendente di ogni proposta."""
    chunks = make_chunks(text, block_chars)
    items, failed, discarded, unverified, checked = [], [], 0, 0, 0
    seen = set()
    bar = st.progress(0, text="Controllo grammaticale AI in due passaggi...")
    total = len(chunks)
    for index, chunk in enumerate(chunks, 1):
        if not chunk["text"].strip():
            checked += 1
            bar.progress(index / total, text=f"Blocco {index}/{total}: vuoto")
            continue
        candidates = []
        try:
            raw = ask_ai(grammar_user_prompt(chunk, total), grammar_system_prompt(language), model, 0.0, json_mode=True)
            for item in parse_json_list(raw):
                error = str(item.get("errore", "")).strip()
                fix = str(item.get("correzione", "")).strip()
                if not error or not fix or error == fix or len(error) > 100 or error not in chunk["text"]:
                    discarded += 1
                    continue
                occurrences = [m.start() for m in re.finditer(re.escape(error), chunk["text"])]
                if len(occurrences) != 1:
                    discarded += 1
                    continue
                signature = (chunk["start"] + occurrences[0], error, fix)
                if signature in seen:
                    discarded += 1
                    continue
                seen.add(signature)
                candidates.append({"pos": chunk["start"] + occurrences[0], "errore": error,
                                   "correzione": fix, "tipo": str(item.get("tipo", ""))[:30],
                                   "spiegazione": str(item.get("spiegazione", ""))[:200]})
        except QuotaExhaustedError as exc:
            failed.append(f"Blocco {index}, analisi: {str(exc)}")
            # Circuit breaker: non martellare l'API con tutti gli altri blocchi.
            for remaining in range(index + 1, total + 1):
                failed.append(f"Blocco {remaining}: non eseguito per quota Gemini esaurita.")
            bar.progress(1.0, text="Controllo fermato: quota Gemini esaurita")
            bar.empty()
            items.sort(key=lambda item: item["pos"])
            return {"sig": hashlib.sha1(text.encode("utf-8", "ignore")).hexdigest(), "items": items,
                    "blocks": total, "failed": failed, "discarded": discarded, "unverified": unverified,
                    "checked": checked, "complete": False, "quota_exhausted": True}
        except Exception as exc:  # noqa: BLE001
            failed.append(f"Blocco {index}, analisi: {str(exc)[:250]}")
            bar.progress(index / total, text=f"Blocco {index}/{total}: analisi fallita")
            continue

        if not candidates:
            checked += 1
            bar.progress(index / total, text=f"Blocco {index}/{total}: nessun candidato")
            continue
        try:
            verdicts = verify_candidates(chunk, candidates, model, language)
            if len(verdicts) != len(candidates):
                raise RuntimeError(f"Verifica incompleta: {len(verdicts)}/{len(candidates)} risposte.")
        except QuotaExhaustedError as exc:
            failed.append(f"Blocco {index}, verifica: {str(exc)}")
            for candidate in candidates:
                candidate.update(auto=False, fonte="AI non verificata",
                                 spiegazione="Verifica non eseguita per quota esaurita: controllo umano obbligatorio.")
                items.append(candidate)
                unverified += 1
            for remaining in range(index + 1, total + 1):
                failed.append(f"Blocco {remaining}: non eseguito per quota Gemini esaurita.")
            bar.progress(1.0, text="Controllo fermato: quota Gemini esaurita")
            bar.empty()
            items.sort(key=lambda item: item["pos"])
            return {"sig": hashlib.sha1(text.encode("utf-8", "ignore")).hexdigest(), "items": items,
                    "blocks": total, "failed": failed, "discarded": discarded, "unverified": unverified,
                    "checked": checked, "complete": False, "quota_exhausted": True}
        except Exception as exc:  # noqa: BLE001
            failed.append(f"Blocco {index}, verifica: {str(exc)[:250]}")
            for candidate in candidates:
                candidate.update(auto=False, fonte="AI non verificata",
                                 spiegazione="Verifica automatica non riuscita: controllo umano obbligatorio.")
                items.append(candidate)
                unverified += 1
            bar.progress(index / total, text=f"Blocco {index}/{total}: verifica fallita")
            continue

        for candidate_index, candidate in enumerate(candidates):
            valid, reason = verdicts[candidate_index]
            if valid:
                candidate.update(auto=True, fonte="AI verificata 2 passaggi",
                                 spiegazione=(candidate["spiegazione"] + (" · Verifica: " + reason if reason else ""))[:300])
                items.append(candidate)
            else:
                discarded += 1
        checked += 1
        bar.progress(index / total, text=f"Blocco {index}/{total}: analisi e verifica completate")
    bar.empty()
    items.sort(key=lambda item: item["pos"])
    return {"sig": hashlib.sha1(text.encode("utf-8", "ignore")).hexdigest(), "items": items,
            "blocks": total, "failed": failed, "discarded": discarded, "unverified": unverified,
            "checked": checked, "complete": not failed}


# ---------------------------------------------------------------------
# REGOLE LOCALI: nessun falso positivo per caporali divise tra pagine
# ---------------------------------------------------------------------
WORD_FIXES = {
    "perchè": "perché", "poichè": "poiché", "affinchè": "affinché", "benchè": "benché",
    "finchè": "finché", "cosicchè": "cosicché", "giacchè": "giacché", "pultroppo": "purtroppo",
    "sopratutto": "soprattutto", "propio": "proprio", "daccordo": "d'accordo",
    "accellerare": "accelerare", "eccezzione": "eccezione", "qual'è": "qual è",
    "quì": "qui", "quà": "qua", "fà": "fa", "stà": "sta", "pò": "po'", "sè": "sé",
}
_WORD_FIX_RE = re.compile(r"(?<![\w'’])(" + "|".join(re.escape(k) for k in sorted(WORD_FIXES, key=len, reverse=True)) + r")(?!\w)", re.I)
_BAD_ELISION_RE = re.compile(r"\b(un|l|dell|nell|all|sull|quest|quell|d|dov|com)['’] +(?=[a-zà-ÿ])", re.I)
_REPEAT_OK = {"bene", "piano", "così", "forte", "presto", "adagio", "poco", "giù", "sempre", "tanto", "pian"}


def _case_like(source, fix):
    if len(source) > 1 and source.isupper():
        return fix.upper()
    return fix[0].upper() + fix[1:] if source[:1].isupper() else fix


def local_grammar_check(text, limit=3000):
    found = []

    def add(match, fix, kind, explanation, auto=True):
        found.append({"pos": match.start(), "errore": match.group(0), "correzione": fix,
                      "tipo": kind, "spiegazione": explanation, "auto": auto, "fonte": "regole"})

    for match in _WORD_FIX_RE.finditer(text):
        add(match, _case_like(match.group(1), WORD_FIXES[match.group(1).lower()]), "ortografia", "Forma errata: accento, apostrofo o grafia.")
    for match in re.finditer(r"\bun po(?![\w'’])", text, re.I):
        add(match, match.group(0) + "'", "ortografia", "«un po'» richiede l'apostrofo di troncamento.")
    for match in _BAD_ELISION_RE.finditer(text):
        add(match, match.group(0).rstrip(), "punteggiatura", "Dopo l'apostrofo di elisione non va lo spazio.")
    for match in re.finditer(r"[A-Za-zÀ-ÿ]{2,}[,;:][A-Za-zÀ-ÿ]{2,}", text):
        punct = re.search(r"[,;:]", match.group(0))
        add(match, match.group(0)[:punct.end()] + " " + match.group(0)[punct.end():], "punteggiatura", "Manca lo spazio dopo la punteggiatura.")
    for match in re.finditer(r"[,;]{2,}", text):
        add(match, match.group(0)[0], "punteggiatura", "Punteggiatura doppia.")
    for match in re.finditer(r"(?<![.])\.\.(?![.])", text):
        add(match, ".", "punteggiatura", "Due punti consecutivi: usa il punto singolo o i puntini di sospensione.")
    for match in re.finditer(r"(?<=\w) +([,;:!?]|\.(?!\.))", text):
        add(match, match.group(1), "punteggiatura", "In italiano non va lo spazio prima della punteggiatura.")
    for match in re.finditer(r"\b(\w{3,})\s+\1\b", text, re.I):
        if match.group(1).lower() not in _REPEAT_OK:
            add(match, match.group(1), "grammatica", "Parola ripetuta consecutivamente: verifica che non sia intenzionale.", auto=False)

    # Parentesi: controllo per paragrafo. Caporali: soltanto bilancio GLOBALE.
    pos = 0
    for para in re.split(r"(\n\s*\n)", text):
        if para.count("(") != para.count(")") and para.strip():
            excerpt = para.strip()[:60]
            found.append({"pos": pos + para.index(excerpt), "errore": excerpt, "correzione": "—",
                          "tipo": "punteggiatura", "spiegazione": "Parentesi non bilanciate nel paragrafo: verifica manuale.",
                          "auto": False, "fonte": "regole"})
        pos += len(para)
    if text.count("«") != text.count("»"):
        found.append({"pos": 0, "errore": "Bilancio complessivo delle caporali", "correzione": "—",
                      "tipo": "punteggiatura",
                      "spiegazione": f"Nel documento: {text.count('«')} aperture e {text.count('»')} chiusure. Verifica manuale sull'intero testo.",
                      "auto": False, "fonte": "regole"})
    found.sort(key=lambda item: item["pos"])
    output, last_end = [], -1
    for item in found:
        if item["pos"] >= last_end or not item["auto"]:
            output.append(item)
            if item["auto"]:
                last_end = item["pos"] + len(item["errore"])
    return output[:limit]


def merge_items(ai_items, local_items):
    merged = list(ai_items)
    for item in local_items:
        start, end = item["pos"], item["pos"] + len(item["errore"])
        if not any(start < other["pos"] + len(other["errore"]) and other["pos"] < end for other in merged):
            merged.append(item)
    return sorted(merged, key=lambda item: item["pos"])


def correction_is_safe(item, text):
    """Ultimo controllo deterministico prima di modificare il documento."""
    if not item.get("auto", False):
        return False
    error, fix = item.get("errore", ""), item.get("correzione", "")
    start = item.get("pos", -1)
    if not error or not fix or start < 0 or text[start:start + len(error)] != error:
        return False
    if len(fix) > 200 or len(fix) > max(40, len(error) * 4):
        return False
    if "```" in fix or CONTROL_RE.search(fix) or "\n" in fix or "\r" in fix:
        return False
    if Counter(re.findall(r"\d+", error)) != Counter(re.findall(r"\d+", fix)):
        return False
    if (error.count("«"), error.count("»")) != (fix.count("«"), fix.count("»")):
        return False
    return True


# ---------------------------------------------------------------------
# REVISIONE AI DEL TESTO
# ---------------------------------------------------------------------
def build_revision_system(task, language, intensity, preserve_style):
    style = "Conserva rigorosamente voce, tono, significato, fatti, trama, nomi propri e ordine degli eventi." if preserve_style else "Puoi riformulare per chiarezza senza cambiare fatti o significato."
    return f"Sei un revisore editoriale professionista. Lingua: {language}. Intervento: {task}. Intensità: {intensity}. {style} Correggi soltanto il testo compreso tra i delimitatori TESTO_DA_REVISIONARE. Usa il contesto soltanto per evitare tagli o interpretazioni errate. Non modificare numeri, date e nomi propri. Non riassumere, non aggiungere commenti e non usare Markdown. Mantieni paragrafi e interruzioni di riga. Restituisci esclusivamente il testo revisionato."


def validate_block(source, output, intensity):
    low, high = {"Leggera": (0.92, 1.08), "Standard": (0.85, 1.15), "Approfondita": (0.75, 1.30)}[intensity]
    notes = []
    ratio = len(output) / max(1, len(source))
    if not low <= ratio <= high:
        notes.append(f"lunghezza {ratio:.0%} dell'originale")
    if Counter(re.findall(r"\d+", source)) != Counter(re.findall(r"\d+", output)):
        notes.append("numeri modificati")
    if "```" in output or re.match(r"\s*(ecco|certo[,!]|sure[,!])", output, re.I):
        notes.append("commento o Markdown dell'AI")
    return ratio, notes


def run_revision(text, ranges, task, language, intensity, preserve_style, model, strict):
    chunks = []
    for range_index, (start, end) in enumerate(ranges):
        for chunk in make_chunks(text[start:end], 9000):
            chunks.append((range_index, start + chunk["start"], start + chunk["end"], chunk))
    outputs = {index: [] for index in range(len(ranges))}
    log = []
    quota_stop = False
    bar = st.progress(0, text="Revisione AI...")
    for number, (range_index, start, end, chunk) in enumerate(chunks, 1):
        source = chunk["text"]
        final = source
        status, notes, ratio = "", [], None
        if quota_stop:
            status, notes = "non eseguito (quota esaurita)", ["Testo originale conservato senza chiamare l'API."]
            outputs[range_index].append(final)
            log.append({"blocco": number, "stato": status, "lunghezza_%": None, "note": "; ".join(notes)})
            bar.progress(number / len(chunks), text=f"Blocco {number}/{len(chunks)}: saltato per quota")
            continue
        try:
            prompt = f"CONTESTO PRECEDENTE (non revisionare):\n{chunk['previous']}\n\n<<<TESTO_DA_REVISIONARE\n{source}\nTESTO_DA_REVISIONARE>>>\n\nCONTESTO SUCCESSIVO (non revisionare):\n{chunk['next']}"
            raw = ask_ai(prompt, build_revision_system(task, language, intensity, preserve_style), model, 0.2)
            candidate = source[:len(source) - len(source.lstrip())] + strip_fences(raw).strip() + source[len(source.rstrip()):]
            ratio, notes = validate_block(source, candidate, intensity)
            if notes and strict:
                status = "scartato (originale conservato)"
            else:
                status = "accettato con avviso" if notes else "accettato"
                final = candidate
        except QuotaExhaustedError as exc:
            status, notes, quota_stop = "errore quota (originale conservato)", [str(exc)], True
        except Exception as exc:  # noqa: BLE001
            status, notes = "errore (originale conservato)", [str(exc)[:250]]
        outputs[range_index].append(final)
        log.append({"blocco": number, "stato": status, "lunghezza_%": None if ratio is None else round(ratio * 100), "note": "; ".join(notes)})
        bar.progress(number / len(chunks), text=f"Revisione blocco {number}/{len(chunks)}")
    bar.empty()
    new_text = text
    for index in reversed(range(len(ranges))):
        start, end = ranges[index]
        new_text = new_text[:start] + "".join(outputs[index]) + new_text[end:]
    return new_text, log


# ---------------------------------------------------------------------
# ESPORTAZIONE
# ---------------------------------------------------------------------
def sanitize(text):
    return XML_BAD_RE.sub("", text)


def build_docx(text, title, author, extra):
    doc = Document()
    if title:
        doc.add_heading(sanitize(title), 0)
    heading_lines = {chapter["line"] for chapter in detect_chapters(text, extra)}
    for index, line in enumerate(text.split("\n")):
        stripped = sanitize(line).strip()
        if not stripped:
            continue
        if index in heading_lines:
            doc.add_heading(stripped.lstrip("# ").strip(), level=1)
        else:
            doc.add_paragraph(stripped)
    doc.core_properties.title = title or ""
    doc.core_properties.author = author or ""
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def build_pdf(text, title, author, extra):
    output = io.BytesIO()
    pdf = SimpleDocTemplate(output, pagesize=A4, title=title or "", author=author or "",
                            rightMargin=20 * mm, leftMargin=20 * mm, topMargin=20 * mm, bottomMargin=20 * mm)
    styles = getSampleStyleSheet()
    story = [Paragraph(escape(sanitize(title)), styles["Title"]), Spacer(1, 12)] if title else []
    heading_lines = {chapter["line"] for chapter in detect_chapters(text, extra)}
    for index, line in enumerate(text.split("\n")):
        stripped = sanitize(line).strip()
        if not stripped:
            continue
        if index in heading_lines:
            if story:
                story.append(PageBreak())
            story.append(Paragraph(escape(stripped.lstrip("# ").strip()), styles["Heading1"]))
        else:
            story.append(Paragraph(escape(stripped), styles["BodyText"]))
            story.append(Spacer(1, 4))
    pdf.build(story or [Paragraph(" ", styles["BodyText"])],
              onFirstPage=lambda c, d: c.drawCentredString(A4[0] / 2, 10 * mm, str(d.page)),
              onLaterPages=lambda c, d: c.drawCentredString(A4[0] / 2, 10 * mm, str(d.page)))
    return output.getvalue()


def build_epub(text, title, author, language, extra):
    book = epub.EpubBook()
    book.set_identifier(f"libromaster-{datetime.now(timezone.utc).timestamp()}")
    book.set_title(title or "Libro revisionato")
    book.set_language(language)
    book.add_author(author or "Autore")
    chapters = detect_chapters(text, extra)
    spans = chapter_spans(text, chapters) or [{"title": title or "Testo", "start": 0, "end": len(text), "front": True}]
    items = []
    for index, span in enumerate(spans, 1):
        segment = text[span["start"]:span["end"]]
        body = []
        for line_index, line in enumerate(line for line in segment.split("\n") if line.strip()):
            clean = html.escape(sanitize(line.strip()))
            if line_index == 0 and not span["front"]:
                body.append(f"<h1>{html.escape(span['title'].lstrip('# ').strip())}</h1>")
            else:
                body.append(f"<p>{clean}</p>")
        item = epub.EpubHtml(title=span["title"][:60], file_name=f"cap_{index:03d}.xhtml", lang=language)
        item.content = "\n".join(body) or "<p> </p>"
        book.add_item(item)
        items.append(item)
    book.toc = tuple(items)
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    book.spine = ["nav"] + items
    output = io.BytesIO()
    epub.write_epub(output, book)
    return output.getvalue()


# ---------------------------------------------------------------------
# BANDIERA ITALIANA ANIMATA — nessuna risorsa esterna
# ---------------------------------------------------------------------
# ---------------------------------------------------------------------
# BANDIERA FOTOREALISTICA INCORPORATA NEL FILE
# ---------------------------------------------------------------------
FLAG_IMAGE_BASE64 = "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"

def render_italian_flag():
    encoded = FLAG_IMAGE_BASE64

    components.html(f"""
<div class="flag-wrap" role="img" aria-label="Fotografia animata della bandiera italiana in tessuto">
  <canvas id="flag-canvas"></canvas>
  <img id="flag-fallback" src="data:image/jpeg;base64,{encoded}" alt="Bandiera italiana in tessuto">
</div>
<style>
  html,body {{ margin:0; padding:0; background:transparent; overflow:hidden; }}
  .flag-wrap {{ width:100%; height:164px; position:relative; border-radius:10px; overflow:hidden;
    box-shadow:0 5px 15px rgba(0,0,0,.20); background:#07111f; }}
  #flag-canvas,#flag-fallback {{ width:100%; height:164px; display:block; object-fit:cover; }}
  #flag-fallback {{ display:none; }}
</style>
<script>
(function() {{
  const canvas = document.getElementById('flag-canvas');
  const fallback = document.getElementById('flag-fallback');
  const showFallback = function() {{ canvas.style.display='none'; fallback.style.display='block'; }};
  const gl = canvas.getContext('webgl', {{ alpha:false, antialias:true, preserveDrawingBuffer:false }});
  if (!gl) {{ showFallback(); return; }}

  const vertexSource = `
    attribute vec2 a_position;
    varying vec2 v_uv;
    void main() {{
      v_uv = a_position * 0.5 + 0.5;
      gl_Position = vec4(a_position, 0.0, 1.0);
    }}`;
  const fragmentSource = `
    precision mediump float;
    varying vec2 v_uv;
    uniform sampler2D u_image;
    uniform float u_time;
    void main() {{
      vec2 uv = v_uv;
      float anchor = smoothstep(0.15, 0.42, uv.x);
      float primary = sin(uv.x * 17.0 - u_time * 2.7 + uv.y * 2.4);
      float secondary = sin(uv.x * 31.0 - u_time * 4.1 - uv.y * 3.0);
      uv.y += (primary * 0.009 + secondary * 0.0025) * anchor;
      uv.x += sin(uv.x * 13.0 - u_time * 2.1) * 0.0025 * anchor;
      vec4 color = texture2D(u_image, clamp(uv, 0.001, 0.999));
      float light = 1.0 + anchor * (primary * 0.045 + secondary * 0.018);
      color.rgb *= light;
      gl_FragColor = color;
    }}`;

  function compileShader(type, source) {{
    const shader = gl.createShader(type);
    gl.shaderSource(shader, source); gl.compileShader(shader);
    if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) {{
      const message = gl.getShaderInfoLog(shader); gl.deleteShader(shader); throw new Error(message);
    }}
    return shader;
  }}

  let program;
  try {{
    program = gl.createProgram();
    gl.attachShader(program, compileShader(gl.VERTEX_SHADER, vertexSource));
    gl.attachShader(program, compileShader(gl.FRAGMENT_SHADER, fragmentSource));
    gl.linkProgram(program);
    if (!gl.getProgramParameter(program, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(program));
  }} catch (error) {{ showFallback(); return; }}
  gl.useProgram(program);

  const buffer = gl.createBuffer();
  gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
  gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1,-1, 1,-1, -1,1, -1,1, 1,-1, 1,1]), gl.STATIC_DRAW);
  const position = gl.getAttribLocation(program, 'a_position');
  gl.enableVertexAttribArray(position); gl.vertexAttribPointer(position, 2, gl.FLOAT, false, 0, 0);
  const timeLocation = gl.getUniformLocation(program, 'u_time');

  const texture = gl.createTexture();
  gl.bindTexture(gl.TEXTURE_2D, texture);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
  gl.pixelStorei(gl.UNPACK_FLIP_Y_WEBGL, true);

  function resize() {{
    const ratio = Math.min(window.devicePixelRatio || 1, 2);
    const width = Math.max(1, Math.floor(canvas.clientWidth * ratio));
    const height = Math.max(1, Math.floor(canvas.clientHeight * ratio));
    if (canvas.width !== width || canvas.height !== height) {{
      canvas.width = width; canvas.height = height; gl.viewport(0, 0, width, height);
    }}
  }}
  window.addEventListener('resize', resize);

  const image = new Image();
  let ready = false, animationId = 0;
  const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  image.onload = function() {{
    gl.bindTexture(gl.TEXTURE_2D, texture);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, image);
    ready = true;
    function draw(milliseconds) {{
      resize();
      if (ready) {{
        gl.uniform1f(timeLocation, reducedMotion ? 0 : milliseconds * 0.001);
        gl.drawArrays(gl.TRIANGLES, 0, 6);
      }}
      if (!reducedMotion) animationId = requestAnimationFrame(draw);
    }}
    animationId = requestAnimationFrame(draw);
  }};
  image.onerror = showFallback;
  image.src = 'data:image/jpeg;base64,{encoded}';
  document.addEventListener('visibilitychange', function() {{
    if (document.hidden) cancelAnimationFrame(animationId);
    else if (!reducedMotion && ready) animationId = requestAnimationFrame(function restart(t) {{
      gl.uniform1f(timeLocation, t * 0.001); gl.drawArrays(gl.TRIANGLES, 0, 6);
      animationId = requestAnimationFrame(restart);
    }});
  }});
}})();
</script>
""", height=170)


# ---------------------------------------------------------------------
# INTERFACCIA
# ---------------------------------------------------------------------
st.markdown("""<style>
.stApp{background:#f4f6fb}.stButton>button,.stDownloadButton>button{border:1px solid #c9a227;border-radius:10px;color:#0a1f44}
div[data-testid='stMetric']{background:white;border-left:4px solid #c9a227;border-radius:10px;padding:10px}
</style>""", unsafe_allow_html=True)
st.title("📚 LIBRO MASTER AI")
st.caption("Revisione editoriale con controlli verificabili, blocchi AI tracciati e nessuna correzione silenziosa.")
show_flashes()

with st.sidebar:
    render_italian_flag()
    st.header("Impostazioni")
    model_options = list(dict.fromkeys([DEFAULT_MODEL, "gemini-3.8-flash"]))
    selected_model = st.selectbox("Modello AI", model_options, index=0)
    model = st.text_input("Nome modello effettivo", value=selected_model,
                          help="Modificabile se Google cambia nuovamente il nome del modello.")
    language = st.selectbox("Lingua", list(LANG_CODES))
    intensity = st.select_slider("Intensità revisione", ["Leggera", "Standard", "Approfondita"], value="Standard")
    preserve_style = st.toggle("Mantieni stile e voce dell'autore", value=True)
    strict = st.toggle("Protezione anti-alterazione AI", value=True)
    custom_pat = st.text_input("Pattern capitoli personalizzato (regex)")
    if custom_pat:
        try:
            re.compile(custom_pat)
        except re.error as exc:
            st.error(f"Regex non valida: {exc}")
            custom_pat = ""
    key_input = st.text_input("GOOGLE_API_KEY", type="password", help="Non salvare la chiave nel codice.")
    if key_input:
        st.session_state.google_api_key = key_input
    st.success("Chiave AI configurata") if get_api_key() else st.warning("Chiave AI non configurata")
    if st.button("Testa connessione AI", disabled=not get_api_key()):
        _, error = safe_call(ask_ai, "Rispondi solo OK.", "Rispondi brevemente.", model)
        st.error(str(error)) if error else st.success(f"Connessione riuscita con {model}.")
    st.divider()
    if st.button("🗑️ Ripulisci tutto"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_state()
        shutil.rmtree(VERSIONS_DIR, ignore_errors=True)
        st.rerun()

text_now = get_text()
info = st.session_state.doc_info
original = st.session_state.original_text
chapters_now = detect_chapters(text_now, custom_pat)

tabs = st.tabs(["✍️ Importa", "🧪 Verifica", "🧹 Pulizia", "✨ Revisione AI", "📝 Errori", "🔎 SEO", "🕘 Versioni", "⬇️ Esporta"])

with tabs[0]:
    uploaded = st.file_uploader("Carica PDF, EPUB, DOCX, TXT o MD", type=["pdf", "epub", "docx", "txt", "md"], key=f"uploader_{st.session_state.uploader_gen}")
    if st.button("Importa e verifica", type="primary", disabled=uploaded is None):
        result, error = safe_call(extract_document, uploaded)
        if error:
            st.error(str(error))
        elif not result[0].strip():
            st.error("Nessun testo estraibile: per una scansione serve l'OCR.")
        else:
            extracted, new_info = result
            st.session_state.original_text = extracted
            st.session_state.doc_info = new_info
            st.session_state.filename = os.path.splitext(uploaded.name)[0]
            st.session_state.revision_log = []
            apply_text(extracted, "Importazione originale", f"Importato {uploaded.name}: {new_info['parole']:,} parole.")
    if st.button("Ripristina originale", disabled=not bool(original)):
        apply_text(original, "Ripristino originale")
    edited = st.text_area("Testo completo", height=480, key="book_editor")
    col1, col2, col3 = st.columns(3)
    col1.metric("Parole", f"{count_words(edited):,}")
    col2.metric("Caratteri", f"{len(edited):,}")
    col3.metric("Capitoli", len(chapters_now))

with tabs[1]:
    if not text_now:
        st.info("Importa un documento.")
    else:
        st.write(f"**Parole:** {count_words(text_now):,} · **Capitoli:** {len(chapters_now)}")
        if info.get("formato") == "PDF":
            st.write(f"**Pagine PDF reali:** {info['pagine_reali']} · **senza testo:** {len(info['pagine_vuote'])}")
            if info.get("tabella_pagine"):
                st.dataframe(pd.DataFrame(info["tabella_pagine"]), hide_index=True)
        if chapters_now:
            rows = []
            for span in chapter_spans(text_now, chapters_now):
                rows.append({"Titolo": span["title"], "Parole": count_words(text_now[span["start"]:span["end"]])})
            st.dataframe(pd.DataFrame(rows), hide_index=True)
        st.code(info.get("sha256", "Nessun file importato"), language=None)

with tabs[2]:
    st.subheader("Pulizia tecnica, senza AI")
    labels = {"ctrl": "Rimuovi caratteri invisibili", "pagenum": "Rimuovi numeri di pagina isolati",
              "headers": "Rimuovi intestazioni ripetute", "dehyph": "Ricongiungi parole con trattino",
              "spaces": "Elimina spazi doppi", "punct": "Correggi spazi prima della punteggiatura",
              "blank": "Riduci righe vuote multiple"}
    options = {key: st.checkbox(label, value=True, key=f"clean_{key}") for key, label in labels.items()}
    if st.button("Applica pulizia", disabled=not text_now):
        cleaned, report = clean_text(text_now, options)
        save_version("Backup prima della pulizia", text_now)
        apply_text(cleaned, "Pulizia automatica", json.dumps(report, ensure_ascii=False))

with tabs[3]:
    task = st.selectbox("Tipo di revisione", ["Solo grammatica e ortografia", "Correzione completa", "Punteggiatura e sintassi", "Chiarezza e leggibilità"])
    spans = chapter_spans(text_now, chapters_now)
    selected = st.multiselect("Capitoli (lascia vuoto per tutto il testo)", [span["title"] for span in spans])
    ranges = [(span["start"], span["end"]) for span in spans if span["title"] in selected] if selected else ([(0, len(text_now))] if text_now else [])
    calls = sum(len(make_chunks(text_now[a:b], 9000)) for a, b in ranges)
    st.write(f"Blocchi da inviare: **{calls}**")
    if st.button("Avvia revisione AI", type="primary", disabled=not calls):
        if not get_api_key():
            st.error("Configura la chiave API.")
        else:
            save_version("Backup prima della revisione", text_now)
            result, error = safe_call(run_revision, text_now, ranges, task, language, intensity, preserve_style, model, strict)
            if error:
                st.error(str(error))
            else:
                new_text, log = result
                st.session_state.revision_log = log
                if new_text == text_now:
                    st.warning("Nessuna modifica applicata.")
                else:
                    apply_text(new_text, f"Revisione AI: {task}", "Revisione completata. Controlla il report.")
    if st.session_state.revision_log:
        log = st.session_state.revision_log
        accepted = sum(item["stato"].startswith("accettato") for item in log)
        st.write(f"**{accepted}/{len(log)} blocchi applicati; {len(log)-accepted} conservati originali.**")
        st.dataframe(pd.DataFrame(log), hide_index=True)

with tabs[4]:
    st.subheader("Controllo errori senza modificare il testo")
    signature_now = hashlib.sha1(text_now.encode("utf-8", "ignore")).hexdigest()
    block_count = len(make_chunks(text_now, 6000)) if text_now else 0
    left, right = st.columns(2)
    if left.button("⚡ Solo regole locali", disabled=not text_now or language != "italiano"):
        st.session_state.grammar = {"sig": signature_now, "items": local_grammar_check(text_now), "blocks": 0,
                                    "failed": [], "discarded": 0, "unverified": 0, "checked": 0, "complete": True, "source": "regole locali"}
    if right.button("📝 AI + regole locali", type="primary", disabled=not block_count):
        if not get_api_key():
            st.error("Configura la chiave API.")
        else:
            result, error = safe_call(check_grammar_ai, text_now, model, language)
            if error:
                st.error(str(error))
            else:
                result["items"] = merge_items(result["items"], local_grammar_check(text_now) if language == "italiano" else [])
                result["source"] = "AI + regole locali" if language == "italiano" else "AI"
                st.session_state.grammar = result
    grammar = st.session_state.grammar
    if grammar:
        st.markdown(f"**Risultato · fonte: {grammar.get('source', 'AI')}**")
        if grammar["sig"] != signature_now:
            st.warning("Il testo è cambiato: rilancia il controllo.")
        confirmed_count = sum(bool(item.get("auto", False)) for item in grammar["items"])
        review_count = len(grammar["items"]) - confirmed_count
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Errori confermati", confirmed_count)
        m2.metric("Blocchi AI controllati", "—" if not grammar["blocks"] else f"{grammar['checked']}/{grammar['blocks']}")
        m3.metric("Proposte AI scartate", grammar["discarded"])
        m4.metric("Errori confermati / 1000 parole", round(confirmed_count / max(1, count_words(text_now)) * 1000, 1))
        if review_count:
            st.info(f"{review_count} segnalazioni richiedono verifica umana e NON possono essere applicate automaticamente.")
        if grammar.get("quota_exhausted"):
            st.error("QUOTA GEMINI ESAURITA — controllo fermato subito per non sprecare altre chiamate. "
                     "Il testo non è stato modificato. Controlla piano e fatturazione su Google AI Studio "
                     "oppure attendi il reset della quota. Le regole locali restano disponibili.")
            st.link_button("Apri la guida Google su quote e limiti", "https://ai.google.dev/gemini-api/docs/rate-limits")
            with st.expander("Dettaglio dei blocchi non analizzati"):
                st.write("\n".join(grammar["failed"]))
        elif grammar["failed"]:
            st.error("Controllo AI INCOMPLETO: " + " · ".join(grammar["failed"][:8]))
        elif grammar["blocks"]:
            st.success("Tutti i blocchi AI sono stati effettivamente controllati.")
        rows = []
        offsets = info.get("offset_pagine") if info.get("formato") == "PDF" and text_now == original else None
        for item in grammar["items"]:
            row = {"Applica": False}
            if offsets:
                row["Pagina PDF"] = bisect_right(offsets, item["pos"])
            previous = [chapter["title"] for chapter in chapters_now if chapter["start"] <= item["pos"]]
            row["Capitolo"] = previous[-1][:50] if previous else "—"
            row.update({"Fonte": item["fonte"], "Applicabile": "Sì" if item.get("auto", False) else "No, revisione umana",
                        "Tipo": item["tipo"], "Errore": item["errore"],
                        "Correzione": item["correzione"], "Spiegazione": item["spiegazione"]})
            rows.append(row)
        if rows:
            frame = pd.DataFrame(rows)
            edited_frame = st.data_editor(frame, hide_index=True, key="grammar_editor",
                                          disabled=[column for column in frame.columns if column != "Applica"])
            st.download_button("Scarica errori CSV", frame.drop(columns=["Applica"]).to_csv(index=False).encode("utf-8-sig"), "errori_grammaticali.csv", "text/csv")
            chosen = [index for index, value in enumerate(edited_frame["Applica"].tolist()) if value]
            if st.button(f"Applica {len(chosen)} correzioni", disabled=not chosen):
                if grammar["sig"] != signature_now:
                    st.error("Testo cambiato: rilancia il controllo.")
                else:
                    new_text, done, skipped = text_now, 0, 0
                    for index in sorted(chosen, key=lambda i: grammar["items"][i]["pos"], reverse=True):
                        item = grammar["items"][index]
                        start, end = item["pos"], item["pos"] + len(item["errore"])
                        if correction_is_safe(item, new_text):
                            new_text = new_text[:start] + item["correzione"] + new_text[end:]
                            done += 1
                        else:
                            skipped += 1
                    apply_text(new_text, "Correzioni grammaticali", f"{done} correzioni applicate; {skipped} saltate.")

with tabs[5]:
    seo_title = st.text_input("Titolo", key="seo_title")
    seo_author = st.text_input("Autore", key="seo_author")
    if st.button("Genera metadati SEO", disabled=not text_now):
        if not get_api_key():
            st.error("Configura la chiave API.")
        else:
            sample = text_now if len(text_now) <= 14000 else text_now[:6000] + "\n[...]\n" + text_now[len(text_now)//2:len(text_now)//2+4000] + "\n[...]\n" + text_now[-4000:]
            prompt = f"Crea metadati editoriali per questo libro. Titolo: {seo_title}. Autore: {seo_author}. Testo: {sample}\nRestituisci solo JSON con descrizione_breve, descrizione_lunga, meta_title (massimo 60 caratteri), meta_description (massimo 160 caratteri), keyword_principali (7 elementi), tag (12 elementi). Non inventare premi o contenuti."
            raw, error = safe_call(ask_ai, prompt, "Sei un consulente editoriale SEO. Rispondi solo con JSON valido.", model)
            if error:
                st.error(str(error))
            else:
                parsed, parse_error = safe_call(parse_json_object, raw)
                if parse_error:
                    st.error(str(parse_error))
                else:
                    st.session_state.metadata = parsed
    if st.session_state.metadata:
        st.json(st.session_state.metadata)

with tabs[6]:
    if not st.session_state.versions:
        st.info("Nessuna versione salvata.")
    for index, version in reversed(list(enumerate(st.session_state.versions))):
        with st.expander(f"{version['label']} — {version['created_at']}"):
            if st.button("Ripristina", key=f"restore_{index}"):
                apply_text(version["text"], f"Ripristino: {version['label']}")

with tabs[7]:
    if not text_now:
        st.info("Importa un testo.")
    else:
        export_title = st.text_input("Titolo esportazione", value=st.session_state.seo_title or st.session_state.filename)
        export_author = st.text_input("Autore esportazione", value=st.session_state.seo_author)
        st.download_button("TXT", text_now.encode("utf-8"), f"{st.session_state.filename}_revisionato.txt")
        builders = {
            "DOCX": (lambda: build_docx(text_now, export_title, export_author, custom_pat), "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            "PDF": (lambda: build_pdf(text_now, export_title, export_author, custom_pat), "pdf", "application/pdf"),
            "EPUB": (lambda: build_epub(text_now, export_title, export_author, LANG_CODES[language], custom_pat), "epub", "application/epub+zip"),
        }
        for label, (builder, extension, mime) in builders.items():
            if st.button(f"Prepara {label}", key=f"prepare_{extension}"):
                data, error = safe_call(builder)
                if error:
                    st.error(str(error))
                else:
                    st.download_button(f"Scarica {label}", data, f"{st.session_state.filename}_revisionato.{extension}", mime, key=f"download_{extension}")

st.caption("L'AI può sbagliare. Un controllo è completo soltanto quando tutti i blocchi risultano effettivamente analizzati.")
