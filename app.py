import io
import os
import re
import json
import html
from datetime import datetime, timezone
from xml.sax.saxutils import escape

import streamlit as st
from openai import OpenAI
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
.stApp { background: #f7f8fc; }
.hero {
  padding: 1.6rem 1.8rem; border-radius: 18px; color: white;
  background: linear-gradient(120deg, #111827 0%, #312e81 55%, #6d28d9 100%);
  margin-bottom: 1.2rem;
}
.hero h1 { color: white; margin: 0; font-size: 2.1rem; }
.hero p { color: #e0e7ff; margin: .4rem 0 0 0; }
div[data-testid="stMetric"] {
  background: white; border: 1px solid #e5e7eb;
  border-radius: 12px; padding: 12px;
}
</style>
""", unsafe_allow_html=True)

EDITOR_KEY = "book_editor"

def init_state():
    defaults = {
        "original_text": "",
        "working_text": "",
        "versions": [],
        "metadata": {},
        "filename": "libro",
        "last_upload": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_state()

def sync_editor():
    """Aggiorna la text_area dopo modifiche programmatiche del testo."""
    st.session_state[EDITOR_KEY] = st.session_state.working_text

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
# AI E REVISIONE
# ------------------------------------------------------------
def get_api_key() -> str:
    try:
        return st.secrets.get("OPENAI_API_KEY", "") or os.getenv("OPENAI_API_KEY", "")
    except Exception:
        return os.getenv("OPENAI_API_KEY", "")

def get_client():
    key = get_api_key()
    if not key:
        return None
    return OpenAI(api_key=key)

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
    client = get_client()
    if client is None:
        raise RuntimeError(
            "Chiave API non configurata. Inserisci OPENAI_API_KEY nei Secrets "
            "di Streamlit oppure come variabile d'ambiente."
        )
    response = client.chat.completions.create(
        model=model,
        temperature=temperature,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
    )
    return response.choices[0].message.content.strip()

def save_version(label: str):
    st.session_state.versions.append({
        "label": label,
        "text": st.session_state.working_text,
        "created_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    })

def run_revision(text: str, task: str, language: str, intensity: str,
                 preserve_style: bool, model: str):
    pieces = split_text(text)
    output = []
    progress = st.progress(0, text="Preparazione revisione...")
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
        output.append(ask_ai(piece, system, model))
        progress.progress(
            (index + 1) / len(pieces),
            text=f"Revisione blocco {index + 1}/{len(pieces)}",
        )
    progress.empty()
    # Unisce con "\n" per preservare esattamente gli a capo originales
    return "\n".join(output)

def parse_json_response(raw: str) -> dict:
    """Estrae il JSON anche se l'AI lo racchiude in blocchi markdown."""
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", raw.strip(), flags=re.I)
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Nessun oggetto JSON valido trovato nella risposta.")
    return json.loads(cleaned[start:end + 1])

# ------------------------------------------------------------
# INTERFACCIA
# ------------------------------------------------------------
st.markdown("""
<div class="hero">
  <h1>📚 LIBRO MASTER AI</h1>
  <p>La tua redazione editoriale intelligente: correggi, migliora, organizza e prepara i tuoi libri.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Impostazioni")
    model = st.selectbox("Modello AI", ["gpt-4.1-mini", "gpt-4.1"], index=0)
    language = st.selectbox("Lingua", ["italiano", "inglese", "francese", "spagnolo", "tedesco"])
    intensity = st.select_slider("Intensità revisione", ["Leggera", "Standard", "Approfondita"], value="Standard")
    preserve_style = st.toggle("Mantieni stile e voce dell'autore", value=True)
    if get_api_key():
        st.success("Chiave AI configurata")
    else:
        st.warning("Chiave AI non configurata")

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
            try:
                extracted = extract_document(uploaded)
                if not extracted.strip():
                    st.error("Non è stato trovato testo. Se il PDF è una scansione, serve OCR.")
                else:
                    st.session_state.original_text = extracted
                    st.session_state.working_text = extracted
                    st.session_state.filename = os.path.splitext(uploaded.name)[0]
                    st.session_state.last_upload = uploaded.name
                    save_version("Importazione originale")
                    sync_editor()  # mostra subito il testo importato nell'editor
                    st.success(f"Importato: {uploaded.name} ({len(extracted):,} caratteri)")
            except Exception as error:
                st.error(f"Errore durante l'importazione: {error}")
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
    if st.button("✨ Avvia correzione AI", type="primary", disabled=not bool(st.session_state.working_text)):
        try:
            save_version("Backup prima della revisione")
            with st.spinner("L'AI sta revisionando il testo..."):
                revised = run_revision(
                    st.session_state.working_text, task, language,
                    intensity, preserve_style, model
                )
            st.session_state.working_text = revised
            save_version(f"Revisione AI: {task}")
            sync_editor()  # aggiorna la text_area con il testo revisionato
            st.success("Revisione terminata. Controlla sempre il risultato prima della pubblicazione.")
        except Exception as error:
            st.error(f"Revisione non completata: {error}")

    st.divider()
    st.subheader("Testo completo")
    edited = st.text_area(
        "Puoi modificare il testo direttamente qui",
        value=st.session_state.working_text,
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
    m1, m2, m3 = st.columns(3)
    m1.metric("Caratteri", f"{len(st.session_state.working_text):,}")
    m2.metric("Parole", f"{words:,}")
    m3.metric("Capitoli rilevati", chapters)

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
        try:
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

Restituisci esclusivamente JSON valido in italiano con queste chiavi:
titolo_suggerito, sottotitolo_suggerito, descrizione_breve,
descrizione_lunga, meta_title (massimo 60 caratteri),
meta_description (massimo 160 caratteri),
keyword_principali (array di 7 stringhe), tag (array di 12 stringhe),
categorie_suggerite (array di 3 stringhe).
Non inventare contenuti, premi, recensioni o caratteristiche non presenti nel testo.
"""
            raw = ask_ai(
                prompt,
                "Sei un consulente editoriale ed esperto SEO. Rispondi solo con JSON valido, senza markdown.",
                model,
            )
            data = parse_json_response(raw)
            data.update({
                "titolo_input": title,
                "sottotitolo_input": subtitle,
                "autore_input": author,
                "genere_input": genre,
                "pubblico_input": audience,
            })
            st.session_state.metadata = data
            st.success("Metadati generati.")
        except Exception as error:
            st.error(f"Non è stato possibile generare i metadati: {error}")

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
    st.caption("Le versioni sono conservate nella sessione corrente dell'app. Scarica una copia per conservarle a lungo termine.")
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
    if st.button("Salva snapshot manuale", disabled=not bool(st.session_state.working_text)):
        save_version("Snapshot manuale")
        st.success("Snapshot salvato.")

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
        st.download_button(
            "Scarica Word (DOCX)",
            data=make_docx(text_to_export),
            file_name=f"{base_name}_revisionato.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.download_button(
            "Scarica PDF",
            data=make_pdf(text_to_export),
            file_name=f"{base_name}_revisionato.pdf",
            mime="application/pdf",
        )
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
        st.download_button(
            "Scarica EPUB",
            data=make_epub(text_to_export, export_title, export_author),
            file_name=f"{base_name}_revisionato.epub",
            mime="application/epub+zip",
        )

st.divider()
st.caption(
    "LIBRO MASTER AI — versione iniziale. I PDF scannerizzati richiedono OCR. "
    "L'importazione e l'esportazione ricostruiscono il testo e non conservano necessariamente "
    "impaginazione, immagini, note o grafica originale. Controlla sempre il testo revisionato."
)
