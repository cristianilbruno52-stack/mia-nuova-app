"""Generatore di metadati editoriali/SEO per ciascun manoscritto analizzato.

Da usare nel tab SEO di LIBRO MASTER AI. I suggerimenti sono basati sul testo,
non su volumi di ricerca in tempo reale.
"""
from seo_kdp import render_seo_kdp_tab
from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime
from typing import Any, Callable

import streamlit as st


def _sample_text(text: str, limit: int = 14000) -> str:
    if len(text) <= limit:
        return text
    return (text[:6000] + "\n[...estratto centrale...]\n" +
            text[len(text) // 2 - 2000:len(text) // 2 + 2000] +
            "\n[...estratto finale...]\n" + text[-4000:])


def _as_list(value: Any, maximum: int) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(item).strip() for item in value if str(item).strip()][:maximum]


def _clean_metadata(data: dict[str, Any]) -> dict[str, Any]:
    """Normalizza i campi principali senza inventare contenuto mancante."""
    result = dict(data)
    for key in ("titoli_alternativi", "keyword_kdp", "tag_seo", "categorie_kdp", "note_verifica"):
        result[key] = _as_list(result.get(key), 12)
    for key in ("titolo_kdp", "sottotitolo_kdp", "meta_title", "meta_description",
                "descrizione_kdp", "descrizione_breve", "genere", "pubblico_target",
                "promessa_lettore", "motivazione_titolo"):
        result[key] = str(result.get(key, "") or "").strip()
    return result


def _validate_metadata(data: dict[str, Any]) -> list[dict[str, str]]:
    checks: list[dict[str, str]] = []
    title = data.get("titolo_kdp", "")
    subtitle = data.get("sottotitolo_kdp", "")
    meta_title = data.get("meta_title", "")
    meta_description = data.get("meta_description", "")
    description = data.get("descrizione_kdp", "")
    keywords = data.get("keyword_kdp", [])

    if title and subtitle and len(title) + len(subtitle) + 1 > 200:
        checks.append({"Gravità": "Avviso", "Controllo": "Titolo + sottotitolo",
                       "Risultato": f"{len(title) + len(subtitle) + 1} caratteri complessivi.",
                       "Cosa verificare": "Accorcia e verifica il limite mostrato nel flusso KDP; titolo/sottotitolo devono corrispondere ai dati editoriali e alla copertina."})
    if len(meta_title) > 60:
        checks.append({"Gravità": "Avviso", "Controllo": "Meta title web",
                       "Risultato": f"{len(meta_title)} caratteri.",
                       "Cosa verificare": "Riduci idealmente a circa 60 caratteri per lo snippet SEO web. Non è il campo Titolo del libro in KDP."})
    if len(meta_description) > 160:
        checks.append({"Gravità": "Avviso", "Controllo": "Meta description web",
                       "Risultato": f"{len(meta_description)} caratteri.",
                       "Cosa verificare": "Riduci idealmente a circa 160 caratteri per lo snippet SEO web. Non è la descrizione KDP."})
    if len(description) > 4000:
        checks.append({"Gravità": "Errore", "Controllo": "Descrizione KDP",
                       "Risultato": f"{len(description)} caratteri, oltre il limite di 4.000 caratteri indicato da KDP.",
                       "Cosa verificare": "Accorcia la descrizione e ricontrolla il conteggio nel pannello KDP; eventuale HTML contribuisce al conteggio."})
    if len(keywords) > 7:
        checks.append({"Gravità": "Errore", "Controllo": "Keyword KDP",
                       "Risultato": f"Generate {len(keywords)} frasi; KDP consente fino a 7 voci.",
                       "Cosa verificare": "Mantieni al massimo sette frasi pertinenti e inseriscile separatamente nei campi Keyword."})
    elif len(keywords) < 7:
        checks.append({"Gravità": "Info", "Controllo": "Keyword KDP",
                       "Risultato": f"Generate {len(keywords)} frasi su un massimo di 7.",
                       "Cosa verificare": "Aggiungi altre frasi solo se accurate; è meglio lasciare un campo vuoto che inserire una keyword non pertinente."})
    for index, keyword in enumerate(keywords, 1):
        if len(keyword) > 50:
            checks.append({"Gravità": "Avviso", "Controllo": f"Keyword KDP {index}",
                           "Risultato": f"{len(keyword)} caratteri.",
                           "Cosa verificare": "Accorcia la frase e controlla il limite visualizzato nel campo KDP prima di incollarla."})
    if description and re.search(r"https?://|www\.|\b\S+@\S+\.\S+\b", description, re.I):
        checks.append({"Gravità": "Avviso", "Controllo": "Descrizione KDP",
                       "Risultato": "Rilevato un possibile URL o indirizzo email.",
                       "Cosa verificare": "Rimuovi link e recapiti dalla descrizione KDP e controlla le restrizioni aggiornate."})
    if description and re.search(r"\b(bestseller|best seller|gratis|free|in offerta|sconto|review|recensione|lascia una recensione)\b", description, re.I):
        checks.append({"Gravità": "Avviso", "Controllo": "Descrizione KDP",
                       "Risultato": "Possibile frase promozionale o invito a recensione.",
                       "Cosa verificare": "Controlla la conformità alle regole KDP; elimina affermazioni non dimostrabili, promozioni o richieste di recensioni."})
    if not checks:
        checks.append({"Gravità": "OK", "Controllo": "Controlli automatici", "Risultato": "Nessuna anomalia formale rilevata nei campi generati.",
                       "Cosa verificare": "Conferma che titolo, sottotitolo, parole chiave e descrizione rappresentino davvero il libro e siano coerenti con la copertina."})
    return checks


def _render_record(record: dict[str, Any], compact: bool = False) -> None:
    data = record.get("metadati", {})
    record_id = str(record.get("id", "record"))[:24]
    st.caption(f"Libro: {record.get('file', 'libro')} · {record.get('creato_il', '')} · store: {record.get('marketplace', 'Amazon.it')}")
    st.info("**Titolo KDP:** usa il titolo effettivo del libro e assicurati che coincida con la copertina. Se il campo è stato lasciato vuoto, il titolo generato è una proposta da approvare e riportare anche sulla copertina.")

    if compact:
        with st.expander("Apri i metadati consigliati"):
            _render_fields(data, record_id)
        return
    _render_fields(data, record_id)


def _render_fields(data: dict[str, Any], record_id: str) -> None:
    st.download_button("Scarica questa scheda SEO in JSON",
                       json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
                       f"metadati_seo_{record_id}.json", "application/json",
                       key=f"seo_download_{record_id}")
    st.markdown("#### Titolo e posizionamento")
    st.text_input("Titolo da inserire in KDP (verifica che coincida con la copertina)",
                  value=data.get("titolo_kdp", ""), key=f"seo_title_out_{record_id}")
    st.text_input("Sottotitolo consigliato (facoltativo)", value=data.get("sottotitolo_kdp", ""),
                  key=f"seo_subtitle_out_{record_id}")
    alternatives = data.get("titoli_alternativi", [])
    if alternatives:
        st.markdown("**Alternative di titolo da valutare:**")
        for i, item in enumerate(alternatives, 1):
            st.write(f"{i}. {item}")
    st.write(f"**Genere/posizionamento:** {data.get('genere', '—') or '—'}")
    st.write(f"**Lettore ideale:** {data.get('pubblico_target', '—') or '—'}")
    st.write(f"**Promessa al lettore:** {data.get('promessa_lettore', '—') or '—'}")
    if data.get("motivazione_titolo"):
        st.caption(data["motivazione_titolo"])

    st.markdown("#### Descrizione pronta per la scheda Amazon KDP")
    description = data.get("descrizione_kdp", "")
    st.caption(f"{len(description)} / 4.000 caratteri · testo semplice da rivedere prima di incollarlo in KDP")
    st.text_area("Descrizione KDP", value=description, height=240, key=f"seo_desc_out_{record_id}")
    if data.get("descrizione_breve"):
        st.text_area("Descrizione breve / quarta di copertina", value=data["descrizione_breve"], height=100,
                     key=f"seo_blurb_out_{record_id}")

    st.markdown("#### Le 7 frasi chiave per i campi Keyword KDP")
    st.caption("Sono frasi da inserire separatamente nei campi Keyword KDP, non una lista da incollare nella descrizione.")
    for i, keyword in enumerate(data.get("keyword_kdp", []), 1):
        st.text_input(f"Keyword KDP {i} · {len(keyword)} caratteri", value=keyword,
                      key=f"seo_keyword_out_{record_id}_{i}")

    tags = data.get("tag_seo", [])
    if tags:
        st.markdown("#### Tag SEO / sito / social (non sono i campi Keyword KDP)")
        st.text_area("Tag suggeriti", value=", ".join(tags), height=90, key=f"seo_tags_out_{record_id}")
    categories = data.get("categorie_kdp", [])
    if categories:
        st.markdown("#### Categorie KDP da verificare nel pannello")
        for category in categories:
            st.write(f"- {category}")
    st.markdown("#### SEO per sito o blog (se ti serve)")
    st.text_input(f"Meta title web · {len(data.get('meta_title', ''))} caratteri",
                  value=data.get("meta_title", ""), key=f"seo_meta_title_out_{record_id}")
    st.text_area(f"Meta description web · {len(data.get('meta_description', ''))} caratteri",
                 value=data.get("meta_description", ""), height=75,
                 key=f"seo_meta_desc_out_{record_id}")
    checks = _validate_metadata(data)
    st.markdown("#### Controlli formali e cose da verificare")
    st.dataframe(checks, hide_index=True, use_container_width=True)
    notes = data.get("note_verifica", [])
    if notes:
        st.markdown("**Note editoriali:**")
        for note in notes:
            st.write(f"- {note}")


def render_seo_kdp_tab(text: str, filename: str, model: str,
                       get_api_key: Callable[[], str], ask_ai: Callable[..., str],
                       parse_json_object: Callable[[str], dict[str, Any]]) -> None:
    """Tab SEO per libro corrente, con cronologia delle schede generate in sessione."""
    st.subheader("Metadati SEO e scheda Amazon KDP")
    st.caption("Per ogni testo genera un titolo/sottotitolo da valutare, descrizione, sette frasi Keyword KDP e tag SEO separati. I suggerimenti si basano sul manoscritto; non sono dati di volume di ricerca in tempo reale.")
    st.session_state.setdefault("seo_history", [])

    title = st.text_input("Titolo attuale o titolo già stampato in copertina", key="seo_title")
    author = st.text_input("Autore", key="seo_author")
    marketplace = st.selectbox("Store/lingua della scheda", ["Amazon.it · italiano", "Amazon.com · inglese", "Amazon.co.uk · inglese", "Amazon.es · spagnolo", "Amazon.fr · francese", "Amazon.de · tedesco"], key="seo_marketplace")
    st.caption("Se hai già un titolo definitivo, inseriscilo: il generatore lo mantiene e propone un sottotitolo. Se il titolo è vuoto, genera una proposta e alcune alternative.")

    if st.button("Genera titolo, keyword, tag e descrizione per questo libro", type="primary", disabled=not text, key="seo_generate_kdp"):
        if not get_api_key():
            st.error("Configura la chiave API nella sidebar.")
        else:
            signature = hashlib.sha1(text.encode("utf-8", "ignore")).hexdigest()
            sample = _sample_text(text)
            prompt = f"""Crea una proposta di metadati editoriali e SEO per il libro qui sotto.

DATI FORNITI DALL'AUTORE
Titolo già scelto/stampato: {title or "NON INDICATO — proponi un titolo candidato"}
Autore: {author or "non indicato"}
Store e lingua: {marketplace}
Nome file: {filename}

TESTO DEL LIBRO (estratto):
<<<LIBRO
{sample}
LIBRO>>>

Regole di lavoro:
- Basa ogni affermazione esclusivamente sul testo e sui dati forniti. Non inventare trama, pubblico, risultati, contenuti, credenziali, premi o fatti.
- Se il titolo è già stato fornito, riportalo identico in titolo_kdp: non cambiarlo. Proponi il sottotitolo soltanto se utile. Se non è stato fornito, titolo_kdp è una proposta da verificare e deve essere breve, naturale e non una lista di keyword. Fornisci comunque tre alternative, chiaramente solo candidate.
- La scheda KDP deve essere nella lingua adatta allo store selezionato. Il titolo/sottotitolo devono essere coerenti con copertina e contenuto; segnala nelle note che una proposta va riportata sulla copertina prima della pubblicazione.
- Scrivi descrizione_kdp come testo promozionale onesto e leggibile, circa 120–180 parole, senza spoiler importanti; massimo 3.800 caratteri per lasciare margine al controllo, mai oltre 4.000. Non inserire URL, email, telefono, recensioni/testimonianze, richieste di recensioni, prezzi, promozioni, emoji o un elenco di parole chiave.
- Genera fino a sette frasi keyword_kdp, pertinenti al contenuto e scritte come le cercherebbe un lettore. Ogni frase deve avere massimo 50 caratteri; non inserire autore, titoli di libri/altri autori, affermazioni promozionali o temi non presenti. Non inventare volumi di ricerca. Se il testo non supporta sette frasi affidabili, restituiscine meno e spiegalo nelle note.
- tag_seo contiene 8–12 tag tematici per sito/blog/social, separati concettualmente dai sette campi Keyword di KDP.
- categorie_kdp contiene al massimo tre categorie/aree editoriali candidate da verificare nel catalogo KDP; non inventare nomi di categorie non verificabili.
- meta_title è per SEO web, massimo 60 caratteri; meta_description è per SEO web, massimo 160 caratteri. Non presentarli come campi KDP.
- descrizione_breve è una quarta di copertina breve, coerente e senza spoiler.

Rispondi ESCLUSIVAMENTE con un oggetto JSON valido con queste chiavi:
titolo_kdp, titoli_alternativi (lista di 3), sottotitolo_kdp, motivazione_titolo, genere, pubblico_target, promessa_lettore, descrizione_kdp, descrizione_breve, keyword_kdp (lista di massimo 7 stringhe), tag_seo (lista), categorie_kdp (lista), meta_title, meta_description, note_verifica (lista)."""
            system = ("Sei un consulente editoriale e di metadati per libri. Scrivi nella lingua richiesta, "
                      "senza keyword stuffing, rispettando le regole Amazon KDP e distinguendo i campi KDP dai metadati SEO web. "
                      "Sii prudente: quando un'informazione non è nel manoscritto, non inventarla.")
            with st.spinner("Analisi del libro e creazione dei metadati..."):
                raw, error = _safe_call(ask_ai, prompt, system, model, 0.0, json_mode=True)
                if error:
                    st.error(str(error))
                else:
                    parsed, parse_error = _safe_call(parse_json_object, raw)
                    if parse_error:
                        st.error(str(parse_error))
                    else:
                        parsed = _clean_metadata(parsed)
                        record_seed = f"{signature}|{title}|{author}|{marketplace}|{datetime.now().isoformat()}"
                        record_id = hashlib.sha1(record_seed.encode("utf-8", "ignore")).hexdigest()[:16]
                        record = {
                            "id": record_id,
                            "firma_testo": signature,
                            "file": filename or "libro",
                            "titolo_inserito": title,
                            "autore": author,
                            "marketplace": marketplace,
                            "creato_il": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                            "metadati": parsed,
                        }
                        st.session_state["seo_current"] = record
                        history = st.session_state["seo_history"]
                        history.append(record)
                        st.session_state["seo_history"] = history[-40:]
                        # Conserva anche la struttura JSON già usata dalla versione precedente dell'app.
                        st.session_state["metadata"] = parsed

    current = st.session_state.get("seo_current")
    if current:
        st.markdown("### Raccomandazioni per il libro corrente")
        if current.get("firma_testo") != hashlib.sha1(text.encode("utf-8", "ignore")).hexdigest():
            st.warning("Il testo è cambiato dopo la generazione: rigenera i metadati per questa versione del libro.")
        _render_record(current)

    history = st.session_state.get("seo_history", [])
    if history:
        st.markdown("### Metadati generati per i libri analizzati in questa sessione")
        for record in reversed(history):
            if current and record.get("id") == current.get("id"):
                continue
            label = f"{record.get('file', 'libro')} · {record.get('creato_il', '')} · {record.get('metadati', {}).get('titolo_kdp', 'senza titolo')}"
            with st.expander(label):
                _render_record(record, compact=True)


def _safe_call(func: Callable[..., Any], *args: Any, **kwargs: Any) -> tuple[Any, Exception | None]:
    try:
        return func(*args, **kwargs), None
    except Exception as exc:  # noqa: BLE001
        return None, exc
        """Controlli tecnici e visivi di preflight per copertine e illustrazioni Amazon KDP.

Modulo aggiuntivo per LIBRO MASTER AI. Non invia né modifica file: produce solo un report.
I controlli automatici non sostituiscono il Print Previewer o l'approvazione KDP.
"""
from __future__ import annotations

import io
import json
import math
import os
import re
from typing import Any

import streamlit as st

try:
    import fitz  # PyMuPDF
except ImportError:  # il tab mostra un messaggio chiaro se il pacchetto manca
    fitz = None

try:
    from PIL import Image, ImageOps
except ImportError:
    Image = None
    ImageOps = None

try:
    from google.genai import types
except ImportError:
    types = None

BLEED_IN = 0.125
PAPER_FACTORS = {
    "Bianco B/N": 0.002252,
    "Crema B/N": 0.0025,
    "Colore standard": 0.002252,
    "Colore premium": 0.002347,
}
TRIM_PRESETS = {
    "5 × 8 in": (5.0, 8.0),
    "5.5 × 8.5 in": (5.5, 8.5),
    "6 × 9 in": (6.0, 9.0),
    "7 × 10 in": (7.0, 10.0),
    "8.5 × 11 in": (8.5, 11.0),
    "Personalizzato": (6.0, 9.0),
}


def _issue(rows: list[dict[str, Any]], severity: str, area: str, problem: str,
           fix: str, where: str = "") -> None:
    rows.append({
        "Gravità": severity,
        "Elemento": area,
        "Dove": where or "—",
        "Problema rilevato": problem,
        "Cosa correggere": fix,
    })


def _in(value_points: float) -> float:
    return float(value_points) / 72.0


def _fmt_in(value: float) -> str:
    return f"{value:.3f} in ({value * 25.4:.1f} mm)"


def _safe_float(value: Any) -> float | None:
    try:
        value = float(value)
        return value if math.isfinite(value) and value > 0 else None
    except (TypeError, ValueError):
        return None


def pdf_page_count(data: bytes) -> int | None:
    if fitz is None:
        return None
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            return doc.page_count
    except Exception:
        return None


def _pdf_dimensions(data: bytes, page_index: int = 0) -> tuple[float, float]:
    if fitz is None:
        raise RuntimeError("PyMuPDF non installato. Aggiungi PyMuPDF ai requirements.")
    with fitz.open(stream=data, filetype="pdf") as doc:
        if doc.is_encrypted:
            raise ValueError("Il PDF è protetto da password.")
        if doc.page_count <= page_index:
            raise ValueError("Il PDF non contiene la pagina richiesta.")
        rect = doc.load_page(page_index).mediabox
        return _in(rect.width), _in(rect.height)


def _check_pdf_size(rows: list[dict[str, Any]], data: bytes, label: str,
                    expected_w: float | None = None, expected_h: float | None = None,
                    tolerance_in: float = 0.02) -> tuple[int, float | None, float | None]:
    if fitz is None:
        _issue(rows, "Errore", label, "Non posso leggere il PDF: manca PyMuPDF.",
               "Installa il pacchetto con `pip install PyMuPDF` e riavvia Streamlit.")
        return 0, None, None
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            if doc.is_encrypted:
                _issue(rows, "Errore", label, "PDF protetto da password.",
                       "Rimuovi la password e carica un PDF leggibile.")
                return 0, None, None
            count = doc.page_count
            if count != 1 and label == "Copertina stampa":
                _issue(rows, "Errore", label, f"La copertina contiene {count} pagine.",
                       "Esporta un unico PDF di una pagina con retro, dorso e fronte sullo stesso foglio.")
            if count < 1:
                _issue(rows, "Errore", label, "Il PDF non contiene pagine.", "Esporta nuovamente il PDF.")
                return count, None, None
            page = doc.load_page(0)
            box = page.mediabox
            actual_w, actual_h = _in(box.width), _in(box.height)
            if page.rotation:
                _issue(rows, "Avviso", label, f"La prima pagina ha rotazione PDF di {page.rotation}°.",
                       "Appiattisci la rotazione ed esporta la pagina con orientamento e dimensioni già corretti.")
            crop = page.cropbox
            if abs(crop.width - box.width) > 1 or abs(crop.height - box.height) > 1:
                _issue(rows, "Avviso", label, "MediaBox e CropBox del PDF non coincidono.",
                       "Controlla che non ci siano segni di taglio o aree esterne al formato finale.")
            if expected_w is not None and expected_h is not None:
                diff_w = abs(actual_w - expected_w)
                diff_h = abs(actual_h - expected_h)
                if diff_w > tolerance_in or diff_h > tolerance_in:
                    _issue(rows, "Errore", label,
                           f"Dimensione PDF {actual_w:.3f} × {actual_h:.3f} in; attesa {expected_w:.3f} × {expected_h:.3f} in.",
                           f"Imposta il documento a {_fmt_in(expected_w)} × {_fmt_in(expected_h)} (secondo i parametri scelti) ed esporta senza bordi extra.")
                else:
                    _issue(rows, "OK", label,
                           f"Dimensione PDF {actual_w:.3f} × {actual_h:.3f} in: corrisponde alle dimensioni attese entro {tolerance_in:.2f} in.",
                           "Nessuna correzione dimensionale rilevata.")
            return count, actual_w, actual_h
    except Exception as exc:
        _issue(rows, "Errore", label, f"PDF non leggibile: {str(exc)[:180]}.",
               "Esporta o salva di nuovo il file in PDF e verifica che non sia danneggiato.")
        return 0, None, None


def _image_info(name: str, data: bytes) -> dict[str, Any]:
    if Image is None:
        raise RuntimeError("Pillow non installato. Aggiungi Pillow ai requirements.")
    try:
        image = Image.open(io.BytesIO(data))
        image = ImageOps.exif_transpose(image) if ImageOps is not None else image
        width, height = image.size
        dpi_info = image.info.get("dpi")
        dpi_x = _safe_float(dpi_info[0]) if isinstance(dpi_info, (tuple, list)) and dpi_info else None
        dpi_y = _safe_float(dpi_info[1]) if isinstance(dpi_info, (tuple, list)) and len(dpi_info) > 1 else dpi_x
        info = {
            "name": name,
            "format": (image.format or os.path.splitext(name)[1].lstrip(".")).upper(),
            "width": width,
            "height": height,
            "mode": image.mode,
            "dpi_x": dpi_x,
            "dpi_y": dpi_y,
            "frames": int(getattr(image, "n_frames", 1)),
            "alpha": "A" in image.getbands() or "transparency" in image.info,
        }
        image.close()
        return info
    except Exception as exc:
        raise ValueError(f"Immagine non leggibile: {str(exc)[:160]}") from exc


def _check_raster(rows: list[dict[str, Any]], name: str, data: bytes, role: str,
                  target_w_in: float | None = None, target_h_in: float | None = None) -> dict[str, Any] | None:
    try:
        info = _image_info(name, data)
    except Exception as exc:
        _issue(rows, "Errore", role, str(exc), "Esporta nuovamente l'immagine in JPG, PNG o TIFF non danneggiato.", name)
        return None
    width, height = info["width"], info["height"]
    fmt = info["format"]
    ext = os.path.splitext(name)[1].lower()
    if fmt not in {"JPEG", "JPG", "PNG", "TIFF", "TIF", "WEBP", "GIF"}:
        _issue(rows, "Errore", role, f"Formato immagine {fmt or ext or 'sconosciuto'} non riconosciuto.",
               "Per Kindle usa JPG/PNG/TIFF; per stampa carica il PDF completo della copertina.", name)
    if role == "Copertina ebook":
        if width < 625 or height < 1000:
            _issue(rows, "Errore", role, f"Immagine {width} × {height} px: sotto il minimo consigliato di 625 × 1000 px.",
                   "Crea una nuova copertina alla risoluzione originale; non ingrandire artificialmente quella piccola.", name)
        elif width < 1600 or height < 2560:
            _issue(rows, "Avviso", role, f"Immagine {width} × {height} px: inferiore alla dimensione ideale 1600 × 2560 px.",
                   "Se possibile esporta l'originale a 1600 × 2560 px o più, mantenendo le proporzioni.", name)
        if width > 10000 or height > 10000:
            _issue(rows, "Errore", role, f"Dimensione {width} × {height} px: supera il limite di 10.000 px per lato.",
                   "Ridimensiona/esporta l'immagine entro 10.000 px per lato.", name)
        ratio = height / max(1, width)
        if abs(ratio - 1.6) > 0.12:
            _issue(rows, "Avviso", role, f"Rapporto altezza/larghezza {ratio:.2f}:1; il riferimento Kindle è circa 1.6:1.",
                   "Verifica che la copertina mantenga le proporzioni desiderate e non venga stirata o ritagliata.", name)
        if len(data) >= 50 * 1024 * 1024:
            _issue(rows, "Errore", role, f"File di {len(data) / 1024 / 1024:.1f} MB: raggiunge/supera 50 MB.",
                   "Esporta un JPG/TIFF ottimizzato sotto 50 MB senza ridurre eccessivamente la qualità.", name)
        if info["mode"] == "CMYK":
            _issue(rows, "Avviso", role, "Immagine CMYK; per Kindle è preferibile RGB/sRGB.",
                   "Converti una copia in RGB/sRGB e controlla che i colori non cambino.", name)
        if info["alpha"]:
            _issue(rows, "Avviso", role, "L'immagine contiene trasparenza.",
                   "Appiattisci la copertina su uno sfondo opaco e controlla il risultato sui dispositivi Kindle.", name)
        _issue(rows, "OK", role,
               f"Risoluzione pixel: {width} × {height}; formato {fmt}; profilo/modalità {info['mode']}.",
               "Controllo raster completato. Il DPI incorporato è informativo: per ebook contano soprattutto i pixel.", name)
    else:
        if target_w_in and target_h_in:
            dpi_x = width / target_w_in
            dpi_y = height / target_h_in
            eff = min(dpi_x, dpi_y)
            if eff < 300:
                _issue(rows, "Errore", role,
                       f"Risoluzione effettiva stimata {eff:.0f} DPI alla dimensione di stampa {target_w_in:.2f} × {target_h_in:.2f} in.",
                       "Usa un originale più grande o riduci la dimensione di inserimento finché entrambe le direzioni raggiungono almeno 300 DPI.", name)
            else:
                _issue(rows, "OK", role, f"Risoluzione effettiva stimata {eff:.0f} DPI alla dimensione indicata.",
                       "Nessun problema di risoluzione rilevato a questa dimensione.", name)
        elif info["dpi_x"] and info["dpi_y"]:
            minimum = min(info["dpi_x"], info["dpi_y"])
            severity = "Avviso" if minimum < 300 else "OK"
            msg = f"Metadati immagine: {info['dpi_x']:.0f} × {info['dpi_y']:.0f} DPI; {width} × {height} px."
            fix = ("Il DPI nei metadati non garantisce il DPI effettivo: controlla la dimensione di stampa nel PDF finale." if minimum >= 300
                   else "Aumenta la risoluzione originale oppure riduci la dimensione di stampa; verifica nel PDF finale.")
            _issue(rows, severity, role, msg, fix, name)
        else:
            _issue(rows, "Avviso", role, f"Immagine {width} × {height} px senza metadati DPI attendibili.",
                   "Indica la dimensione finale nel PDF; solo lì si può calcolare il DPI effettivo. KDP raccomanda almeno 300 DPI.", name)
        if info["mode"] == "CMYK":
            _issue(rows, "Info", role, "Immagine in modalità CMYK.",
                   "Verifica l'opzione di stampa scelta e controlla una prova colore; la gestione colore varia tra ebook e stampa.", name)
        if info["alpha"]:
            _issue(rows, "Avviso", role, "Immagine con trasparenza/livello alpha.",
                   "Appiattisci gli elementi trasparenti nel file finale e controlla il PDF esportato.", name)
        if info["frames"] > 1:
            _issue(rows, "Avviso", role, f"Immagine con {info['frames']} fotogrammi/pagine.",
                   "Esporta l'illustrazione destinata al libro come immagine statica.", name)
    return info


def _image_placements(doc: Any, page: Any, page_number: int, rows: list[dict[str, Any]], area: str,
                      low_dpi_limit: int = 30) -> tuple[int, list[dict[str, Any]]]:
    count = 0
    extracted: list[dict[str, Any]] = []
    try:
        images = page.get_images(full=True)
    except Exception:
        images = []
    for item in images:
        try:
            xref = int(item[0])
            smask = int(item[1]) if len(item) > 1 else 0
            wpx, hpx = int(item[2]), int(item[3])
            if xref <= 0 or wpx <= 0 or hpx <= 0:
                continue
            rects = page.get_image_rects(xref)
            if not rects:
                continue
            if smask > 0:
                _issue(rows, "Avviso", area, "Immagine con maschera di trasparenza nel PDF.",
                       "Appiattisci le trasparenze e ricontrolla il PDF esportato.", f"pagina {page_number}")
            for rect in rects:
                rw, rh = float(rect.width), float(rect.height)
                if rw <= 0 or rh <= 0:
                    continue
                rect_ratio = rw / rh
                img_ratio = wpx / hpx
                if abs(math.log(max(1e-9, rect_ratio / img_ratio))) <= abs(math.log(max(1e-9, rect_ratio * img_ratio))):
                    dpi_x, dpi_y = wpx / (rw / 72), hpx / (rh / 72)
                else:
                    dpi_x, dpi_y = hpx / (rw / 72), wpx / (rh / 72)
                effective = min(dpi_x, dpi_y)
                count += 1
                if effective < 300:
                    _issue(rows, "Errore", area,
                           f"Illustrazione a {effective:.0f} DPI effettivi (minimo raccomandato 300); immagine sorgente {wpx} × {hpx} px.",
                           "Sostituisci con un originale a risoluzione maggiore oppure riduci la dimensione con cui è collocata nella pagina.",
                           f"pagina {page_number}")
                elif effective >= 600:
                    _issue(rows, "Info", area, f"Illustrazione a {effective:.0f} DPI effettivi; risoluzione alta.",
                           "Nessuna correzione richiesta; immagini eccessivamente pesanti possono aumentare il PDF.",
                           f"pagina {page_number}")
            # Conserva una sola copia per xref/pagina per la revisione visiva AI.
            if len(extracted) < low_dpi_limit:
                try:
                    raw = doc.extract_image(xref)
                    blob = raw.get("image", b"")
                    ext = raw.get("ext", "").lower()
                    if blob:
                        mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                                "webp": "image/webp", "tif": "image/tiff", "tiff": "image/tiff"}.get(ext, "application/octet-stream")
                        extracted.append({"id": f"pagina {page_number} · immagine {xref}", "data": blob, "mime": mime})
                except Exception:
                    pass
        except Exception:
            continue
    return count, extracted


def _check_pdf_images(rows: list[dict[str, Any]], data: bytes, area: str,
                      expected_size: tuple[float, float] | None = None,
                      is_cover: bool = False) -> tuple[int, int]:
    if fitz is None:
        _issue(rows, "Errore", area, "Controllo PDF non disponibile: manca PyMuPDF.",
               "Installa PyMuPDF e ripeti il controllo.")
        return 0, 0
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            if doc.is_encrypted:
                _issue(rows, "Errore", area, "PDF protetto da password.", "Rimuovi la password e ricarica il PDF.")
                return 0, 0
            if not is_cover and expected_size:
                exp_w, exp_h = expected_size
                wrong = []
                for number, page in enumerate(doc, 1):
                    rect = page.mediabox
                    actual_w, actual_h = _in(rect.width), _in(rect.height)
                    direct = abs(actual_w - exp_w) <= 0.03 and abs(actual_h - exp_h) <= 0.03
                    rotated = abs(actual_w - exp_h) <= 0.03 and abs(actual_h - exp_w) <= 0.03
                    if not direct and not rotated:
                        wrong.append((number, actual_w, actual_h))
                if wrong:
                    first = ", ".join(f"p.{n}: {w:.3f}×{h:.3f} in" for n, w, h in wrong[:8])
                    _issue(rows, "Errore", area,
                           f"Formato pagina diverso dal trim impostato su {len(wrong)}/{doc.page_count} pagine ({first}).",
                           f"Imposta tutte le pagine a {_fmt_in(exp_w)} × {_fmt_in(exp_h)} (oppure l'orientamento ruotato se previsto) e riesporta il PDF.")
                else:
                    _issue(rows, "OK", area,
                           f"Formato pagina coerente con il trim impostato su tutte le {doc.page_count} pagine.",
                           "Nessuna correzione dimensionale rilevata.")
            placements = 0
            ai_images: list[dict[str, Any]] = []
            low_count_before = sum(1 for r in rows if r.get("Gravità") == "Errore" and r.get("Elemento") == area and "DPI effettivi" in r.get("Problema rilevato", ""))
            for page_no, page in enumerate(doc, 1):
                page_count, page_imgs = _image_placements(doc, page, page_no, rows, area)
                placements += page_count
                ai_images.extend(page_imgs)
            if placements == 0:
                _issue(rows, "Info", area, "Non sono state trovate immagini raster estraibili nel PDF.",
                       "Se le illustrazioni sono vettoriali, il controllo DPI raster non si applica; controlla le pagine nel Print Previewer.")
            else:
                low_count_after = sum(1 for r in rows if r.get("Gravità") == "Errore" and r.get("Elemento") == area and "DPI effettivi" in r.get("Problema rilevato", ""))
                _issue(rows, "Info", area,
                       f"Rilevate {placements} collocazioni raster; {max(0, low_count_after - low_count_before)} sotto 300 DPI.",
                       "Il DPI è stimato dalle dimensioni di ciascuna immagine nel PDF; illustrazioni vettoriali non hanno DPI raster.")
            # Ritorna solo immagini uniche per questa pagina/documento, senza duplicare xref già estratti.
            unique: dict[str, dict[str, Any]] = {}
            for item in ai_images:
                unique.setdefault(item["id"].split("· immagine")[-1].strip(), item)
            return placements, list(unique.values())
    except Exception as exc:
        _issue(rows, "Errore", area, f"PDF non leggibile durante l'analisi immagini: {str(exc)[:160]}.",
               "Esporta nuovamente il PDF e verifica che le immagini siano incorporate.")
        return 0, 0


def _scan_fonts_and_annotations(rows: list[dict[str, Any]], data: bytes, area: str) -> None:
    if fitz is None:
        return
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            missing_fonts: set[str] = set()
            annotations: list[int] = []
            for page_no, page in enumerate(doc, 1):
                try:
                    for font in doc.get_page_fonts(page_no - 1, full=True):
                        xref = int(font[0])
                        basefont = str(font[3] or "font senza nome")
                        if xref <= 0:
                            missing_fonts.add(basefont)
                            continue
                        try:
                            extracted = doc.extract_font(xref)
                            content = extracted[3] if len(extracted) > 3 else b""
                            if not content:
                                missing_fonts.add(basefont)
                        except Exception:
                            missing_fonts.add(basefont)
                except Exception:
                    pass
                try:
                    if list(page.annots() or []):
                        annotations.append(page_no)
                except Exception:
                    pass
            if missing_fonts:
                names = ", ".join(sorted(missing_fonts)[:6])
                _issue(rows, "Avviso", area, f"Possibili font non incorporati: {names}.",
                       "Esporta il PDF con tutti i font incorporati o converti il testo in tracciati, quindi verifica il risultato visivo.")
            else:
                _issue(rows, "Info", area, "Nessun font non incorporato rilevato dai metadati PDF.",
                       "Controllo automatico indicativo: verifica comunque l'esportazione e i font nel Print Previewer.")
            if annotations:
                _issue(rows, "Avviso", area, f"Annotazioni presenti nelle pagine: {annotations[:12]}.",
                       "Rimuovi commenti, annotazioni o markup prima di caricare il PDF.")
    except Exception:
        pass


def _expected_paperback(trim_w: float, trim_h: float, pages: int, paper: str) -> tuple[float, float, float]:
    spine = pages * PAPER_FACTORS[paper]
    total_w = 2 * trim_w + spine + 2 * BLEED_IN
    total_h = trim_h + 2 * BLEED_IN
    return total_w, total_h, spine


def _render_pdf_page(data: bytes, page_no: int = 0, max_side: int = 1800) -> bytes:
    if fitz is None:
        raise RuntimeError("PyMuPDF necessario per creare l'anteprima AI del PDF.")
    with fitz.open(stream=data, filetype="pdf") as doc:
        if doc.page_count <= page_no:
            raise ValueError("Pagina PDF non trovata.")
        page = doc.load_page(page_no)
        rect = page.rect
        scale = min(max_side / max(float(rect.width), 1), max_side / max(float(rect.height), 1))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        return pix.tobytes("png")


def _vision_ready(data: bytes, max_side: int = 1800) -> bytes:
    if Image is None:
        return data
    with Image.open(io.BytesIO(data)) as image:
        image = ImageOps.exif_transpose(image) if ImageOps is not None else image
        image.thumbnail((max_side, max_side))
        if image.mode in ("RGBA", "LA") or "transparency" in image.info:
            rgba = image.convert("RGBA")
            bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            bg.alpha_composite(rgba)
            image = bg.convert("RGB")
        else:
            image = image.convert("RGB")
        out = io.BytesIO()
        image.save(out, format="JPEG", quality=88, optimize=True)
        return out.getvalue()


def _parse_ai_list(raw: str) -> list[dict[str, Any]]:
    text = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", str(raw).strip(), flags=re.I)
    try:
        obj = json.loads(text)
    except Exception:
        start, end = text.find("["), text.rfind("]")
        if start < 0 or end < start:
            start, end = text.find("{"), text.rfind("}")
            if start < 0 or end < start:
                raise ValueError("La risposta visiva AI non contiene JSON valido.")
            obj = json.loads(text[start:end + 1])
        else:
            obj = json.loads(text[start:end + 1])
    if isinstance(obj, dict):
        obj = obj.get("problemi", obj.get("issues", []))
    if not isinstance(obj, list):
        raise ValueError("La risposta AI non è una lista JSON.")
    return [item for item in obj if isinstance(item, dict)]


def _run_visual_ai(rows: list[dict[str, Any]], images: list[dict[str, Any]], model: str,
                   ask_ai_fn: Any, language: str) -> None:
    if not images:
        _issue(rows, "Info", "Controllo visivo AI", "Nessuna immagine disponibile per il controllo visivo.",
               "Carica una copertina o immagini del libro.")
        return
    if types is None:
        _issue(rows, "Errore", "Controllo visivo AI", "SDK Google GenAI non disponibile.",
               "Verifica l'installazione del pacchetto google-genai.")
        return
    system = (
        f"Sei un revisore visivo editoriale che valuta immagini per Amazon KDP. Lingua delle segnalazioni: {language}. "
        "Analizza solo problemi effettivamente visibili: testo che sembra tagliato o illeggibile, refusi chiaramente leggibili, "
        "basso contrasto, sovrapposizioni, oggetti importanti troppo vicini al bordo, artefatti o immagini sfocate. "
        "Non inventare errori, non riscrivere l'opera e non dichiarare mai che il file è garantito conforme/approvato da KDP. "
        "Se non puoi leggere una parola, dillo invece di indovinare. Rispondi solo con JSON valido: "
        "lista di oggetti {image_id, gravita (Errore|Avviso|Info), problema, cosa_correggere, testo_letto}. "
        "Se non rilevi problemi concreti, rispondi []."
    )
    total = (len(images) + 3) // 4
    progress = st.progress(0, text="Analisi visiva AI delle immagini...")
    for batch_no in range(total):
        batch = images[batch_no * 4:(batch_no + 1) * 4]
        parts: list[Any] = []
        for item in batch:
            parts.append(f"ID IMMAGINE: {item['id']}")
            parts.append(types.Part.from_bytes(data=item["data"], mime_type=item.get("mime", "image/jpeg")))
        prompt = (
            "Esamina ogni immagine allegata separatamente e collega ogni rilievo al suo ID. "
            "Per una copertina controlla testo, titolo/autore leggibili, tagli e contrasto; per illustrazioni controlla "
            "solo difetti visivi oggettivi o testo visibile. Non valutare DPI o dimensioni, già controllati a parte. "
            "Restituisci esclusivamente la lista JSON specificata."
        )
        parts.append(prompt)
        try:
            raw = ask_ai_fn(parts, system, model, 0.0, json_mode=True)
            findings = _parse_ai_list(raw)
            for finding in findings:
                image_id = str(finding.get("image_id", "immagine non identificata"))[:120]
                severity = str(finding.get("gravita", "Avviso"))
                if severity not in {"Errore", "Avviso", "Info"}:
                    severity = "Avviso"
                problem = str(finding.get("problema", "Problema visivo da verificare."))[:500]
                fix = str(finding.get("cosa_correggere", "Controlla l'immagine ingrandita e correggi il punto indicato."))[:500]
                quote = str(finding.get("testo_letto", "")).strip()
                if quote:
                    problem += f" · testo visibile: «{quote[:180]}»"
                _issue(rows, severity, "Controllo visivo AI", problem, fix, image_id)
            if not findings:
                _issue(rows, "OK", "Controllo visivo AI", "Nessun difetto visivo certo rilevato nel gruppo.",
                       "Controllo automatico indicativo: fai comunque una verifica manuale al 100%.",
                       ", ".join(item["id"] for item in batch))
        except Exception as exc:
            msg = str(exc)
            _issue(rows, "Avviso", "Controllo visivo AI",
                   f"Analisi AI non completata: {msg[:240]}.",
                   "Controlla la chiave/quota/modello Gemini e riprova; i controlli tecnici già completati restano validi.",
                   ", ".join(item["id"] for item in batch))
            if any(token in msg.upper() for token in ("429", "RESOURCE_EXHAUSTED", "QUOTA")):
                for remaining in range(batch_no + 1, total):
                    rest = images[(remaining) * 4:(remaining + 1) * 4]
                    _issue(rows, "Avviso", "Controllo visivo AI", "Gruppo non analizzato perché la quota API è esaurita.",
                           "Ripeti l'analisi dopo il ripristino della quota oppure controlla manualmente.",
                           ", ".join(item["id"] for item in rest))
                break
        progress.progress(min(1.0, (batch_no + 1) / total), text=f"Gruppo AI {batch_no + 1}/{total}")
    progress.empty()


def _report_view(rows: list[dict[str, Any]], summary: dict[str, Any]) -> None:
    errors = sum(row["Gravità"] == "Errore" for row in rows)
    warnings = sum(row["Gravità"] == "Avviso" for row in rows)
    infos = sum(row["Gravità"] in {"Info", "OK"} for row in rows)
    c1, c2, c3 = st.columns(3)
    c1.metric("Errori da correggere", errors)
    c2.metric("Avvisi da verificare", warnings)
    c3.metric("Controlli informativi", infos)
    if errors:
        st.error(f"Rilevati {errors} problemi tecnici/visivi. Consulta la colonna **Cosa correggere**.")
    elif warnings:
        st.warning("Non sono stati rilevati errori bloccanti automatici, ma restano avvisi da verificare manualmente.")
    else:
        st.success("Nessun problema rilevato dai controlli eseguiti.")
    st.warning("Questo preflight non garantisce l'approvazione KDP: controlla sempre il PDF nel Print Previewer e ordina una prova di stampa quando possibile.")
    if summary:
        st.json(summary)
    if rows:
        st.dataframe(rows, hide_index=True, use_container_width=True)
        import pandas as pd
        csv = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")
        st.download_button("Scarica report KDP CSV", csv, "report_preflight_kdp.csv", "text/csv", key="kdp_report_csv")
        st.download_button("Scarica report KDP JSON", json.dumps({"riepilogo": summary, "controlli": rows}, ensure_ascii=False, indent=2).encode("utf-8"),
                           "report_preflight_kdp.json", "application/json", key="kdp_report_json")


def render_kdp_tab(model: str, ask_ai_fn: Any, api_key_configured: bool, language: str = "italiano") -> None:
    """Renderizza il tab KDP. Da chiamare dentro `with tabs[8]:` nell'app principale."""
    st.subheader("Preflight copertina e illustrazioni Amazon KDP")
    st.caption("Controlla dimensioni, risoluzione, file PDF e immagini; indica in modo esplicito cosa correggere. I file originali non vengono modificati.")
    if Image is None:
        st.error("Manca Pillow. Aggiungi `Pillow` ai requirements e riavvia l'app.")
        return
    if fitz is None:
        st.error("Manca PyMuPDF. Aggiungi `PyMuPDF` ai requirements e riavvia l'app.")
        return

    edition = st.selectbox("Formato KDP", ["Kindle eBook", "Paperback", "Hardcover"], key="kdp_edition")
    is_ebook = edition == "Kindle eBook"
    is_paperback = edition == "Paperback"
    file_types = ["jpg", "jpeg", "png", "tif", "tiff"] if is_ebook else ["pdf"]
    cover = st.file_uploader(
        "Copertina finale" + (" (solo fronte)" if is_ebook else " (PDF completo: retro + dorso + fronte)"),
        type=file_types, key=f"kdp_cover_{edition}",
        help=("Per ebook carica un JPG, PNG o TIFF." if is_ebook else "Per paperback/hardcover KDP richiede un PDF completo di una pagina.")
    )
    interior = None
    template = None
    if not is_ebook:
        interior = st.file_uploader("Interno impaginato (PDF) — per controllare pagine e DPI delle illustrazioni", type=["pdf"], key=f"kdp_interior_{edition}")
        if not is_paperback:
            template = st.file_uploader("Template PDF ufficiale generato dal Cover Calculator KDP (consigliato per hardcover)", type=["pdf"], key="kdp_hardcover_template")
    art_files = st.file_uploader(
        "Illustrazioni originali da controllare (facoltativo, seleziona più file)",
        type=["jpg", "jpeg", "png", "tif", "tiff", "webp"], accept_multiple_files=True,
        key=f"kdp_art_{edition}",
    )

    trim_w, trim_h = 6.0, 9.0
    page_count = 120
    paper = "Bianco B/N"
    bleed = False
    art_w = art_h = 0.0
    template_w = template_h = 0.0
    summary_inputs: dict[str, Any] = {"formato KDP": edition}

    if not is_ebook:
        suggested_pages = pdf_page_count(interior.getvalue()) if interior else None
        suggested_pages = suggested_pages or 120
        even_pages = suggested_pages + (suggested_pages % 2)
        with st.expander("Parametri di stampa per il controllo del formato", expanded=True):
            size_name = st.selectbox("Formato pagina / trim", list(TRIM_PRESETS), index=2, key=f"kdp_trim_{edition}")
            preset_w, preset_h = TRIM_PRESETS[size_name]
            if size_name == "Personalizzato":
                tw, th = st.columns(2)
                trim_w = tw.number_input("Larghezza trim (pollici)", min_value=3.0, max_value=12.0, value=6.0, step=0.01, key=f"kdp_trim_w_{edition}")
                trim_h = th.number_input("Altezza trim (pollici)", min_value=4.0, max_value=14.0, value=9.0, step=0.01, key=f"kdp_trim_h_{edition}")
            else:
                trim_w, trim_h = preset_w, preset_h
                st.caption(f"Trim selezionato: {trim_w:g} × {trim_h:g} in ({trim_w * 25.4:.1f} × {trim_h * 25.4:.1f} mm). Verifica che questo sia il formato scelto nella configurazione KDP.")
            page_count = st.number_input("Pagine interne per il calcolo del dorso", min_value=1, max_value=1000, value=int(even_pages), step=1, key=f"kdp_pages_{edition}", help="Per il paperback il dorso dipende da pagine, tipo di carta e inchiostro. Usa il conteggio approvato/previsto su KDP.")
            if is_paperback:
                paper = st.selectbox("Inchiostro/carta paperback", list(PAPER_FACTORS), key="kdp_paper")
            else:
                st.info("Per l'hardcover la larghezza del dorso e l'area di wrap dipendono dal template KDP: carica il PDF del template oppure inserisci le dimensioni finali mostrate dal Cover Calculator.")
                if template is not None:
                    try:
                        template_w, template_h = _pdf_dimensions(template.getvalue())
                        st.success(f"Template letto: {template_w:.3f} × {template_h:.3f} in.")
                    except Exception as exc:
                        st.error(f"Template non leggibile: {exc}")
                manual_cols = st.columns(2)
                template_w = manual_cols[0].number_input("Larghezza totale da template (in)", min_value=0.0, max_value=30.0, value=float(template_w), step=0.01, key="kdp_hard_w", help="Lascia 0 solo se carichi il template ufficiale PDF.")
                template_h = manual_cols[1].number_input("Altezza totale da template (in)", min_value=0.0, max_value=30.0, value=float(template_h), step=0.01, key="kdp_hard_h", help="Lascia 0 solo se carichi il template ufficiale PDF.")
            bleed = st.checkbox("L'interno è impaginato con bleed (abbondanza)", value=False, key=f"kdp_bleed_{edition}", help="Per un interno con bleed, KDP richiede un PDF; le dimensioni pagina includono 0,125 in sul lato esterno e 0,125 in sopra e sotto.")
            if art_files:
                art_cols = st.columns(2)
                art_w = art_cols[0].number_input("Larghezza finale illustrazioni (in, opzionale)", min_value=0.0, max_value=20.0, value=0.0, step=0.1, key="kdp_art_w", help="Serve per stimare il DPI di file immagine separati. Se non lo sai, lascia 0: il PDF finale verrà misurato direttamente.")
                art_h = art_cols[1].number_input("Altezza finale illustrazioni (in, opzionale)", min_value=0.0, max_value=20.0, value=0.0, step=0.1, key="kdp_art_h")

    if is_paperback:
        expected_w, expected_h, spine = _expected_paperback(float(trim_w), float(trim_h), int(page_count), paper)
        summary_inputs.update({"pagine per il calcolo": int(page_count), "carta/colore": paper,
                               "dorso stimato in": round(spine, 4), "dimensione copertina attesa in": [round(expected_w, 4), round(expected_h, 4)]})
        st.info(f"Dimensione paperback stimata: **{expected_w:.3f} × {expected_h:.3f} in** ({expected_w * 25.4:.1f} × {expected_h * 25.4:.1f} mm); dorso stimato **{spine:.3f} in**. Conferma sempre con il Cover Calculator KDP.")
    elif not is_ebook and template_w > 0 and template_h > 0:
        expected_w, expected_h = float(template_w), float(template_h)
        summary_inputs.update({"dimensione hardcover attesa in": [round(expected_w, 4), round(expected_h, 4)]})
    else:
        expected_w = expected_h = None

    ai_enabled = st.checkbox("Aggiungi analisi visiva AI/OCR su copertina e illustrazioni", value=True, key=f"kdp_ai_{edition}", help="Può individuare testo visibilmente tagliato, refusi leggibili, contrasto e difetti evidenti. Non certifica l'approvazione KDP.")
    ai_limit = st.number_input("Massimo illustrazioni interne da inviare all'AI", min_value=1, max_value=100, value=12, step=1, key="kdp_ai_limit", help="Tutte le immagini nel PDF vengono comunque controllate tecnicamente; il limite riguarda solo l'analisi visiva AI.")
    st.caption("Il controllo DPI del PDF analizza tutte le collocazioni raster. Per le illustrazioni separate, il DPI effettivo si calcola solo se indichi la dimensione finale di stampa.")

    run = st.button("🔎 Controlla copertina e illustrazioni", type="primary", key="kdp_run")
    if run:
        rows: list[dict[str, Any]] = []
        summary = dict(summary_inputs)
        vision_images: list[dict[str, Any]] = []
        cover_data = cover.getvalue() if cover else None
        interior_data = interior.getvalue() if interior else None
        template_data = template.getvalue() if template else None
        if not cover_data and not interior_data and not art_files:
            st.session_state.pop("kdp_report", None)
            st.error("Carica almeno una copertina, un PDF interno o un'illustrazione.")
        else:
            if cover_data:
                name = cover.name
                summary["copertina"] = {"nome": name, "dimensione MB": round(len(cover_data) / 1024 / 1024, 2)}
                if is_ebook:
                    image_info = _check_raster(rows, name, cover_data, "Copertina ebook")
                    if image_info:
                        try:
                            vision_images.append({"id": f"COPERTINA EBOOK · {name}", "data": _vision_ready(cover_data), "mime": "image/jpeg"})
                        except Exception:
                            pass
                else:
                    if len(cover_data) > 650 * 1024 * 1024:
                        _issue(rows, "Errore", "Copertina stampa", "Il PDF supera 650 MB.", "Ottimizza il PDF; KDP indica che file cover oltre 650 MB non vengono convertiti.")
                    elif len(cover_data) > 40 * 1024 * 1024:
                        _issue(rows, "Avviso", "Copertina stampa", f"Il PDF pesa {len(cover_data) / 1024 / 1024:.1f} MB, sopra la dimensione consigliata di circa 40 MB.", "Ottimizza le immagini senza scendere sotto la risoluzione richiesta.")
                    count, actual_w, actual_h = _check_pdf_size(rows, cover_data, "Copertina stampa", expected_w, expected_h)
                    if count == 1:
                        _check_pdf_images(rows, cover_data, "Copertina stampa", is_cover=True)
                        _scan_fonts_and_annotations(rows, cover_data, "Copertina stampa")
                        try:
                            preview = _render_pdf_page(cover_data)
                            vision_images.append({"id": f"COPERTINA STAMPA · {name}", "data": _vision_ready(preview), "mime": "image/jpeg"})
                        except Exception as exc:
                            _issue(rows, "Avviso", "Copertina stampa", f"Anteprima AI non creata: {str(exc)[:160]}.", "Verifica il PDF nel visualizzatore/Print Previewer.")
                    if actual_w and actual_h:
                        summary["dimensione copertina rilevata in"] = [round(actual_w, 4), round(actual_h, 4)]
                    if is_paperback and int(page_count) < 24:
                        _issue(rows, "Errore", "Copertina paperback", f"Pagine indicate: {int(page_count)}; il paperback KDP ha normalmente minimo 24 pagine.", "Controlla il numero di pagine e i requisiti aggiornati per il formato selezionato.")
                    if not is_paperback:
                        if template_data:
                            try:
                                t_w, t_h = _pdf_dimensions(template_data)
                                if expected_w is None or expected_h is None:
                                    expected_w, expected_h = t_w, t_h
                                summary["dimensione template hardcover in"] = [round(t_w, 4), round(t_h, 4)]
                                if actual_w and actual_h and (abs(actual_w - t_w) > 0.02 or abs(actual_h - t_h) > 0.02):
                                    _issue(rows, "Errore", "Copertina hardcover", f"PDF {actual_w:.3f} × {actual_h:.3f} in; template KDP {t_w:.3f} × {t_h:.3f} in.", "Adatta il PDF alle dimensioni esatte del template KDP, comprese aree di wrap e dorso.")
                            except Exception as exc:
                                _issue(rows, "Avviso", "Template hardcover", f"Template non utilizzabile: {str(exc)[:160]}.", "Scarica di nuovo il template PDF dal Cover Calculator KDP.")
                        elif not (expected_w and expected_h):
                            _issue(rows, "Avviso", "Copertina hardcover", "Dimensioni di wrap/dorso non verificate: manca il template PDF o la misura completa del Cover Calculator.", "Carica il template ufficiale KDP o inserisci larghezza e altezza finali mostrate dal calcolatore.")
                    # Il testo sul dorso non è deducibile con certezza dai soli metadati; l'AI può segnalarlo se visibile.
                    if is_paperback and int(page_count) < 79:
                        _issue(rows, "Info", "Copertina paperback", "Con meno di 79 pagine KDP non consente testo sul dorso secondo la guida paperback.", "Se sul dorso è presente testo, rimuovilo oppure verifica l'ultima soglia nella guida KDP.")

            if interior_data:
                summary["interno PDF"] = {"nome": interior.name, "pagine effettive": pdf_page_count(interior_data),
                                           "dimensione MB": round(len(interior_data) / 1024 / 1024, 2)}
                page_total = pdf_page_count(interior_data)
                if page_total is not None and is_paperback and not (24 <= page_total <= 830):
                    _issue(rows, "Avviso", "Interno PDF", f"Il PDF contiene {page_total} pagine; il limite varia per formato, carta e inchiostro (il calcolatore KDP riporta in genere 24–830 per paperback).",
                           "Verifica i limiti disponibili per il trim e l'opzione di stampa selezionati su KDP.")
                if page_total is not None and not is_paperback and not is_ebook and not (75 <= page_total <= 550):
                    _issue(rows, "Avviso", "Interno PDF", f"Il PDF contiene {page_total} pagine; la guida hardcover riporta normalmente 75–550.",
                           "Controlla i limiti aggiornati per l'edizione hardcover e il trim scelto.")
                interior_expected = (float(trim_w) + (BLEED_IN if bleed else 0),
                                     float(trim_h) + (2 * BLEED_IN if bleed else 0))
                if bleed:
                    summary["dimensione pagina attesa con bleed in"] = [round(interior_expected[0], 4), round(interior_expected[1], 4)]
                _check_pdf_images(rows, interior_data, "Illustrazioni interno", expected_size=interior_expected)
                _scan_fonts_and_annotations(rows, interior_data, "Interno PDF")
                if page_total is not None and page_total % 2:
                    _issue(rows, "Info", "Interno PDF", f"Il PDF ha un numero dispari di pagine ({page_total}).",
                           "KDP può aggiungere una pagina bianca per ottenere un conteggio pari; verifica il file elaborato nel previewer.")
                # Prendi immagini raster incorporate: l'analisi tecnica ha già controllato tutte le collocazioni.
                try:
                    with fitz.open(stream=interior_data, filetype="pdf") as doc:
                        seen_xrefs: set[int] = set()
                        for page_no, page in enumerate(doc, 1):
                            for item in page.get_images(full=True):
                                xref = int(item[0])
                                if xref <= 0 or xref in seen_xrefs:
                                    continue
                                seen_xrefs.add(xref)
                                raw = doc.extract_image(xref)
                                blob = raw.get("image", b"")
                                ext = raw.get("ext", "").lower()
                                if blob:
                                    try:
                                        vision_blob = _vision_ready(blob)
                                    except Exception:
                                        continue
                                    vision_images.append({"id": f"ILLUSTRAZIONE INTERNO · pagina {page_no} · {xref}", "data": vision_blob, "mime": "image/jpeg"})
                except Exception:
                    pass

            if art_files:
                for image_file in art_files:
                    data = image_file.getvalue()
                    summary.setdefault("illustrazioni separate", []).append({"nome": image_file.name, "dimensione MB": round(len(data) / 1024 / 1024, 2)})
                    info = _check_raster(rows, image_file.name, data, "Illustrazione separata",
                                         float(art_w) if art_w > 0 else None,
                                         float(art_h) if art_h > 0 else None)
                    if info:
                        try:
                            vision_images.append({"id": f"ILLUSTRAZIONE · {image_file.name}", "data": _vision_ready(data), "mime": "image/jpeg"})
                        except Exception:
                            pass

            if not cover_data:
                cover_label = "Copertina ebook" if is_ebook else f"Copertina {edition.lower()}"
                expected_file = "un'immagine JPG/PNG/TIFF del fronte" if is_ebook else "il PDF completo di una pagina (retro + dorso + fronte)"
                _issue(rows, "Avviso", cover_label, "Non è stata caricata la copertina finale.",
                       f"Carica {expected_file} per controllare la copertina.")
            if not is_ebook and not interior_data:
                _issue(rows, "Info", "Interno PDF", "Nessun PDF interno caricato: le illustrazioni incorporate e le dimensioni pagina non sono state controllate.",
                       "Carica il PDF interno per controllare tutte le pagine e i DPI effettivi delle immagini.")

            if ai_enabled:
                if not api_key_configured:
                    _issue(rows, "Avviso", "Controllo visivo AI", "Chiave Gemini non configurata: OCR e controllo visivo AI saltati.",
                           "Configura GOOGLE_API_KEY nella sidebar; il controllo tecnico è indipendente dall'AI.")
                else:
                    # Copertina sempre inclusa; limita le illustrazioni per contenere le chiamate/costi.
                    cover_images = [item for item in vision_images if item["id"].startswith("COPERTINA")]
                    interior_images = [item for item in vision_images if not item["id"].startswith("COPERTINA")]
                    chosen = cover_images + interior_images[:int(ai_limit)]
                    remaining = max(0, len(interior_images) - int(ai_limit))
                    if remaining:
                        _issue(rows, "Info", "Controllo visivo AI", f"L'AI esaminerà {len(chosen)} immagini; {remaining} illustrazioni restano escluse dall'AI per il limite scelto.",
                               "Aumenta il limite AI e rilancia per estendere il controllo; tutte le immagini interne sono già state controllate tecnicamente.")
                    _run_visual_ai(rows, chosen, model, ask_ai_fn, language)
            summary["controllo AI visivo"] = bool(ai_enabled and api_key_configured)
            st.session_state["kdp_report"] = {"riepilogo": summary, "controlli": rows}

    report = st.session_state.get("kdp_report")
    if report:
        st.markdown("### Report del controllo")
        _report_view(report.get("controlli", []), report.get("riepilogo", {}))

    with st.expander("Cosa non può certificare questo controllo"):
        st.markdown(
            "- Non sostituisce il **Print Previewer KDP**, che verifica il file rispetto alla configurazione effettiva dell'account.\n"
            "- Per paperback il calcolo del dorso è stimato da pagina, carta e inchiostro; confrontalo con il Cover Calculator KDP.\n"
            "- Per hardcover serve il template KDP corretto: il wrap e le cerniere non si possono dedurre con affidabilità dal solo conteggio pagine.\n"
            "- L'AI può non riconoscere refusi o tagli; non garantisce la qualità editoriale, i diritti d'autore, il barcode o l'approvazione KDP.\n"
            "- La verifica del DPI delle immagini interne è precisa quando il PDF contiene immagini raster e le dimensioni di collocazione sono leggibili; elementi vettoriali non hanno DPI raster."
       
        """Controlli tecnici e visivi di preflight per copertine e illustrazioni Amazon KDP.

Modulo aggiuntivo per LIBRO MASTER AI. Non invia né modifica file: produce solo un report.
I controlli automatici non sostituiscono il Print Previewer o l'approvazione KDP.
"""
from __future__ import annotations

import io
import json
import math
import os
import re
from typing import Any

import streamlit as st

try:
    import fitz  # PyMuPDF
except ImportError:  # il tab mostra un messaggio chiaro se il pacchetto manca
    fitz = None

try:
    from PIL import Image, ImageOps
except ImportError:
    Image = None
    ImageOps = None

try:
    from google.genai import types
except ImportError:
    types = None

BLEED_IN = 0.125
PAPER_FACTORS = {
    "Bianco B/N": 0.002252,
    "Crema B/N": 0.0025,
    "Colore standard": 0.002252,
    "Colore premium": 0.002347,
}
TRIM_PRESETS = {
    "5 × 8 in": (5.0, 8.0),
    "5.5 × 8.5 in": (5.5, 8.5),
    "6 × 9 in": (6.0, 9.0),
    "7 × 10 in": (7.0, 10.0),
    "8.5 × 11 in": (8.5, 11.0),
    "Personalizzato": (6.0, 9.0),
}


def _issue(rows: list[dict[str, Any]], severity: str, area: str, problem: str,
           fix: str, where: str = "") -> None:
    rows.append({
        "Gravità": severity,
        "Elemento": area,
        "Dove": where or "—",
        "Problema rilevato": problem,
        "Cosa correggere": fix,
    })


def _in(value_points: float) -> float:
    return float(value_points) / 72.0


def _fmt_in(value: float) -> str:
    return f"{value:.3f} in ({value * 25.4:.1f} mm)"


def _safe_float(value: Any) -> float | None:
    try:
        value = float(value)
        return value if math.isfinite(value) and value > 0 else None
    except (TypeError, ValueError):
        return None


def pdf_page_count(data: bytes) -> int | None:
    if fitz is None:
        return None
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            return doc.page_count
    except Exception:
        return None


def _pdf_dimensions(data: bytes, page_index: int = 0) -> tuple[float, float]:
    if fitz is None:
        raise RuntimeError("PyMuPDF non installato. Aggiungi PyMuPDF ai requirements.")
    with fitz.open(stream=data, filetype="pdf") as doc:
        if doc.is_encrypted:
            raise ValueError("Il PDF è protetto da password.")
        if doc.page_count <= page_index:
            raise ValueError("Il PDF non contiene la pagina richiesta.")
        rect = doc.load_page(page_index).mediabox
        return _in(rect.width), _in(rect.height)


def _check_pdf_size(rows: list[dict[str, Any]], data: bytes, label: str,
                    expected_w: float | None = None, expected_h: float | None = None,
                    tolerance_in: float = 0.02) -> tuple[int, float | None, float | None]:
    if fitz is None:
        _issue(rows, "Errore", label, "Non posso leggere il PDF: manca PyMuPDF.",
               "Installa il pacchetto con `pip install PyMuPDF` e riavvia Streamlit.")
        return 0, None, None
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            if doc.is_encrypted:
                _issue(rows, "Errore", label, "PDF protetto da password.",
                       "Rimuovi la password e carica un PDF leggibile.")
                return 0, None, None
            count = doc.page_count
            if count != 1 and label == "Copertina stampa":
                _issue(rows, "Errore", label, f"La copertina contiene {count} pagine.",
                       "Esporta un unico PDF di una pagina con retro, dorso e fronte sullo stesso foglio.")
            if count < 1:
                _issue(rows, "Errore", label, "Il PDF non contiene pagine.", "Esporta nuovamente il PDF.")
                return count, None, None
            page = doc.load_page(0)
            box = page.mediabox
            actual_w, actual_h = _in(box.width), _in(box.height)
            if page.rotation:
                _issue(rows, "Avviso", label, f"La prima pagina ha rotazione PDF di {page.rotation}°.",
                       "Appiattisci la rotazione ed esporta la pagina con orientamento e dimensioni già corretti.")
            crop = page.cropbox
            if abs(crop.width - box.width) > 1 or abs(crop.height - box.height) > 1:
                _issue(rows, "Avviso", label, "MediaBox e CropBox del PDF non coincidono.",
                       "Controlla che non ci siano segni di taglio o aree esterne al formato finale.")
            if expected_w is not None and expected_h is not None:
                diff_w = abs(actual_w - expected_w)
                diff_h = abs(actual_h - expected_h)
                if diff_w > tolerance_in or diff_h > tolerance_in:
                    _issue(rows, "Errore", label,
                           f"Dimensione PDF {actual_w:.3f} × {actual_h:.3f} in; attesa {expected_w:.3f} × {expected_h:.3f} in.",
                           f"Imposta il documento a {_fmt_in(expected_w)} × {_fmt_in(expected_h)} (secondo i parametri scelti) ed esporta senza bordi extra.")
                else:
                    _issue(rows, "OK", label,
                           f"Dimensione PDF {actual_w:.3f} × {actual_h:.3f} in: corrisponde alle dimensioni attese entro {tolerance_in:.2f} in.",
                           "Nessuna correzione dimensionale rilevata.")
            return count, actual_w, actual_h
    except Exception as exc:
        _issue(rows, "Errore", label, f"PDF non leggibile: {str(exc)[:180]}.",
               "Esporta o salva di nuovo il file in PDF e verifica che non sia danneggiato.")
        return 0, None, None


def _image_info(name: str, data: bytes) -> dict[str, Any]:
    if Image is None:
        raise RuntimeError("Pillow non installato. Aggiungi Pillow ai requirements.")
    try:
        image = Image.open(io.BytesIO(data))
        image = ImageOps.exif_transpose(image) if ImageOps is not None else image
        width, height = image.size
        dpi_info = image.info.get("dpi")
        dpi_x = _safe_float(dpi_info[0]) if isinstance(dpi_info, (tuple, list)) and dpi_info else None
        dpi_y = _safe_float(dpi_info[1]) if isinstance(dpi_info, (tuple, list)) and len(dpi_info) > 1 else dpi_x
        info = {
            "name": name,
            "format": (image.format or os.path.splitext(name)[1].lstrip(".")).upper(),
            "width": width,
            "height": height,
            "mode": image.mode,
            "dpi_x": dpi_x,
            "dpi_y": dpi_y,
            "frames": int(getattr(image, "n_frames", 1)),
            "alpha": "A" in image.getbands() or "transparency" in image.info,
        }
        image.close()
        return info
    except Exception as exc:
        raise ValueError(f"Immagine non leggibile: {str(exc)[:160]}") from exc


def _check_raster(rows: list[dict[str, Any]], name: str, data: bytes, role: str,
                  target_w_in: float | None = None, target_h_in: float | None = None) -> dict[str, Any] | None:
    try:
        info = _image_info(name, data)
    except Exception as exc:
        _issue(rows, "Errore", role, str(exc), "Esporta nuovamente l'immagine in JPG, PNG o TIFF non danneggiato.", name)
        return None
    width, height = info["width"], info["height"]
    fmt = info["format"]
    ext = os.path.splitext(name)[1].lower()
    if fmt not in {"JPEG", "JPG", "PNG", "TIFF", "TIF", "WEBP", "GIF"}:
        _issue(rows, "Errore", role, f"Formato immagine {fmt or ext or 'sconosciuto'} non riconosciuto.",
               "Per Kindle usa JPG/PNG/TIFF; per stampa carica il PDF completo della copertina.", name)
    if role == "Copertina ebook":
        if width < 625 or height < 1000:
            _issue(rows, "Errore", role, f"Immagine {width} × {height} px: sotto il minimo consigliato di 625 × 1000 px.",
                   "Crea una nuova copertina alla risoluzione originale; non ingrandire artificialmente quella piccola.", name)
        elif width < 1600 or height < 2560:
            _issue(rows, "Avviso", role, f"Immagine {width} × {height} px: inferiore alla dimensione ideale 1600 × 2560 px.",
                   "Se possibile esporta l'originale a 1600 × 2560 px o più, mantenendo le proporzioni.", name)
        if width > 10000 or height > 10000:
            _issue(rows, "Errore", role, f"Dimensione {width} × {height} px: supera il limite di 10.000 px per lato.",
                   "Ridimensiona/esporta l'immagine entro 10.000 px per lato.", name)
        ratio = height / max(1, width)
        if abs(ratio - 1.6) > 0.12:
            _issue(rows, "Avviso", role, f"Rapporto altezza/larghezza {ratio:.2f}:1; il riferimento Kindle è circa 1.6:1.",
                   "Verifica che la copertina mantenga le proporzioni desiderate e non venga stirata o ritagliata.", name)
        if len(data) >= 50 * 1024 * 1024:
            _issue(rows, "Errore", role, f"File di {len(data) / 1024 / 1024:.1f} MB: raggiunge/supera 50 MB.",
                   "Esporta un JPG/TIFF ottimizzato sotto 50 MB senza ridurre eccessivamente la qualità.", name)
        if info["mode"] == "CMYK":
            _issue(rows, "Avviso", role, "Immagine CMYK; per Kindle è preferibile RGB/sRGB.",
                   "Converti una copia in RGB/sRGB e controlla che i colori non cambino.", name)
        if info["alpha"]:
            _issue(rows, "Avviso", role, "L'immagine contiene trasparenza.",
                   "Appiattisci la copertina su uno sfondo opaco e controlla il risultato sui dispositivi Kindle.", name)
        _issue(rows, "OK", role,
               f"Risoluzione pixel: {width} × {height}; formato {fmt}; profilo/modalità {info['mode']}.",
               "Controllo raster completato. Il DPI incorporato è informativo: per ebook contano soprattutto i pixel.", name)
    else:
        if target_w_in and target_h_in:
            dpi_x = width / target_w_in
            dpi_y = height / target_h_in
            eff = min(dpi_x, dpi_y)
            if eff < 300:
                _issue(rows, "Errore", role,
                       f"Risoluzione effettiva stimata {eff:.0f} DPI alla dimensione di stampa {target_w_in:.2f} × {target_h_in:.2f} in.",
                       "Usa un originale più grande o riduci la dimensione di inserimento finché entrambe le direzioni raggiungono almeno 300 DPI.", name)
            else:
                _issue(rows, "OK", role, f"Risoluzione effettiva stimata {eff:.0f} DPI alla dimensione indicata.",
                       "Nessun problema di risoluzione rilevato a questa dimensione.", name)
        elif info["dpi_x"] and info["dpi_y"]:
            minimum = min(info["dpi_x"], info["dpi_y"])
            severity = "Avviso" if minimum < 300 else "OK"
            msg = f"Metadati immagine: {info['dpi_x']:.0f} × {info['dpi_y']:.0f} DPI; {width} × {height} px."
            fix = ("Il DPI nei metadati non garantisce il DPI effettivo: controlla la dimensione di stampa nel PDF finale." if minimum >= 300
                   else "Aumenta la risoluzione originale oppure riduci la dimensione di stampa; verifica nel PDF finale.")
            _issue(rows, severity, role, msg, fix, name)
        else:
            _issue(rows, "Avviso", role, f"Immagine {width} × {height} px senza metadati DPI attendibili.",
                   "Indica la dimensione finale nel PDF; solo lì si può calcolare il DPI effettivo. KDP raccomanda almeno 300 DPI.", name)
        if info["mode"] == "CMYK":
            _issue(rows, "Info", role, "Immagine in modalità CMYK.",
                   "Verifica l'opzione di stampa scelta e controlla una prova colore; la gestione colore varia tra ebook e stampa.", name)
        if info["alpha"]:
            _issue(rows, "Avviso", role, "Immagine con trasparenza/livello alpha.",
                   "Appiattisci gli elementi trasparenti nel file finale e controlla il PDF esportato.", name)
        if info["frames"] > 1:
            _issue(rows, "Avviso", role, f"Immagine con {info['frames']} fotogrammi/pagine.",
                   "Esporta l'illustrazione destinata al libro come immagine statica.", name)
    return info


def _image_placements(doc: Any, page: Any, page_number: int, rows: list[dict[str, Any]], area: str,
                      low_dpi_limit: int = 30) -> tuple[int, list[dict[str, Any]]]:
    count = 0
    extracted: list[dict[str, Any]] = []
    try:
        images = page.get_images(full=True)
    except Exception:
        images = []
    for item in images:
        try:
            xref = int(item[0])
            smask = int(item[1]) if len(item) > 1 else 0
            wpx, hpx = int(item[2]), int(item[3])
            if xref <= 0 or wpx <= 0 or hpx <= 0:
                continue
            rects = page.get_image_rects(xref)
            if not rects:
                continue
            if smask > 0:
                _issue(rows, "Avviso", area, "Immagine con maschera di trasparenza nel PDF.",
                       "Appiattisci le trasparenze e ricontrolla il PDF esportato.", f"pagina {page_number}")
            for rect in rects:
                rw, rh = float(rect.width), float(rect.height)
                if rw <= 0 or rh <= 0:
                    continue
                rect_ratio = rw / rh
                img_ratio = wpx / hpx
                if abs(math.log(max(1e-9, rect_ratio / img_ratio))) <= abs(math.log(max(1e-9, rect_ratio * img_ratio))):
                    dpi_x, dpi_y = wpx / (rw / 72), hpx / (rh / 72)
                else:
                    dpi_x, dpi_y = hpx / (rw / 72), wpx / (rh / 72)
                effective = min(dpi_x, dpi_y)
                count += 1
                if effective < 300:
                    _issue(rows, "Errore", area,
                           f"Illustrazione a {effective:.0f} DPI effettivi (minimo raccomandato 300); immagine sorgente {wpx} × {hpx} px.",
                           "Sostituisci con un originale a risoluzione maggiore oppure riduci la dimensione con cui è collocata nella pagina.",
                           f"pagina {page_number}")
                elif effective >= 600:
                    _issue(rows, "Info", area, f"Illustrazione a {effective:.0f} DPI effettivi; risoluzione alta.",
                           "Nessuna correzione richiesta; immagini eccessivamente pesanti possono aumentare il PDF.",
                           f"pagina {page_number}")
            # Conserva una sola copia per xref/pagina per la revisione visiva AI.
            if len(extracted) < low_dpi_limit:
                try:
                    raw = doc.extract_image(xref)
                    blob = raw.get("image", b"")
                    ext = raw.get("ext", "").lower()
                    if blob:
                        mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                                "webp": "image/webp", "tif": "image/tiff", "tiff": "image/tiff"}.get(ext, "application/octet-stream")
                        extracted.append({"id": f"pagina {page_number} · immagine {xref}", "data": blob, "mime": mime})
                except Exception:
                    pass
        except Exception:
            continue
    return count, extracted


def _check_pdf_images(rows: list[dict[str, Any]], data: bytes, area: str,
                      expected_size: tuple[float, float] | None = None,
                      is_cover: bool = False) -> tuple[int, int]:
    if fitz is None:
        _issue(rows, "Errore", area, "Controllo PDF non disponibile: manca PyMuPDF.",
               "Installa PyMuPDF e ripeti il controllo.")
        return 0, 0
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            if doc.is_encrypted:
                _issue(rows, "Errore", area, "PDF protetto da password.", "Rimuovi la password e ricarica il PDF.")
                return 0, 0
            if not is_cover and expected_size:
                exp_w, exp_h = expected_size
                wrong = []
                for number, page in enumerate(doc, 1):
                    rect = page.mediabox
                    actual_w, actual_h = _in(rect.width), _in(rect.height)
                    direct = abs(actual_w - exp_w) <= 0.03 and abs(actual_h - exp_h) <= 0.03
                    rotated = abs(actual_w - exp_h) <= 0.03 and abs(actual_h - exp_w) <= 0.03
                    if not direct and not rotated:
                        wrong.append((number, actual_w, actual_h))
                if wrong:
                    first = ", ".join(f"p.{n}: {w:.3f}×{h:.3f} in" for n, w, h in wrong[:8])
                    _issue(rows, "Errore", area,
                           f"Formato pagina diverso dal trim impostato su {len(wrong)}/{doc.page_count} pagine ({first}).",
                           f"Imposta tutte le pagine a {_fmt_in(exp_w)} × {_fmt_in(exp_h)} (oppure l'orientamento ruotato se previsto) e riesporta il PDF.")
                else:
                    _issue(rows, "OK", area,
                           f"Formato pagina coerente con il trim impostato su tutte le {doc.page_count} pagine.",
                           "Nessuna correzione dimensionale rilevata.")
            placements = 0
            ai_images: list[dict[str, Any]] = []
            low_count_before = sum(1 for r in rows if r.get("Gravità") == "Errore" and r.get("Elemento") == area and "DPI effettivi" in r.get("Problema rilevato", ""))
            for page_no, page in enumerate(doc, 1):
                page_count, page_imgs = _image_placements(doc, page, page_no, rows, area)
                placements += page_count
                ai_images.extend(page_imgs)
            if placements == 0:
                _issue(rows, "Info", area, "Non sono state trovate immagini raster estraibili nel PDF.",
                       "Se le illustrazioni sono vettoriali, il controllo DPI raster non si applica; controlla le pagine nel Print Previewer.")
            else:
                low_count_after = sum(1 for r in rows if r.get("Gravità") == "Errore" and r.get("Elemento") == area and "DPI effettivi" in r.get("Problema rilevato", ""))
                _issue(rows, "Info", area,
                       f"Rilevate {placements} collocazioni raster; {max(0, low_count_after - low_count_before)} sotto 300 DPI.",
                       "Il DPI è stimato dalle dimensioni di ciascuna immagine nel PDF; illustrazioni vettoriali non hanno DPI raster.")
            # Ritorna solo immagini uniche per questa pagina/documento, senza duplicare xref già estratti.
            unique: dict[str, dict[str, Any]] = {}
            for item in ai_images:
                unique.setdefault(item["id"].split("· immagine")[-1].strip(), item)
            return placements, list(unique.values())
    except Exception as exc:
        _issue(rows, "Errore", area, f"PDF non leggibile durante l'analisi immagini: {str(exc)[:160]}.",
               "Esporta nuovamente il PDF e verifica che le immagini siano incorporate.")
        return 0, 0


def _scan_fonts_and_annotations(rows: list[dict[str, Any]], data: bytes, area: str) -> None:
    if fitz is None:
        return
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            missing_fonts: set[str] = set()
            annotations: list[int] = []
            for page_no, page in enumerate(doc, 1):
                try:
                    for font in doc.get_page_fonts(page_no - 1, full=True):
                        xref = int(font[0])
                        basefont = str(font[3] or "font senza nome")
                        if xref <= 0:
                            missing_fonts.add(basefont)
                            continue
                        try:
                            extracted = doc.extract_font(xref)
                            content = extracted[3] if len(extracted) > 3 else b""
                            if not content:
                                missing_fonts.add(basefont)
                        except Exception:
                            missing_fonts.add(basefont)
                except Exception:
                    pass
                try:
                    if list(page.annots() or []):
                        annotations.append(page_no)
                except Exception:
                    pass
            if missing_fonts:
                names = ", ".join(sorted(missing_fonts)[:6])
                _issue(rows, "Avviso", area, f"Possibili font non incorporati: {names}.",
                       "Esporta il PDF con tutti i font incorporati o converti il testo in tracciati, quindi verifica il risultato visivo.")
            else:
                _issue(rows, "Info", area, "Nessun font non incorporato rilevato dai metadati PDF.",
                       "Controllo automatico indicativo: verifica comunque l'esportazione e i font nel Print Previewer.")
            if annotations:
                _issue(rows, "Avviso", area, f"Annotazioni presenti nelle pagine: {annotations[:12]}.",
                       "Rimuovi commenti, annotazioni o markup prima di caricare il PDF.")
    except Exception:
        pass


def _expected_paperback(trim_w: float, trim_h: float, pages: int, paper: str) -> tuple[float, float, float]:
    spine = pages * PAPER_FACTORS[paper]
    total_w = 2 * trim_w + spine + 2 * BLEED_IN
    total_h = trim_h + 2 * BLEED_IN
    return total_w, total_h, spine


def _render_pdf_page(data: bytes, page_no: int = 0, max_side: int = 1800) -> bytes:
    if fitz is None:
        raise RuntimeError("PyMuPDF necessario per creare l'anteprima AI del PDF.")
    with fitz.open(stream=data, filetype="pdf") as doc:
        if doc.page_count <= page_no:
            raise ValueError("Pagina PDF non trovata.")
        page = doc.load_page(page_no)
        rect = page.rect
        scale = min(max_side / max(float(rect.width), 1), max_side / max(float(rect.height), 1))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        return pix.tobytes("png")


def _vision_ready(data: bytes, max_side: int = 1800) -> bytes:
    if Image is None:
        return data
    with Image.open(io.BytesIO(data)) as image:
        image = ImageOps.exif_transpose(image) if ImageOps is not None else image
        image.thumbnail((max_side, max_side))
        if image.mode in ("RGBA", "LA") or "transparency" in image.info:
            rgba = image.convert("RGBA")
            bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            bg.alpha_composite(rgba)
            image = bg.convert("RGB")
        else:
            image = image.convert("RGB")
        out = io.BytesIO()
        image.save(out, format="JPEG", quality=88, optimize=True)
        return out.getvalue()


def _parse_ai_list(raw: str) -> list[dict[str, Any]]:
    text = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", str(raw).strip(), flags=re.I)
    try:
        obj = json.loads(text)
    except Exception:
        start, end = text.find("["), text.rfind("]")
        if start < 0 or end < start:
            start, end = text.find("{"), text.rfind("}")
            if start < 0 or end < start:
                raise ValueError("La risposta visiva AI non contiene JSON valido.")
            obj = json.loads(text[start:end + 1])
        else:
            obj = json.loads(text[start:end + 1])
    if isinstance(obj, dict):
        obj = obj.get("problemi", obj.get("issues", []))
    if not isinstance(obj, list):
        raise ValueError("La risposta AI non è una lista JSON.")
    return [item for item in obj if isinstance(item, dict)]


def _run_visual_ai(rows: list[dict[str, Any]], images: list[dict[str, Any]], model: str,
                   ask_ai_fn: Any, language: str) -> None:
    if not images:
        _issue(rows, "Info", "Controllo visivo AI", "Nessuna immagine disponibile per il controllo visivo.",
               "Carica una copertina o immagini del libro.")
        return
    if types is None:
        _issue(rows, "Errore", "Controllo visivo AI", "SDK Google GenAI non disponibile.",
               "Verifica l'installazione del pacchetto google-genai.")
        return
    system = (
        f"Sei un revisore visivo editoriale che valuta immagini per Amazon KDP. Lingua delle segnalazioni: {language}. "
        "Analizza solo problemi effettivamente visibili: testo che sembra tagliato o illeggibile, refusi chiaramente leggibili, "
        "basso contrasto, sovrapposizioni, oggetti importanti troppo vicini al bordo, artefatti o immagini sfocate. "
        "Non inventare errori, non riscrivere l'opera e non dichiarare mai che il file è garantito conforme/approvato da KDP. "
        "Se non puoi leggere una parola, dillo invece di indovinare. Rispondi solo con JSON valido: "
        "lista di oggetti {image_id, gravita (Errore|Avviso|Info), problema, cosa_correggere, testo_letto}. "
        "Se non rilevi problemi concreti, rispondi []."
    )
    total = (len(images) + 3) // 4
    progress = st.progress(0, text="Analisi visiva AI delle immagini...")
    for batch_no in range(total):
        batch = images[batch_no * 4:(batch_no + 1) * 4]
        parts: list[Any] = []
        for item in batch:
            parts.append(f"ID IMMAGINE: {item['id']}")
            parts.append(types.Part.from_bytes(data=item["data"], mime_type=item.get("mime", "image/jpeg")))
        prompt = (
            "Esamina ogni immagine allegata separatamente e collega ogni rilievo al suo ID. "
            "Per una copertina controlla testo, titolo/autore leggibili, tagli e contrasto; per illustrazioni controlla "
            "solo difetti visivi oggettivi o testo visibile. Non valutare DPI o dimensioni, già controllati a parte. "
            "Restituisci esclusivamente la lista JSON specificata."
        )
        parts.append(prompt)
        try:
            raw = ask_ai_fn(parts, system, model, 0.0, json_mode=True)
            findings = _parse_ai_list(raw)
            for finding in findings:
                image_id = str(finding.get("image_id", "immagine non identificata"))[:120]
                severity = str(finding.get("gravita", "Avviso"))
                if severity not in {"Errore", "Avviso", "Info"}:
                    severity = "Avviso"
                problem = str(finding.get("problema", "Problema visivo da verificare."))[:500]
                fix = str(finding.get("cosa_correggere", "Controlla l'immagine ingrandita e correggi il punto indicato."))[:500]
                quote = str(finding.get("testo_letto", "")).strip()
                if quote:
                    problem += f" · testo visibile: «{quote[:180]}»"
                _issue(rows, severity, "Controllo visivo AI", problem, fix, image_id)
            if not findings:
                _issue(rows, "OK", "Controllo visivo AI", "Nessun difetto visivo certo rilevato nel gruppo.",
                       "Controllo automatico indicativo: fai comunque una verifica manuale al 100%.",
                       ", ".join(item["id"] for item in batch))
        except Exception as exc:
            msg = str(exc)
            _issue(rows, "Avviso", "Controllo visivo AI",
                   f"Analisi AI non completata: {msg[:240]}.",
                   "Controlla la chiave/quota/modello Gemini e riprova; i controlli tecnici già completati restano validi.",
                   ", ".join(item["id"] for item in batch))
            if any(token in msg.upper() for token in ("429", "RESOURCE_EXHAUSTED", "QUOTA")):
                for remaining in range(batch_no + 1, total):
                    rest = images[(remaining) * 4:(remaining + 1) * 4]
                    _issue(rows, "Avviso", "Controllo visivo AI", "Gruppo non analizzato perché la quota API è esaurita.",
                           "Ripeti l'analisi dopo il ripristino della quota oppure controlla manualmente.",
                           ", ".join(item["id"] for item in rest))
                break
        progress.progress(min(1.0, (batch_no + 1) / total), text=f"Gruppo AI {batch_no + 1}/{total}")
    progress.empty()


def _report_view(rows: list[dict[str, Any]], summary: dict[str, Any]) -> None:
    errors = sum(row["Gravità"] == "Errore" for row in rows)
    warnings = sum(row["Gravità"] == "Avviso" for row in rows)
    infos = sum(row["Gravità"] in {"Info", "OK"} for row in rows)
    c1, c2, c3 = st.columns(3)
    c1.metric("Errori da correggere", errors)
    c2.metric("Avvisi da verificare", warnings)
    c3.metric("Controlli informativi", infos)
    if errors:
        st.error(f"Rilevati {errors} problemi tecnici/visivi. Consulta la colonna **Cosa correggere**.")
    elif warnings:
        st.warning("Non sono stati rilevati errori bloccanti automatici, ma restano avvisi da verificare manualmente.")
    else:
        st.success("Nessun problema rilevato dai controlli eseguiti.")
    st.warning("Questo preflight non garantisce l'approvazione KDP: controlla sempre il PDF nel Print Previewer e ordina una prova di stampa quando possibile.")
    if summary:
        st.json(summary)
    if rows:
        st.dataframe(rows, hide_index=True, use_container_width=True)
        import pandas as pd
        csv = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")
        st.download_button("Scarica report KDP CSV", csv, "report_preflight_kdp.csv", "text/csv", key="kdp_report_csv")
        st.download_button("Scarica report KDP JSON", json.dumps({"riepilogo": summary, "controlli": rows}, ensure_ascii=False, indent=2).encode("utf-8"),
                           "report_preflight_kdp.json", "application/json", key="kdp_report_json")


def render_kdp_tab(model: str, ask_ai_fn: Any, api_key_configured: bool, language: str = "italiano") -> None:
    """Renderizza il tab KDP. Da chiamare dentro `with tabs[8]:` nell'app principale."""
    st.subheader("Preflight copertina e illustrazioni Amazon KDP")
    st.caption("Controlla dimensioni, risoluzione, file PDF e immagini; indica in modo esplicito cosa correggere. I file originali non vengono modificati.")
    if Image is None:
        st.error("Manca Pillow. Aggiungi `Pillow` ai requirements e riavvia l'app.")
        return
    if fitz is None:
        st.error("Manca PyMuPDF. Aggiungi `PyMuPDF` ai requirements e riavvia l'app.")
        return

    edition = st.selectbox("Formato KDP", ["Kindle eBook", "Paperback", "Hardcover"], key="kdp_edition")
    is_ebook = edition == "Kindle eBook"
    is_paperback = edition == "Paperback"
    file_types = ["jpg", "jpeg", "png", "tif", "tiff"] if is_ebook else ["pdf"]
    cover = st.file_uploader(
        "Copertina finale" + (" (solo fronte)" if is_ebook else " (PDF completo: retro + dorso + fronte)"),
        type=file_types, key=f"kdp_cover_{edition}",
        help=("Per ebook carica un JPG, PNG o TIFF." if is_ebook else "Per paperback/hardcover KDP richiede un PDF completo di una pagina.")
    )
    interior = None
    template = None
    if not is_ebook:
        interior = st.file_uploader("Interno impaginato (PDF) — per controllare pagine e DPI delle illustrazioni", type=["pdf"], key=f"kdp_interior_{edition}")
        if not is_paperback:
            template = st.file_uploader("Template PDF ufficiale generato dal Cover Calculator KDP (consigliato per hardcover)", type=["pdf"], key="kdp_hardcover_template")
    art_files = st.file_uploader(
        "Illustrazioni originali da controllare (facoltativo, seleziona più file)",
        type=["jpg", "jpeg", "png", "tif", "tiff", "webp"], accept_multiple_files=True,
        key=f"kdp_art_{edition}",
    )

    trim_w, trim_h = 6.0, 9.0
    page_count = 120
    paper = "Bianco B/N"
    bleed = False
    art_w = art_h = 0.0
    template_w = template_h = 0.0
    summary_inputs: dict[str, Any] = {"formato KDP": edition}

    if not is_ebook:
        suggested_pages = pdf_page_count(interior.getvalue()) if interior else None
        suggested_pages = suggested_pages or 120
        even_pages = suggested_pages + (suggested_pages % 2)
        with st.expander("Parametri di stampa per il controllo del formato", expanded=True):
            size_name = st.selectbox("Formato pagina / trim", list(TRIM_PRESETS), index=2, key=f"kdp_trim_{edition}")
            preset_w, preset_h = TRIM_PRESETS[size_name]
            if size_name == "Personalizzato":
                tw, th = st.columns(2)
                trim_w = tw.number_input("Larghezza trim (pollici)", min_value=3.0, max_value=12.0, value=6.0, step=0.01, key=f"kdp_trim_w_{edition}")
                trim_h = th.number_input("Altezza trim (pollici)", min_value=4.0, max_value=14.0, value=9.0, step=0.01, key=f"kdp_trim_h_{edition}")
            else:
                trim_w, trim_h = preset_w, preset_h
                st.caption(f"Trim selezionato: {trim_w:g} × {trim_h:g} in ({trim_w * 25.4:.1f} × {trim_h * 25.4:.1f} mm). Verifica che questo sia il formato scelto nella configurazione KDP.")
            page_count = st.number_input("Pagine interne per il calcolo del dorso", min_value=1, max_value=1000, value=int(even_pages), step=1, key=f"kdp_pages_{edition}", help="Per il paperback il dorso dipende da pagine, tipo di carta e inchiostro. Usa il conteggio approvato/previsto su KDP.")
            if is_paperback:
                paper = st.selectbox("Inchiostro/carta paperback", list(PAPER_FACTORS), key="kdp_paper")
            else:
                st.info("Per l'hardcover la larghezza del dorso e l'area di wrap dipendono dal template KDP: carica il PDF del template oppure inserisci le dimensioni finali mostrate dal Cover Calculator.")
                if template is not None:
                    try:
                        template_w, template_h = _pdf_dimensions(template.getvalue())
                        st.success(f"Template letto: {template_w:.3f} × {template_h:.3f} in.")
                    except Exception as exc:
                        st.error(f"Template non leggibile: {exc}")
                manual_cols = st.columns(2)
                template_w = manual_cols[0].number_input("Larghezza totale da template (in)", min_value=0.0, max_value=30.0, value=float(template_w), step=0.01, key="kdp_hard_w", help="Lascia 0 solo se carichi il template ufficiale PDF.")
                template_h = manual_cols[1].number_input("Altezza totale da template (in)", min_value=0.0, max_value=30.0, value=float(template_h), step=0.01, key="kdp_hard_h", help="Lascia 0 solo se carichi il template ufficiale PDF.")
            bleed = st.checkbox("L'interno è impaginato con bleed (abbondanza)", value=False, key=f"kdp_bleed_{edition}", help="Per un interno con bleed, KDP richiede un PDF; le dimensioni pagina includono 0,125 in sul lato esterno e 0,125 in sopra e sotto.")
            if art_files:
                art_cols = st.columns(2)
                art_w = art_cols[0].number_input("Larghezza finale illustrazioni (in, opzionale)", min_value=0.0, max_value=20.0, value=0.0, step=0.1, key="kdp_art_w", help="Serve per stimare il DPI di file immagine separati. Se non lo sai, lascia 0: il PDF finale verrà misurato direttamente.")
                art_h = art_cols[1].number_input("Altezza finale illustrazioni (in, opzionale)", min_value=0.0, max_value=20.0, value=0.0, step=0.1, key="kdp_art_h")

    if is_paperback:
        expected_w, expected_h, spine = _expected_paperback(float(trim_w), float(trim_h), int(page_count), paper)
        summary_inputs.update({"pagine per il calcolo": int(page_count), "carta/colore": paper,
                               "dorso stimato in": round(spine, 4), "dimensione copertina attesa in": [round(expected_w, 4), round(expected_h, 4)]})
        st.info(f"Dimensione paperback stimata: **{expected_w:.3f} × {expected_h:.3f} in** ({expected_w * 25.4:.1f} × {expected_h * 25.4:.1f} mm); dorso stimato **{spine:.3f} in**. Conferma sempre con il Cover Calculator KDP.")
    elif not is_ebook and template_w > 0 and template_h > 0:
        expected_w, expected_h = float(template_w), float(template_h)
        summary_inputs.update({"dimensione hardcover attesa in": [round(expected_w, 4), round(expected_h, 4)]})
    else:
        expected_w = expected_h = None

    ai_enabled = st.checkbox("Aggiungi analisi visiva AI/OCR su copertina e illustrazioni", value=True, key=f"kdp_ai_{edition}", help="Può individuare testo visibilmente tagliato, refusi leggibili, contrasto e difetti evidenti. Non certifica l'approvazione KDP.")
    ai_limit = st.number_input("Massimo illustrazioni interne da inviare all'AI", min_value=1, max_value=100, value=12, step=1, key="kdp_ai_limit", help="Tutte le immagini nel PDF vengono comunque controllate tecnicamente; il limite riguarda solo l'analisi visiva AI.")
    st.caption("Il controllo DPI del PDF analizza tutte le collocazioni raster. Per le illustrazioni separate, il DPI effettivo si calcola solo se indichi la dimensione finale di stampa.")

    run = st.button("🔎 Controlla copertina e illustrazioni", type="primary", key="kdp_run")
    if run:
        rows: list[dict[str, Any]] = []
        summary = dict(summary_inputs)
        vision_images: list[dict[str, Any]] = []
        cover_data = cover.getvalue() if cover else None
        interior_data = interior.getvalue() if interior else None
        template_data = template.getvalue() if template else None
        if not cover_data and not interior_data and not art_files:
            st.session_state.pop("kdp_report", None)
            st.error("Carica almeno una copertina, un PDF interno o un'illustrazione.")
        else:
            if cover_data:
                name = cover.name
                summary["copertina"] = {"nome": name, "dimensione MB": round(len(cover_data) / 1024 / 1024, 2)}
                if is_ebook:
                    image_info = _check_raster(rows, name, cover_data, "Copertina ebook")
                    if image_info:
                        try:
                            vision_images.append({"id": f"COPERTINA EBOOK · {name}", "data": _vision_ready(cover_data), "mime": "image/jpeg"})
                        except Exception:
                            pass
                else:
                    if len(cover_data) > 650 * 1024 * 1024:
                        _issue(rows, "Errore", "Copertina stampa", "Il PDF supera 650 MB.", "Ottimizza il PDF; KDP indica che file cover oltre 650 MB non vengono convertiti.")
                    elif len(cover_data) > 40 * 1024 * 1024:
                        _issue(rows, "Avviso", "Copertina stampa", f"Il PDF pesa {len(cover_data) / 1024 / 1024:.1f} MB, sopra la dimensione consigliata di circa 40 MB.", "Ottimizza le immagini senza scendere sotto la risoluzione richiesta.")
                    count, actual_w, actual_h = _check_pdf_size(rows, cover_data, "Copertina stampa", expected_w, expected_h)
                    if count == 1:
                        _check_pdf_images(rows, cover_data, "Copertina stampa", is_cover=True)
                        _scan_fonts_and_annotations(rows, cover_data, "Copertina stampa")
                        try:
                            preview = _render_pdf_page(cover_data)
                            vision_images.append({"id": f"COPERTINA STAMPA · {name}", "data": _vision_ready(preview), "mime": "image/jpeg"})
                        except Exception as exc:
                            _issue(rows, "Avviso", "Copertina stampa", f"Anteprima AI non creata: {str(exc)[:160]}.", "Verifica il PDF nel visualizzatore/Print Previewer.")
                    if actual_w and actual_h:
                        summary["dimensione copertina rilevata in"] = [round(actual_w, 4), round(actual_h, 4)]
                    if is_paperback and int(page_count) < 24:
                        _issue(rows, "Errore", "Copertina paperback", f"Pagine indicate: {int(page_count)}; il paperback KDP ha normalmente minimo 24 pagine.", "Controlla il numero di pagine e i requisiti aggiornati per il formato selezionato.")
                    if not is_paperback:
                        if template_data:
                            try:
                                t_w, t_h = _pdf_dimensions(template_data)
                                if expected_w is None or expected_h is None:
                                    expected_w, expected_h = t_w, t_h
                                summary["dimensione template hardcover in"] = [round(t_w, 4), round(t_h, 4)]
                                if actual_w and actual_h and (abs(actual_w - t_w) > 0.02 or abs(actual_h - t_h) > 0.02):
                                    _issue(rows, "Errore", "Copertina hardcover", f"PDF {actual_w:.3f} × {actual_h:.3f} in; template KDP {t_w:.3f} × {t_h:.3f} in.", "Adatta il PDF alle dimensioni esatte del template KDP, comprese aree di wrap e dorso.")
                            except Exception as exc:
                                _issue(rows, "Avviso", "Template hardcover", f"Template non utilizzabile: {str(exc)[:160]}.", "Scarica di nuovo il template PDF dal Cover Calculator KDP.")
                        elif not (expected_w and expected_h):
                            _issue(rows, "Avviso", "Copertina hardcover", "Dimensioni di wrap/dorso non verificate: manca il template PDF o la misura completa del Cover Calculator.", "Carica il template ufficiale KDP o inserisci larghezza e altezza finali mostrate dal calcolatore.")
                    # Il testo sul dorso non è deducibile con certezza dai soli metadati; l'AI può segnalarlo se visibile.
                    if is_paperback and int(page_count) < 79:
                        _issue(rows, "Info", "Copertina paperback", "Con meno di 79 pagine KDP non consente testo sul dorso secondo la guida paperback.", "Se sul dorso è presente testo, rimuovilo oppure verifica l'ultima soglia nella guida KDP.")

            if interior_data:
                summary["interno PDF"] = {"nome": interior.name, "pagine effettive": pdf_page_count(interior_data),
                                           "dimensione MB": round(len(interior_data) / 1024 / 1024, 2)}
                page_total = pdf_page_count(interior_data)
                if page_total is not None and is_paperback and not (24 <= page_total <= 830):
                    _issue(rows, "Avviso", "Interno PDF", f"Il PDF contiene {page_total} pagine; il limite varia per formato, carta e inchiostro (il calcolatore KDP riporta in genere 24–830 per paperback).",
                           "Verifica i limiti disponibili per il trim e l'opzione di stampa selezionati su KDP.")
                if page_total is not None and not is_paperback and not is_ebook and not (75 <= page_total <= 550):
                    _issue(rows, "Avviso", "Interno PDF", f"Il PDF contiene {page_total} pagine; la guida hardcover riporta normalmente 75–550.",
                           "Controlla i limiti aggiornati per l'edizione hardcover e il trim scelto.")
                interior_expected = (float(trim_w) + (BLEED_IN if bleed else 0),
                                     float(trim_h) + (2 * BLEED_IN if bleed else 0))
                if bleed:
                    summary["dimensione pagina attesa con bleed in"] = [round(interior_expected[0], 4), round(interior_expected[1], 4)]
                _check_pdf_images(rows, interior_data, "Illustrazioni interno", expected_size=interior_expected)
                _scan_fonts_and_annotations(rows, interior_data, "Interno PDF")
                if page_total is not None and page_total % 2:
                    _issue(rows, "Info", "Interno PDF", f"Il PDF ha un numero dispari di pagine ({page_total}).",
                           "KDP può aggiungere una pagina bianca per ottenere un conteggio pari; verifica il file elaborato nel previewer.")
                # Prendi immagini raster incorporate: l'analisi tecnica ha già controllato tutte le collocazioni.
                try:
                    with fitz.open(stream=interior_data, filetype="pdf") as doc:
                        seen_xrefs: set[int] = set()
                        for page_no, page in enumerate(doc, 1):
                            for item in page.get_images(full=True):
                                xref = int(item[0])
                                if xref <= 0 or xref in seen_xrefs:
                                    continue
                                seen_xrefs.add(xref)
                                raw = doc.extract_image(xref)
                                blob = raw.get("image", b"")
                                ext = raw.get("ext", "").lower()
                                if blob:
                                    try:
                                        vision_blob = _vision_ready(blob)
                                    except Exception:
                                        continue
                                    vision_images.append({"id": f"ILLUSTRAZIONE INTERNO · pagina {page_no} · {xref}", "data": vision_blob, "mime": "image/jpeg"})
                except Exception:
                    pass

            if art_files:
                for image_file in art_files:
                    data = image_file.getvalue()
                    summary.setdefault("illustrazioni separate", []).append({"nome": image_file.name, "dimensione MB": round(len(data) / 1024 / 1024, 2)})
                    info = _check_raster(rows, image_file.name, data, "Illustrazione separata",
                                         float(art_w) if art_w > 0 else None,
                                         float(art_h) if art_h > 0 else None)
                    if info:
                        try:
                            vision_images.append({"id": f"ILLUSTRAZIONE · {image_file.name}", "data": _vision_ready(data), "mime": "image/jpeg"})
                        except Exception:
                            pass

            if not cover_data:
                cover_label = "Copertina ebook" if is_ebook else f"Copertina {edition.lower()}"
                expected_file = "un'immagine JPG/PNG/TIFF del fronte" if is_ebook else "il PDF completo di una pagina (retro + dorso + fronte)"
                _issue(rows, "Avviso", cover_label, "Non è stata caricata la copertina finale.",
                       f"Carica {expected_file} per controllare la copertina.")
            if not is_ebook and not interior_data:
                _issue(rows, "Info", "Interno PDF", "Nessun PDF interno caricato: le illustrazioni incorporate e le dimensioni pagina non sono state controllate.",
                       "Carica il PDF interno per controllare tutte le pagine e i DPI effettivi delle immagini.")

            if ai_enabled:
                if not api_key_configured:
                    _issue(rows, "Avviso", "Controllo visivo AI", "Chiave Gemini non configurata: OCR e controllo visivo AI saltati.",
                           "Configura GOOGLE_API_KEY nella sidebar; il controllo tecnico è indipendente dall'AI.")
                else:
                    # Copertina sempre inclusa; limita le illustrazioni per contenere le chiamate/costi.
                    cover_images = [item for item in vision_images if item["id"].startswith("COPERTINA")]
                    interior_images = [item for item in vision_images if not item["id"].startswith("COPERTINA")]
                    chosen = cover_images + interior_images[:int(ai_limit)]
                    remaining = max(0, len(interior_images) - int(ai_limit))
                    if remaining:
                        _issue(rows, "Info", "Controllo visivo AI", f"L'AI esaminerà {len(chosen)} immagini; {remaining} illustrazioni restano escluse dall'AI per il limite scelto.",
                               "Aumenta il limite AI e rilancia per estendere il controllo; tutte le immagini interne sono già state controllate tecnicamente.")
                    _run_visual_ai(rows, chosen, model, ask_ai_fn, language)
            summary["controllo AI visivo"] = bool(ai_enabled and api_key_configured)
            st.session_state["kdp_report"] = {"riepilogo": summary, "controlli": rows}

    report = st.session_state.get("kdp_report")
    if report:
        st.markdown("### Report del controllo")
        _report_view(report.get("controlli", []), report.get("riepilogo", {}))

    with st.expander("Cosa non può certificare questo controllo"):
      ,,,  st.markdown(
            "- Non sostituisce il **Print Previewer KDP**, che verifica il file rispetto alla configurazione effettiva dell'account.\n"
            "- Per paperback il calcolo del dorso è stimato da pagina, carta e inchiostro; confrontalo con il Cover Calculator KDP.\n"
            "- Per hardcover serve il template KDP corretto: il wrap e le cerniere non si possono dedurre con affidabilità dal solo conteggio pagine.\n"
            "- L'AI può non riconoscere refusi o tagli; non garantisce la qualità editoriale, i diritti d'autore, il barcode o l'approvazione KDP.\n"
            "- La verifica del DPI delle immagini interne è precisa quando il PDF contiene immagini raster e le dimensioni di collocazione sono leggibili; elementi vettoriali non hanno DPI raster.",,,)
       
        
      
        
        
    

        
        
        
        
        
        
        

