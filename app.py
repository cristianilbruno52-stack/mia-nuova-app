"""Generatore di metadati editoriali/SEO per ciascun manoscritto analizzato.

Da usare nel tab SEO di LIBRO MASTER AI. I suggerimenti sono basati sul testo,
non su volumi di ricerca in tempo reale.
"""
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
