import io
import os
import re
import tempfile

import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE

st.set_page_config(page_title="Traducător Documente", page_icon="📄", layout="centered")

st.title("📄 Traducător Documente")
st.write(
    "Încarcă un document (PDF scanat, imagine, Word sau text) și primești "
    "traducerea completă, ca document Word. Gemini detectează automat limba "
    "originală și traduce, cu OCR pentru documentele scanate."
)

LANGUAGES = [
    "Română", "Engleză", "Franceză", "Germană", "Italiană", "Spaniolă",
    "Portugheză", "Suedeză", "Norvegiană", "Daneză", "Olandeză", "Polonă",
    "Cehă", "Slovacă", "Maghiară", "Bulgară", "Greacă", "Rusă", "Ucraineană",
    "Turcă", "Arabă", "Chineză", "Japoneză", "Coreeană",
]

# Extensii acceptate și cum sunt tratate:
# - trimise ca fișier către Gemini (necesită OCR): pdf, imagini
# - text extras local, apoi trimis ca text (fără OCR): docx, txt
FILE_TYPES = ["pdf", "png", "jpg", "jpeg", "docx", "txt"]
OCR_EXTENSIONS = {"pdf", "png", "jpg", "jpeg"}


def get_secret_keys():
    """Ia lista de chei API din st.secrets, dacă există."""
    if "GEMINI_API_KEYS" in st.secrets:
        return [k for k in st.secrets["GEMINI_API_KEYS"] if k]
    if "GEMINI_API_KEY" in st.secrets:
        return [st.secrets["GEMINI_API_KEY"]]
    return []


secret_keys = get_secret_keys()

with st.sidebar:
    st.header("Configurare")

    with st.expander("❓ Cum obțin o cheie? (gratuit)"):
        st.markdown(
            """
**Ai nevoie de un cont Google** (Gmail). E complet gratuit.

**Pasul 1** — Deschide Google AI Studio:
"""
        )
        st.link_button("🔵 Mergi la aistudio.google.com", "https://aistudio.google.com/apikey")
        st.markdown(
            """
**Pasul 2** — Autentifică-te cu contul Google.

**Pasul 3** — Apasă „Create API key" (dreapta sus).

**Pasul 4** — Dacă ți se cere, alege „Create API key in new project".

**Pasul 5** — Copiază cheia afișată.
- Poate arăta ca `AIzaSy...` sau, în formatul mai nou, ca `AQ.Ab8R...`
- Apasă iconița 📋 de lângă cheie

**Pasul 6** — Lipește cheia mai jos (sau în Secrets, dacă publici aplicația) și apasă Salvează.

---
💡 **Limită gratuită:** 15 cereri/minut, 1 milion tokeni/zi — suficient pentru volum moderat de documente.

Poți repeta pașii de mai sus cu conturi Google diferite ca să obții **mai multe chei** și să crești numărul de documente pe care le poți traduce zilnic.
"""
        )

    if secret_keys:
        st.success(f"{len(secret_keys)} cheie/chei API încărcate din Secrets.")
        manual_keys_input = ""
    else:
        st.info("Nicio cheie găsită în Secrets — introdu una sau mai multe mai jos (separate prin virgulă).")
        manual_keys_input = st.text_input("Gemini API Key(s)", type="password")

    model_name = st.selectbox(
        "Model Gemini",
        ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-3.1-flash-lite", "gemini-3.5-flash-lite"],
        index=0,
        help=(
            "Pe planul gratuit: gemini-3.5-flash și gemini-2.5-flash au limită de 20 cereri/zi "
            "per cheie API, dar calitate mai bună. Variantele „flash-lite” permit 500 cereri/zi, "
            "dar sunt mai puțin precise pe documente complexe (contracte). "
            "gemini-2.5-flash va fi retras de Google pe 16 octombrie 2026."
        ),
    )

api_keys = secret_keys or [k.strip() for k in manual_keys_input.split(",") if k.strip()]

target_language = st.selectbox("Tradu în limba", LANGUAGES, index=0)

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

uploaded_file = st.file_uploader(
    "Alege fișierul (PDF, imagine, Word sau text)",
    type=FILE_TYPES,
    key=f"uploader_{st.session_state['uploader_key']}",
)


STRUCTURE_RULES = """PĂSTREAZĂ STRUCTURA EXACTĂ a originalului:
   - Fiecare paragraf din original trebuie să rămână un paragraf separat în traducere — nu uni, nu împărți, nu omite niciun paragraf.
   - Păstrează exact numerotarea articolelor/clauzelor (ex: "Art. 1", "1.1", "1.2", "(a)", "(b)") așa cum apare în original.
   - Păstrează titlurile de secțiuni, listele, liniile goale dintre paragrafe și ordinea exactă a conținutului.
   - Tabelele: redă-le rând cu rând, cu celulele separate prin " | ", păstrând numărul de coloane.
   - Semnături, ștampile, date, numere de referință: transcrie-le exact așa cum apar (nu traduce numele proprii, denumirile de companii sau numerele de înregistrare)."""

NO_EXTRA_CONTENT_RULE = """NU ADĂUGA ABSOLUT NIMIC ÎN PLUS față de conținutul tradus, cu SINGURA excepție a liniei
"###TIP: ..." cerute explicit la punctul următor. În afară de acea linie, este interzis strict:
   - nicio linie de tipul "Limba originală: ...", "Limba sursă: ...", "Detected language: ..." sau echivalent
   - niciun titlu, etichetă sau notă introductivă adăugată de tine ("Traducere:", "Document tradus", etc.)
   - niciun comentariu, explicație, rezumat sau observație despre document sau despre traducere
   - niciun text în altă limbă decât {target_language} (cu excepția numelor proprii/denumirilor care rămân netraduse, conform regulilor de mai sus)
   Ieșirea trebuie să conțină DOAR linia "###TIP: ..." urmată de paragrafele traduse ale documentului, nimic altceva."""

DOCUMENT_TYPE_RULE = """La începutul răspunsului, înainte de orice altceva, clasifică documentul printr-o
singură linie exact în acest format (fără explicații suplimentare):
###TIP: formular sau ###TIP: continut
- Folosește "formular" dacă documentul conține câmpuri/casete goale menite să fie completate
  de mână (semnături, date, nume, casete de bifat etc.) — adică e un document necompletat sau
  parțial completat, ca un formular, o cerere, un act de identitate cu casete goale.
- Folosește "continut" dacă documentul e integral completat/redactat — un contract, un act,
  un tabel cu date reale (prețuri, cifre, liste), fără câmpuri goale de completat manual.

În plus, în interiorul tabelelor (rânduri cu " | "), dacă o celulă corespunde unui câmp gol,
menit să fie completat de mână (fără conținut real în original, doar eticheta câmpului, cu
spațiu liber dedesubt în original), marcheaz-o adăugând " [GOL]" la finalul textului celulei.
Nu marca astfel celulele care conțin deja date reale completate."""

OUTPUT_FORMAT = """Format de răspuns OBLIGATORIU (respectă-l strict, câte un paragraf tradus pe fiecare linie nouă,
exact în ordinea din original):
###TIP: formular sau ###TIP: continut
--- Pagina 1 ---
<paragraf 1>
<paragraf 2>
...

--- Pagina 2 ---
<paragraf 1>
...

... continuă pentru fiecare pagină din document, în ordine."""


def build_prompt_ocr(target_language: str) -> str:
    return f"""Ești un traducător profesionist, specializat în documente juridice
și contracte comerciale. Acest document este probabil un contract sau un act oficial,
așa că fidelitatea structurală este esențială — poate fi folosit ca referință legală.

Sarcina ta:
1. Detectează automat limba originală a documentului.
2. Citește (OCR) tot textul din acest document, pagină cu pagină (dacă e o singură imagine, tratează-o ca pagina 1).
3. Tradu fiecare pagină integral în limba {target_language}, păstrând sensul exact și tonul oficial/juridic.
4. {STRUCTURE_RULES}
5. Nu rezuma, nu parafraza liber, nu adăuga comentarii sau explicații proprii — este o traducere fidelă, nu un rezumat.
6. Dacă un cuvânt sau nume propriu nu poate fi tradus, lasă-l în original.
7. {NO_EXTRA_CONTENT_RULE.format(target_language=target_language)}
8. {DOCUMENT_TYPE_RULE}

{OUTPUT_FORMAT}"""


def build_prompt_text(target_language: str) -> str:
    return f"""Ești un traducător profesionist, specializat în documente juridice
și contracte comerciale. Textul de mai jos a fost deja extras dintr-un document oficial
(Word sau text), fără nevoie de OCR. Fidelitatea structurală este esențială — poate fi
folosit ca referință legală.

Sarcina ta:
1. Detectează automat limba originală a textului.
2. Tradu textul integral în limba {target_language}, păstrând sensul exact și tonul oficial/juridic.
3. {STRUCTURE_RULES}
4. Nu rezuma, nu parafraza liber, nu adăuga comentarii sau explicații proprii — este o traducere fidelă, nu un rezumat.
5. Dacă un cuvânt sau nume propriu nu poate fi tradus, lasă-l în original.
6. Tratează tot textul ca fiind pe o singură "pagină" logică, decât dacă vezi marcaje clare de pagină nouă în text.
7. {NO_EXTRA_CONTENT_RULE.format(target_language=target_language)}
8. {DOCUMENT_TYPE_RULE}

{OUTPUT_FORMAT}

TEXT ORIGINAL:
\"\"\"
{{document_text}}
\"\"\""""


def extract_docx_text(file_bytes: bytes) -> str:
    """Extrage textul (paragrafe, în ordine) dintr-un fișier .docx."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
        lines = [p.text for p in doc.paragraphs]
        return "\n".join(lines)
    finally:
        os.unlink(tmp_path)


def call_gemini_with_fallback(uploaded_file, keys: list, model_name: str, target_language: str) -> str:
    """Încearcă fiecare cheie API pe rând. Dacă una eșuează (cotă depășită, cheie
    invalidă etc.), trece automat la următoarea. Alege fluxul (fișier + OCR, sau
    text extras local) în funcție de extensia fișierului."""
    ext = os.path.splitext(uploaded_file.name)[1].lower().lstrip(".")
    file_bytes = uploaded_file.getvalue()
    last_error = None

    if ext in OCR_EXTENSIONS:
        prompt = build_prompt_ocr(target_language)
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        try:
            for key in keys:
                try:
                    client = genai.Client(api_key=key)
                    gemini_file = client.files.upload(file=tmp_path)
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[gemini_file, prompt],
                    )
                    return response.text
                except Exception as e:
                    last_error = e
                    continue
            raise RuntimeError(f"Toate cheile API au eșuat. Ultima eroare: {last_error}")
        finally:
            os.unlink(tmp_path)

    else:
        if ext == "docx":
            document_text = extract_docx_text(file_bytes)
        else:  # txt
            document_text = file_bytes.decode("utf-8", errors="replace")

        prompt = build_prompt_text(target_language).replace("{document_text}", document_text)

        for key in keys:
            try:
                client = genai.Client(api_key=key)
                response = client.models.generate_content(
                    model=model_name,
                    contents=[prompt],
                )
                return response.text
            except Exception as e:
                last_error = e
                continue
        raise RuntimeError(f"Toate cheile API au eșuat. Ultima eroare: {last_error}")


def build_docx(translated_text: str, document_type: str = "formular") -> io.BytesIO:
    """Construiește un document Word, cu fiecare pagină tradusă ca secțiune separată.
    Celulele de tabel marcate cu "[GOL]" (câmpuri de completat manual) primesc spațiu
    liber și rând mai înalt; celelalte celule rămân compacte."""
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    parts = re.split(r"-{2,}\s*Pagina\s+(\d+)\s*-{2,}", translated_text)
    gol_marker = re.compile(r"\s*\[GOL\]\s*$", re.IGNORECASE)

    def add_table(table_lines):
        raw_rows = [[cell.strip() for cell in re.split(r"\s*\|\s*", ln.strip())] for ln in table_lines]
        num_cols = max(len(r) for r in raw_rows)
        any_marker_in_table = any(gol_marker.search(c) for r in raw_rows for c in r)
        fallback_all_blank = document_type == "formular" and not any_marker_in_table

        table = doc.add_table(rows=len(raw_rows), cols=num_cols)
        table.style = "Table Grid"
        for r_idx, row_cells in enumerate(raw_rows):
            row_has_blank_field = False
            for c_idx in range(num_cols):
                raw_text = row_cells[c_idx] if c_idx < len(row_cells) else ""
                is_blank_field = bool(gol_marker.search(raw_text)) or fallback_all_blank
                text = gol_marker.sub("", raw_text).strip()
                cell = table.cell(r_idx, c_idx)
                cell.text = text
                if is_blank_field:
                    cell.add_paragraph("")  # spațiu liber pentru completare manuală
                    row_has_blank_field = True
            if row_has_blank_field:
                table.rows[r_idx].height = Cm(1.4)
                table.rows[r_idx].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        doc.add_paragraph("")  # spațiu după tabel

    def add_body_lines(text_block: str):
        lines = text_block.split("\n")
        pending_table = []
        for idx, line in enumerate(lines):
            is_table_row = "|" in line and line.strip()
            if is_table_row:
                pending_table.append(line)
                continue
            if pending_table:
                add_table(pending_table)
                pending_table = []
            if line.strip():
                doc.add_paragraph(line.rstrip())
            elif idx not in (0, len(lines) - 1):
                doc.add_paragraph("")
        if pending_table:
            add_table(pending_table)

    if len(parts) > 1:
        preamble = parts[0].strip("\n")
        if preamble.strip():
            add_body_lines(preamble)
        for i in range(1, len(parts), 2):
            page_text = parts[i + 1] if i + 1 < len(parts) else ""
            add_body_lines(page_text.strip("\n"))
            if i + 2 < len(parts):
                doc.add_page_break()
    else:
        add_body_lines(translated_text)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def clean_translated_text(text: str) -> str:
    """Plasă de siguranță: elimină liniile meta pe care modelul le-ar putea adăuga
    din greșeală (ex: "Limba originală: ..."), în ciuda instrucțiunilor din prompt."""
    meta_pattern = re.compile(
        r"^\s*(limba\s+(originală|sursă)|detected\s+language|source\s+language|"
        r"traducere\s*:|document\s+tradus)\b",
        re.IGNORECASE,
    )
    lines = text.split("\n")
    cleaned = [ln for ln in lines if not meta_pattern.match(ln)]
    return "\n".join(cleaned)


def extract_document_type(text: str) -> tuple:
    """Extrage marcajul '###TIP: formular|continut' de la începutul răspunsului,
    dacă există, și returnează (tip_document, text_fără_marcaj)."""
    match = re.match(r"^\s*###\s*TIP\s*:\s*(formular|continut)\s*\n?", text, re.IGNORECASE)
    if match:
        doc_type = match.group(1).lower()
        return doc_type, text[match.end():]
    return "formular", text  # implicit: comportament conservator (spațiu de completat)


if uploaded_file and not api_keys:
    st.warning("Adaugă cel puțin o cheie API Gemini (în Secrets sau manual) pentru a continua.")

if uploaded_file and api_keys and st.button("🔄 Tradu documentul", type="primary"):
    with st.spinner("Se procesează documentul... poate dura câteva minute pentru documente mari."):
        try:
            translated_text = call_gemini_with_fallback(uploaded_file, api_keys, model_name, target_language)
            translated_text = clean_translated_text(translated_text)
            document_type, translated_text = extract_document_type(translated_text)
            st.session_state["translated_text"] = translated_text
            st.session_state["document_type"] = document_type
            st.session_state["source_name"] = uploaded_file.name
            st.session_state["target_language"] = target_language
        except Exception as e:
            st.error(f"A apărut o eroare: {e}")

if "translated_text" in st.session_state:
    st.subheader("✅ Rezultat")
    st.text_area("Previzualizare text tradus", st.session_state["translated_text"], height=400)

    docx_buffer = build_docx(
        st.session_state["translated_text"],
        st.session_state.get("document_type", "formular"),
    )
    base_name = os.path.splitext(st.session_state.get("source_name", "document"))[0]
    lang_suffix = st.session_state.get("target_language", target_language)[:2].lower()

    col1, col2 = st.columns([3, 1])
    with col1:
        st.download_button(
            "⬇️ Descarcă traducerea (.docx)",
            data=docx_buffer,
            file_name=f"{base_name}_{lang_suffix}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    with col2:
        if st.button("🔄 Document nou"):
            for key in ("translated_text", "source_name", "target_language", "document_type"):
                st.session_state.pop(key, None)
            st.session_state["uploader_key"] += 1
            st.rerun()
