import io
import os
import re
import tempfile

import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Cm

st.set_page_config(page_title="Traducător Documente", page_icon="📄", layout="centered")

st.title("📄 Traducător Documente")
st.write(
    "Încarcă un document PDF scanat (contract, act, formular etc.) și primești "
    "traducerea completă, ca document Word. Gemini detectează automat limba "
    "originală, face OCR și traduce într-un singur pas."
)

LANGUAGES = [
    "Română", "Engleză", "Franceză", "Germană", "Italiană", "Spaniolă",
    "Portugheză", "Suedeză", "Norvegiană", "Daneză", "Olandeză", "Polonă",
    "Cehă", "Slovacă", "Maghiară", "Bulgară", "Greacă", "Rusă", "Ucraineană",
    "Turcă", "Arabă", "Chineză", "Japoneză", "Coreeană",
]


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
    if secret_keys:
        st.success(f"{len(secret_keys)} cheie/chei API încărcate din Secrets.")
        manual_keys_input = ""
    else:
        st.info("Nicio cheie găsită în Secrets — introdu una sau mai multe mai jos (separate prin virgulă).")
        manual_keys_input = st.text_input("Gemini API Key(s)", type="password")

    model_name = st.selectbox(
        "Model Gemini",
        ["gemini-2.5-flash", "gemini-2.5-pro"],
        index=0,
        help="Flash e mai rapid și mai ieftin; Pro poate fi mai precis pe documente complicate.",
    )

api_keys = secret_keys or [k.strip() for k in manual_keys_input.split(",") if k.strip()]

target_language = st.selectbox("Tradu în limba", LANGUAGES, index=0)

uploaded_file = st.file_uploader("Alege fișierul PDF", type=["pdf"])

def build_prompt(target_language: str) -> str:
    return f"""Ești un traducător profesionist, specializat în documente juridice
și contracte comerciale. Acest document este probabil un contract sau un act oficial,
așa că fidelitatea structurală este esențială — poate fi folosit ca referință legală.

Sarcina ta:
1. Detectează automat limba originală a documentului.
2. Citește (OCR) tot textul din acest document PDF scanat, pagină cu pagină.
3. Tradu fiecare pagină integral în limba {target_language}, păstrând sensul exact și tonul oficial/juridic.
4. PĂSTREAZĂ STRUCTURA EXACTĂ a originalului:
   - Fiecare paragraf din original trebuie să rămână un paragraf separat în traducere — nu uni, nu împărți, nu omite niciun paragraf.
   - Păstrează exact numerotarea articolelor/clauzelor (ex: "Art. 1", "1.1", "1.2", "(a)", "(b)") așa cum apare în original.
   - Păstrează titlurile de secțiuni, listele, liniile goale dintre paragrafe și ordinea exactă a conținutului.
   - Tabelele: redă-le rând cu rând, cu celulele separate prin " | ", păstrând numărul de coloane.
   - Semnături, ștampile, date, numere de referință: transcrie-le exact așa cum apar (nu traduce numele proprii, denumirile de companii sau numerele de înregistrare).
5. Nu rezuma, nu parafraza liber, nu adăuga comentarii sau explicații proprii — este o traducere fidelă, nu un rezumat.
6. Dacă un cuvânt sau nume propriu nu poate fi tradus, lasă-l în original.

Format de răspuns OBLIGATORIU (respectă-l strict, câte un paragraf tradus pe fiecare linie nouă,
exact în ordinea din original):
--- Pagina 1 ---
<paragraf 1>
<paragraf 2>
...

--- Pagina 2 ---
<paragraf 1>
...

... continuă pentru fiecare pagină din document, în ordine."""


def call_gemini_with_fallback(file_bytes: bytes, keys: list, model_name: str, target_language: str) -> str:
    """Încearcă fiecare cheie API pe rând. Dacă una eșuează (cotă depășită, cheie
    invalidă etc.), trece automat la următoarea."""
    last_error = None
    prompt = build_prompt(target_language)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
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


def build_docx(translated_text: str) -> io.BytesIO:
    """Construiește un document Word, cu fiecare pagină tradusă ca secțiune separată."""
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

    def add_body_lines(text_block: str, first_page: bool = False):
        lines = text_block.split("\n")
        for idx, line in enumerate(lines):
            if line.strip():
                doc.add_paragraph(line.rstrip())
            elif idx not in (0, len(lines) - 1):
                # linie goală în mijlocul textului = spațiere intenționată în original
                doc.add_paragraph("")

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


if uploaded_file and not api_keys:
    st.warning("Adaugă cel puțin o cheie API Gemini (în Secrets sau manual) pentru a continua.")

if uploaded_file and api_keys and st.button("🔄 Tradu documentul", type="primary"):
    with st.spinner("Se procesează documentul (OCR + traducere)... poate dura câteva minute pentru documente mari."):
        try:
            translated_text = call_gemini_with_fallback(uploaded_file.read(), api_keys, model_name, target_language)
            st.session_state["translated_text"] = translated_text
            st.session_state["source_name"] = uploaded_file.name
            st.session_state["target_language"] = target_language
        except Exception as e:
            st.error(f"A apărut o eroare: {e}")

if "translated_text" in st.session_state:
    st.subheader("✅ Rezultat")
    st.text_area("Previzualizare text tradus", st.session_state["translated_text"], height=400)

    docx_buffer = build_docx(st.session_state["translated_text"])
    base_name = os.path.splitext(st.session_state.get("source_name", "document"))[0]
    lang_suffix = st.session_state.get("target_language", target_language)[:2].lower()
    st.download_button(
        "⬇️ Descarcă traducerea (.docx)",
        data=docx_buffer,
        file_name=f"{base_name}_{lang_suffix}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
