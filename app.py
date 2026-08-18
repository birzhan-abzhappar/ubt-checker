import streamlit as st
import docx
import pptx
import pdfplumber
import json
from google import genai
from google.genai import types

st.set_page_config(page_title="ҰБТ Материалдарын Тексеру Жүйесі", layout="wide", page_icon="🎓")

st.title("🎓 ҰБТ Материалдарын Автоматты Тексеру Жүйесі (Gemini AI)")
st.caption("Word (.docx), PPTX (.pptx) және PDF (.pdf) файлдарындағы қателерді AI методист арқылы табу")

# Secrets немесе қолдан енгізілген API Key
api_key = st.secrets.get("GEMINI_API_KEY", "")

if not api_key:
    with st.sidebar:
        st.header("⚙️ Баптаулар")
        api_key = st.text_input("Google Gemini API Key білдіріңіз:", type="password")

with st.sidebar:
    st.subheader("📋 Тексеру бағыттары:")
    check_ortho = st.checkbox("Орфография & Грамматика", value=True)
    check_tech = st.checkbox("Техникалық формат (A-E жауаптар, нөмірлеу)", value=True)
    check_logic = st.checkbox("Мазмұндық & Логикалық сәйкестік", value=True)

# Файлдарды оқу функциялары
def extract_text_from_docx(file):
    doc = docx.Document(file)
    text = []
    for p in doc.paragraphs:
        if p.text.strip():
            text.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_text:
                text.append(" | ".join(row_text))
    return "\n".join(text)

def extract_text_from_pptx(file):
    prs = pptx.Presentation(file)
    text = []
    for i, slide in enumerate(prs.slides):
        slide_text = [f"--- Слайд {i+1} ---"]
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text.append(shape.text)
        text.append("\n".join(slide_text))
    return "\n".join(text)

def extract_text_from_pdf(file):
    text = []
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            p_text = page.extract_text()
            if p_text:
                text.append(f"--- Бет {i+1} ---\n" + p_text)
    return "\n".join(text)

uploaded_file = st.file_uploader("Файлды таңдаңыз (.docx, .pptx, .pdf):", type=["docx", "pptx", "pdf"])

if uploaded_file and api_key:
    file_type = uploaded_file.name.split(".")[-1].lower()
    
    with st.spinner("Файлдан мәтін оқылуда..."):
        if file_type == "docx":
            content = extract_text_from_docx(uploaded_file)
        elif file_type == "pptx":
            content = extract_text_from_pptx(uploaded_file)
        elif file_type == "pdf":
            content = extract_text_from_pdf(uploaded_file)
        else:
            content = ""

    if not content.strip():
        st.error("Файлдан мәтін оқылмады немесе файл бос.")
    else:
        st.success(f"Мәтін сәтті оқылды! Жалпы ұзындығы: {len(content)} символ.")
        
        if st.button("🔍 Файлды AI Методистпен Тексеру", type="primary"):
            with st.spinner("Gemini AI қателерді талдауда... Соны күтіңіз..."):
                try:
                    client = genai.Client(api_key=api_key)
                    
                    prompt = f"""
Сіз — ҰБТ (Ұлттық Бірыңғай Тестілеу) материалдарын тексеретін тәжірибелі методистсіз.
Төменде берілген тест материалын мұқият талдап, қателерді табыңыз.

Тексеру талаптары:
1. Орфографиялық, грамматикалық және стилистикалық қателер (Қазақ тілінің нормаларына сай).
2. Техникалық қателер (Нөмірлеу реті, A-E/A-H нұсқаларының толықтығы, сұрақтардың дұрыс құрылуы).
3. Логикалық/Мазмұндық қателер (Тест сұрағындағы фактілердің дұрыстығы, жауаптардың сәйкестігі).

Тексерілетін мәтін:
\"\"\"
{content[:15000]}
\"\"\"

Талдау нәтижесін МІНДЕТТІ ТҮРДЕ мынадай JSON форматында қайтарыңыз:
{{
  "summary": "Жалпы тексеру қорытындысы мен материал сапасы туралы қысқаша пікір",
  "errors": [
    {{
      "type": "Орфография / Техникалық / Мазмұндық",
      "location": "Қате табылған орын (мысалы: 5-сұрақ немесе 2-слайд)",
      "original": "Мәтіндегі қате жазылған нұсқа",
      "correction": "Дұрыс нұсқасы",
      "explanation": "Неліктен қате екендігіне түсініктеме"
    }}
  ]
}}
"""

                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )
                    
                    result = json.loads(response.text)
                    
                    st.subheader("📊 Тексеру Қорытындысы")
                    st.info(result.get("summary", "Тексеру аяқталды."))
                    
                    errors = result.get("errors", [])
                    if not errors:
                        st.balloons()
                        st.success("🎉 Тамаша! Файлдан ешқандай қате табылмады.")
                    else:
                        st.warning(f"Жалпы табылған қателер саны: {len(errors)}")
                        
                        for err in errors:
                            with st.expander(f"❌ [{err.get('type', 'Қате')}] {err.get('location', 'Мәтін ішінде')}"):
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.error(f"**Қате нұсқасы:**\n{err.get('original')}")
                                with col2:
                                    st.success(f"**Дұрыс нұсқасы:**\n{err.get('correction')}")
                                st.write(f"**Түсініктеме:** {err.get('explanation')}")

                except Exception as e:
                    st.error(f"Тексеру кезінде қате шықты: {e}")
elif uploaded_file and not api_key:
    st.info("💡 Тексеруді бастау үшін сол жақтағы менюге Google Gemini API кілтін енгізіңіз немесе Secrets бөлімінде сақтаңыз.")
