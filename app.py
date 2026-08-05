import streamlit as st
from gigachat import GigaChat
import os
import io
import re

# Библиотеки для чтения документов
from docx import Document
import PyPDF2
from pptx import Presentation
import fitz  # PyMuPDF
from PIL import Image
import pytesseract

st.set_page_config(page_title="ИИ-Анализатор договоров", page_icon="📄", layout="wide")
st.title("🤖 ИИ-Анализатор договоров")
st.markdown("Автоматический анализ договоров с помощью искусственного интеллекта")
st.markdown("📄 **Загрузите договоры** → 🤖 **ИИ найдёт риски** → 📊 **Получите отчёты**")

# --- Извлечение текста из файлов ---
def extract_text_from_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join([p.text for p in doc.paragraphs])

def extract_text_from_pdf(file_bytes):
    text = ""
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            text += page.get_text() + "\n"
        if len(text.strip()) > 20:
            return text
    except:
        pass
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            text += page.extract_text() + "\n"
    except:
        pass
    return text

def extract_text_from_pptx(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

def extract_text(file):
    file_bytes = file.read()
    filename = file.name.lower()
    if filename.endswith('.docx'):
        return extract_text_from_docx(file_bytes)
    elif filename.endswith('.pdf'):
        return extract_text_from_pdf(file_bytes)
    elif filename.endswith('.pptx'):
        return extract_text_from_pptx(file_bytes)
    elif filename.endswith(('.png', '.jpg', '.jpeg', '.tiff')):
        try:
            img = Image.open(io.BytesIO(file_bytes))
            return pytesseract.image_to_string(img, lang='rus+eng')
        except:
            return ""
    else:
        try:
            return file_bytes.decode('utf-8', errors='ignore')
        except:
            return ""

# --- Анализ через GigaChat ---
def analyze_contract(text):
    credentials = st.secrets.get("GIGACHAT_CREDENTIALS", os.getenv("GIGACHAT_CREDENTIALS", ""))
    if not credentials:
        st.error("⚠️ Не найден токен GIGACHAT_CREDENTIALS в Settings → Secrets!")
        return None

    prompt = f"""Проанализируй текст договора и верни данные СТРОГО в формате:
НОМЕР: [номер договора]
ДАТА: [дата заключения]
ТИП ДОГОВОРА: [тип]
СУБЪЕКТНЫЙ СОСТАВ: [стороны]
СУММА: [сумма]
ИНН: [ИНН]
ПЕНИ: [пени/штрафы]
РИСКИ: [основные юридические риски]

ТЕКСТ ДОГОВОРА:
{text[:15000]}"""

    try:
        giga = GigaChat(
            credentials=credentials,
            scope="GIGACHAT_API_PERS",
            model="GigaChat",
            verify_ssl_certs=False
        )
        response = giga.chat(prompt)
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"❌ Ошибка анализа: {e}")
        return None

# --- Разбор ответа ИИ ---
def parse_response(ai_text):
    data = {"Номер": "не указан", "Дата": "не указана", "Тип договора": "не определён",
            "Субъектный состав": "не определён", "Сумма": "не указана", "ИНН": "0",
            "Пени": "не указаны", "Риски": "не найдены"}
    if not ai_text:
        return data
    for line in ai_text.split('\n'):
        line = line.strip()
        if ":" in line:
            key, val = line.split(":", 1)
            key = key.strip().upper()
            val = val.strip()
            if "НОМЕР" in key: data["Номер"] = val
            elif "ДАТА" in key: data["Дата"] = val
            elif "ТИП" in key: data["Тип договора"] = val
            elif "СУБЪЕКТ" in key or "СТОРОНЫ" in key: data["Субъектный состав"] = val
            elif "СУММА" in key or "ЦЕНА" in key: data["Сумма"] = val
            elif "ИНН" in key: data["ИНН"] = val
            elif "ПЕНИ" in key or "ШТРАФ" in key: data["Пени"] = val
            elif "РИСКИ" in key: data["Риски"] = val
    return data

# --- Основной интерфейс ---
st.markdown("---")
st.subheader("📤 Загрузка договоров")
uploaded_files = st.file_uploader(
    "Перетащите файлы или нажмите для выбора",
    type=['docx', 'pdf', 'pptx', 'txt', 'png', 'jpg', 'jpeg'],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ Загружено файлов: {len(uploaded_files)}")
    st.markdown(f"### 📂 Список файлов ({len(uploaded_files)})")

    for file in uploaded_files:
        st.markdown(f"#### 📄 {file.name}")
        text = extract_text(file)

        if not text.strip():
            st.warning("⚠️ Не удалось извлечь текст из файла.")
            continue

        with st.spinner("🤖 ИИ анализирует документ..."):
            ai_response = analyze_contract(text)

        if ai_response:
            parsed = parse_response(ai_response)
            st.success(f"✅ Обработано: {file.name}")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Номер", parsed["Номер"])
                st.metric("Дата", parsed["Дата"])
            with col2:
                st.metric("Сумма", parsed["Сумма"])
                st.metric("ИНН", parsed["ИНН"])
            with col3:
                st.metric("Пени", parsed["Пени"])

            st.markdown(f"📋 **Тип договора:** {parsed['Тип договора']}")
            st.markdown(f"👥 🏢 **Субъектный состав:** {parsed['Субъектный состав']}")

            with st.expander("🤖 ИИ-анализ и Риски"):
                st.info(parsed["Риски"])

            st.download_button(
                label="📥 Скачать отчёт (TXT)",
                data=f"Отчёт по {file.name}\n\n{ai_response}".encode('utf-8'),
                file_name=f"Report_{file.name}.txt",
                mime="text/plain"
            )
            st.markdown("---")
