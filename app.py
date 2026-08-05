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

st.set_page_config(page_title="ИИ-Анализатор договоров", page_icon="🤖", layout="wide")

# ============================================================
# НАСТРОЙКИ ТОКЕНА И ВЫБОР МОДЕЛИ
# ============================================================
GIGACHAT_TOKEN_HARDCODE = ""  # ← можно вставить токен сюда, если лень вводить каждый раз

st.title("🤖 ИИ-Анализатор договоров")
st.markdown("Автоматический анализ договоров с помощью искусственного интеллекта")
st.markdown("📄 **Загрузите договоры** → 🤖 **ИИ найдёт риски** → 📊 **Получите отчёты**")

with st.sidebar:
    st.header("⚙️ Настройки GigaChat")
    sidebar_token = st.text_input(
        "Токен (client_id:client_secret)",
        type="password",
        help="Получить токен: developers.sber.ru → Studio → ваш проект → Credentials"
    )
    if sidebar_token.strip():
        st.session_state["giga_token"] = sidebar_token.strip()
        st.success("✅ Токен сохранён")
    elif st.session_state.get("giga_token"):
        st.success("✅ Токен активен")
        
    st.markdown("---")
    
    # ИСПРАВЛЕНИЕ ОШИБКИ 404: Выбор модели из списка
    model_name = st.selectbox(
        "Модель нейросети",
        ["GigaChat-Max", "GigaChat-Pro", "GigaChat-Plus", "GigaChat", "GigaChat-Lite"],
        index=0,
        help="Если при анализе вылетает ошибка 404 'No such model', просто выберите другую модель из списка. Чаще всего работает GigaChat-Max."
    )
    st.session_state["giga_model"] = model_name


def get_credentials():
    """Ищем токен GigaChat во всех возможных местах"""
    if GIGACHAT_TOKEN_HARDCODE:
        return GIGACHAT_TOKEN_HARDCODE.strip()
    if st.session_state.get("giga_token"):
        return st.session_state["giga_token"]
    try:
        for key in ("GIGACHAT_CREDENTIALS", "GIGACHAT_TOKEN", "GIGACHAT_ACCESS_TOKEN", "GIGACHAT_API_KEY"):
            if key in st.secrets:
                return str(st.secrets[key]).strip()
    except Exception:
        pass
    for var in ("GIGACHAT_CREDENTIALS", "GIGACHAT_TOKEN", "GIGACHAT_ACCESS_TOKEN", "GIGACHAT_API_KEY"):
        val = os.getenv(var)
        if val:
            return val.strip()
    return None


# ============================================================
# ИЗВЛЕЧЕНИЕ ТЕКСТА ИЗ ФАЙЛОВ
# ============================================================
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
    except Exception:
        pass
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            text += page.extract_text() + "\n"
    except Exception:
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
        except Exception:
            return ""
    else:
        try:
            return file_bytes.decode('utf-8', errors='ignore')
        except Exception:
            return ""


# ============================================================
# ИЗВЛЕЧЕНИЕ НОМЕРА И ДАТЫ
# ============================================================
def extract_number(text):
    head = text[:3000]
    m = re.search(r'[№N]\s*(\d+[а-яА-Я/\-]*)', head)
    if m:
        return m.group(1)
    m = re.search(r'(?:Договор|Контракт|Соглашение)\s+[№N]?\s*(\d+)', head, re.IGNORECASE)
    if m:
        return m.group(1)
    return "не указан"

def extract_date(text):
    head = text[:3000]
    months = r'января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря'
    m = re.search(rf'[«"]?(\d{{1,2}})[»"]?\s+({months})\s+(\d{{4}})', head, re.IGNORECASE)
    if m:
        return f"{m.group(1)} {m.group(2)} {m.group(3)} г"
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', head)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    return "не указана"


# ============================================================
# ИИ-АНАЛИЗ
# ============================================================
def analyze_contract(text):
    credentials = get_credentials()
    if not credentials:
        raise ValueError(
            "Не найден токен GigaChat. Вставьте его в боковую панель (⚙️) "
            "или добавьте GIGACHAT_CREDENTIALS в Settings → Secrets."
        )

    # Берём модель, которую выбрал пользователь в сайдбаре
    model_to_use = st.session_state.get("giga_model", "GigaChat-Max")

    giga = GigaChat(
        credentials=credentials,
        scope="GIGACHAT_API_PERS",
        model=model_to_use,       # <--- ЗДЕСЬ ИСПРАВЛЕНА ОШИБКА 404
        verify_ssl_certs=False
    )

    prompt = f"""Проанализируй текст договора и верни данные СТРОГО в формате:
ТИП ДОГОВОРА: [тип]
СУБЪЕКТНЫЙ СОСТАВ: [стороны]
СУММА: [сумма]
ИНН: [ИНН]
ПЕНИ: [пени]
РИСКИ: [основные риски]

ТЕКСТ ДОГОВОРА:
{text[:15000]}"""

    response = giga.chat(prompt)
    return response.choices[0].message.content


def empty_parsed():
    return {
        "Тип договора": "не определён",
        "Субъектный состав": "не определён",
        "Сумма": "не указана",
        "ИНН": "0",
        "Пени": "не указаны",
        "Риски": "не найдены"
    }

def parse_response(ai_text):
    data = empty_parsed()
    if not ai_text:
        return data
    for line in ai_text.split("\n"):
        line = line.strip().lstrip("-•* ").replace("**", "")
        if ":" not in line:
            continue
        key, val = line.split(":", 1)
        key = key.strip().upper()
        val = val.strip()
        if not val:
            continue
        if "ТИП" in key:
            data["Тип договора"] = val
        elif "СУБЪЕКТ" in key or "СОСТАВ" in key or "СТОРОН" in key:
            data["Субъектный состав"] = val
        elif "СУММА" in key or "ЦЕНА" in key or "СТОИМОСТЬ" in key:
            data["Сумма"] = val
        elif "ИНН" in key:
            data["ИНН"] = val
        elif "ПЕН" in key or "ШТРАФ" in key or "НЕУСТОЙК" in key:
            data["Пени"] = val
        elif "РИСК" in key:
            data["Риски"] = val
    return data


# ============================================================
# ОСНОВНОЙ ИНТЕРФЕЙС
# ============================================================
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

    results = []
    processed = 0

    for file in uploaded_files:
        st.markdown(f"📄 **{file.name}**")
        text = extract_text(file)

        if not text.strip():
            st.warning("⚠️ Не удалось извлечь текст из файла.")
            continue

        with st.spinner(f"🤖 ИИ ({st.session_state.get('giga_model', 'GigaChat-Max')}) анализирует документ..."):
            ai_response = None
            ai_error = None
            try:
                ai_response = analyze_contract(text)
            except Exception as e:
                ai_error = str(e)

        parsed = parse_response(ai_response)
        results.append((file.name, parsed, ai_response, ai_error, text))
        processed += 1

    if processed:
        st.success(f"✅ Обработано: {processed} договоров")
        st.balloons()
        st.success(f"🎉 Готово! Проанализировано: {processed} договоров")

        st.markdown("---")
        st.subheader("📊 Результаты анализа")

        for fname, parsed, ai_response, ai_error, raw_text in results:
            st.markdown(f"#### 📄 {fname}")

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Номер", extract_number(raw_text))
            c2.metric("Дата", extract_date(raw_text))
            c3.metric("Сумма", parsed["Сумма"])
            c4.metric("ИНН", parsed["ИНН"])
            c5.metric("Пени", parsed["Пени"])

            st.markdown(f"📋 **Тип договора:** {parsed['Тип договора']}")
            st.markdown(f"👥 🏢 **Субъектный состав:** {parsed['Субъектный состав']}")

            st.markdown("**🤖 ИИ-анализ:**")
            if ai_response:
                st.info(ai_response)
                if parsed["Риски"] != "не найдены":
                    with st.expander("⚠️ Юридические риски"):
                        st.write(parsed["Риски"])
            else:
                st.text(f"ТИП ДОГОВОРА: {parsed['Тип договора']}  СУБЪЕКТНЫЙ СОСТАВ: {parsed['Субъектный состав']}")

            if ai_error:
                st.error(f"❌ Ошибка анализа: {ai_error}")
                st.warning("💡 Подсказка: Попробуйте выбрать другую модель в боковой панели слева (например, GigaChat-Pro или GigaChat-Plus).")

            st.download_button(
                label="📥 Скачать отчёт (TXT)",
                data=f"Отчёт по {fname}\n\n{ai_response or 'Анализ не выполнен'}".encode('utf-8'),
                file_name=f"Report_{fname}.txt",
                mime="text/plain"
            )
            st.markdown("---")
