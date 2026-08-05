import streamlit as st
from gigachat import GigaChat
import os
import io
import re

from docx import Document
import PyPDF2
from pptx import Presentation
import fitz
from PIL import Image
import pytesseract

st.set_page_config(page_title="ИИ-Анализатор договоров", page_icon="🤖", layout="wide")

st.title("🤖 ИИ-Анализатор договоров")
st.markdown("Автоматический анализ договоров с помощью искусственного интеллекта")
st.markdown("📄 **Загрузите договоры** → 🤖 **ИИ найдёт риски** → 📊 **Получите отчёты**")

def get_credentials():
    if st.session_state.get("giga_token"): return st.session_state["giga_token"]
    try:
        for key in ("GIGACHAT_CREDENTIALS", "GIGACHAT_TOKEN", "GIGACHAT_ACCESS_TOKEN"):
            if key in st.secrets: return str(st.secrets[key]).strip()
    except Exception: pass
    return None

def parse_model_id(m):
    """Безопасно достаем чистый ID модели из объекта Сбера (учитываем id_, id, name и repr)"""
    model_id = getattr(m, 'id_', None) or getattr(m, 'id', None) or getattr(m, 'name', None)
    if not model_id:
        s = str(m)
        match = re.search(r"id_='([^']+)'", s) or re.search(r"id='([^']+)'", s)
        if match:
            model_id = match.group(1)
    return model_id

def fetch_models(creds):
    try:
        temp_giga = GigaChat(credentials=creds, scope="GIGACHAT_API_PERS", verify_ssl_certs=False)
        models_resp = temp_giga.get_models()
        available = []
        for m in models_resp.data:
            mid = parse_model_id(m)
            if mid and mid not in available:
                available.append(mid)
        return available
    except Exception:
        return []

with st.sidebar:
    st.header("⚙️ Настройки GigaChat")
    sidebar_token = st.text_input(
        "Токен (client_id:client_secret)",
        type="password",
        help="Получить токен: developers.sber.ru"
    )
    if sidebar_token.strip():
        st.session_state["giga_token"] = sidebar_token.strip()
        # Сбрасываем кэш моделей при смене токена
        if "available_models" in st.session_state:
            del st.session_state["available_models"]
        if "working_model" in st.session_state:
            del st.session_state["working_model"]
        st.success("✅ Токен сохранён")
    elif st.session_state.get("giga_token"):
        st.success("✅ Токен активен")
    
    st.markdown("---")
    
    creds = get_credentials()
    # Автоматически запрашиваем модели при первом запуске
    if creds and "available_models" not in st.session_state:
        with st.spinner("🔍 Запрашиваю доступные модели у Сбера..."):
            fetched = fetch_models(creds)
            if fetched:
                st.session_state["available_models"] = fetched
                
    if st.button("🔄 Обновить список моделей"):
        if not creds:
            st.warning("Сначала введите токен выше!")
        else:
            with st.spinner("Запрашиваю список..."):
                fetched = fetch_models(creds)
                if fetched:
                    st.session_state["available_models"] = fetched
                    st.success(f"Найдено: {fetched}")
                else:
                    st.error("Не удалось получить. Проверьте токен.")

    default_models = ["GigaChat-2", "GigaChat", "GigaChat-Plus", "GigaChat-Pro", "GigaChat-Max"]
    models_to_show = st.session_state.get("available_models", default_models)
    
    # Очищаем список от мусора
    models_to_show = [m for m in models_to_show if m and "object_" not in str(m) and "x_headers" not in str(m)]
    if not models_to_show:
        models_to_show = default_models
        
    model_name = st.selectbox("Выберите модель", models_to_show, index=0)
    st.session_state["giga_model"] = model_name

# --- Извлечение текста ---
def extract_text_from_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join([p.text for p in doc.paragraphs])

def extract_text_from_pdf(file_bytes):
    text = ""
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc: text += page.get_text() + "\n"
        if len(text.strip()) > 20: return text
    except: pass
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages: text += page.extract_text() + "\n"
    except: pass
    return text

def extract_text_from_pptx(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    return "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])

def extract_text(file):
    file_bytes = file.read()
    filename = file.name.lower()
    if filename.endswith('.docx'): return extract_text_from_docx(file_bytes)
    elif filename.endswith('.pdf'): return extract_text_from_pdf(file_bytes)
    elif filename.endswith('.pptx'): return extract_text_from_pptx(file_bytes)
    elif filename.endswith(('.png', '.jpg', '.jpeg', '.tiff')):
        try: return pytesseract.image_to_string(Image.open(io.BytesIO(file_bytes)), lang='rus+eng')
        except: return ""
    else:
        try: return file_bytes.decode('utf-8', errors='ignore')
        except: return ""

def extract_number(text):
    m = re.search(r'[№N]\s*(\d+[а-яА-Я/\-]*)', text[:3000])
    if m: return m.group(1)
    m = re.search(r'(?:Договор|Контракт)\s+[№N]?\s*(\d+)', text[:3000], re.IGNORECASE)
    return m.group(1) if m else "не указан"

def extract_date(text):
    months = r'января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря'
    m = re.search(rf'[«"]?(\d{{1,2}})[»"]?\s+({months})\s+(\d{{4}})', text[:3000], re.IGNORECASE)
    if m: return f"{m.group(1)} {m.group(2)} {m.group(3)} г"
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', text[:3000])
    return f"{m.group(1)}.{m.group(2)}.{m.group(3)}" if m else "не указана"

# --- Анализ (С УМНЫМ ПЕРЕБОРОМ И КЭШИРОВАНИЕМ) ---
def analyze_contract(text):
    credentials = get_credentials()
    if not credentials:
        raise ValueError("Не найден токен GigaChat. Вставьте его в боковую панель (⚙️).")

    # Если мы уже нашли рабочую модель в этой сессии — используем только её!
    if st.session_state.get("working_model"):
        models_to_try = [st.session_state["working_model"]]
    else:
        selected = st.session_state.get("giga_model", "")
        available = st.session_state.get("available_models", [])
        fallback = ["GigaChat-2", "GigaChat", "GigaChat-Plus", "GigaChat-Pro", "GigaChat-Max"]
        
        models_to_try = []
        if selected and "object_" not in str(selected): models_to_try.append(selected)
        for m in available:
            if m not in models_to_try and "object_" not in str(m): models_to_try.append(m)
        for m in fallback:
            if m not in models_to_try: models_to_try.append(m)
        
    last_error = None
    for model in models_to_try:
        try:
            giga = GigaChat(
                credentials=credentials,
                scope="GIGACHAT_API_PERS",
                model=model,
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
            # 🎉 Успех! Запоминаем модель, чтобы следующие файлы летали мгновенно
            st.session_state["working_model"] = model
            return response.choices[0].message.content
        except Exception as e:
            last_error = str(e)
            if "No such model" in last_error or "404" in last_error:
                continue
            else:
                raise
    
    raise ValueError(f"Ни одна модель не подошла. Нажмите '🔄 Обновить список' в меню слева. Ошибка: {last_error}")

def empty_parsed():
    return {"Тип договора": "не определён", "Субъектный состав": "не определён",
            "Сумма": "не указана", "ИНН": "0", "Пени": "не указаны", "Риски": "не найдены"}

def parse_response(ai_text):
    data = empty_parsed()
    if not ai_text: return data
    for line in ai_text.split("\n"):
        line = line.strip().lstrip("-•* ").replace("**", "")
        if ":" not in line: continue
        key, val = line.split(":", 1)
        key, val = key.strip().upper(), val.strip()
        if not val: continue
        if "ТИП" in key: data["Тип договора"] = val
        elif "СУБЪЕКТ" in key or "СТОРОН" in key: data["Субъектный состав"] = val
        elif "СУММА" in key or "ЦЕНА" in key: data["Сумма"] = val
        elif "ИНН" in key: data["ИНН"] = val
        elif "ПЕН" in key or "ШТРАФ" in key: data["Пени"] = val
        elif "РИСК" in key: data["Риски"] = val
    return data

# --- Интерфейс ---
st.markdown("---")
st.subheader("📤 Загрузка договоров")
uploaded_files = st.file_uploader("Перетащите файлы", type=['docx', 'pdf', 'pptx', 'txt', 'png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ Загружено файлов: {len(uploaded_files)}")
    results, processed = [], 0

    for file in uploaded_files:
        st.markdown(f"📄 **{file.name}**")
        text = extract_text(file)
        if not text.strip():
            st.warning("⚠️ Не удалось извлечь текст.")
            continue

        model_hint = st.session_state.get('working_model', 'подбираю модель...')
        with st.spinner(f"🤖 ({model_hint}) анализирует документ..."):
            ai_response, ai_error = None, None
            try:
                ai_response = analyze_contract(text)
            except Exception as e:
                ai_error = str(e)

        parsed = parse_response(ai_response)
        results.append((file.name, parsed, ai_response, ai_error, text))
        processed += 1

    if processed:
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

            if ai_response:
                with st.expander("🤖 ИИ-анализ и Риски", expanded=True):
                    st.info(ai_response)
            if ai_error:
                st.error(f"❌ Ошибка: {ai_error}")

            st.download_button("📥 Скачать отчёт (TXT)", f"Отчёт по {fname}\n\n{ai_response or ''}".encode('utf-8'), file_name=f"Report_{fname}.txt", mime="text/plain")
            st.markdown("---")
