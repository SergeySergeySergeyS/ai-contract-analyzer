import streamlit as st
from gigachat import GigaChat
import os
import io
import re

# Библиотеки для чтения и создания документов
from docx import Document
import PyPDF2
from pptx import Presentation
import fitz  # PyMuPDF
from PIL import Image
import pytesseract

st.set_page_config(page_title="ИИ-Анализатор договоров", page_icon="🤖", layout="wide")

st.title("🤖 ИИ-Анализатор договоров")
st.markdown("Автоматический анализ договоров с помощью искусственного интеллекта")
st.markdown("📄 **Загрузите договоры** → 🤖 **ИИ найдёт риски** → 📊 **Получите отчёты и презентации**")

with st.sidebar:
    st.header("⚙️ Настройки GigaChat")
    sidebar_token = st.text_input(
        "Токен (client_id:client_secret)",
        type="password",
        help="Получить токен: developers.sber.ru"
    )
    if sidebar_token.strip():
        st.session_state["giga_token"] = sidebar_token.strip()
        if "available_models" in st.session_state:
            del st.session_state["available_models"]
        if "working_model" in st.session_state:
            del st.session_state["working_model"]
        st.success("✅ Токен сохранён")
    elif st.session_state.get("giga_token"):
        st.success("✅ Токен активен")
    
    st.markdown("---")
    
    creds = st.session_state.get("giga_token", "")
    if creds and "available_models" not in st.session_state:
        with st.spinner("🔍 Запрашиваю доступные модели у Сбера..."):
            try:
                temp_giga = GigaChat(credentials=creds, scope="GIGACHAT_API_PERS", verify_ssl_certs=False)
                models_resp = temp_giga.get_models()
                available = []
                for m in models_resp.data:
                    mid = getattr(m, 'id_', None) or getattr(m, 'id', None) or getattr(m, 'name', None)
                    if not mid:
                        match = re.search(r"id_='([^']+)'", str(m)) or re.search(r"id='([^']+)'", str(m))
                        if match: mid = match.group(1)
                    if mid and mid not in available: available.append(mid)
                st.session_state["available_models"] = available
            except Exception:
                pass
                
    default_models = ["GigaChat", "GigaChat-Max", "GigaChat-Pro", "GigaChat-Plus", "GigaChat-Lite"]
    models_to_show = st.session_state.get("available_models", default_models)
    models_to_show = [m for m in models_to_show if m and "object_" not in str(m)]
    if not models_to_show: models_to_show = default_models
        
    model_name = st.selectbox("Выберите модель", models_to_show, index=0)
    st.session_state["giga_model"] = model_name

def get_credentials():
    if st.session_state.get("giga_token"): return st.session_state["giga_token"]
    try:
        for key in ("GIGACHAT_CREDENTIALS", "GIGACHAT_TOKEN", "GIGACHAT_ACCESS_TOKEN"):
            if key in st.secrets: return str(st.secrets[key]).strip()
    except Exception: pass
    return None

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

# --- Анализ ---
def analyze_contract(text):
    credentials = get_credentials()
    if not credentials: raise ValueError("Не найден токен GigaChat.")

    if st.session_state.get("working_model"):
        models_to_try = [st.session_state["working_model"]]
    else:
        selected = st.session_state.get("giga_model", "")
        available = st.session_state.get("available_models", [])
        fallback = ["GigaChat", "GigaChat-Max", "GigaChat-Pro", "GigaChat-Plus", "GigaChat-Lite"]
        models_to_try = []
        if selected and "object_" not in str(selected): models_to_try.append(selected)
        for m in available:
            if m not in models_to_try and "object_" not in str(m): models_to_try.append(m)
        for m in fallback:
            if m not in models_to_try: models_to_try.append(m)
        
    last_error = None
    for model in models_to_try:
        try:
            giga = GigaChat(credentials=credentials, scope="GIGACHAT_API_PERS", model=model, verify_ssl_certs=False)
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
            st.session_state["working_model"] = model
            return response.choices[0].message.content
        except Exception as e:
            last_error = str(e)
            if "No such model" in last_error or "404" in last_error: continue
            else: raise
    raise ValueError(f"Ни одна модель не подошла. Ошибка: {last_error}")

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
        elif "СУБЪЕКТ" in key or "СОСТАВ" in key or "СТОРОН" in key: data["Субъектный состав"] = val
        elif "СУММА" in key or "ЦЕНА" in key: data["Сумма"] = val
        elif "ИНН" in key: data["ИНН"] = val
        elif "ПЕН" in key or "ШТРАФ" in key: data["Пени"] = val
        elif "РИСК" in key: data["Риски"] = val
    return data

# --- ГЕНЕРАЦИЯ ОТЧЕТОВ И ПРЕЗЕНТАЦИЙ ---
def generate_docx_report(fname, parsed, ai_response):
    doc = Document()
    doc.add_heading(f'Отчёт по анализу: {fname}', 0)
    
    doc.add_heading('📋 Основные параметры', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Параметр'
    hdr_cells[1].text = 'Значение'
    
    for key, val in parsed.items():
        if key != 'Риски':
            row_cells = table.add_row().cells
            row_cells[0].text = key
            row_cells[1].text = str(val)
            
    if parsed.get("Тип договора") and parsed["Тип договора"] != "не определён":
        doc.add_paragraph(f"\nТип договора: {parsed['Тип договора']}")
    if parsed.get("Субъектный состав") and parsed["Субъектный состав"] != "не определён":
        doc.add_paragraph(f"Субъектный состав: {parsed['Субъектный состав']}")
            
    doc.add_heading('🤖 ИИ-анализ', level=1)
    doc.add_paragraph(ai_response if ai_response else "Анализ не выполнен")
    
    if parsed.get("Риски") and parsed["Риски"] != "не найдены":
        doc.add_heading('⚠️ Юридические риски', level=1)
        doc.add_paragraph(parsed["Риски"])
        
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_pptx_report(fname, parsed, ai_response):
    prs = Presentation()
    
    # Слайд 1: Титульный
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = f"Анализ договора\n{fname}"
    slide.placeholders[1].text = "Подготовлено ИИ-Анализатором"
    
    # Слайд 2: Основные параметры
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    slide.shapes.title.text = "📋 Основные параметры"
    tf = slide.placeholders[1].text_frame
    
    keys_to_show = ["Номер", "Дата", "Сумма", "ИНН", "Пени"]
    for i, key in enumerate(keys_to_show):
        if i == 0:
            tf.text = f"{key}: {parsed.get(key, 'не указано')}"
        else:
            p = tf.add_paragraph()
            p.text = f"{key}: {parsed.get(key, 'не указано')}"
            
    # Слайд 3: Стороны и тип
    slide = prs.slides.add_slide(bullet_slide_layout)
    slide.shapes.title.text = "👥 Тип и стороны"
    tf = slide.placeholders[1].text_frame
    tf.text = f"Тип договора: {parsed.get('Тип договора', 'не определен')}"
    p = tf.add_paragraph()
    p.text = f"Состав: {parsed.get('Субъектный состав', 'не определен')}"
    
    # Слайд 4: Риски
    if parsed.get("Риски") and parsed["Риски"] != "не найдены":
        slide = prs.slides.add_slide(bullet_slide_layout)
        slide.shapes.title.text = "⚠️ Юридические риски"
        slide.placeholders[1].text_frame.text = parsed["Риски"]
        
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# --- Основной интерфейс ---
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

            # КНОПКИ СКАЧИВАНИЯ ОТЧЕТОВ И ПРЕЗЕНТАЦИЙ
            st.markdown("##### 📥 Скачать результаты:")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                docx_buffer = generate_docx_report(fname, parsed, ai_response)
                st.download_button(
                    label="📄 Отчёт (Word)",
                    data=docx_buffer,
                    file_name=f"Report_{fname}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with col2:
                pptx_buffer = generate_pptx_report(fname, parsed, ai_response)
                st.download_button(
                    label="📊 Презентация (PPTX)",
                    data=pptx_buffer,
                    file_name=f"Presentation_{fname}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            with col3:
                st.download_button(
                    label="📝 Текст (TXT)",
                    data=f"Отчёт по {fname}\n\n{ai_response or ''}".encode('utf-8'),
                    file_name=f"Report_{fname}.txt",
                    mime="text/plain",
                    use_container_width=True
                )
            st.markdown("---")
