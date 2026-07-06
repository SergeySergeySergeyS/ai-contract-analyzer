import streamlit as st
import tempfile
import os
from pathlib import Path
import time
from datetime import datetime

# Импортируем наши модули
from gigachat import GigaChat
from report_generator import create_contract_report
from presentation_generator import create_presentation
from ai_contract_analyzer import (
    analyze_contract,
    analyze_with_ai,
    read_docx,
    read_pdf,
    read_pdf_with_ocr,
    read_txt,
    save_to_excel,
    convert_doc_to_docx,
    safe_decode
)

# === НАСТРОЙКИ СТРАНИЦЫ ===
st.set_page_config(
    page_title="🤖 ИИ-Анализатор договоров",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === ЗАГОЛОВОК ===
st.title("🤖 ИИ-Анализатор договоров")
st.markdown("""
**Автоматический анализ договоров с помощью искусственного интеллекта**

📄 Загрузите договоры → 🤖 ИИ найдёт риски → 📊 Получите отчёты
""")

# === БОКОВАЯ ПАНЕЛЬ ===
with st.sidebar:
    st.header("⚙️ Настройки")

    # Поле для ввода ключа GigaChat
    auth_key = st.text_input(
        "🔑 Ключ GigaChat API",
        type="password",
        help="Получите ключ на developers.sber.ru"
    )

    st.markdown("---")
    st.markdown("### 📊 Возможности системы:")
    st.markdown("""
    - ✅ Чтение .txt, .docx, .pdf, .doc
    - 🔍 OCR для сканированных PDF
    - 🤖 ИИ-анализ по 10 пунктам
    - 👥 Определение субъектного состава
    - ⚖️ Проверка правильности подсудности
    - 🔥 Критические риски со ссылками на ГК РФ
    - 📊 Excel-отчёт
    - 📄 Word-отчёт с чек-листом
    - 📽️ Презентация
    """)

    st.markdown("---")
    st.markdown(f"📅 **Дата:** {datetime.now().strftime('%d.%m.%Y')}")

# === ОСНОВНАЯ ОБЛАСТЬ ===
st.header("📤 Загрузка договоров")

uploaded_files = st.file_uploader(
    "Перетащите файлы или нажмите для выбора",
    type=['txt', 'docx', 'pdf', 'doc'],
    accept_multiple_files=True,
    help="Поддерживаемые форматы: .txt, .docx, .pdf, .doc"
)

# === КНОПКА ЗАПУСКА ===
if uploaded_files:
    st.success(f"✅ Загружено файлов: {len(uploaded_files)}")

    # Показываем список файлов
    with st.expander(f"📂 Список файлов ({len(uploaded_files)})"):
        for f in uploaded_files:
            st.write(f"📄 {f.name} ({f.size / 1024:.1f} КБ)")

    if st.button("🚀 Начать анализ", type="primary", use_container_width=True):

        # Проверка ключа
        if not auth_key:
            st.error("❌ Введите ключ GigaChat API в боковой панели!")
            st.stop()

        # Создаём временную папку для работы
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            # Подключаем GigaChat
            with st.spinner("🔗 Подключаюсь к GigaChat..."):
                try:
                    llm = GigaChat(
                        credentials=auth_key,
                        scope="GIGACHAT_API_PERS",
                        verify_ssl_certs=False
                    )
                except Exception as e:
                    st.error(f"❌ Ошибка подключения к GigaChat: {e}")
                    st.stop()

            # === ПРОГРЕСС-БАР ===
            progress_bar = st.progress(0)
            status_text = st.empty()
            results = []

            # === ОБРАБОТКА ФАЙЛОВ ===
            for idx, uploaded_file in enumerate(uploaded_files):
                progress = (idx) / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"🔍 Обрабатываю: {uploaded_file.name}")

                # Сохраняем файл во временную папку
                file_path = temp_path / uploaded_file.name
                with open(file_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())

                try:
                    # Конвертация .doc если нужно
                    if file_path.suffix.lower() == '.doc':
                        docx_path = convert_doc_to_docx(file_path)
                        if docx_path is None:
                            st.warning(f"⚠️ Не удалось конвертировать: {uploaded_file.name}")
                            continue
                        file_path = docx_path

                    # Чтение файла
                    if file_path.suffix == '.docx':
                        text = read_docx(file_path)
                    elif file_path.suffix == '.pdf':
                        text = read_pdf(file_path)
                        if not text.strip():
                            text = read_pdf_with_ocr(file_path)
                    else:
                        text = read_txt(file_path)

                    if not text.strip():
                        st.warning(f"⚠️ Пустой файл: {uploaded_file.name}")
                        continue

                    # Анализ данных
                    data = analyze_contract(text, uploaded_file.name)

                    # ИИ-анализ
                    ai_analysis = analyze_with_ai(text, llm)
                    data['ai_analysis'] = ai_analysis

                    # Извлекаем тип договора и субъектный состав
                    type_match = re.search(r'ТИП ДОГОВОРА[:\s]+(.+)', ai_analysis, re.IGNORECASE)
                    if type_match:
                        data['contract_type'] = type_match.group(1).strip()

                    subject_match = re.search(r'СУБЪЕКТНЫЙ СОСТАВ[:\s]+(.+)', ai_analysis, re.IGNORECASE)
                    if subject_match:
                        data['subject_type'] = subject_match.group(1).strip()

                    results.append(data)

                except Exception as e:
                    st.error(f"❌ Ошибка с {uploaded_file.name}: {e}")

            progress_bar.progress(1.0)
            status_text.text(f"✅ Обработано: {len(results)} договоров")

            # === РЕЗУЛЬТАТЫ ===
            if results:
                st.success(f"🎉 Готово! Проанализировано: {len(results)} договоров")

                # Показываем результаты
                st.header("📊 Результаты анализа")

                for data in results:
                    with st.expander(f"📄 {data['filename']}", expanded=True):
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Номер", data['number'])
                            st.metric("Дата", data['date'])
                        with col2:
                            st.metric("Сумма", data['amount'])
                            st.metric("ИНН", len(data['inn_list']))
                        with col3:
                            st.metric("Пени", data['peni'])
                            if 'ВЫШЕ НОРМЫ' in data['peni_risk']:
                                st.error(data['peni_risk'])
                            elif 'норма' in data['peni_risk']:
                                st.success(data['peni_risk'])

                        # Тип договора и субъектный состав
                        col1, col2 = st.columns(2)
                        with col1:
                            st.info(f"📋 Тип договора: {data.get('contract_type', 'не определён')}")
                        with col2:
                            subject = data.get('subject_type', 'не определён')
                            if subject == 'физлица':
                                st.info(f"👥 Субъектный состав: физические лица")
                            elif subject == 'юрлица':
                                st.info(f"🏢 Субъектный состав: юридические лица")
                            else:
                                st.info(f"👥🏢 Субъектный состав: {subject}")

                        st.markdown("### 🤖 ИИ-анализ:")
                        st.markdown(data.get('ai_analysis', 'Нет данных'))

                # === ГЕНЕРАЦИЯ ОТЧЁТОВ ===
                st.header("📥 Скачать отчёты")

                # Excel
                excel_path = temp_path / 'report.xlsx'
                save_to_excel(results, excel_path)
                with open(excel_path, 'rb') as f:
                    st.download_button(
                        "📊 Скачать Excel-отчёт",
                        f,
                        file_name="ai_contracts_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                # Word-отчёты
                reports_dir = temp_path / 'reports'
                reports_dir.mkdir(exist_ok=True)

                col1, col2 = st.columns(2)
                with col1:
                    for data in results:
                        try:
                            report_path = create_contract_report(data, reports_dir)
                            with open(report_path, 'rb') as f:
                                st.download_button(
                                    f"📄 Отчёт: {data['filename'][:30]}...",
                                    f,
                                    file_name=report_path.name,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                        except Exception as e:
                            st.error(f"❌ {e}")

                # Презентация
                with col2:
                    try:
                        pres_path = temp_path / 'presentation.pptx'
                        create_presentation(results, pres_path)
                        with open(pres_path, 'rb') as f:
                            st.download_button(
                                "📽️ Скачать презентацию",
                                f,
                                file_name="presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"❌ Ошибка презентации: {e}")

else:
    st.info("👆 Загрузите договоры, чтобы начать анализ")

    # Демо-блок
    st.markdown("---")
    st.subheader("ℹ️ Как это работает")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("### 1️⃣")
        st.markdown("Загрузите договоры в любом формате")
    with col2:
        st.markdown("### 2️⃣")
        st.markdown("ИИ проанализирует каждый документ")
    with col3:
        st.markdown("### 3️⃣")
        st.markdown("Найдёт скрытые риски и проблемы")
    with col4:
        st.markdown("### 4️⃣")
        st.markdown("Скачайте готовые отчёты")