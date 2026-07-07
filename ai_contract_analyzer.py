import re
from pathlib import Path
from datetime import datetime
from docx import Document
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from gigachat import GigaChat


def convert_doc_to_docx(doc_path):
    """Конвертирует .doc в .docx через MS Word"""
    try:
        import win32com.client
        doc_path = Path(doc_path).resolve()
        docx_path = doc_path.with_suffix('.docx')
        if docx_path.exists() and docx_path.stat().st_mtime >= doc_path.stat().st_mtime:
            return docx_path
        print(f"   🔄 Конвертирую {doc_path.name} → .docx ...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(doc_path))
            doc.SaveAs(str(docx_path), FileFormat=16)
            doc.Close()
        finally:
            word.Quit()
        return docx_path
    except Exception as e:
        print(f"   ❌ Ошибка конвертации: {e}")
        return None


def read_docx(file_path):
    """Читает текст из .docx"""
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])


def read_pdf(file_path):
    """Читает текст из .pdf"""
    reader = PdfReader(file_path)
    return '\n'.join([page.extract_text() or '' for page in reader.pages])


def read_pdf_with_ocr(file_path):
    """Читает текст из сканированного PDF через OCR"""
    try:
        import fitz
        import pytesseract
        from PIL import Image
        import io

        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        ]
        for path in possible_paths:
            if Path(path).exists():
                pytesseract.pytesseract.tesseract_cmd = path
                break

        pdf_document = fitz.open(str(file_path))
        text = ''
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))
            text += pytesseract.image_to_string(image, lang='rus+eng') + '\n'
        pdf_document.close()
        return text
    except Exception as e:
        print(f"   ❌ Ошибка OCR: {e}")
        return ''


def read_txt(file_path):
    """Читает текст из .txt с автоопределением кодировки"""
    encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'windows-1251', 'latin-1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return f.read()
        except (UnicodeDecodeError, UnicodeEncodeError):
            continue
    with open(file_path, 'rb') as f:
        return f.read().decode('utf-8', errors='ignore')


def analyze_contract(text, filename):
    """Извлекает реквизиты договора через regex"""
    data = {
        'filename': filename,
        'number': 'не указан',
        'date': 'не указана',
        'amount': 'не указана',
        'inn_list': [],
        'peni': 'не указаны',
        'peni_risk': 'нет данных',
        'ai_analysis': '',
        'contract_type': 'не определён'
    }
    
    if match := re.search(r'№\s*(\d+)', text):
        data['number'] = match.group(1)
    if match := re.search(r'(\d{1,2})\s+_*([а-яё]+)_*\s+(\d{4})\s*г', text, re.IGNORECASE):
        data['date'] = f"{match.group(1)} {match.group(2)} {match.group(3)} г"
    if match := re.search(r'(\d[\d\s]*)\s*(рубл|руб)', text, re.IGNORECASE):
        data['amount'] = match.group(1).replace(' ', '') + ' руб.'
    data['inn_list'] = re.findall(r'ИНН\s*(\d{10,12})', text)
    if match := re.search(r'пени.*?(\d+[,.]?\d*)\s*%', text, re.IGNORECASE):
        peni_value = float(match.group(1).replace(',', '.'))
        data['peni'] = f"{peni_value}%"
        data['peni_risk'] = '⚠️ ВЫШЕ НОРМЫ' if peni_value > 0.1 else '✅ норма'
    
    return data


def safe_decode(content):
    """Безопасно декодирует байты в строку"""
    if isinstance(content, bytes):
        return content.decode('utf-8', errors='ignore')
    return content


def normalize_contract_type(raw_type):
    """Нормализует тип договора к стандартному списку"""
    raw_lower = raw_type.lower()
    
    type_mapping = {
        'субаренда': 'аренда',
        'аренда': 'аренда',
        'наём': 'аренда',
        'найм': 'аренда',
        'поставка': 'поставка',
        'услуга': 'услуги',
        'услуги': 'услуги',
        'оказание услуг': 'услуги',
        'купля-продажа': 'купля-продажа',
        'продажа': 'купля-продажа',
        'дкп': 'купля-продажа',
        'подряд': 'подряд',
        'выполнение работ': 'подряд',
    }
    
    if raw_lower in type_mapping:
        return type_mapping[raw_lower]
    
    for key, value in type_mapping.items():
        if key in raw_lower:
            return value
    
    return 'не определён'


def analyze_with_ai(text, llm):
    """Отправляет договор в GigaChat для анализа"""
    max_length = 3500
    if len(text) > max_length:
        text = text[:max_length] + "...[текст обрезан]"

    # === ЭТАП 1: Тип договора и субъектный состав ===
    prompt_type = f"""Определи тип договора ОДНИМ СЛОВОМ из списка: аренда, поставка, услуги, купля-продажа, подряд, иное.
Также определи субъектный состав: физлица, юрлица, смешанный.

Текст:
{text[:500]}

Ответь СТРОГО в формате: тип_договора, субъектный состав
Пример: аренда, юрлица"""

    try:
        response = llm.chat(prompt_type)
        result = safe_decode(response.choices[0].message.content)
        
        parts = [p.strip() for p in result.split(',')]
        raw_type = parts[0].lower().split()[0] if parts else 'не определён'
        contract_type = normalize_contract_type(raw_type)
        
        subject_type = parts[1].lower().split()[0] if len(parts) > 1 else 'не определён'
        
        valid_subjects = ['физлица', 'юрлица', 'смешанный']
        if subject_type not in valid_subjects:
            subject_type = 'не определён'
    except Exception:
        contract_type = 'не определён'
        subject_type = 'не определён'

    # === ЭТАП 2: Анализ 10 пунктов ===
    prompt_analysis = f"""Ты — юрист. Проанализируй договор по 10 пунктам.
Субъектный состав: {subject_type}

ВАЖНО: отвечай СТРОГО в формате ниже. БЕЗ markdown, БЕЗ заголовков ###, БЕЗ списков -.
Каждый пункт — ОДНА СТРОКА в формате: НОМЕР. НАЗВАНИЕ: СТАТУС КОММЕНТАРИЙ

СТАТУСЫ:
✅ — пункт есть и корректен
⚠️ — пункт есть, но требует уточнения
❌ — пункт отсутствует

Пункты:
1. Предмет договора
2. Цена и порядок расчётов
3. Сроки исполнения
4. Ответственность
5. Форс-мажор
6. Подсудность
7. Расторжение
8. Персональные данные
9. Существенные условия
10. Односторонние изменения

ПРИМЕР ПРАВИЛЬНОГО ОТВЕТА (скопируй формат!):
1. Предмет договора: ✅ В договоре указан адрес объекта, площадь 150 кв.м., кадастровый номер 78:00:1234567:890, назначение — нежилое помещение
2. Цена и порядок расчётов: ❌ В договоре полностью отсутствует раздел о цене и порядке расчётов
3. Сроки исполнения: ⚠️ Указан срок передачи имущества, но конкретные даты начала и окончания аренды не определены
4. Ответственность: ✅ Определены меры ответственности сторон за нарушение условий договора
5. Форс-мажор: ⚠️ Перечислены обстоятельства непреодолимой силы, но отсутствует порядок уведомления
6. Подсудность: ✅ Установлена подсудность споров в Арбитражном суде Санкт-Петербурга
7. Расторжение: ✅ Определены основания и процедура расторжения договора
8. Персональные данные: ❌ Полностью отсутствует раздел об обработке персональных данных
9. Существенные условия: ⚠️ Основные условия указаны, но требуют детализации
10. Односторонние изменения: ❌ Отсутствуют положения об одностороннем изменении условий

ИТОГО: X из 10 в порядке, Y требуют внимания, Z отсутствуют.

Договор:
{text}"""

    analysis = ""
    try:
        response = llm.chat(prompt_analysis)
        analysis = safe_decode(response.choices[0].message.content)
    except Exception as e:
        analysis = f"❌ Ошибка анализа: {str(e)[:100]}"

    # === ЭТАП 3: Критические риски ===
    prompt_risks = f"""Ты — юрист. Тип договора: {contract_type}. Субъектный состав: {subject_type}.
Найди РОВНО 5 КРИТИЧЕСКИХ рисков.

ФОРМАТ (каждый риск — СТРОГО ОДНА СТРОКА):
КРИТИЧЕСКИЕ РИСКИ:
1. [риск] — [статья закона] — [конкретная рекомендация]
2. [риск] — [статья закона] — [конкретная рекомендация]
3. [риск] — [статья закона] — [конкретная рекомендация]
4. [риск] — [статья закона] — [конкретная рекомендация]
5. [риск] — [статья закона] — [конкретная рекомендация]

ВАЖНО ДЛЯ ПОДСУДНОСТИ:
- Если субъекты ФИЗЛИЦА, а указан арбитражный суд — это КРИТИЧЕСКИЙ РИСК (ст.24-27 ГПК РФ)

ВАЖНО ДЛЯ НЕДВИЖИМОСТИ:
- Упоминай: кадастровый номер, площадь, этажность, адрес

Договор:
{text}

Твой ответ (РОВНО 5 строк):"""

    try:
        response = llm.chat(prompt_risks)
        risks = safe_decode(response.choices[0].message.content)
    except Exception:
        risks = ""

    return f"ТИП ДОГОВОРА: {contract_type}\nСУБЪЕКТНЫЙ СОСТАВ: {subject_type}\n\n{analysis}\n\n{risks}"


def save_to_excel(results, output_file):
    """Сохраняет результаты в Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Анализ договоров"
    headers = ['Файл', 'Тип договора', 'Номер', 'Дата', 'Сумма', 'Кол-во ИНН', 'Пени', 'Оценка риска', 'ИИ-анализ']
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    risk_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = border
    for row_idx, data in enumerate(results, 2):
        ws.cell(row=row_idx, column=1, value=data['filename']).border = border
        ws.cell(row=row_idx, column=2, value=data.get('contract_type', 'не определён')).border = border
        ws.cell(row=row_idx, column=3, value=data['number']).border = border
        ws.cell(row=row_idx, column=4, value=data['date']).border = border
        ws.cell(row=row_idx, column=5, value=data['amount']).border = border
        ws.cell(row=row_idx, column=6, value=len(data['inn_list'])).border = border
        ws.cell(row=row_idx, column=7, value=data['peni']).border = border
        risk_cell = ws.cell(row=row_idx, column=8, value=data['peni_risk'])
        risk_cell.border = border
        if 'ВЫШЕ НОРМЫ' in data['peni_risk']:
            risk_cell.fill = risk_fill
            risk_cell.font = Font(color="9C0006", bold=True)
        elif 'норма' in data['peni_risk']:
            risk_cell.fill = ok_fill
            risk_cell.font = Font(color="006100")
        ai_cell = ws.cell(row=row_idx, column=9, value=data.get('ai_analysis', ''))
        ai_cell.border = border
        ai_cell.alignment = Alignment(wrap_text=True)
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 60
    wb.save(output_file)


def main():
    """Основная функция для локального запуска"""
    print("=" * 70)
    print("🤖 ИИ-АНАЛИЗАТОР ДОГОВОРОВ v2.0")
    print("=" * 70)
    
    script_dir = Path(__file__).parent / 'real_contracts'
    if not script_dir.exists():
        script_dir = Path(__file__).parent
        print("⚠️  Папка 'real_contracts' не найдена, использую основную директорию")
    
    print(f"📁 Папка для анализа: {script_dir}")

    # Ключ GigaChat — замените на свой!
    AUTH_KEY = "твой_ключ_сюда"
    
    llm = GigaChat(
        credentials=AUTH_KEY,
        scope="GIGACHAT_API_PERS",
        verify_ssl_certs=False
    )

    files = []
    for ext in ['*.txt', '*.docx', '*.pdf', '*.doc']:
        files.extend(script_dir.glob(ext))

    print(f"\n📂 НАЙДЕНО ФАЙЛОВ: {len(files)}")
    if not files:
        return

    results = []

    for file_path in files:
        if file_path.name.startswith('~$') or file_path.name == 'ai_contracts_report.xlsx':
            continue
        if file_path.suffix.lower() == '.doc':
            docx_version = file_path.with_suffix('.docx')
            if docx_version.exists():
                continue

        print(f"\n🔍 {file_path.name}")

        try:
            if file_path.suffix.lower() == '.doc':
                docx_path = convert_doc_to_docx(file_path)
                if docx_path is None:
                    continue
                file_path = docx_path

            if file_path.suffix == '.docx':
                text = read_docx(file_path)
            elif file_path.suffix == '.pdf':
                text = read_pdf(file_path)
                if not text.strip():
                    text = read_pdf_with_ocr(file_path)
            else:
                text = read_txt(file_path)

            if not text.strip():
                continue

            data = analyze_contract(text, file_path.name)
            print(f"   📄 Номер: {data['number']} | 💰 {data['amount']}")

            ai_analysis = analyze_with_ai(text, llm)
            data['ai_analysis'] = ai_analysis
            
            type_match = re.search(r'ТИП ДОГОВОРА[:\s]+(.+)', ai_analysis, re.IGNORECASE)
            if type_match:
                data['contract_type'] = type_match.group(1).strip()

            results.append(data)

        except Exception as e:
            print(f"   ❌ Ошибка: {type(e).__name__}: {e}")

    if results:
        output_excel = script_dir / 'ai_contracts_report.xlsx'
        save_to_excel(results, output_excel)
        print(f"\n📊 Excel-отчёт сохранён: {output_excel}")

        # Импорт только для локального запуска
        try:
            from report_generator import create_contract_report
            from presentation_generator import create_presentation

            reports_dir = script_dir / 'reports'
            reports_dir.mkdir(exist_ok=True)
            for data in results:
                try:
                    report_path = create_contract_report(data, reports_dir)
                    print(f"   ✅ {report_path.name}")
                except Exception as e:
                    print(f"   ❌ Ошибка для {data['filename']}: {e}")

            presentation_path = script_dir / 'presentation.pptx'
            try:
                create_presentation(results, presentation_path)
                print(f"   ✅ Презентация сохранена")
            except Exception as e:
                print(f"   ❌ Ошибка презентации: {e}")
        except ImportError:
            print("⚠️  Модули отчётов не найдены — пропускаем Word/PPTX")

    print(f"\n🎯 ОБРАБОТАНО: {len(results)} договоров")


if __name__ == "__main__":
    main()
