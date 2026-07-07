from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path
import re
from datetime import datetime


def clean_markdown(text):
    """Удаляет markdown-разметку и нормализует эмодзи"""
    # Удаляем заголовки markdown
    text = re.sub(r'^#+\s*', '', text, flags=re.MULTILINE)
    # Удаляем горизонтальные разделители
    text = re.sub(r'^---+\s*$', '', text, flags=re.MULTILINE)
    text = text.replace('**', '')
    text = text.replace('*', '')
    text = text.replace('__', '')
    text = text.replace('_', '')
    text = text.replace('`', '')

    # Нормализация эмодзи
    replacements = {
        '✌': '❌', '✖': '❌', '✕': '❌', '✘': '❌',
        '✗': '❌', '×': '❌', '❎': '❌',
        '✔': '✅', '✓': '✅', '☑': '✅',
        '⚡': '⚠️', '⚠': '⚠️',
    }

    for wrong, correct in replacements.items():
        text = text.replace(wrong, correct)

    return text.strip()


def create_contract_report(contract_data, output_dir):
    """Создаёт Word-отчёт с типом договора и критическими рисками"""

    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # === ЗАГОЛОВОК ===
    title = doc.add_heading('АНАЛИЗ ДОГОВОРА', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # === ТИП ДОГОВОРА ===
    contract_type = contract_data.get('contract_type', 'не определён')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'📋 Тип договора: {contract_type}')
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(68, 114, 196)

    # === ИНФОРМАЦИЯ О ФАЙЛЕ ===
    doc.add_heading('📄 Информация о документе', level=1)
    p = doc.add_paragraph()
    p.add_run('Файл: ').bold = True
    p.add_run(contract_data['filename'])

    # === РЕКВИЗИТЫ ===
    doc.add_heading('📋 Реквизиты договора', level=1)
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Light Grid Accent 1'

    data_rows = [
        ('Номер договора', contract_data['number']),
        ('Дата заключения', contract_data['date']),
        ('Сумма договора', contract_data['amount']),
        ('Найдено ИНН', f"{len(contract_data['inn_list'])} шт. ({', '.join(contract_data['inn_list']) or 'не найдено'})"),
        ('Размер пеней', f"{contract_data['peni']} — {contract_data['peni_risk']}"),
    ]

    for i, (label, value) in enumerate(data_rows):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[i].cells[1].text = value

    # === ОЦЕНКА РИСКА ПО ПЕНЯМ ===
    doc.add_heading('⚠️ Оценка риска по пеням', level=1)
    if 'ВЫШЕ НОРМЫ' in contract_data['peni_risk']:
        p = doc.add_paragraph()
        run = p.add_run('🔴 ВЫСОКИЙ РИСК: ')
        run.bold = True
        run.font.color.rgb = RGBColor(156, 0, 6)
        p.add_run('Размер пеней превышает рыночную норму (0.1% в день).')
    elif 'норма' in contract_data['peni_risk']:
        p = doc.add_paragraph()
        run = p.add_run('🟢 НОРМА: ')
        run.bold = True
        run.font.color.rgb = RGBColor(0, 97, 0)
        p.add_run('Размер пеней в пределах рыночной нормы.')
    else:
        doc.add_paragraph('⚪ Информация о пенях не найдена в договоре.')

    # === ЮРИДИЧЕСКИЙ ЧЕК-ЛИСТ ===
    doc.add_heading('📋 Юридический чек-лист (ИИ-анализ)', level=1)

    ai_text = contract_data.get('ai_analysis', 'Анализ не проводился')
    checklist_items = parse_checklist(ai_text)

    if checklist_items:
        checklist_table = doc.add_table(rows=len(checklist_items) + 1, cols=3)
        checklist_table.style = 'Light Grid Accent 1'

        headers = ['Статус', 'Пункт', 'Комментарий']
        for i, header in enumerate(headers):
            cell = checklist_table.rows[0].cells[i]
            cell.text = header
            cell.paragraphs[0].runs[0].bold = True

        for row_idx, item in enumerate(checklist_items, 1):
            status_cell = checklist_table.rows[row_idx].cells[0]
            status_cell.text = item['status']

            if '✅' in item['status']:
                status_cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
            elif '⚠️' in item['status']:
                status_cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(196, 120, 0)
            elif '❌' in item['status']:
                status_cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(156, 0, 6)

            checklist_table.rows[row_idx].cells[1].text = item['title']
            checklist_table.rows[row_idx].cells[2].text = item['comment']

        doc.add_paragraph()
        summary = calculate_summary(checklist_items) or extract_summary(ai_text)
        if summary:
            p = doc.add_paragraph()
            p.add_run('📊 ИТОГ: ').bold = True
            p.add_run(summary)
    else:
        doc.add_paragraph(clean_markdown(ai_text))

    # === КРИТИЧЕСКИЕ РИСКИ ===
    critical_risks = extract_critical_risks(ai_text)
    if critical_risks:
        doc.add_heading('🔥 Критические риски', level=1)
        for i, risk in enumerate(critical_risks, 1):
            p = doc.add_paragraph(style='List Number')
            run = p.add_run(risk)
            run.font.color.rgb = RGBColor(156, 0, 6)

    # === ЧЕК-ЛИСТ ДЛЯ ИСПРАВЛЕНИЯ ===
    doc.add_heading('✅ Чек-лист для исправления', level=1)

    checklist = generate_checklist(contract_data, checklist_items)

    for i, item in enumerate(checklist, 1):
        p = doc.add_paragraph(style='List Number')
        checkbox = p.add_run('☐ ')
        checkbox.font.size = Pt(14)
        p.add_run(item['task'])

        if item['priority'] == 'высокий':
            p.add_run(f"  [{item['priority'].upper()}]").bold = True
            p.runs[-1].font.color.rgb = RGBColor(156, 0, 6)
        elif item['priority'] == 'средний':
            p.add_run(f"  [{item['priority'].upper()}]")
            p.runs[-1].font.color.rgb = RGBColor(196, 120, 0)

    # === ПОДПИСЬ ===
    doc.add_paragraph()
    doc.add_paragraph('_' * 50)
    p = doc.add_paragraph()
    p.add_run('Отчёт сформирован автоматически системой ИИ-анализа договоров').italic = True
    p = doc.add_paragraph()
    p.add_run(f'Дата формирования: {datetime.now().strftime("%d.%m.%Y %H:%M")}').italic = True

    # === СОХРАНЕНИЕ ===
    safe_filename = re.sub(r'[^\w\-.]', '_', contract_data['filename'].rsplit('.', 1)[0])
    output_path = Path(output_dir) / f'Отчёт_{safe_filename}.docx'
    doc.save(output_path)

    return output_path


def parse_checklist(ai_text):
    """Умный парсер с логикой статусов"""
    items = []
def parse_checklist(ai_text):
    """Умный парсер с разделением склеенных пунктов"""
    
    # === ШАГ 0: Разделяем склеенные пункты ===
    # Если ИИ склеил несколько пунктов в одну строку — разделяем их
    for _ in range(5):  # Повторяем 5 раз на случай множественного склеивания
        ai_text = re.sub(
            r'([✅⚠️❌])\s*(\d+)\.\s*([^:]+?):\s*(.+?)\s*([✅⚠️❌])\s*(\d+)\.',
            r'\1 \2. \3: \4\n\5 \6.',
            ai_text,
            flags=re.DOTALL
        )
    
    items = []
    # === СЛОВА ОБ ОТСУТСТВИИ → ❌ ===
    absence_words = [
        'отсутствует', 'отсутствуют', 'отсутствие', 'отсутствием',
        'не указан', 'не указана', 'не указано', 'не упомянут',
        'не определён', 'не определена', 'не определено',
        'не содержится', 'не включён',
        'не описан', 'не предусмотрен', 'полностью отсутствует',
        'нет пункта', 'нет раздела', 'нет указания', 'нет положения',
        'нет четких', 'нет конкретных', 'нет отдельного', 'нет конкретной',
        'нет запрета', 'нет условий', 'нет основания', 'нет оснований',
        'опущено условие', 'опущены условия',
        'никаких указаний', 'никаких условий',
        'не урегулирован', 'не урегулировано',
        'не оговорены', 'не оговорено',
        'не согласована', 'не согласовано',
        'не указан непосредственно', 'отсутствует информация',
        'отсутствует подробное', 'отсутствует раздел',
        'отсутствует специальный', 'нет специального',
        'нет прямого указания', 'нет прямого упоминания',
        'нет конкретного указания', 'нет конкретного упоминания',
        'не урегулирован вопрос', 'не урегулирована возможность',
        'не упомянуты', 'не упомянуто'
    ]
    absence_words = [
        # ... существующие слова ...
        # === НОВЫЕ СЛОВА ===
        'не имеется', 'не имеются', 'не имелось',
        'отсутствуют положения', 'отсутствует положение',
        'не прописаны положения', 'не прописано положение',
        'не предусмотрено', 'не предусмотрены',
        'нет положений', 'нет положения'
    ]
    # === СЛОВА ОБ УТОЧНЕНИИ → ⚠️ ===
    clarification_words = [
        'требуется уточнение', 'требуется уточнить', 'лучше конкретизировать',
        'следует уточнить', 'требует уточнения', 'нужно уточнить',
        'требует доработки', 'неполная', 'недостаточно подробно',
        'ограничено', 'неясно', 'лишь общее', 'только общее',
        'общие случаи', 'поверхностно', 'описаны неполно',
        'не детализированы', 'общие формулировки', 'общие положения',
        'условия не раскрыты', 'требует конкретизации',
        'требует уточнений', 'необходимо уточнить',
        'желательно включить', 'рекомендуется указать',
        'не уточнено', 'не уточнены', 'не уточнена',
        'частично указано', 'частично указаны', 'частично указана',
        'указано частично', 'указаны частично', 'указана частично',
        'требуется конкретизация',
        'неполный перечень', 'неполная информация',
        'отсутствие требований', 'отсутствие конкретных',
        'неправильно назван', 'недостаточно детализированы',
        # === НОВЫЕ СЛОВА ===
        'не прописан', 'не прописана', 'не прописано',
        'подробно не прописан', 'подробно не прописана',
        'способ не прописан', 'детали не прописаны',
        'конкретные даты отсутствуют', 'конкретная дата отсутствует',
        'отсутствует определение', 'отсутствует конкретика',
        'не подробно', 'не детализирован'
    ]

    # === СЛОВА О НАЛИЧИИ → ✅ ===
    positive_words = [
        'предусмотрен', 'предусмотрено', 'предусматривает',
        'указан', 'указана', 'указано', 'указывает',
        'прописан', 'прописана', 'прописано',
        'определён', 'определена', 'определено',
        'возможно', 'допускается', 'разрешено',
        'согласован', 'согласована', 'согласовано',
        'закреплен', 'закреплено', 'закрепляет',
        'включен', 'включена', 'включено',
        'установлен', 'установлена', 'установлено',
        'достаточно полная', 'достаточно подробн',
        'четкая', 'четко', 'чёткая', 'чётко',
        'чётко определён', 'четко определён',
        'указаны', 'указаны меры', 'указаны сроки',
        'прописаны', 'прописаны основания',
        'определены', 'определены меры'
    ]

    # === НОРМАЛИЗАЦИЯ ТЕКСТА ===
    normalized_lines = []
    for line in ai_text.split('\n'):
        line = clean_markdown(line)
        line = line.strip()
        # Убираем "Комментарий:" в начале строки
        line = re.sub(r'^[Кк]омментарий:\s*', '', line)
        # Убираем квадратные скобки
        line = re.sub(r'\[(.+)\]', r'\1', line)
        if line and not line.startswith('---'):
            normalized_lines.append(line)

    # === ФОРМАТ 1: Стандартный (статус + номер + название: комментарий) ===
    for line in normalized_lines:
        match = re.match(r'^([✅⚠️❌])\s*(\d+)\.\s*([^:]+?):\s*(.+)$', line)
        if match:
            status = match.group(1)
            title = match.group(3).strip()
            comment = match.group(4).strip()

            # Постобработка статусов
            comment_lower = comment.lower()
            has_absence = any(word in comment_lower for word in absence_words)
            has_clarification = any(word in comment_lower for word in clarification_words)
            has_positive = any(word in comment_lower for word in positive_words)

            # === УМНАЯ ЛОГИКА ===
            # Если есть И позитивные слова (указан/предусмотрен), И негативные (отсутствует/не прописан)
            # → значит пункт ЕСТЬ, но с замечаниями → ⚠️
            if has_positive and (has_absence or has_clarification):
                status = '⚠️'
            # Если только негативные слова без позитивных → ❌
            elif has_absence and not has_positive:
                status = '❌'
            # Если только слова уточнения → ⚠️
            elif has_clarification and status == '✅':
                status = '⚠️'
            # Если только позитивные слова → ✅
            elif has_positive and status == '❌':
                status = '✅'

            items.append({
                'status': status,
                'number': match.group(2),
                'title': title,
                'comment': comment
            })

    # === ФОРМАТ 2: Если формат 1 не сработал — ищем пункты по номерам ===
    if len(items) < 5:
        items = []
        i = 0
        while i < len(normalized_lines):
            line = normalized_lines[i]

            # Ищем строку со статусом и номером
            match = re.match(r'^([✅⚠️❌])\s*(\d+)\.\s*([^:]+?):?\s*$', line)
            if match:
                status = match.group(1)
                number = match.group(2)
                title = match.group(3).strip()

                # Собираем комментарий со следующих строк
                comment_lines = []
                i += 1
                while i < len(normalized_lines):
                    next_line = normalized_lines[i]

                    # Останавливаемся на новом пункте
                    if re.match(r'^[✅⚠️❌]\s*\d+\.', next_line):
                        break
                    # Останавливаемся на итогах/рисках
                    if 'итого' in next_line.lower():
                        break
                    if 'критические риск' in next_line.lower():
                        break
                    if 'тип договора' in next_line.lower():
                        break

                    comment_lines.append(next_line)
                    i += 1

                comment = ' '.join(comment_lines).strip()

                if comment:
                    # Постобработка
                    comment_lower = comment.lower()
                    has_absence = any(word in comment_lower for word in absence_words)
                    has_clarification = any(word in comment_lower for word in clarification_words)
                    has_positive = any(word in comment_lower for word in positive_words)

                    # === УМНАЯ ЛОГИКА ===
                    if has_positive and (has_absence or has_clarification):
                        status = '⚠️'
                    elif has_absence and not has_positive:
                        status = '❌'
                    elif has_clarification and status == '✅':
                        status = '⚠️'
                    elif has_positive and status == '❌':
                        status = '✅'

                    items.append({
                        'status': status,
                        'number': number,
                        'title': title,
                        'comment': comment
                    })
                continue

            i += 1

    # === ФИЛЬТРАЦИЯ: убираем пункты с пустыми комментариями ===
    filtered_items = []
    for item in items:
        # Увеличиваем минимум до 15 символов
        if len(item['comment']) < 15:
            continue
        filtered_items.append(item)
    # === ОЧИСТКА ОТ МУСОРНЫХ ФРАЗ ===
    for item in filtered_items:
        comment = item['comment']
        # Удаляем фразы, которые ИИ мог скопировать из промта
        junk_phrases = [
            r'Убираем\s*[«"]?\(?\d+\s*слов\)?[»"]?',
            r'\(\d+\s*слов\)',
            r'твой\s+анализ',
            r'минимум\s+\d+\s*слов',
            r'минимум\s+слов',
            r'\[твой\s+анализ\]',
            r'твой\s+комментарий',
            r'\(отсутствие\s+данного\s+пункта\s+отмечено\s*[✅⚠️❌]\)',
            r'\(недостаток\s+отмечен\s*[✅⚠️❌]\)',
            r'\(недочет\s+указан\s*[✅⚠️❌]\)',
            r'\(недостающий\s+пункт\s*[✅⚠️❌]\)',
            r'\(отмечено\s*[✅⚠️❌]\)',
            r'\(указано\s*[✅⚠️❌]\)',
            r'недостаток\s+отмечен',
            r'недочет\s+указан',
            r'отсутствие\s+отмечено',
            r'недостающий\s+пункт',
        ]
        for pattern in junk_phrases:
            comment = re.sub(pattern, '', comment, flags=re.IGNORECASE)
        # Убираем лишние пробелы
        comment = re.sub(r'\s+', ' ', comment).strip()
        # Убираем висящие знаки препинания
        comment = comment.strip('.,;: ')
        item['comment'] = comment
    
    return filtered_items


def extract_summary(ai_text):
    """Извлекает итоговую строку и очищает от markdown и эмодзи"""
    lines = ai_text.split('\n')
    for line in lines:
        line_lower = line.lower()
        if 'итого' in line_lower:
            summary = line.split(':', 1)[-1].strip() if ':' in line else line.strip()
            summary = summary.replace('**', '').replace('*', '').replace('__', '').replace('_', '')
            summary = re.sub(r'[✅⚠️❌✌✖✕✘✗×✔✓☑🔹🔸🔴🟢🟡]', '', summary).strip()
            summary = re.sub(r'[XYZ]\s*=\s*', '', summary)
            summary = re.sub(r'^[Оо]бщий\s*', '', summary)
            summary = re.sub(r'^[Ии]того[:\s]*', '', summary)
            summary = re.sub(r'\s+', ' ', summary).strip()
            if summary:
                return summary
    return None
    
def calculate_summary(checklist_items):
    """Пересчитывает итог на основе реальных статусов"""
    if not checklist_items:
        return None
    
    ok_count = sum(1 for item in checklist_items if item['status'] == '✅')
    warn_count = sum(1 for item in checklist_items if item['status'] == '⚠️')
    bad_count = sum(1 for item in checklist_items if item['status'] == '❌')
    total = len(checklist_items)
    
    return f"{ok_count} из {total} в порядке, {warn_count} требуют внимания, {bad_count} отсутствуют"

def extract_critical_risks(ai_text):
    """Умный парсер: собирает разбитые строки в один риск"""
    risks = []
    lines = ai_text.split('\n')
    in_risks_block = False
    current_risk = ""

    for line in lines:
        line = line.strip()
        line = clean_markdown(line)

        # Начало блока
        if 'КРИТИЧЕСКИЕ РИСКИ' in line.upper():
            in_risks_block = True
            continue

        # Конец блока
        if in_risks_block:
            if line == '':
                # Сохраняем накопленный риск
                if current_risk and len(current_risk) > 20:
                    risks.append(current_risk.strip())
                    current_risk = ""
                    if len(risks) >= 5:
                        break
                continue
            if re.match(r'^[✅⚠️❌]\s*\d+\.', line):
                break
            if 'ИТОГО' in line.upper():
                break
            if 'ТИП ДОГОВОРА' in line.upper():
                break

        # Собираем риски
        if in_risks_block:
            # Убираем номер в начале строки
            risk_part = re.sub(r'^\d+[\.\)]\s*', '', line)
            risk_part = re.sub(r'^[✅⚠️❌📍🔹🔸🔴🟢🟡]+\s*', '', risk_part)
            risk_part = risk_part.strip()

            # Пропускаем мусорные строки
            junk_phrases = [
                'что именно опасно', 'нарушается статья', 'рекомендация:',
                'описание риска', 'риск + статья', '[риск', '###',
                'наркотики', 'оружие'
            ]
            if any(junk in risk_part.lower() for junk in junk_phrases):
                continue

            # Если строка начинается с тире — это продолжение предыдущего риска
            if risk_part.startswith('—') or risk_part.startswith('-'):
                if current_risk:
                    # Убираем тире в начале и добавляем к текущему риску
                    risk_part = risk_part.lstrip('—-').strip()
                    current_risk += ' — ' + risk_part
            else:
                # Это новый риск — сохраняем предыдущий
                if current_risk and len(current_risk) > 20:
                    risks.append(current_risk.strip())
                current_risk = risk_part

            # Проверяем лимит
            if len(risks) >= 5:
                break

    # Сохраняем последний накопленный риск
    if current_risk and len(current_risk) > 20 and len(risks) < 5:
        risks.append(current_risk.strip())

    return risks[:5]  # Возвращаем максимум 5 рисков


def generate_checklist(contract_data, checklist_items):
    """Генерирует чек-лист для исправления на основе распарсенных данных ИИ"""
    checklist = []

    # Проверка реквизитов (из regex-анализа)
    if contract_data['number'] == 'не указан':
        checklist.append({'task': 'Указать номер договора в преамбуле', 'priority': 'высокий'})

    if contract_data['date'] == 'не указана':
        checklist.append({'task': 'Указать дату заключения договора', 'priority': 'высокий'})

    if contract_data['amount'] == 'не указана':
        price_item = next((i for i in checklist_items if 'цен' in i['title'].lower() or 'расч' in i['title'].lower() or 'оплат' in i['title'].lower()), None)
        if price_item and price_item['status'] in ['⚠️', '❌']:
            checklist.append({'task': 'Чётко прописать сумму договора и порядок расчётов', 'priority': 'высокий'})

    if len(contract_data['inn_list']) < 2:
        checklist.append({'task': 'Проверить реквизиты сторон — должны быть ИНН обеих сторон', 'priority': 'высокий'})

    if 'ВЫШЕ НОРМЫ' in contract_data['peni_risk']:
        checklist.append({'task': 'Снизить размер пеней до рыночной нормы (0.05-0.1% в день)', 'priority': 'высокий'})

    # Анализ чек-листа ИИ — проверяем КОНКРЕТНЫЕ пункты
    for item in checklist_items:
        title_lower = item['title'].lower()
        status = item['status']

        if status == '✅':
            continue

        if 'форс' in title_lower and status in ['⚠️', '❌']:
            checklist.append({'task': 'Добавить/уточнить раздел о форс-мажорных обстоятельствах', 'priority': 'средний'})

        if ('подсудн' in title_lower or 'спор' in title_lower) and status in ['⚠️', '❌']:
            checklist.append({'task': 'Добавить раздел о подсудности споров', 'priority': 'средний'})

        if 'расторж' in title_lower and status in ['⚠️', '❌']:
            checklist.append({'task': 'Уточнить условия расторжения договора', 'priority': 'средний'})

        if ('перс' in title_lower or '152' in title_lower) and status in ['⚠️', '❌']:
            checklist.append({'task': 'Добавить согласие на обработку персональных данных (152-ФЗ)', 'priority': 'высокий'})

        if 'односторонн' in title_lower and status in ['⚠️', '❌']:
            checklist.append({'task': 'Уточнить условия односторонних изменений договора', 'priority': 'высокий'})

        if 'предмет' in title_lower and status in ['⚠️', '❌']:
            checklist.append({'task': 'Уточнить формулировки предмета договора', 'priority': 'высокий'})

        if 'срок' in title_lower and status in ['⚠️', '❌']:
            checklist.append({'task': 'Уточнить сроки исполнения обязательств', 'priority': 'средний'})

        if 'ответ' in title_lower and status in ['⚠️', '❌']:
            if not any('ответственност' in t['task'].lower() for t in checklist):
                checklist.append({'task': 'Добавить/уточнить раздел об ответственности сторон', 'priority': 'средний'})

    # Базовые пункты, если чек-лист пустой
    if not checklist:
        checklist = [
            {'task': 'Проверить соответствие реквизитов сторон', 'priority': 'средний'},
            {'task': 'Убедиться в наличии всех существенных условий', 'priority': 'средний'},
            {'task': 'Проверить сроки действия договора', 'priority': 'средний'},
        ]

    return checklist
