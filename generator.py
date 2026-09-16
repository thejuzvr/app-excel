import os
import re
import logging
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
import openpyxl
from openpyxl.cell.cell import MergedCell

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from config import TEMPLATES_DIR, LOGS_DIR, LOG_FILE, CLEAN_TEMP_FILES

# ==========================================
# НАСТРОЙКА ЛОГИРОВАНИЯ
# ==========================================
os.makedirs(LOGS_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ==========================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==========================================
def sanitize_filename(filename):
    """Очищает имя файла от недопустимых символов"""
    bad_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '(', ')']
    for char in bad_chars:
        filename = filename.replace(char, '_')
    
    # Убираем двойные подчеркивания и пробелы по краям
    while '__' in filename:
        filename = filename.replace('__', '_')
    return filename.strip('_ ')


def format_date(value):
    """Форматирует дату в читаемый вид ДД.ММ.ГГГГ"""
    if pd.isna(value):
        return ''
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime('%d.%m.%Y')
    val_str = str(value).strip()
    try:
        dt = pd.to_datetime(val_str)
        if pd.notna(dt):
            return dt.strftime('%d.%m.%Y')
    except Exception:
        pass
    return val_str


def format_context(row, fill_date=""):
    """
    Формирует словарь-контекст для шаблона.
    fill_date: Дата заполнения, переданная из интерфейса.
    """
    context = {}
    
    for key, value in row.to_dict().items():
        key_raw = str(key)
        # Очищаем заголовок от переносов строк (Alt+Enter в Excel) и лишних пробелов
        key_clean = " ".join(key_raw.replace('\r', ' ').replace('\n', ' ').split())
        key_underscore = key_clean.replace(' ', '_')
        
        # Вариант с полным удалением переносов строк (на случай если Alt+Enter нажат внутри слова, напр. выд\nачи)
        key_no_nl = " ".join(key_raw.replace('\r', '').replace('\n', '').split())
        key_no_nl_underscore = key_no_nl.replace(' ', '_')
        
        key_lower = key_clean.lower()
        
        # Форматирование дат из Excel (включая 'действует', 'действителен', 'дата', 'date')
        if isinstance(value, (datetime, pd.Timestamp)) or 'дата' in key_lower or 'date' in key_lower or 'действителен' in key_lower or 'действует' in key_lower:
            formatted_val = format_date(value)
        else:
            if pd.isna(value):
                formatted_val = ''
            else:
                val = value
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                formatted_val = str(val).strip()
        
        # Добавляем все варианты ключей для полной совместимости
        variants = {key, key_clean, key_underscore, key_no_nl, key_no_nl_underscore}
        for var_name in list(variants):
            context[var_name] = formatted_val
            # Поддержка замены «полюс» <-> «полис»
            var_lower = var_name.lower()
            if 'полюс' in var_lower:
                context[var_name.replace('полюс', 'полис').replace('Полюс', 'Полис')] = formatted_val
            elif 'полис' in var_lower:
                context[var_name.replace('полис', 'полюс').replace('Полис', 'Полюс')] = formatted_val
    
    # Добавляем дату заполнения из интерфейса (тег {{дата}})
    context['дата'] = fill_date
    context['Дата'] = fill_date
    
    # Должность: {{Должность}}, {{ДОЛЖНОСТЬ}}, {{должность}}
    dolg = ''
    for k in ['Должность', 'должность', 'ДОЛЖНОСТЬ', 'Профессия', 'Профессия, должность']:
        if k in context and context[k]:
            dolg = context[k]
            break
        elif k in row and pd.notna(row[k]):
            dolg = str(row[k]).strip()
            break
            
    context['Должность'] = dolg
    context['ДОЛЖНОСТЬ'] = dolg
    context['должность'] = dolg
    
    # Подразделение: {{Подразделение}}, {{ПОДРАЗДЕЛЕНИЕ}}, {{подразделение}}, {{стр_Участок}}
    podr = ''
    for k in ['Подразделение', 'подразделение', 'ПОДРАЗДЕЛЕНИЕ', 'стр_Участок', 'Участок', 'Отдел']:
        if k in context and context[k]:
            podr = context[k]
            break
        elif k in row and pd.notna(row[k]):
            podr = str(row[k]).strip()
            break
            
    context['Подразделение'] = podr
    context['ПОДРАЗДЕЛЕНИЕ'] = podr
    context['подразделение'] = podr
    if not context.get('стр_Участок'):
        context['стр_Участок'] = podr

    # Номер договора: взаимозаменяемость Номер_договора и Номер_трудового_договора
    num_dog = ''
    for k in ['Номер_трудового_договора', 'Номер трудового договора', 'Номер_договора', 'Номер договора', '№_договора', '№ договора']:
        if k in context and context[k]:
            num_dog = context[k]
            break
    if num_dog:
        context['Номер_трудового_договора'] = num_dog
        context['Номер_договора'] = num_dog
        context['номер_договора'] = num_dog
    
    return context


def replace_excel_tags(text, context):
    """
    Заменяет теги {{имя_переменной}} в тексте ячейки на значения из context.
    Поддерживает поиск с учетом регистра и нечувствительный fallback.
    """
    if not text or not isinstance(text, str):
        return text
    if '{{' not in text or '}}' not in text:
        return text

    def repl(match):
        key = match.group(1).strip()
        if key in context:
            return str(context[key])
        for k, v in context.items():
            if str(k).strip().lower() == key.lower():
                return str(v)
        return ''

    return re.sub(r'\{\{\s*([^{}]+?)\s*\}\}', repl, text)


def render_excel_template(template_path, context, output_path):
    """
    Заполняет шаблон Excel (.xlsx), заменяя {{переменные}} на значения из context.
    Сохраняет стили ячеек, формулы, объединенные диапазоны и структуру страниц.
    """
    wb = openpyxl.load_workbook(template_path)
    
    for ws in wb.worksheets:
        # Обновляем теги в названии листа
        if '{{' in ws.title and '}}' in ws.title:
            new_title = replace_excel_tags(ws.title, context)[:31].strip()
            if new_title:
                ws.title = new_title

        # Обходим ячейки листа
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is not None and isinstance(cell.value, str):
                    val = cell.value
                    if '{{' in val and '}}' in val:
                        cell.value = replace_excel_tags(val, context)

    wb.save(output_path)
    wb.close()


def convert_excel_to_pdf(excel_path, pdf_path):
    """
    Конвертирует файл .xlsx в .pdf через Excel COM API (Windows).
    """
    if not HAS_WIN32COM:
        raise RuntimeError("Модуль win32com не установлен, конвертация Excel в PDF невозможна.")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_excel = os.path.abspath(excel_path)
        abs_pdf = os.path.abspath(pdf_path)
        
        wb = excel.Workbooks.Open(abs_excel)
        wb.ExportAsFixedFormat(0, abs_pdf)  # 0 = xlTypePDF
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ==========================================
# РАБОТА С ФАЙЛАМИ
# ==========================================
def get_template_files():
    """Возвращает список всех .docx и .xlsx шаблонов, исключая временные файлы (~$)"""
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
        logger.warning(f"Папка '{TEMPLATES_DIR}' не существовала, создана пустая папка")
        return []
    
    templates = []
    for f in os.listdir(TEMPLATES_DIR):
        if f.startswith('~$'):
            continue  # Пропускаем временные файлы Word/Excel
        if f.endswith('.docx') or f.endswith('.xlsx'):
            templates.append(f)
        elif f.endswith('.doc'):
            logger.warning(f"Файл '{f}' в старом формате .doc. Конвертируйте его в .docx.")
        elif f.endswith('.xls'):
            logger.warning(f"Файл '{f}' в старом формате .xls. Конвертируйте его в .xlsx.")
    
    logger.info(f"Найдено шаблонов: {len(templates)}")
    return templates


def load_excel(file_path):
    """Загружает данные из Excel файла"""
    if not os.path.exists(file_path):
        logger.error(f"Файл не найден: {file_path}")
        return None
    
    try:
        df = pd.read_excel(file_path)
        df = df.dropna(how='all')  # Удаляем полностью пустые строки
        df = df.reset_index(drop=True)
        
        # Гарантируем наличие колонок Должность и Подразделение
        if 'Должность' not in df.columns:
            df['Должность'] = ''
        if 'Подразделение' not in df.columns:
            df['Подразделение'] = ''
        
        logger.info(f"Загружен файл: {os.path.basename(file_path)}, строк: {len(df)}")
        return df
    except Exception as e:
        logger.error(f"Ошибка чтения Excel: {e}")
        return None


# ==========================================
# ГЕНЕРАЦИЯ ДОКУМЕНТОВ
# ==========================================
def generate_documents(df, selected_indices, template_files, output_folder, output_format="word", fill_date="", progress_callback=None):
    """
    Генерирует документы для выбранных людей.
    
    Args:
        output_format: "word", "pdf" или "both"
        fill_date: Дата заполнения, выбранная в интерфейсе
    """
    os.makedirs(output_folder, exist_ok=True)
    
    if df is None or df.empty:
        return {'success': 0, 'errors': ['Нет данных для обработки']}
    if not template_files:
        return {'success': 0, 'errors': ['Не найдено ни одного шаблона']}
    if not selected_indices:
        return {'success': 0, 'errors': ['Не выбрано ни одной строки']}
    
    df_selected = df.iloc[selected_indices]
    
    # Определяем количество операций в зависимости от формата
    if output_format == "both":
        total_docs = len(df_selected) * len(template_files) * 2
    else:
        total_docs = len(df_selected) * len(template_files)
    
    current_doc = 0
    results = {'success': 0, 'errors': []}
    
    logger.info(f"=== НАЧАЛО ГЕНЕРАЦИИ ===")
    logger.info(f"Папка сохранения: {output_folder}")
    logger.info(f"Формат вывода: {output_format}")
    logger.info(f"Дата заполнения для документов: {fill_date}") # Логируем дату
    
    for idx_in_selection, (index, row) in enumerate(df_selected.iterrows(), 1):
        fio_raw = str(row['ФИО']) if 'ФИО' in df.columns and pd.notna(row['ФИО']) else f"Person_{index}"
        fio = sanitize_filename(fio_raw)
        
        logger.info(f"[{idx_in_selection}/{len(df_selected)}] Обработка: {fio_raw}")
        
        # <--- ИЗМЕНЕНИЕ ЗДЕСЬ: передаем fill_date в format_context
        context = format_context(row, fill_date)
        
        # Создаем подпапку для человека
        person_folder = os.path.join(output_folder, fio)
        os.makedirs(person_folder, exist_ok=True)
        
        for template_file in template_files:
            try:
                template_path = os.path.join(TEMPLATES_DIR, template_file)
                base_name = sanitize_filename(os.path.splitext(template_file)[0])
                ext = os.path.splitext(template_file)[1].lower()
                
                if ext == '.docx':
                    doc = DocxTemplate(template_path)
                    doc.render(context)
                    target_path = os.path.join(person_folder, f"{fio}_{base_name}.docx")
                    doc.save(target_path)
                    logger.info(f"  Создан Word: {target_path}")
                    
                    if output_format in ["pdf", "both"]:
                        pdf_path = os.path.join(person_folder, f"{fio}_{base_name}.pdf")
                        convert(target_path, pdf_path)
                        logger.info(f"  Создан PDF (из Word): {pdf_path}")
                    
                    if output_format == "pdf" and CLEAN_TEMP_FILES and os.path.exists(target_path):
                        os.remove(target_path)

                elif ext == '.xlsx':
                    target_path = os.path.join(person_folder, f"{fio}_{base_name}.xlsx")
                    render_excel_template(template_path, context, target_path)
                    logger.info(f"  Создан Excel: {target_path}")
                    
                    if output_format in ["pdf", "both"]:
                        pdf_path = os.path.join(person_folder, f"{fio}_{base_name}.pdf")
                        convert_excel_to_pdf(target_path, pdf_path)
                        logger.info(f"  Создан PDF (из Excel): {pdf_path}")
                    
                    if output_format == "pdf" and CLEAN_TEMP_FILES and os.path.exists(target_path):
                        os.remove(target_path)
                else:
                    logger.warning(f"  Пропущен неподдерживаемый файл шаблона: {template_file}")
                    current_doc += 1
                    if progress_callback:
                        progress_callback(current_doc, total_docs)
                    continue
                
                results['success'] += 1
                current_doc += 1
                
                if progress_callback:
                    progress_callback(current_doc, total_docs)
                    
            except Exception as e:
                error_msg = f"Ошибка при обработке {fio_raw} - {template_file}: {str(e)}"
                logger.error(error_msg)
                results['errors'].append(error_msg)
                current_doc += 1
                
                if progress_callback:
                    progress_callback(current_doc, total_docs)
    
    logger.info(f"=== ГЕНЕРАЦИЯ ЗАВЕРШЕНА === Успешно: {results['success']}, Ошибок: {len(results['errors'])}")
    return results