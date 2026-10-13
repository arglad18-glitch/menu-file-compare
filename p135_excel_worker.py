import os
import sys
import tempfile
import subprocess
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Определение папки, в которой физически находится данный скрипт в Linux
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

def _show_error_window(title, message):
    """Выводит ошибку в виде графической формы на экран и завершает работу."""
    root = tk.Tk()
    root.withdraw()  # Скрываем основное пустое окно Tkinter
    messagebox.showerror(title, message)
    root.destroy()

def _scan_directory_for_files():
    """
    Сканирует BASE_DIR на наличие файлов .xlsx, .xls, .ods.
    При превышении лимита файлов прерывает программу графическим окном.
    """
    valid_extensions = ('.xlsx', '.xls', '.ods')
    found_files = []
    
    # Считываем абсолютно все файлы в директории запуска
    for file in os.listdir(BASE_DIR):
        # Игнорируем скрытые временные файлы (~$)
        if file.lower().endswith(valid_extensions) and not file.startswith('~$'):
            found_files.append(file)
            
    if len(found_files) == 0:
        _show_error_window(
            "Ошибка выполнения", 
            f"В папке запуска не найдено файлов для сверки (.xlsx, .xls, .ods)!\n\n"
            f"Путь сканирования: {BASE_DIR}\n\n"
            f"Пожалуйста, добавьте файлы и запустите программу снова."
        )
        sys.exit(1)
        
    elif len(found_files) > 2:
        # Формируем список файлов для вывода в графическое окно
        files_list_str = "\n".join([f"- {f}" for f in found_files])
        _show_error_window(
            "Ошибка выполнения", 
            f"В папке обнаружено слишком много файлов ({len(found_files)}). Должно быть не более двух!\n\n"
            f"Найденные файлы:\n{files_list_str}\n\n"
            f"Устраните проблему (удалите или переместите лишние файлы из этой папки) и запустите программу снова."
        )
        sys.exit(1)
        
    return found_files

def _load_file_to_matrix(file_name):
    
    import zipfile
    import xml.etree.ElementTree as ET

    full_path = os.path.join(BASE_DIR, file_name)
    ext = os.path.splitext(file_name)[1].lower()
    
    # 1. Если это современный формат .xlsx (основной для Excel)
    if ext == '.xlsx':
        import openpyxl
        wb = openpyxl.load_workbook(full_path, data_only=False)
        ws = wb.active
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([str(cell) if cell is not None else '' for cell in row])
        wb.close()
        df = pd.DataFrame(data)
        
    # 2. Если это формат LibreOffice / OpenOffice Calc (.ods)
    elif ext == '.ods':
        active_table_name = None
        
        # Шаг 2.1: Ищем имя активного листа в settings.xml
        if zipfile.is_zipfile(full_path):
            try:
                with zipfile.ZipFile(full_path, 'r') as z:
                    if 'settings.xml' in z.namelist():
                        root = ET.fromstring(z.read('settings.xml'))
                        for prop in root.iter('{urn:openoffice:document:settings:1.0}config-item'):
                            if prop.attrib.get('{urn:openoffice:document:settings:1.0}name') == 'ActiveTable':
                                if prop.text:
                                    active_table_name = prop.text.strip()
                                    break
                                    
                    # Шаг 2.2: Извлекаем данные конкретно этого листа напрямую из content.xml
                    if 'content.xml' in z.namelist():
                        content_root = ET.fromstring(z.read('content.xml'))
                        
                        target_table_element = None
                        all_tables = list(content_root.iter('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table'))
                        
                        # Если нашли имя активного листа, ищем именно его элемент
                        if active_table_name:
                            for table in all_tables:
                                if table.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name') == active_table_name:
                                    target_table_element = table
                                    break
                        
                        # Если имя не нашлось — берем первый лист
                        if target_table_element is None and all_tables:
                            target_table_element = all_tables[0]
                        
                        # Шаг 2.3: Парсим строки и ячейки выбранного листа
                        data = []
                        if target_table_element is not None:
                            for row_node in target_table_element.iter('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-row'):
                                row_data = []
                                for cell_node in row_node.iter('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-cell'):
                                    repeated = cell_node.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated')
                                    repeat_count = int(repeated) if repeated else 1
                                    
                                    # Собираем текст из <text:p>
                                    text_nodes = list(cell_node.iter('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p'))
                                    cell_text = "".join([t.text if t.text else "" for t in text_nodes]) if text_nodes else ""
                                    
                                    # Если ячейка пустая и повторяется слишком много раз (хвост листа Calc) — не раздуваем память
                                    if not cell_text and repeat_count > 100:
                                        # Добавляем только одну пустую ячейку, чтобы зафиксировать структуру, но не плодить тысячи элементов
                                        row_data.append('')
                                        continue
                                        
                                    for _ in range(repeat_count):
                                        row_data.append(cell_text)
                                data.append(row_data)
                        
                        # Шаг 2.4: Фильтрация структуры матрицы
                        if data:
                            # Удаляем полностью пустые строки с конца (мусорные строки Calc)
                            while data and not any(cell != '' for cell in data[-1]):
                                data.pop()
                                
                            if data:
                                # Выравниваем длину строк до максимальной фактически заполненной
                                max_cols = max(len(r) for r in data)
                                for r in data:
                                    while len(r) < max_cols:
                                        r.append('')
                                        
                                df = pd.DataFrame(data)
                                df = df.fillna('')
                                return df.values.tolist()
                                
            except Exception as e:
                print(f"Прямой парсинг ODS не удался, откатываемся на pandas: {e}")
        
        # Резервный вариант, если XML-структура нарушена
        df = pd.read_excel(full_path, sheet_name=0, header=None, dtype=str, engine='odf')
        df = df.fillna('')



        
    # 3. Для старых форматов .xls
    else:
        df = pd.read_excel(full_path, sheet_name=0, header=None, dtype=str, engine='xlrd')
        df = df.fillna('')

    return df.values.tolist()

def _open_in_background_calc(matrix, original_name):
    import time # Добавляем для уникальности времени
    ext = os.path.splitext(original_name)[1].lower()
    base_name = os.path.splitext(os.path.basename(original_name))[0]
    
    # Добавляем к имени метку времени запуска, чтобы файлы никогда не конфликтовали в /tmp
    unique_prefix = f"{base_name}_{int(time.time())}_"
    
    with tempfile.NamedTemporaryFile(prefix=unique_prefix, suffix=ext, delete=False) as tmp_file:
        tmp_path = tmp_file.name
        
    df = pd.DataFrame(matrix)
    engine = 'openpyxl' if ext == '.xlsx' else ('xlwt' if ext == '.xls' else 'odf')
    df.to_excel(tmp_path, index=False, header=False, engine=engine)
    
    # Запускаем LibreOffice Calc в Linux асинхронно. Файл откроется на экране и «не закроется»
    try:
        subprocess.Popen(['xdg-open', tmp_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        try:
            subprocess.Popen(['libreoffice', tmp_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            print(f"Не удалось автоматически открыть редактор таблиц: {e}")

def p135_run_excel_comparator(*args, **kwargs):
    """
    Главная точка входа для внешних модулей.
    """
    # 1. Автоматический поиск файлов с жестким контролем количества
    files = _scan_directory_for_files()
    
    if not files:
        return None
        
    if len(files) == 1:
        _show_error_window(
            "Ошибка выполнения", 
            f"Найден только один файл '{files[0]}'.\n\n"
            f"Для корректной сверки в папке должна находиться пара файлов. Добавьте второй файл и запустите программу снова."
        )
        return None
        
    file1_name = files[0]
    file2_name = files[1]

    # 2. Чтение данных 
    try:
        data1 = _load_file_to_matrix(file1_name)
        data2 = _load_file_to_matrix(file2_name)
    except Exception as e:
        _show_error_window(
            "Ошибка чтения файлов", 
            f"Не удалось загрузить данные из таблиц.\n"
            f"Убедитесь, что файлы не повреждены и не открыты в другой программе.\n\nДетали: {e}"
        )
        return None

    # 3. Генерация слепков строк для сквозного сопоставления
    keys1 = ["|||".join(row[:2]) for row in data1] #keys1 = ["|||".join(row) for row in data1]
    keys2 = ["|||".join(row[:2]) for row in data2] #keys2 = ["|||".join(row) for row in data2]

    # 4. Сверка и сборка результатов в 3-й столбец (индекс 2)
    memory_table1 = []
    for row, key in zip(data1, keys1):
        status = "Расхождений нет" if key in keys2 else "Есть расхождения"
        new_row = row[:2] + [status] + row[2:]
        memory_table1.append(new_row)

    memory_table2 = []
    for row, key in zip(data2, keys2):
        status = "Расхождений нет" if key in keys1 else "Есть расхождения"
        new_row = row[:2] + [status] + row[2:]
        memory_table2.append(new_row)
    
    # 5. открываем оба файла не изменяя исходники на диске
    # Безопасное добавление заголовка для первой таблицы
    if not memory_table1:
        # Если таблица была совсем пустая, создаем первую строку с заголовком
        memory_table1.append(["", "", "Результат сверки"])
    else:
        # Если строки есть, расширяем первую строку до 3-х столбцов, если нужно
        while len(memory_table1[0]) < 3:
            memory_table1[0].append("")
        memory_table1[0][2] = "Результат сверки"

    # Безопасное добавление заголовка для второй таблицы
    if not memory_table2:
        # Если таблица была совсем пустая, создаем первую строку с заголовком
        memory_table2.append(["", "", "Результат сверки"])
    else:
        # Если строки есть, расширяем первую строку до 3-х столбцов, если нужно
        while len(memory_table2[0]) < 3:
            memory_table2[0].append("")
        memory_table2[0][2] = "Результат сверки"

    _open_in_background_calc(memory_table1, file1_name)
    _open_in_background_calc(memory_table2, file2_name)
    
    # Возвращаем структуры
    return {
        file1_name: memory_table1,
        file2_name: memory_table2
    }

if __name__ == "__main__":
    # Локальный тест работы модуля
    p135_run_excel_comparator()
