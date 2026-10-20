import gspread
from google.oauth2.service_account import Credentials
import json
import logging
import time
import os
from googleapiclient.discovery import build
from gspread.exceptions import WorksheetNotFound, APIError

# --- Настройка логирования ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- КОНФИГУРАЦИЯ ---
SERVICE_ACCOUNT_FILE = 'key_sheet.json'

# Источник: Готовая Легкая Таблица (IMPORT-PROM-UA-PLUS)
SOURCE_SPREADSHEET_ID = os.environ.get('IMPORT_PROM_UA_SPREADSHEET_ID', '1WKlCqKReIp7ZByP6O8ctmjJ3Nm7eFsrcVv77y-9XKd4')

# Приемник: Тяжелая Контент-Таблица (IMPORT-PROM-UA-PLUS-HEAVY)
TARGET_SPREADSHEET_ID = os.environ.get('IMPORT_PROM_UA_HEAVY_SPREADSHEET_ID', '1Rkbne3OmHOC3vZtvnuvW_eINL1dBIKdNI7sYkMxATmM')

# Минимальное количество пустых строк
MIN_EMPTY_ROWS = 1000
SHEETS_TO_ENSURE_EMPTY_ROWS = [
    'Export Products Sheet',
    'Export Groups Sheet',
]

# Ключевые колонки для синхронизации
SHEET_DATA_CONFIG = {
    'Export Products Sheet': {
        'key_cols': ['Код_товару', 'Ідентифікатор_товару'],
        'sort_col': 'Особисті_нотатки',
    },
    'Export Groups Sheet': {
        'key_cols': ['Номер_групи', 'Назва_групи', 'Ідентифікатор_групи', 'Номер_батьківської_групи'],
        'sort_col': None,
    },
}

# 🛡️ КОЛОНКИ, КОТОРЫЕ КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО ЗАТИРАТЬ В HEAVY-ТАБЛИЦЕ
# Если в Heavy-таблице в этих колонках уже есть значение, оно НИКОГДА не перезаписывается легкой таблицей!
PROTECTED_COLUMNS = [
    'Пошукові_запити',
    'Пошукові_запити_укр',
    # Описания и названия не перезаписываем, если они уже обогащены нейросетью/GAS
    'Опис',
    'Опис_укр',
]


def get_google_sheet_client(credentials):
    try:
        client = gspread.authorize(credentials)
        logging.info("Успешная авторизация в Google Sheets API через gspread.")
        return client
    except Exception as e:
        logging.error(f"Ошибка авторизации gspread: {e}")
        raise


def get_sheets_service(credentials):
    try:
        service = build('sheets', 'v4', credentials=credentials)
        logging.info("Успешное создание сервисного клиента sheets v4.")
        return service
    except Exception as e:
        logging.error(f"Ошибка создания сервисного клиента sheets: {e}")
        raise


def get_column_letter(col_idx_0based):
    result = ""
    col_num = col_idx_0based + 1
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def find_column_index(headers, col_name):
    for i, h in enumerate(headers):
        if h and h.strip().lower() == col_name.strip().lower():
            return i
    return -1


def sync_sheet_structure_and_data(source_sheet, target_sheet, sheets_service, target_spreadsheet_id, sheet_title):
    logging.info(f"--- Начало синхронизации листа: {sheet_title} ---")
    
    # Чтение данных из источника (Легкая таблица)
    source_all_values = source_sheet.get_all_values()
    if not source_all_values:
        logging.warning(f"Исходный лист {sheet_title} пуст. Пропуск.")
        return

    source_headers = source_all_values[0]
    source_rows = source_all_values[1:]
    logging.info(f"В источнике {sheet_title}: колонок {len(source_headers)}, строк {len(source_rows)}")

    # Чтение целевого листа (Heavy таблица)
    target_all_values = target_sheet.get_all_values()
    if not target_all_values:
        # Лист абсолютно пустой: записываем заголовки
        logging.info(f"Целевой лист {sheet_title} абсолютно пустой. Инициализируем заголовки...")
        target_sheet.update(range_name='A1', values=[source_headers])
        target_headers = source_headers
        target_rows = []
    else:
        target_headers = target_all_values[0]
        target_rows = target_all_values[1:]
        # Проверяем структуру заголовков
        if target_headers != source_headers:
            logging.info(f"Обновляем заголовки в целевом листе {sheet_title} под структуру источника...")
            target_sheet.update(range_name='A1', values=[source_headers])
            target_headers = source_headers

    logging.info(f"В приемнике {sheet_title}: колонок {len(target_headers)}, строк {len(target_rows)}")

    # Синхронизация данных
    if sheet_title not in SHEET_DATA_CONFIG:
        logging.info(f"Лист {sheet_title} не требует синхронизации строк по правилам. Пропуск.")
        return

    config = SHEET_DATA_CONFIG[sheet_title]
    key_cols = config['key_cols']

    # Индексы ключевых колонок в источнике и приемнике
    src_key_indices = [find_column_index(source_headers, c) for c in key_cols]
    tgt_key_indices = [find_column_index(target_headers, c) for c in key_cols]

    if -1 in src_key_indices or -1 in tgt_key_indices:
        logging.error(f"Не все ключевые колонки найдены в листе {sheet_title}!")
        return

    # Индексы защищенных колонок в целевой таблице
    protected_indices = set()
    for col_name in PROTECTED_COLUMNS:
        idx = find_column_index(target_headers, col_name)
        if idx != -1:
            protected_indices.add(idx)

    logging.info(f"Защищенные колонки в {sheet_title}: {[target_headers[i] for i in protected_indices]}")

    # Создаем карту строк целевой таблицы: key -> index
    target_map = {}
    for idx, row in enumerate(target_rows):
        key = tuple(row[k_idx].strip() if k_idx < len(row) else '' for k_idx in tgt_key_indices)
        if any(key):
            target_map[key] = idx

    update_operations = []
    new_rows_to_add = []
    seen_keys = set()

    for src_row in source_rows:
        key = tuple(src_row[k_idx].strip() if k_idx < len(src_row) else '' for k_idx in src_key_indices)
        if not any(key):
            continue

        seen_keys.add(key)

        if key in target_map:
            # ТОВАР УЖЕ ЕСТЬ В HEAVY: Проверяем, что нужно обновить, НЕ затирая защищенные поля
            tgt_row_idx = target_map[key]
            tgt_row = target_rows[tgt_row_idx]
            tgt_row_num = tgt_row_idx + 2  # 1-based + 1 за заголовок

            num_cols = min(len(source_headers), len(target_headers))
            for col_idx in range(num_cols):
                # 🛡️ ЕСЛИ КОЛОНКА ЗАЩИЩЕНА:
                if col_idx in protected_indices:
                    # Если в Heavy уже есть текст — НЕ ТРОГАЕМ!
                    existing_val = tgt_row[col_idx] if col_idx < len(tgt_row) else ''
                    if existing_val.strip():
                        continue  # Оставляем старый накопленный контент!

                src_val = src_row[col_idx] if col_idx < len(src_row) else ''
                tgt_val = tgt_row[col_idx] if col_idx < len(tgt_row) else ''

                if src_val != tgt_val and src_val.strip():
                    col_letter = get_column_letter(col_idx)
                    cell_range = f"'{sheet_title}'!{col_letter}{tgt_row_num}"
                    update_operations.append({
                        'range': cell_range,
                        'values': [[src_val]]
                    })
        else:
            # НОВЫЙ ТОВАР: Добавляем в Heavy-таблицу
            new_row = [''] * len(target_headers)
            num_cols = min(len(src_row), len(target_headers))
            for col_idx in range(num_cols):
                new_row[col_idx] = src_row[col_idx]
            new_rows_to_add.append(new_row)

    # Поиск строк на удаление (товары, пропавшие у поставщиков)
    rows_to_delete_1based = []
    for key, tgt_row_idx in target_map.items():
        if key not in seen_keys:
            rows_to_delete_1based.append(tgt_row_idx + 2)

    logging.info(f"Итог анализа {sheet_title}: обновлений ячеек {len(update_operations)}, добавлений новых строк {len(new_rows_to_add)}, удалений {len(rows_to_delete_1based)}")

    # 1. Применяем пакетные обновления существующих ячеек
    BATCH_SIZE = 10000
    if update_operations:
        logging.info(f"Применяем {len(update_operations)} обновлений пачками по {BATCH_SIZE}...")
        for i in range(0, len(update_operations), BATCH_SIZE):
            chunk = update_operations[i:i + BATCH_SIZE]
            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': chunk
            }
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=target_spreadsheet_id,
                body=body
            ).execute()
            time.sleep(1)

    # 2. Добавляем новые строки
    if new_rows_to_add:
        logging.info(f"Добавляем {len(new_rows_to_add)} новых строк в целевую таблицу...")
        target_sheet.append_rows(new_rows_to_add, value_input_option='USER_ENTERED')

    # 3. Удаляем выбывшие строки
    if rows_to_delete_1based:
        logging.info(f"Удаляем {len(rows_to_delete_1based)} выбывших строк...")
        # Удаляем с конца, чтобы индексы не смещались
        rows_to_delete_1based.sort(reverse=True)
        # Получаем sheetId
        target_sheet_id = target_sheet.id
        requests = []
        for r_num in rows_to_delete_1based:
            requests.append({
                'deleteDimension': {
                    'range': {
                        'sheetId': target_sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': r_num - 1,
                        'endIndex': r_num
                    }
                }
            })
            if len(requests) >= 500:
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=target_spreadsheet_id,
                    body={'requests': requests}
                ).execute()
                requests = []
                time.sleep(1)
        if requests:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=target_spreadsheet_id,
                body={'requests': requests}
            ).execute()

    logging.info(f"Синхронизация листа {sheet_title} успешно завершена!")


def main():
    start_time = time.time()
    logging.info("=== Запуск скрипта умной синхронизации SYNC_HEAVY_PROM_UA ===")
    logging.info(f"Исходная легкая таблица (SOURCE): {SOURCE_SPREADSHEET_ID}")
    logging.info(f"Целевая тяжелая таблица (TARGET): {TARGET_SPREADSHEET_ID}")

    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)

    gspread_client = get_google_sheet_client(credentials)
    sheets_service = get_sheets_service(credentials)

    source_spreadsheet = gspread_client.open_by_key(SOURCE_SPREADSHEET_ID)
    target_spreadsheet = gspread_client.open_by_key(TARGET_SPREADSHEET_ID)

    for sheet_title in ['Export Products Sheet', 'Export Groups Sheet']:
        try:
            src_sheet = source_spreadsheet.worksheet(sheet_title)
            try:
                tgt_sheet = target_spreadsheet.worksheet(sheet_title)
            except WorksheetNotFound:
                logging.info(f"Лист {sheet_title} не найден в целевой. Создаем...")
                tgt_sheet = target_spreadsheet.add_worksheet(title=sheet_title, rows=1000, cols=130)

            sync_sheet_structure_and_data(src_sheet, tgt_sheet, sheets_service, TARGET_SPREADSHEET_ID, sheet_title)
        except Exception as e:
            logging.error(f"Ошибка при синхронизации листа {sheet_title}: {e}", exc_info=True)

    elapsed = time.time() - start_time
    logging.info(f"=== Синхронизация завершена за {elapsed:.2f} сек ===")


if __name__ == '__main__':
    main()
