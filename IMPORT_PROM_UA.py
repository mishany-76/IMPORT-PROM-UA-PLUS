import gspread
from google.oauth2.service_account import Credentials
import json
import re
import logging
import time
import os  # ДОБАВЛЕНО: для работы с секретами GitHub
from googleapiclient.discovery import build  # Импортируем функцию build для создания сервисного клиента
from gspread.exceptions import WorksheetNotFound, APIError  # Импортируем исключения gspread

# --- Настройка логирования ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- КОНФИГУРАЦИЯ ---
SERVICE_ACCOUNT_FILE = 'key_sheet.json'  

# ИСПРАВЛЕНО: Получаем ID из секретов GitHub (вторым параметром оставлены старые ID как запасные)
SOURCE_SPREADSHEET_ID = os.environ.get('IMPORT_TEMPLATE_PROM_SPREADSHEET_ID', 'IMPORT_TEMPLATE_PROM_SPREADSHEET_ID')
TARGET_SPREADSHEET_ID = os.environ.get('IMPORT_PROM_UA_SPREADSHEET_ID', 'IMPORT_PROM_UA_SPREADSHEET_ID')

# Базовые названия столбцов характеристик без суффикса, как они ДОЛЖНЫ БЫТЬ в целевой таблице,
# если в исходнике найдены характеристики. Используется для формирования структуры и маппинга данных.
BASE_TARGET_CHARACTERISTIC_COLUMNS = [
    'Назва_Характеристики',
    'Одиниця_виміру_Характеристики',
    'Значення_Характеристики',
]

# Столбцы, которые должны присутствовать в целевой таблице
# в дополнение к (столбцам из исходника с учетом правил характеристик).
# Используется только для формирования структуры целевой таблицы.
# ОБНОВЛЕНО: Добавлены новые столбцы согласно требованиям пользователя.
ADDITIONAL_TARGET_COLUMNS = [
    'Ярлик',
    'Де_знаходиться_товар',
    'Тип_товару',
]

# Конфигурация минимального количества пустых строк (новая)
MIN_EMPTY_ROWS = 1000
SHEETS_TO_ENSURE_EMPTY_ROWS = [
    'Export Products Sheet',
    'Export Groups Sheet',
]

# --- КОНФИГУРАЦИЯ КЛЮЧЕВЫХ СТОЛБЦОВ И СОРТИРОВКИ ПО ЛИСТАМ ---
# Определите ключевые столбцы (для идентификации/сравнения) и столбец для сортировки
# для каждого листа, который должен синхронизироваться по данным.
# Ключи словаря - названия листов.
# Значения - словари с ключами 'key_cols' (список имен ключевых столбцов) и 'sort_col' (имя столбца для сортировки исходника, или None).
# Если лист не указан здесь, синхронизация ДАННЫХ для него будет пропущена, только структура скопируется как есть (без правил характеристик).
SHEET_DATA_CONFIG = {
    'Export Products Sheet': {
        'key_cols': ['Код_товару', 'Ідентифікатор_товару'],
        'sort_col': 'Особисті_нотатки',
    },
    'Export Groups Sheet': {
        'key_cols': ['Номер_групи', 'Назва_групи', 'Ідентифікатор_групи', 'Номер_батьківської_групи'],
        'sort_col': None,  # Например, если сортировка для групп не нужна
    },
    # Добавьте сюда другие листы, если они требуют синхронизации данных
    # Для листов не указанных здесь, будет выполнена только синхронизация структуры заголовков.
}


# --- АВТОРИЗАЦИЯ ---
def get_google_sheet_client(credentials):
    """Авторизуется в Google Sheets API и возвращает gspread Client, используя готовые учетные данные."""
    try:
        client = gspread.authorize(credentials)
        logging.info("Авторизация gspread Client успешна.")
        return client
    except Exception as e:
        logging.error(f"Ошибка при авторизации gspread Client: {e}")
        raise


def get_sheets_api_v4_service(credentials):
    """Создает и возвращает клиент Sheets API v4 (googleapiclient), используя готовые учетные данные."""
    try:
        service = build('sheets', 'v4', credentials=credentials)
        logging.info("Клиент Google Sheets API v4 (googleapiclient) успешно создан.")
        return service
    except Exception as e:
        logging.error(f"Ошибка при создании клиента Google Sheets API v4: {e}")
        raise


# --- УТИЛИТЫ ДЛЯ РАБОТЫ С ТАБЛИЦАМИ ---
def get_sheet_by_name(spreadsheet, sheet_name):
    """Получает лист по имени, создает, если не существует."""
    try:
        sheet = spreadsheet.worksheet(sheet_name)
        logging.info(f"Лист '{sheet_name}' найден в таблице '{spreadsheet.title}'.")
        return sheet
    except WorksheetNotFound:
        logging.warning(f"Лист '{sheet_name}' не найден в таблице '{spreadsheet.title}'. Создаем новый...")
        try:
            # gspread создает лист с небольшим начальным размером, при записи он расширится
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
            logging.info(f"Лист '{sheet_name}' успешно создан.")
            return sheet
        except Exception as e:
            logging.error(f"Ошибка при создании листа '{sheet_name}': {e}")
            # При ошибке создания листа - выбрасываем исключение, т.к. на этот лист нужно синхронизировать
            raise


def get_headers(worksheet):
    """Получает заголовки столбцов из первой строки листа."""
    try:
        headers = worksheet.row_values(1)
        if not headers:
            return []
        # Очистка заголовков от возможных пробелов и пустых строк в конце
        headers = [header.strip() for header in headers]
        while headers and not headers[-1]:
            headers.pop()
        return headers
    except Exception as e:
        logging.error(f"Ошибка при чтении заголовков листа '{worksheet.title}': {e}")
        raise


def find_column_index(headers, column_name):
    """Находит индекс столбца по имени (0-based). Возвращает -1, если не найден."""
    try:
        return headers.index(column_name)
    except ValueError:
        return -1


# --- Собственная функция для преобразования индекса столбца в букву ---
def get_column_letter(col_index_0based):
    """Преобразует 0-базовый индекс столбца в его буквенное представление (A, B, AA, etc.)."""
    letter = ''
    while col_index_0based >= 0:
        remainder = col_index_0based % 26
        letter = chr(ord('A') + remainder) + letter
        col_index_0based = col_index_0based // 26 - 1
    return letter


def build_target_headers_with_char_rules(source_headers, sheet_title):
    """
    Формирует список требуемых заголовков для целевой таблицы,
    применяя правила переименования характеристик и сохраняя ОРИГИНАЛЬНЫЙ ОТНОСИТЕЛЬНЫЙ порядок.
    Характеристические блоки добавляются ТОЛЬКО если характеристики найдены в исходном листе.
    """
    max_n = 0
    char_pattern = re.compile(r'^(Назва|Одиниця_виміру|Значення)_Характеристики(?:_(\d+))?$')

    source_has_any_characteristic_headers = False

    # ИСПРАВЛЕНИЕ 1: Первый проход для определения max_n с учетом повторяющихся без суффикса столбцов
    base_char_counts = {}
    for header in source_headers:
        header_stripped = header.strip()
        if not header_stripped:
            continue
        match = char_pattern.match(header_stripped)
        if match:
            source_has_any_characteristic_headers = True
            base_name_part = match.group(1)
            base_name = base_name_part + '_Характеристики'
            n_str = match.group(2)

            if n_str:
                n = int(n_str)
                base_char_counts[base_name] = max(base_char_counts.get(base_name, 0), n)
            else:
                base_char_counts[base_name] = base_char_counts.get(base_name, 0) + 1
                n = base_char_counts[base_name]

            max_n = max(max_n, n)

    final_char_sets_in_target = 0
    if source_has_any_characteristic_headers and BASE_TARGET_CHARACTERISTIC_COLUMNS:
        if max_n == 0:
            final_char_sets_in_target = 1
        else:
            final_char_sets_in_target = max_n

    elif source_has_any_characteristic_headers and not BASE_TARGET_CHARACTERISTIC_COLUMNS:
        logging.warning(
            "В исходнике найдены характеристики, но BASE_TARGET_CHARACTERISTIC_COLUMNS в конфигурации пуст. Блоки характеристик не будут добавлены в целевую структуру для этого листа.")

    final_target_headers = []
    added_char_block_for_set = set()

    # ИСПРАВЛЕНИЕ 2: Второй проход со счетчиками для точного позиционирования
    base_char_counts_pass2 = {}

    for source_index, source_header in enumerate(source_headers):
        header_stripped = source_header.strip()
        if not header_stripped:
            continue

        match = char_pattern.match(header_stripped)

        if match:
            base_name_part = match.group(1)
            base_name = base_name_part + '_Характеристики'
            n_str = match.group(2)

            if n_str:
                n = int(n_str)
                base_char_counts_pass2[base_name] = max(base_char_counts_pass2.get(base_name, 0), n)
            else:
                base_char_counts_pass2[base_name] = base_char_counts_pass2.get(base_name, 0) + 1
                n = base_char_counts_pass2[base_name]

            if n <= final_char_sets_in_target and n not in added_char_block_for_set:
                if BASE_TARGET_CHARACTERISTIC_COLUMNS:
                    final_target_headers.extend(BASE_TARGET_CHARACTERISTIC_COLUMNS)
                    added_char_block_for_set.add(n)

        else:
            final_target_headers.append(header_stripped)

    if sheet_title != 'Export Groups Sheet':
        seen_final = set(final_target_headers)
        for additional_header in ADDITIONAL_TARGET_COLUMNS:
            if additional_header.strip() and additional_header.strip() not in seen_final:
                final_target_headers.append(additional_header.strip())
                seen_final.add(additional_header.strip())
    else:
        logging.info(
            f"Для листа '{sheet_title}' дополнительные столбцы 'Ярлик', 'Де_знаходиться_товар', 'Тип_товару' не будут добавлены.")

    return final_target_headers


def ensure_target_structure(worksheet, source_headers):
    """
    Проверяет и обновляет заголовки целевого листа, чтобы они соответствовали
    структуре, определенной на основе source_headers с правилами характеристик.
    Возвращает актуальный список заголовков целевой таблицы.
    """
    required_target_headers = build_target_headers_with_char_rules(source_headers, worksheet.title)

    logging.info(f"Чтение текущих заголовков целевого листа '{worksheet.title}'...")
    current_target_headers = get_headers(worksheet)

    logging.debug(f"Требуемые заголовки целевого листа: {required_target_headers}")
    logging.debug(f"Текущие заголовки целевого листа: {current_target_headers}")

    if current_target_headers != required_target_headers:
        logging.warning(
            f"Заголовки в целевом листе '{worksheet.title}' не соответствуют требуемой структуре или порядку. Обновляем заголовки...")
        try:
            num_required_cols = len(required_target_headers)
            if num_required_cols == 0:
                logging.warning(
                    "Результирующий список требуемых заголовков оказался пуст. Пропускаем обновление заголовков.")
                current_values = worksheet.get_all_values()
                current_rows = len(current_values)
                if current_rows > 0:
                    try:
                        worksheet.delete_rows(1, 1)
                        logging.info(f"Заголовки в листе '{worksheet.title}' успешно очищены.")
                    except Exception as e:
                        logging.error(f"Ошибка при очистке заголовков в листе '{worksheet.title}': {e}")

                return required_target_headers

            end_column_letter = get_column_letter(num_required_cols - 1)
            range_name = f'A1:{end_column_letter}1'

            logging.info(f"Запись заголовков в целевой лист '{worksheet.title}' в диапазон '{range_name}'...")
            worksheet.update([required_target_headers], range_name=range_name)
            logging.info(f"Заголовки в листе '{worksheet.title}' успешно обновлены.")
            return required_target_headers

        except Exception as e:
            logging.error(f"Ошибка при записи обновленных заголовков в лист '{worksheet.title}': {e}")
            raise

    else:
        logging.info(f"Заголовки в целевом листе '{worksheet.title}' уже соответствуют требуемой структуре.")
        return current_target_headers


def build_column_mapping(source_headers, target_headers):
    """
    Создает словарь соответствия индексов столбцов: source_index -> target_index.
    Маппит исходные характеристики с _N на базовые имена в целевой.
    Маппит остальные столбцы по имени.
    Использует target_headers, сформированные с правилами характеристик.
    """
    mapping = {}
    target_header_to_index = {header: index for index, header in enumerate(target_headers)}

    char_pattern = re.compile(r'^(Назва|Одиниця_виміру|Значення)_Характеристики(?:_(\d+))?$')
    base_char_pattern = re.compile(r'^(Назва|Одиниця_виміру|Значення)_Характеристики$')

    target_base_char_indices = []
    for idx, header in enumerate(target_headers):
        if header in BASE_TARGET_CHARACTERISTIC_COLUMNS:
            target_base_char_indices.append(idx)

    is_structured_char_block = False
    first_char_target_index = -1
    num_target_char_sets = 0

    if target_base_char_indices and BASE_TARGET_CHARACTERISTIC_COLUMNS:
        base_len = len(BASE_TARGET_CHARACTERISTIC_COLUMNS)
        if base_len > 0 and len(target_base_char_indices) > 0 and len(target_base_char_indices) % base_len == 0:
            all_blocks_sequential_and_match = True
            for i in range(0, len(target_base_char_indices), base_len):
                current_block_start_idx_in_target_indices = i
                start_idx_in_target_headers = target_base_char_indices[current_block_start_idx_in_target_indices]

                expected_sequential_indices = list(
                    range(start_idx_in_target_headers, start_idx_in_target_headers + base_len))
                actual_indices_in_this_block = target_base_char_indices[i: i + base_len]

                if actual_indices_in_this_block != expected_sequential_indices:
                    all_blocks_sequential_and_match = False
                    break

                headers_in_block_in_target = []
                for idx in actual_indices_in_this_block:
                    if idx < len(target_headers):
                        headers_in_block_in_target.append(target_headers[idx])
                    else:
                        all_blocks_sequential_and_match = False
                        break

                if headers_in_block_in_target != BASE_TARGET_CHARACTERISTIC_COLUMNS:
                    all_blocks_sequential_and_match = False
                    break

            if all_blocks_sequential_and_match:
                is_structured_char_block = True
                first_char_target_index = target_base_char_indices[0]
                num_target_char_sets = len(target_base_char_indices) // base_len
                logging.debug(
                    f"Обнаружена структурированная область характеристик в целевых заголовках. Начало: {first_char_target_index}, Наборов: {num_target_char_sets}")
            else:
                logging.warning(
                    "Базовые названия характеристик найдены в целевых заголовках, но структура блоков нарушена. Маппинг характеристик может быть некорректен.")

        else:
            logging.warning(
                "Количество базовых названий характеристик в целевых заголовках не кратно размеру базового набора, или размер базового набора 0. Маппинг характеристик может быть некорректен.")

    else:
        pass

        # ИСПРАВЛЕНИЕ 3: Привязываем маппинг с учетом повторяющихся колонок без индексов
    base_char_counts_map = {}

    for source_index, source_header in enumerate(source_headers):
        source_header_stripped = source_header.strip()
        if not source_header_stripped:
            continue

        match = char_pattern.match(source_header_stripped)
        if match:
            base_name_part = match.group(1)
            base_name = base_name_part + '_Характеристики'
            n_str = match.group(2)

            if n_str:
                n = int(n_str)
                base_char_counts_map[base_name] = max(base_char_counts_map.get(base_name, 0), n)
            else:
                base_char_counts_map[base_name] = base_char_counts_map.get(base_name, 0) + 1
                n = base_char_counts_map[base_name]

            if is_structured_char_block and base_name in BASE_TARGET_CHARACTERISTIC_COLUMNS and n <= num_target_char_sets:
                try:
                    base_index_in_set = BASE_TARGET_CHARACTERISTIC_COLUMNS.index(base_name)
                    target_index = first_char_target_index + (n - 1) * len(
                        BASE_TARGET_CHARACTERISTIC_COLUMNS) + base_index_in_set

                    if target_index < len(target_headers) and target_headers[target_index] == base_name:
                        mapping[source_index] = target_index
                    else:
                        logging.warning(
                            f"Рассчитанный целевой индекс {target_index} для '{source_header_stripped}' не соответствует базовому имени в целевых заголовках или выходит за пределы. Пропущено маппинг.")

                except ValueError:
                    logging.error(
                        f"Критическая ошибка маппинга: Базовое имя характеристики '{base_name}' не найдено в BASE_TARGET_CHARACTERISTIC_COLUMNS при маппинге исходного '{source_header_stripped}'.")
            elif source_header_stripped in target_header_to_index:
                if base_char_pattern.match(source_header_stripped) and n == 1:
                    mapping[source_index] = target_header_to_index[source_header_stripped]

        else:
            if source_header_stripped in target_header_to_index:
                mapping[source_index] = target_header_to_index[source_header_stripped]

    return mapping


def synchronize_single_sheet_with_data(gspread_client, sheets_service, source_sheet, target_spreadsheet):
    """Синхронизирует данные и структуру с одного листа исходной таблицы (на месте)."""
    source_sheet_title = source_sheet.title
    logging.info(f"--- Начинаем синхронизацию для листа: '{source_sheet_title}' ---")

    try:
        # Получаем или создаем целевой лист с тем же именем через gspread_client
        target_sheet = get_sheet_by_name(target_spreadsheet, source_sheet_title)

        # --- Чтение и анализ заголовков ---
        logging.info(f"Чтение заголовков исходного листа '{source_sheet_title}'...")
        source_headers = get_headers(source_sheet)
        if not source_headers:
            logging.warning(
                f"Исходный лист '{source_sheet_title}' пуст или не содержит заголовков. Пропускаем синхронизацию для этого листа.")
            # Если исходный лист пуст или без заголовков, мы все равно создадим целевой лист, но оставим его без заголовков и без данных.
            logging.info(f"--- Синхронизация для листа '{source_sheet_title}' завершена (исходник без заголовков) ---")
            return

        # --- Проверка и обновление структуры целевой таблицы (заголовков) ---
        # Обновляем структуру независимо от наличия конфигурации данных
        updated_target_headers = ensure_target_structure(target_sheet, source_headers)

        # --- Определение ключевых столбцов и столбца сортировки для этого листа ---
        sheet_config = SHEET_DATA_CONFIG.get(source_sheet_title)

        if not sheet_config:
            logging.info(
                f"Для листа '{source_sheet_title}' нет конфигурации данных в SHEET_DATA_CONFIG. Синхронизация данных пропущена.")
            logging.info(f"--- Синхронизация для листа '{source_sheet_title}' завершена (нет конфига данных) ---")
            return

        key_cols_names = sheet_config.get('key_cols', [])
        sort_col_name = sheet_config.get('sort_col')

        if not key_cols_names:
            logging.warning(
                f"В конфигурации для листа '{source_sheet_title}' не указаны ключевые столбцы ('key_cols'). Синхронизация данных пропущена.")
            logging.info(f"--- Синхронизация для листа '{source_sheet_title}' завершена (нет key_cols в конфиге) ---")
            return

        # Получаем индексы ключевых столбцов в исходнике и целевой
        source_key_indices = []
        missing_source_key_cols = []
        for col_name in key_cols_names:
            idx = find_column_index(source_headers, col_name)
            if idx == -1:
                missing_source_key_cols.append(col_name)
            source_key_indices.append(idx)  # Индекс может быть -1

        target_key_indices = []
        missing_target_key_cols = []
        for col_name in key_cols_names:
            # Ищем индекс в обновленных заголовках целевой таблицы
            idx = find_column_index(updated_target_headers, col_name)
            if idx == -1:
                missing_target_key_cols.append(col_name)
            target_key_indices.append(idx)  # Индекс может быть -1

        if missing_source_key_cols:
            logging.error(
                f"Ключевые столбцы для обработки данных отсутствуют в исходном листе '{source_sheet_title}': {missing_source_key_cols}. Синхронизация данных пропущена.")
            logging.info(f"--- Синхронизация для листа '{source_sheet_title}' завершена (нет key_cols в исходнике) ---")
            return  # Нет ключевых колонок в исходнике, не можем синхронизировать данные

        if missing_target_key_cols:
            # Этого не должно произойти, если ensure_target_structure отработал корректно
            # и build_target_headers_with_char_rules включил эти колонки, если они были в исходнике.
            logging.error(
                f"Ключевые столбцы для обработки данных отсутствуют в целевом листе '{source_sheet_title}' после обновления структуры: {missing_target_key_cols}. Критическая ошибка. Синхронизация данных пропущена.")
            return  # Нет ключевых колонок в целевой, не можем синхронизировать данные

        # Получаем индекс столбца сортировки в исходнике
        source_sort_col_index = find_column_index(source_headers, sort_col_name) if sort_col_name else -1

        # Создаем маппинг столбцов из исходной в целевую
        logging.info("Построение маппинга столбцов между исходным и целевым листами...")
        col_mapping = build_column_mapping(source_headers, updated_target_headers)

        # --- Чтение данных ---
        logging.info(f"Чтение данных из исходного листа '{source_sheet_title}'...")
        source_data = source_sheet.get_all_values()[1:]  # Пропускаем заголовки
        logging.info(f"Прочитано {len(source_data)} строк из исходного листа.")

        if not source_data:
            logging.warning(
                f"В исходном листе '{source_sheet_title}' нет данных. Будет выполнено только удаление устаревших/дублирующихся строк из целевого листа.")

        logging.info(f"Чтение данных из целевого листа '{source_sheet_title}'...")
        target_data = target_sheet.get_all_values()[1:]  # Пропускаем заголовки
        logging.info(f"Прочитано {len(target_data)} строк из целевого листа.")

        if not source_data and not target_data:
            logging.info(
                f"Нет данных ни в исходном, ни в целевом листе '{source_sheet_title}'. Синхронизация данных не требуется.")
            logging.info(f"--- Синхронизация для листа '{source_sheet_title}' завершена (нет данных) ---")
            return

        # --- Сортировка исходных данных ---
        if sort_col_name and source_sort_col_index != -1:
            source_data.sort(key=lambda row: row[source_sort_col_index].strip() if len(
                row) > source_sort_col_index else '')  # Сортируем по значению, удаляя пробелы
            logging.info(f"Исходные данные отсортированы по '{sort_col_name}'.")

        elif sort_col_name and source_sort_col_index == -1:
            logging.warning(
                f"Столбец сортировки '{sort_col_name}' указан в конфиге, но отсутствует в исходнике. Сортировка исходных данных пропущена.")

        # --- Подготовка данных для синхронизации на месте и идентификация дубликатов в целевой ---
        target_product_map = {}  # Key: product_key tuple, Value: target_row_index (0-based) первой встречи
        target_row_data_by_index = {}  # Key: target_row_index (0-based), Value: target_row_data list (оригинал из чтения, дополненный до нужной длины)

        rows_to_delete_indices_1based = []

        max_key_idx_target = max(target_key_indices) if target_key_indices else -1
        required_cols_len_target = len(updated_target_headers)

        seen_target_keys = set()

        for row_index_0based, row in enumerate(target_data):
            target_row_index_1based = row_index_0based + 2

            target_key_values = []
            has_any_key_value = False
            if len(row) > max_key_idx_target:
                for key_idx in target_key_indices:
                    key_value = row[key_idx].strip() if key_idx != -1 and key_idx < len(row) else ''
                    target_key_values.append(key_value)
                    if key_value:
                        has_any_key_value = True

            product_key = tuple(target_key_values)

            if not has_any_key_value:
                logging.debug(
                    f"Строка {target_row_index_1based} в целевой таблице не имеет валидного ключа. Помечена на удаление.")
                rows_to_delete_indices_1based.append(target_row_index_1based)
            elif product_key in seen_target_keys:
                logging.debug(
                    f"Строка {target_row_index_1based} в целевой таблице является дубликатом ключа {product_key}. Помечена на удаление.")
                rows_to_delete_indices_1based.append(target_row_index_1based)
            else:
                seen_target_keys.add(product_key)
                target_product_map[product_key] = row_index_0based
                processed_row_data = list(row)
                if len(processed_row_data) < required_cols_len_target:
                    processed_row_data.extend([''] * (required_cols_len_target - len(processed_row_data)))
                elif len(processed_row_data) > required_cols_len_target:
                    processed_row_data = processed_row_data[:required_cols_len_target]
                target_row_data_by_index[row_index_0based] = processed_row_data

        logging.info(
            f"Идентификация дубликатов и строк без ключа в целевой таблице завершена. Найдено {len([idx for idx in rows_to_delete_indices_1based if idx <= len(target_data) + 1])} строк для удаления на этом этапе.")

        # --- Идентификация операций (Обновления, Добавления) ---
        update_operations_batch_format = []
        source_product_keys_seen = set()
        new_rows_to_add_data = []

        max_key_idx_source = max(source_key_indices) if source_key_indices else -1
        source_num_cols = len(source_headers)

        availability_col_source_idx = find_column_index(source_headers, 'Наявність')
        supplier_col_source_idx = find_column_index(source_headers, 'Особисті_нотатки')
        quantity_col_source_idx = find_column_index(source_headers, 'Кількість')

        label_col_target_idx = find_column_index(updated_target_headers, 'Ярлик')
        location_col_target_idx = find_column_index(updated_target_headers, 'Де_знаходиться_товар')
        type_col_target_idx = find_column_index(updated_target_headers, 'Тип_товару')

        # Индексы обязательных полей для подстановки fallback (UA -> RU)
        name_ru_target_idx = find_column_index(updated_target_headers, 'Назва_позиції')
        name_ua_target_idx = find_column_index(updated_target_headers, 'Назва_позиції_укр')
        desc_ru_target_idx = find_column_index(updated_target_headers, 'Опис')
        desc_ua_target_idx = find_column_index(updated_target_headers, 'Опис_укр')

        # Индексы поисковых запросов для очистки в Легком Эшелоне
        kw_ru_target_idx = find_column_index(updated_target_headers, 'Пошукові_запити')
        kw_ua_target_idx = find_column_index(updated_target_headers, 'Пошукові_запити_укр')

        if supplier_col_source_idx == -1:
            logging.warning(
                "Столбец 'Особисті_нотатки' не найден в исходной таблице. Функции 'Ярлик', 'Де_знаходиться_товар' и 'Тип_товару' не будут работать корректно.")

        for source_row_index, source_row in enumerate(source_data):
            source_row_1based = source_row_index + 2

            if len(source_row) <= max_key_idx_source:
                continue

            source_key_values = []
            has_any_key_value_source = False
            for key_idx in source_key_indices:
                key_value = source_row[key_idx].strip() if key_idx != -1 and key_idx < len(source_row) else ''
                source_key_values.append(key_value)
                if key_value:
                    has_any_key_value_source = True

            if not has_any_key_value_source:
                continue

            product_key = tuple(source_key_values)
            source_product_keys_seen.add(product_key)

            current_supplier = ''
            if supplier_col_source_idx != -1 and supplier_col_source_idx < len(source_row):
                current_supplier = source_row[supplier_col_source_idx].strip().upper()

            if product_key in target_product_map:
                target_row_index_0based = target_product_map[product_key]
                target_row_data = target_row_data_by_index[target_row_index_0based]
                target_row_number_1based = target_row_index_0based + 2

                for source_col_index in range(source_num_cols):
                    if source_col_index in col_mapping:
                        target_col_index_0based = col_mapping[source_col_index]

                        if source_col_index < len(source_row) and target_col_index_0based < len(target_row_data):
                            source_value = source_row[source_col_index]

                            if source_col_index == availability_col_source_idx:
                                if source_value.strip().upper() == 'TRUE' or source_value.strip() == '+':
                                    source_value = '!'

                            if source_col_index == quantity_col_source_idx:
                                if source_value != target_row_data[target_col_index_0based]:
                                    target_col_letter = get_column_letter(target_col_index_0based)
                                    cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                                    update_operations_batch_format.append({
                                        'range': cell_range_a1,
                                        'values': [[source_value]]
                                    })
                            else:
                                if source_value.strip() and source_value != target_row_data[target_col_index_0based]:
                                    target_col_letter = get_column_letter(target_col_index_0based)
                                    cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                                    update_operations_batch_format.append({
                                        'range': cell_range_a1,
                                        'values': [[source_value]]
                                    })

                if source_sheet_title != 'Export Groups Sheet':
                    if label_col_target_idx != -1 and current_supplier != 'FOOTBALLERS':
                        new_label_value = 'Топ продажів'
                        if target_row_data[label_col_target_idx] != new_label_value:
                            target_col_letter = get_column_letter(label_col_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [[new_label_value]]
                            })

                    if location_col_target_idx != -1 and current_supplier != 'FOOTBALLERS':
                        location_map = {
                            'AGER': 'Одеса',
                            'IZIDROP': 'Одеса',
                            'MOYDROP': 'Одеса',
                            'SPECULANT': 'Полтава',
                            'KIRS': 'Дніпро',
                            'BAGSROOM': 'Київ',
                            'BAYRAKTAR': 'с.Шаровечка, Хмельницька область',
                        }
                        new_location_value = location_map.get(current_supplier, '')
                        if target_row_data[location_col_target_idx] != new_location_value:
                            target_col_letter = get_column_letter(location_col_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [[new_location_value]]
                            })

                    if type_col_target_idx != -1 and current_supplier != 'FOOTBALLERS':
                        new_type_value = 'r'
                        if target_row_data[type_col_target_idx] != new_type_value:
                            target_col_letter = get_column_letter(type_col_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [[new_type_value]]
                            })

                    # Fallback для обязательных полей Prom (если русское пустое -> копируем украинское)
                    if name_ru_target_idx != -1 and name_ua_target_idx != -1:
                        ru_val = target_row_data[name_ru_target_idx] if name_ru_target_idx < len(target_row_data) else ''
                        ua_val = target_row_data[name_ua_target_idx] if name_ua_target_idx < len(target_row_data) else ''
                        if not ru_val.strip() and ua_val.strip():
                            target_col_letter = get_column_letter(name_ru_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [[ua_val]]
                            })

                    if desc_ru_target_idx != -1 and desc_ua_target_idx != -1:
                        ru_val = target_row_data[desc_ru_target_idx] if desc_ru_target_idx < len(target_row_data) else ''
                        ua_val = target_row_data[desc_ua_target_idx] if desc_ua_target_idx < len(target_row_data) else ''
                        if not ru_val.strip() and ua_val.strip():
                            target_col_letter = get_column_letter(desc_ru_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [[ua_val]]
                            })

                    # Очистка поисковых запросов в легком эшелоне (если там что-то осталось)
                    if kw_ru_target_idx != -1 and kw_ru_target_idx < len(target_row_data):
                        if target_row_data[kw_ru_target_idx].strip():
                            target_col_letter = get_column_letter(kw_ru_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [['']]
                            })
                    if kw_ua_target_idx != -1 and kw_ua_target_idx < len(target_row_data):
                        if target_row_data[kw_ua_target_idx].strip():
                            target_col_letter = get_column_letter(kw_ua_target_idx)
                            cell_range_a1 = f"{source_sheet_title}!{target_col_letter}{target_row_number_1based}"
                            update_operations_batch_format.append({
                                'range': cell_range_a1,
                                'values': [['']]
                            })


            else:
                new_row = [''] * required_cols_len_target
                for source_col_index in range(source_num_cols):
                    if source_col_index in col_mapping:
                        target_col_index_0based = col_mapping[source_col_index]
                        if source_col_index < len(source_row) and target_col_index_0based < len(new_row):
                            source_value = source_row[source_col_index]

                            if source_col_index == availability_col_source_idx:
                                if source_value.strip().upper() == 'TRUE' or source_value.strip() == '+':
                                    source_value = '!'

                            if source_value.strip():
                                new_row[target_col_index_0based] = source_value

                if source_sheet_title != 'Export Groups Sheet':
                    if label_col_target_idx != -1 and current_supplier != 'FOOTBALLERS':
                        new_row[label_col_target_idx] = 'Топ продажів'

                    if location_col_target_idx != -1 and current_supplier != 'FOOTBALLERS':
                        location_map = {
                            'AGER': 'Одеса',
                            'IZIDROP': 'Одеса',
                            'MOYDROP': 'Одеса',
                            'SPECULANT': 'Полтава',
                            'KIRS': 'Дніпро',
                            'BAGSROOM': 'Київ',
                            'BAYRAKTAR': 'с.Шаровечка, Хмельницька область',
                        }
                        new_row[location_col_target_idx] = location_map.get(current_supplier, '')

                    if type_col_target_idx != -1 and current_supplier != 'FOOTBALLERS':
                        new_row[type_col_target_idx] = 'r'

                    # Fallback для обязательных полей Prom (если русское пустое -> копируем украинское)
                    if name_ru_target_idx != -1 and name_ua_target_idx != -1:
                        if not new_row[name_ru_target_idx].strip() and new_row[name_ua_target_idx].strip():
                            new_row[name_ru_target_idx] = new_row[name_ua_target_idx]

                    if desc_ru_target_idx != -1 and desc_ua_target_idx != -1:
                        if not new_row[desc_ru_target_idx].strip() and new_row[desc_ua_target_idx].strip():
                            new_row[desc_ru_target_idx] = new_row[desc_ua_target_idx]

                    # Очистка поисковых запросов в легком эшелоне
                    if kw_ru_target_idx != -1 and kw_ru_target_idx < len(new_row):
                        new_row[kw_ru_target_idx] = ''
                    if kw_ua_target_idx != -1 and kw_ua_target_idx < len(new_row):
                        new_row[kw_ua_target_idx] = ''

                new_rows_to_add_data.append(new_row)

        logging.info(
            f"Идентификация операций завершена. Найдено {len(update_operations_batch_format)} обновлений ячеек, {len(new_rows_to_add_data)} добавлений.")

        for product_key, target_row_index_0based in target_product_map.items():
            if product_key not in source_product_keys_seen:
                target_row_index_1based = target_row_index_0based + 2
                logging.debug(
                    f"Уникальный ключ из целевой {product_key} не найден в исходнике. Строка {target_row_index_1based} помечена на удаление.")
                rows_to_delete_indices_1based.append(target_row_index_1based)

        logging.info(
            f"Найдено всего {len(rows_to_delete_indices_1based)} строк для удаления (включая дубликаты в целевой, строки без ключа и отсутствующие в исходнике).")

        BATCH_CHUNK_SIZE = 10000

        if update_operations_batch_format:
            total_ops = len(update_operations_batch_format)
            total_chunks = (total_ops + BATCH_CHUNK_SIZE - 1) // BATCH_CHUNK_SIZE
            logging.info(
                f"Отправка пакетных обновлений: {total_ops} ячеек, "
                f"разбито на {total_chunks} чанков по {BATCH_CHUNK_SIZE}...")

            batch_max_retries = 5
            batch_retry_delay = 10

            all_chunks_ok = True
            for chunk_idx in range(total_chunks):
                chunk_start = chunk_idx * BATCH_CHUNK_SIZE
                chunk_end = min(chunk_start + BATCH_CHUNK_SIZE, total_ops)
                chunk_data = update_operations_batch_format[chunk_start:chunk_end]

                logging.info(
                    f"  Чанк {chunk_idx + 1}/{total_chunks}: "
                    f"операции {chunk_start + 1}-{chunk_end}...")

                chunk_ok = False
                for attempt in range(batch_max_retries):
                    try:
                        body = {
                            'value_input_option': 'USER_ENTERED',
                            'data': chunk_data
                        }
                        sheets_service.spreadsheets().values().batchUpdate(
                            spreadsheetId=TARGET_SPREADSHEET_ID,
                            body=body
                        ).execute()
                        logging.info(
                            f"  Чанк {chunk_idx + 1}/{total_chunks} выполнен успешно.")
                        chunk_ok = True
                        break

                    except Exception as e:
                        is_retryable = (
                                (hasattr(e, 'resp') and e.resp.status in [429, 500, 502, 503, 504]) or
                                ("temporarily unavailable" in str(e).lower()) or
                                ("bad gateway" in str(e).lower()) or
                                ("backend error" in str(e).lower()) or
                                ("timed out" in str(e).lower()) or
                                ("-1" in str(e))
                        )
                        if attempt < batch_max_retries - 1 and is_retryable:
                            logging.warning(
                                f"  API Error при чанке {chunk_idx + 1}, "
                                f"попытка {attempt + 1}: {e}. "
                                f"Повторная попытка через {batch_retry_delay} сек...")
                            time.sleep(batch_retry_delay)
                        else:
                            logging.error(
                                f"  Критическая ошибка при чанке {chunk_idx + 1} "
                                f"после {attempt + 1} попыток: {e}")
                            raise

                if not chunk_ok:
                    logging.error(
                        f"  Не удалось выполнить чанк {chunk_idx + 1}/{total_chunks}.")
                    all_chunks_ok = False

                if chunk_idx < total_chunks - 1:
                    time.sleep(1)

            if all_chunks_ok:
                logging.info(
                    f"Все {total_chunks} чанков ({total_ops} ячеек) успешно обновлены.")
            else:
                logging.warning(
                    f"Пакетное обновление завершено с ошибками. "
                    f"Проверьте лог выше для деталей.")

        else:
            logging.info("Нет операций обновления ячеек для выполнения.")

        if rows_to_delete_indices_1based:
            logging.info(f"Выполнение {len(rows_to_delete_indices_1based)} удалений строк...")
            rows_to_delete_indices_1based.sort(reverse=True)

            delete_max_retries = 5
            delete_retry_delay = 5

            for row_index_1based in rows_to_delete_indices_1based:
                for delete_attempt in range(delete_max_retries):
                    try:
                        logging.debug(f"Попытка удаления строки {row_index_1based}...")
                        target_sheet.delete_rows(row_index_1based)
                        logging.debug(f"Удалена строка {row_index_1based}.")
                        time.sleep(1.1)
                        break
                    except gspread.exceptions.APIError as e:
                        if delete_attempt < delete_max_retries - 1 and (
                                '429' in str(e) or '500' in str(e) or '502' in str(e) or '503' in str(e) or '-1' in str(
                            e) or 'temporarily unavailable' in str(e).lower() or 'bad gateway' in str(
                            e).lower() or 'backend error' in str(e).lower()):
                            logging.warning(
                                f"API Error (вероятно временная) при удалении строки {row_index_1based}: {e}. Повторная попытка через {delete_retry_delay} сек...")
                            time.sleep(delete_retry_delay)
                        else:
                            logging.error(
                                f"Критическая ошибка API при удалении строки {row_index_1based} после {delete_attempt + 1} попыток: {e}")
                            break

                    except Exception as e:
                        logging.error(f"Неожиданная ошибка при удалении строки {row_index_1based}: {e}")
                        break

            else:
                logging.error(
                    f"Не удалось удалить строку после {delete_max_retries} попыток для листа '{source_sheet_title}'.")

        logging.info("Выполнение удалений завершено.")

        if new_rows_to_add_data:
            logging.info(f"Выполнение {len(new_rows_to_add_data)} добавлений строк...")

            max_retries_append = 5
            retry_delay_seconds_append = 5

            for attempt in range(max_retries_append):
                try:
                    logging.info(f"Попытка записи новых данных {attempt + 1}/{max_retries_append}...")
                    required_cols_len = len(updated_target_headers)
                    data_to_write = []
                    for row in new_rows_to_add_data:
                        processed_row = list(row)
                        if len(processed_row) < required_cols_len:
                            processed_row.extend([''] * (required_cols_len - len(processed_row)))
                        elif len(processed_row) > required_cols_len:
                            processed_row = processed_row[:required_cols_len]
                        data_to_write.append(processed_row)

                    if data_to_write:
                        target_sheet.append_rows(data_to_write,
                                                 value_input_option='USER_ENTERED')
                        logging.info(
                            f"Успешно записано {len(data_to_write)} новых строк в целевой лист '{source_sheet_title}'.")
                    else:
                        logging.info(f"Нет данных для добавления в целевой лист '{source_sheet_title}'.")

                    break

                except gspread.exceptions.APIError as e:
                    if attempt < max_retries_append - 1 and (
                            '429' in str(e) or '500' in str(e) or '502' in str(e) or '503' in str(e) or '-1' in str(
                        e) or 'temporarily unavailable' in str(e).lower() or 'bad gateway' in str(
                        e).lower() or 'backend error' in str(e).lower()):
                        logging.warning(
                            f"API Error (вероятно временная) при добавлении строк: {e}. Повторная попытка через {retry_delay_seconds_append} сек...")
                        time.sleep(retry_delay_seconds_append)
                    else:
                        logging.error(f"Критическая ошибка API при добавлении строк после {attempt + 1} попыток: {e}")
                        raise

                except Exception as e:
                    logging.error(f"Неожиданная ошибка при добавлении строк: {e}")
                    raise

            else:
                logging.error(
                    f"Не удалось добавить строки после {max_retries_append} попыток для листа '{source_sheet_title}'.")

        logging.info(f"--- Синхронизация для листа '{source_sheet_title}' завершена ---")

    except Exception as e:
        logging.error(f"Произошла ошибка при синхронизации листа '{source_sheet_title}'. Error: {e}")
        import traceback
        logging.error(traceback.format_exc())


# --------------------------------------------------------------------------------------------------------------------

def apply_price_discount(sheets_service, source_spreadsheet, target_spreadsheet, sheet_name='Export Products Sheet'):
    """
    Отдельная функция для обработки цен и скидок.
    Запускается ПОСЛЕ основной синхронизации.
    """
    PRICE_CHUNK_SIZE = 10000

    logging.info(f"--- ЗАПУСК ОБРАБОТКИ ЦЕН/СКИДОК для листа '{sheet_name}' ---")
    try:
        try:
            source_sheet = source_spreadsheet.worksheet(sheet_name)
        except WorksheetNotFound:
            logging.warning(
                f"Лист '{sheet_name}' не найден в исходной таблице. "
                f"Обработка цен/скидок пропущена.")
            return

        try:
            target_sheet = target_spreadsheet.worksheet(sheet_name)
        except WorksheetNotFound:
            logging.warning(
                f"Лист '{sheet_name}' не найден в целевой таблице. "
                f"Обработка цен/скидок пропущена.")
            return

        source_headers = get_headers(source_sheet)
        target_headers = get_headers(target_sheet)

        if not source_headers or not target_headers:
            logging.warning(
                f"Заголовки не найдены в одном из листов '{sheet_name}'. "
                f"Обработка цен/скидок пропущена.")
            return

        src_price_idx = find_column_index(source_headers, 'Ціна')
        src_price_from_idx = find_column_index(source_headers, 'Ціна_від')

        if src_price_idx == -1 or src_price_from_idx == -1:
            logging.warning(
                f"Столбцы 'Ціна' и/или 'Ціна_від' не найдены в исходном листе '{sheet_name}'. "
                f"Обработка цен/скидок пропущена.")
            return

        tgt_price_idx = find_column_index(target_headers, 'Ціна')
        tgt_price_from_idx = find_column_index(target_headers, 'Ціна_від')
        tgt_discount_idx = find_column_index(target_headers, 'Знижка')

        if tgt_price_idx == -1:
            logging.warning(
                f"Столбец 'Ціна' не найден в целевом листе '{sheet_name}'. "
                f"Обработка цен/скидок пропущена.")
            return

        logging.info(
            f"Индексы исходника: Ціна={src_price_idx}, Ціна_від={src_price_from_idx}. "
            f"Индексы целевой:  Ціна={tgt_price_idx}, Ціна_від={tgt_price_from_idx}, "
            f"Знижка={tgt_discount_idx}.")

        logging.info(f"Чтение данных исходного листа '{sheet_name}' для обработки цен...")
        source_data = source_sheet.get_all_values()[1:]
        logging.info(f"Прочитано {len(source_data)} строк из исходного листа.")

        if not source_data:
            logging.info(f"Нет данных для обработки цен в листе '{sheet_name}'.")
            return

        key_cols = SHEET_DATA_CONFIG.get(sheet_name, {}).get('key_cols', [])
        if not key_cols:
            logging.warning(
                f"Ключевые столбцы не определены для листа '{sheet_name}' в SHEET_DATA_CONFIG. "
                f"Обработка цен/скидок пропущена.")
            return

        src_key_indices = [find_column_index(source_headers, c) for c in key_cols]
        tgt_key_indices = [find_column_index(target_headers, c) for c in key_cols]

        if -1 in src_key_indices or -1 in tgt_key_indices:
            logging.warning(
                f"Не все ключевые столбцы найдены. "
                f"src_key_indices={src_key_indices}, tgt_key_indices={tgt_key_indices}. "
                f"Обработка цен/скидок пропущена.")
            return

        logging.info(f"Чтение данных целевого листа '{sheet_name}' для построения карты строк...")
        target_data = target_sheet.get_all_values()[1:]
        logging.info(f"Прочитано {len(target_data)} строк из целевого листа.")

        target_row_map = {}
        for row_idx, row in enumerate(target_data):
            key = tuple(
                row[i].strip() if i < len(row) else ''
                for i in tgt_key_indices
            )
            if any(k for k in key) and key not in target_row_map:
                target_row_map[key] = row_idx + 2

        price_updates = []
        rows_updated = 0  # Изменил название переменной, так как теперь мы считаем и обновления без скидок

        for src_row in source_data:
            src_price = src_row[src_price_idx].strip() if src_price_idx < len(src_row) else ''
            src_price_from = src_row[src_price_from_idx].strip() if src_price_from_idx < len(src_row) else ''

            # --- ИЗМЕНЕНИЕ 1: Отмена скидки ---
            if not src_price:
                continue  # Если базовой цены нет вообще, пропускаем

            def parse_price(val):
                val = re.sub(r'\s+', '', val)
                val = val.replace(',', '.')
                return float(val)

            # По умолчанию считаем, что скидки нет.
            # В таком случае базовая цена = цена, цена от = цена, скидка = пусто.
            val_price = src_price
            val_price_from = src_price
            discount_str = ""

            # Если колонка "Ціна від" не пустая, проверяем, есть ли реальная разница
            if src_price_from:
                try:
                    p_base = parse_price(src_price)
                    p_from = parse_price(src_price_from)

                    if p_base != p_from:
                        if p_base > p_from:
                            val_price = src_price
                            val_price_from = src_price_from
                            discount = p_base - p_from
                        else:
                            val_price = src_price_from
                            val_price_from = src_price
                            discount = p_from - p_base

                        discount_str = str(round(discount, 2))
                        if discount_str.endswith('.0'):
                            discount_str = discount_str[:-2]
                except ValueError:
                    continue  # Ошибка при парсинге цены - игнорируем строку, как и было раньше

            src_key = tuple(
                src_row[i].strip() if i < len(src_row) else ''
                for i in src_key_indices
            )
            tgt_row_num = target_row_map.get(src_key)
            if tgt_row_num is None:
                continue

            # --- ИЗМЕНЕНИЕ 2: Защита от «холостых» обновлений (проверка всех 3 колонок) ---
            # Достаем строку из target_data (индекс равен номеру строки минус 2, т.к. 1-базовый и заголовок)
            tgt_idx = tgt_row_num - 2
            tgt_row = target_data[tgt_idx]

            # Получаем то, что сейчас реально записано в целевой таблице (если колонка вообще существует)
            current_tgt_price = tgt_row[tgt_price_idx].strip() if tgt_price_idx != -1 and tgt_price_idx < len(tgt_row) else ''
            current_tgt_price_from = tgt_row[tgt_price_from_idx].strip() if tgt_price_from_idx != -1 and tgt_price_from_idx < len(tgt_row) else ''
            current_tgt_discount = tgt_row[tgt_discount_idx].strip() if tgt_discount_idx != -1 and tgt_discount_idx < len(tgt_row) else ''

            needs_update = False

            # Проверяем колонку "Ціна"
            if tgt_price_idx != -1 and current_tgt_price != val_price:
                tgt_price_col_letter = get_column_letter(tgt_price_idx)
                price_updates.append({
                    'range': f"{sheet_name}!{tgt_price_col_letter}{tgt_row_num}",
                    'values': [[val_price]]
                })
                needs_update = True

            # Проверяем колонку "Ціна_від"
            if tgt_price_from_idx != -1 and current_tgt_price_from != val_price_from:
                tgt_price_from_col_letter = get_column_letter(tgt_price_from_idx)
                price_updates.append({
                    'range': f"{sheet_name}!{tgt_price_from_col_letter}{tgt_row_num}",
                    'values': [[val_price_from]]
                })
                needs_update = True

            # Проверяем колонку "Знижка"
            if tgt_discount_idx != -1 and current_tgt_discount != discount_str:
                tgt_discount_col_letter = get_column_letter(tgt_discount_idx)
                price_updates.append({
                    'range': f"{sheet_name}!{tgt_discount_col_letter}{tgt_row_num}",
                    'values': [[discount_str]]
                })
                needs_update = True

            # Увеличиваем счетчик только если товар реально требует обновления
            if needs_update:
                rows_updated += 1

        logging.info(
            f"Найдено {rows_updated} товаров, требующих обновления цен/скидок. "
            f"Подготовлено {len(price_updates)} операций обновления ячеек.")

        if not price_updates:
            logging.info("Нет операций для обновления цен/скидок. Данные актуальны.")
            logging.info(f"--- Обработка цен/скидок для листа '{sheet_name}' завершена ---")
            return

        total_ops = len(price_updates)
        total_chunks = (total_ops + PRICE_CHUNK_SIZE - 1) // PRICE_CHUNK_SIZE
        logging.info(
            f"Отправка обновлений цен/скидок: {total_ops} операций, "
            f"{total_chunks} чанков по {PRICE_CHUNK_SIZE}...")

        batch_max_retries = 5
        batch_retry_delay = 10
        all_chunks_ok = True

        for chunk_idx in range(total_chunks):
            chunk_start = chunk_idx * PRICE_CHUNK_SIZE
            chunk_end = min(chunk_start + PRICE_CHUNK_SIZE, total_ops)
            chunk_data = price_updates[chunk_start:chunk_end]

            logging.info(
                f"  Цены: чанк {chunk_idx + 1}/{total_chunks}, "
                f"операции {chunk_start + 1}-{chunk_end}...")

            chunk_ok = False
            for attempt in range(batch_max_retries):
                try:
                    body = {
                        'value_input_option': 'USER_ENTERED',
                        'data': chunk_data
                    }
                    sheets_service.spreadsheets().values().batchUpdate(
                        spreadsheetId=TARGET_SPREADSHEET_ID,
                        body=body
                    ).execute()
                    logging.info(
                        f"  Цены: чанк {chunk_idx + 1}/{total_chunks} выполнен успешно.")
                    chunk_ok = True
                    break

                except Exception as e:
                    is_retryable = (
                            (hasattr(e, 'resp') and e.resp.status in [429, 500, 502, 503, 504]) or
                            ("temporarily unavailable" in str(e).lower()) or
                            ("bad gateway" in str(e).lower()) or
                            ("backend error" in str(e).lower()) or
                            ("timed out" in str(e).lower()) or
                            ("-1" in str(e))
                    )
                    if attempt < batch_max_retries - 1 and is_retryable:
                        logging.warning(
                            f"  API Error при чанке цен {chunk_idx + 1}, "
                            f"попытка {attempt + 1}: {e}. "
                            f"Повторная попытка через {batch_retry_delay} сек...")
                        time.sleep(batch_retry_delay)
                    else:
                        logging.error(
                            f"  Критическая ошибка при чанке цен {chunk_idx + 1} "
                            f"после {attempt + 1} попыток: {e}")
                        raise

            if not chunk_ok:
                logging.error(
                    f"  Не удалось выполнить чанк цен {chunk_idx + 1}/{total_chunks}.")
                all_chunks_ok = False

            if chunk_idx < total_chunks - 1:
                time.sleep(1)

        if all_chunks_ok:
            logging.info(
                f"Обработка цен/скидок завершена успешно: "
                f"{rows_updated} строк, {total_ops} операций.")
        else:
            logging.warning(
                f"Обработка цен/скидок завершена с ошибками. "
                f"Проверьте лог выше.")

    except Exception as e:
        logging.error(f"Неожиданная ошибка в apply_price_discount: {e}")
        import traceback
        logging.error(traceback.format_exc())

    logging.info(f"--- Обработка цен/скидок для листа '{sheet_name}' завершена ---")


# --- НОВАЯ ФУНКЦИЯ: Проверка и добавление пустых строк ---
def ensure_minimum_empty_rows(gspread_client, target_spreadsheet, sheet_name, min_empty_rows):
    """
    Проверяет лист в целевой таблице и добавляет строки,
    если количество пустых строк (после последней строки с данными) меньше заданного минимума.
    """
    logging.info(f"--- Проверка и обеспечение минимума пустых строк ({min_empty_rows}) для листа: '{sheet_name}' ---")
    try:
        target_sheet = target_spreadsheet.worksheet(sheet_name)

        all_values = target_sheet.get_all_values()

        total_rows_in_sheet = len(all_values)

        last_data_row_index_0based = -1
        for i in range(len(all_values) - 1, -1, -1):
            if any(cell.strip() for cell in all_values[i]):
                last_data_row_index_0based = i
                break

        last_non_empty_row_1based = last_data_row_index_0based + 1

        if last_data_row_index_0based == -1 and target_sheet.row_count > 0 and len(target_sheet.row_values(1)) > 0:
            last_non_empty_row_1based = 1
        elif last_data_row_index_0based == -1:
            last_non_empty_row_1based = 0

        current_empty_rows = total_rows_in_sheet - last_non_empty_row_1based

        current_rows_count = target_sheet.row_count

        current_empty_rows = current_rows_count - last_non_empty_row_1based

        logging.info(
            f"Лист '{sheet_name}' имеет {current_rows_count} всего строк, последняя строка с данными/заголовком: {last_non_empty_row_1based}. Текущее количество пустых строк: {current_empty_rows}.")

        if current_empty_rows < min_empty_rows:
            rows_to_add = min_empty_rows - current_empty_rows
            logging.info(
                f"Количество пустых строк ({current_empty_rows}) меньше минимума ({min_empty_rows}). Добавляем {rows_to_add} строк.")

            new_total_rows = current_rows_count + rows_to_add

            max_retries = 5
            retry_delay = 5

            for attempt in range(max_retries):
                try:
                    logging.info(
                        f"Попытка изменения размера листа до {new_total_rows} строк ({attempt + 1}/{max_retries})...")
                    target_sheet.resize(rows=new_total_rows)
                    logging.info(f"Размер листа '{sheet_name}' успешно увеличен до {new_total_rows} строк.")
                    break
                except APIError as e:
                    if attempt < max_retries - 1 and (
                            '429' in str(e) or '500' in str(e) or '502' in str(e) or '503' in str(e) or '-1' in str(
                        e) or 'temporarily unavailable' in str(e).lower() or 'bad gateway' in str(
                        e).lower() or 'backend error' in str(e).lower()):
                        logging.warning(
                            f"API Error (вероятно временная) при изменении размера листа: {e}. Повторная попытка через {retry_delay} сек...")
                        time.sleep(retry_delay)
                    else:
                        logging.error(
                            f"Критическая ошибка API при изменении размера листа после {attempt + 1} попыток: {e}")
                        break

                except Exception as e:
                    logging.error(f"Неожиданная ошибка при изменении размера листа: {e}")
                    break
            else:
                logging.error(f"Не удалось изменить размер листа '{sheet_name}' после {max_retries} попыток.")

        else:
            logging.info(
                f"Текущее количество пустых строк ({current_empty_rows}) больше или равно минимуму ({min_empty_rows}). Изменений не требуется.")

    except WorksheetNotFound:
        logging.warning(f"Лист '{sheet_name}' не найден в целевой таблице. Проверка пустых строк пропущена.")
    except Exception as e:
        logging.error(f"Произошла ошибка при проверке/добавлении пустых строк в лист '{sheet_name}'. Error: {e}")
        import traceback
        logging.error(traceback.format_exc())

    logging.info(f"--- Проверка для листа '{sheet_name}' завершена ---")


# --------------------------------------------------------------------------------------------------------------------

def synchronize_all_sheets():
    """Основная функция для синхронизации всех листов (структуры и данных) из исходной таблицы в целевую."""
    logging.info("--- Запуск скрипта синхронизации Google Таблиц (структура и данные на месте) ---")
    try:
        # --- 1. Загрузка учетных данных сервисного аккаунта ---
        try:
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            credentials = Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=scopes)
            logging.info("Учетные данные сервисного аккаунта успешно загружены.")
        except FileNotFoundError:
            logging.error(f"Файл ключа сервисного аккаунта не найден: {SERVICE_ACCOUNT_FILE}")
            return
        except Exception as e:
            logging.error(f"Ошибка при загрузке учетных данных сервисного аккаунта: {e}")
            return

            # --- 2. Авторизация gspread Client с использованием учетных данных ---
        try:
            gspread_client = get_google_sheet_client(credentials)
        except Exception as e:
            return

            # --- 3. Создание клиента Sheets API v4 (googleapiclient) с использованием тех же учетных данных ---
        try:
            sheets_service = get_sheets_api_v4_service(credentials)
        except Exception as e:
            import traceback
            logging.error(traceback.format_exc())
            return

        logging.info(f"Открытие исходной таблицы (ID: {SOURCE_SPREADSHEET_ID})...")
        source_spreadsheet = gspread_client.open_by_key(SOURCE_SPREADSHEET_ID)
        logging.info(f"Открытие целевой таблицы (ID: {TARGET_SPREADSHEET_ID})...")
        target_spreadsheet = gspread_client.open_by_key(TARGET_SPREADSHEET_ID)

        logging.info("Получение списка листов из исходной таблицы...")
        source_worksheets = source_spreadsheet.worksheets()
        logging.info(f"Найдено {len(source_worksheets)} листов в исходной таблице.")

        if not source_worksheets:
            logging.warning("Исходная таблица не содержит листов. Нечего синхронизировать.")
            return

        for i, source_sheet in enumerate(source_worksheets):
            try:
                synchronize_single_sheet_with_data(gspread_client, sheets_service, source_sheet, target_spreadsheet)
            except Exception as e:
                logging.error(
                    f"Произошла НЕОБРАБОТАННАЯ ошибка при синхронизации листа '{source_sheet.title}'. Продолжаем с следующим листом (если есть). Error: {e}")
                import traceback
                logging.error(traceback.format_exc())

            if i < len(source_worksheets) - 1:
                time.sleep(2)

        logging.info("\n--- ЗАПУСК ОБРАБОТКИ ЦЕН/СКИДОК ---")
        apply_price_discount(sheets_service, source_spreadsheet, target_spreadsheet,
                             sheet_name='Export Products Sheet')
        logging.info("--- ОБРАБОТКА ЦЕН/СКИДОК ЗАВЕРШЕНА ---\n")

        logging.info("\n--- ЗАПУСК ПРОВЕРКИ МИНИМАЛЬНОГО КОЛИЧЕСТВА ПУСТЫХ СТРОК ---")
        for sheet_name in SHEETS_TO_ENSURE_EMPTY_ROWS:
            ensure_minimum_empty_rows(gspread_client, target_spreadsheet, sheet_name, MIN_EMPTY_ROWS)
        logging.info("--- ПРОВЕРКА МИНИМАЛЬНОГО КОЛИЧЕСТВА ПУСТЫХ СТРОК ЗАВЕРШЕНА ---\n")

    except gspread.exceptions.APIError as e:
        logging.error(f"Глобальная ошибка Google Sheets API (gspread): {e.response.text}")
        logging.error("Проверьте права доступа сервисного аккаунта к таблицам.")
    except Exception as e:
        logging.error(f"Произошла непредвиденная ошибка в основном процессе: {e}")
        import traceback
        logging.error(traceback.format_exc())
    finally:
        logging.info("--- Скрипт синхронизации Google Таблиц (структура и данные на месте) завершен ---")


# --- ЗАПУСК СКРИПТА ---
if __name__ == '__main__':
    synchronize_all_sheets()
