import pandas as pd
import json
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time
from datetime import datetime
import os  # <--- ВОТ ЭТА СТРОКА ОБЯЗАТЕЛЬНА
import re  # ДОБАВЛЕНО: нужно для обработки строк с ценами
from group_normalizer import resolve_product_group, get_canonical_groups_df

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='google_sheets_import.log',
    encoding='utf-8'
)
logger = logging.getLogger(__name__)

# Конфигурация
CONFIG = {
    "SERVICE_ACCOUNT_FILE": "key_sheet.json",
    "MAPPING_FILE": "column_mapping.json",
    "SUPPLIERS": {
        "AGER": os.environ.get("PARS_AGER_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_AGER_SPREADSHEET_ID", "PARS_AGER_SPREADSHEET_ID"),
        "FOOTBALLERS": os.environ.get("PARS_FOOTBALLERS_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_FOOTBALLERS_SPREADSHEET_ID", "PARS_FOOTBALLERS_SPREADSHEET_ID"),
        "IZIDROP": os.environ.get("PARS_IZIDROP_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_IZIDROP_SPREADSHEET_ID", "PARS_IZIDROP_SPREADSHEET_ID"),
        "MOYDROP": os.environ.get("PARS_MOYDROP_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_MOYDROP_SPREADSHEET_ID", "PARS_MOYDROP_SPREADSHEET_ID"),
        "BAYRAKTAR": os.environ.get("PARS_BAYRAKTAR_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_BAYRAKTAR_SPREADSHEET_ID", "PARS_BAYRAKTAR_SPREADSHEET_ID"),
        "KIRS": os.environ.get("PARS_KIRS_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_KIRS_SPREADSHEET_ID", "PARS_KIRS_SPREADSHEET_ID"),
        "BAGSROOM": os.environ.get("PARS_BAGSROOM_PLUS_SPREADSHEET_ID") or os.environ.get("PARS_BAGSROOM_SPREADSHEET_ID", "PARS_BAGSROOM_SPREADSHEET_ID")
    },
    "OUTPUT_SPREADSHEET_ID": os.environ.get("IMPORT_TEMPLATE_PROM_SPREADSHEET_ID", "IMPORT_TEMPLATE_PROM_SPREADSHEET_ID"),
    "DELAY_BETWEEN_REQUESTS": 2,
    "MAX_RETRIES": 7,
    "BATCH_SIZE": 1000,
    "DELETE_BATCH_SIZE": 1000,
}


class GoogleSheetsManager:
    def __init__(self):
        self.service = self._authenticate()
        self.last_request_time = datetime.now()
        self.spreadsheet_id = CONFIG["OUTPUT_SPREADSHEET_ID"]
        self._sheet_id_cache = {}  # Кэш sheetId по имени листа

    def _authenticate(self):
        try:
            scopes = ['https://www.googleapis.com/auth/spreadsheets']
            credentials = service_account.Credentials.from_service_account_file(
                CONFIG["SERVICE_ACCOUNT_FILE"], scopes=scopes)
            return build('sheets', 'v4', credentials=credentials)
        except Exception as e:
            logger.error(f"Ошибка аутентификации: {e}")
            raise

    def _wait_if_needed(self):
        elapsed = (datetime.now() - self.last_request_time).total_seconds()
        if elapsed < CONFIG["DELAY_BETWEEN_REQUESTS"]:
            time.sleep(CONFIG["DELAY_BETWEEN_REQUESTS"] - elapsed)
        self.last_request_time = datetime.now()

    def _execute_with_retry(self, request):
        for attempt in range(CONFIG["MAX_RETRIES"]):
            try:
                self._wait_if_needed()
                return request.execute()
            except Exception as e:
                logger.warning(f"Сетевая ошибка, попытка {attempt + 1}/{CONFIG['MAX_RETRIES']}: {e}")
                if attempt == CONFIG["MAX_RETRIES"] - 1:
                    raise
                time.sleep(2 ** attempt)

    def get_sheet_data(self, spreadsheet_id, sheet_name):
        for attempt in range(CONFIG["MAX_RETRIES"]):
            try:
                self._wait_if_needed()
                result = self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=sheet_name,
                    valueRenderOption='FORMATTED_VALUE'
                ).execute()

                values = result.get('values', [])

                if not values:
                    logger.warning(f"Лист {sheet_name} пуст или не найден в spreadsheets.values().get()")
                    return pd.DataFrame()

                raw_headers = values[0]
                data = values[1:]

                # --- ДИАГНОСТИЧЕСКИЙ ЛОГ ВНУТРИ get_sheet_data ---
                logger.info(
                    f"GET_SHEET_DATA DEBUG ({sheet_name}): Количество исходных заголовков (len(raw_headers)): {len(raw_headers)}")
                logger.info(
                    f"GET_SHEET_DATA DEBUG ({sheet_name}): Первые 50 исходных заголовков: {raw_headers[:50]}")
                if len(raw_headers) > 50:
                    logger.info(
                        f"GET_SHEET_DATA DEBUG ({sheet_name}): Последние 50 исходных заголовков: {raw_headers[-50:]}")

                from collections import Counter
                header_counts = Counter(raw_headers)
                duplicate_headers_in_raw = {header: count for header, count in header_counts.items() if count > 1}
                if duplicate_headers_in_raw:
                    logger.info(
                        f"GET_SHEET_DATA DEBUG ({sheet_name}): Найдены дублирующиеся заголовки в raw_headers: {duplicate_headers_in_raw}")
                else:
                    logger.info(
                        f"GET_SHEET_DATA DEBUG ({sheet_name}): Дублирующихся заголовков в raw_headers НЕ найдено.")

                processed_headers = []
                counts = {}
                for header_val in raw_headers:
                    if header_val is None: header_val = ''
                    str_header_val = str(header_val)
                    if str_header_val in counts:
                        counts[str_header_val] += 1
                        processed_headers.append(f"{str_header_val}.{counts[str_header_val] - 1}")
                    else:
                        counts[str_header_val] = 1
                        processed_headers.append(str_header_val)

                logger.info(
                    f"GET_SHEET_DATA DEBUG ({sheet_name}): Количество обработанных заголовков (len(processed_headers)): {len(processed_headers)}")
                logger.info(
                    f"GET_SHEET_DATA DEBUG ({sheet_name}): Первые 50 обработанных заголовков: {processed_headers[:50]}")
                if len(processed_headers) > 50:
                    logger.info(
                        f"GET_SHEET_DATA DEBUG ({sheet_name}): Последние 50 обработанных заголовков: {processed_headers[-50:]}")

                normalized_data = []
                for row_values in data:
                    current_row_normalized = list(row_values)
                    if len(current_row_normalized) < len(processed_headers):
                        current_row_normalized.extend([''] * (len(processed_headers) - len(current_row_normalized)))
                    elif len(current_row_normalized) > len(processed_headers):
                        current_row_normalized = current_row_normalized[:len(processed_headers)]
                    normalized_data.append(current_row_normalized)

                if not normalized_data and not processed_headers:
                    return pd.DataFrame()
                if not processed_headers and normalized_data:
                    logger.warning(
                        f"GET_SHEET_DATA WARNING ({sheet_name}): Есть данные, но не удалось определить заголовки.")
                    return pd.DataFrame(normalized_data)

                df = pd.DataFrame(normalized_data, columns=processed_headers)
                logger.info(
                    f"GET_SHEET_DATA INFO ({sheet_name}): DataFrame создан. Количество столбцов в DataFrame: {len(df.columns)}")
                logger.info(
                    f"GET_SHEET_DATA INFO ({sheet_name}): Столбцы DataFrame: {df.columns.tolist()[:50]}")
                if len(df.columns) > 50:
                    logger.info(
                        f"GET_SHEET_DATA INFO ({sheet_name}): Последние 50 столбцов DataFrame: {df.columns.tolist()[-50:]}")

                return df

            except Exception as e:
                logger.warning(f"GET_SHEET_DATA ({sheet_name}): Попытка {attempt + 1} не удалась: {e}")
                if attempt == CONFIG["MAX_RETRIES"] - 1:
                    logger.error(f"GET_SHEET_DATA ({sheet_name}): Не удалось прочитать лист: {e}")
                    return pd.DataFrame()
                time.sleep(2 ** attempt)

        return pd.DataFrame()

    def write_sheet_data(self, spreadsheet_id, sheet_name, df):
        try:
            if df.empty:
                logger.info(f"Пустой DataFrame, нет данных для записи в {sheet_name}")
                return True, None

            is_groups = "Export Groups Sheet" in sheet_name

            def _col_letter(idx):
                letter, n = '', idx + 1
                while n > 0:
                    n, rem = divmod(n - 1, 26)
                    letter = chr(65 + rem) + letter
                return letter

            # ── Шаг 1: читаем только строку заголовков (строка 1) ──────────────
            self._wait_if_needed()
            hdr_result = self._execute_with_retry(
                self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"{sheet_name}!1:1",
                    valueRenderOption='FORMATTED_VALUE'
                )
            )
            header_values = hdr_result.get('values', [])

            # Лист пустой — пишем с нуля
            if not header_values:
                logger.info(f"Таблица {sheet_name} пустая, записываем новые данные")
                self._ensure_sheet_capacity(spreadsheet_id, sheet_name, len(df) + 1)
                for i in range(0, len(df), CONFIG["BATCH_SIZE"]):
                    self._write_batch(spreadsheet_id, sheet_name,
                                      df.iloc[i:i + CONFIG["BATCH_SIZE"]], i == 0)
                return True, None

            headers = header_values[0]

            # Гарантируем наличие критических колонок в заголовке таблицы
            if not is_groups and "Ідентифікатор_підрозділу" not in headers:
                logger.warning(f"'Ідентифікатор_підрозділу' отсутствует в заголовках {sheet_name}. Добавляем в заголовок...")
                headers.append("Ідентифікатор_підрозділу")
                col_letter = _col_letter(len(headers) - 1)
                self._wait_if_needed()
                self._execute_with_retry(
                    self.service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=f"{sheet_name}!{col_letter}1",
                        valueInputOption='USER_ENTERED',
                        body={'values': [["Ідентифікатор_підрозділу"]]}
                    )
                )

            # ── Шаг 2: определяем ключевые столбцы ────────────────────────────
            if is_groups:
                id_column = next((c for c in ["Ідентифікатор_групи", "Номер_групи"]
                                  if c in df.columns), df.columns[0])
                name_column = "Назва_групи" if "Назва_групи" in df.columns else None
            else:
                id_column = next((c for c in ["Ідентифікатор_товару", "Код_товару"]
                                  if c in df.columns), df.columns[0])
                name_column = None

            id_col_idx = headers.index(id_column) if id_column in headers else 0
            name_col_idx = headers.index(name_column) if name_column in headers else -1

            # ── Шаг 3: читаем ключевой столбец (+ name_column для групп) ──────
            #    Это ОДИН узкий запрос вместо чтения всего листа целиком.
            key_col_letter = _col_letter(id_col_idx)
            self._wait_if_needed()
            key_result = self._execute_with_retry(
                self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"{sheet_name}!{key_col_letter}:{key_col_letter}",
                    valueRenderOption='FORMATTED_VALUE'
                )
            )
            key_col_values = key_result.get('values', [])  # [[hdr],[v1],[v2],...]

            # Реальное количество занятых строк берём из метаданных листа (кэш),
            # а НЕ из длины ключевого столбца — ключ может быть пустым в части строк.
            self._ensure_sheet_capacity(spreadsheet_id, sheet_name, 1)  # прогреваем кэш без расширения
            cached = self._sheet_id_cache.get(sheet_name)
            if isinstance(cached, tuple):
                # rowCount — полный размер листа; вычитаем пустые строки в конце.
                # Используем длину ключевого столбца как нижнюю границу, но
                # берём максимум с num_data_rows который вычислим позже.
                _sheet_total_rows = cached[1]
            else:
                _sheet_total_rows = len(key_col_values)

            # После построения col_data пересчитаем точнее (см. ниже)
            total_existing_rows = len(key_col_values)  # временно, уточняется в Шаге 6

            # ── Шаг 4: определяем критические индексы ─────────────────────────
            col_idx_to_df_col = {}
            if is_groups:
                critical_fields = [
                    "Назва_групи", "Назва_групи_укр", "Ідентифікатор_групи",
                    "Номер_батьківської_групи", "Ідентифікатор_батьківської_групи"
                ]
                for f in critical_fields:
                    if f in headers and f in df.columns:
                        idx = headers.index(f)
                        col_idx_to_df_col[idx] = f
            else:
                base_critical_fields = [
                    "Назва_позиції", "Назва_позиції_укр", "Опис", "Опис_укр",
                    "Наявність", "Ціна", "Ціна_від", "Кількість", "Знижка",
                    "Ярлик", "Де_знаходиться_товар", "Тип_товару",
                    "Ідентифікатор_підрозділу", "Номер_групи", "Назва_групи", "Ідентифікатор_групи",
                    "Виробник", "Країна_виробник", "Посилання_зображення"
                ]
                # Добавляем все слоты характеристик (1..24) для полной синхронизации
                for i in range(1, 25):
                    suffix = f"_{i}" if i > 1 else ""
                    base_critical_fields.extend([
                        f"Назва_Характеристики{suffix}",
                        f"Одиниця_виміру_Характеристики{suffix}",
                        f"Значення_Характеристики{suffix}"
                    ])

                for f in base_critical_fields:
                    if f in headers and f in df.columns:
                        idx = headers.index(f)
                        col_idx_to_df_col[idx] = f

            critical_indices = sorted(col_idx_to_df_col.keys())

            # ── Шаг 5: читаем только критические столбцы чанками ─────────────
            #    Вместо загрузки всего листа (20к строк × 127 колонок) читаем
            #    только нужные колонки. Разбиваем на чанки по BATCH_SIZE колонок.
            COL_CHUNK = 20  # столбцов за один запрос
            # Словарь col_idx -> list значений (без заголовка, 0-based по строкам)
            col_data = {}

            unique_col_indices = sorted(set(critical_indices + [id_col_idx]
                                            + ([name_col_idx] if name_col_idx != -1 else [])))

            for chunk_start in range(0, len(unique_col_indices), COL_CHUNK):
                chunk_cols = unique_col_indices[chunk_start:chunk_start + COL_CHUNK]
                # Строим multi-range: A2:A, C2:C, ...
                ranges = [f"{sheet_name}!{_col_letter(ci)}2:{_col_letter(ci)}"
                          for ci in chunk_cols]
                self._wait_if_needed()
                batch_res = self._execute_with_retry(
                    self.service.spreadsheets().values().batchGet(
                        spreadsheetId=spreadsheet_id,
                        ranges=ranges,
                        valueRenderOption='FORMATTED_VALUE'
                    )
                )
                for col_idx, vr in zip(chunk_cols, batch_res.get('valueRanges', [])):
                    vals = vr.get('values', [])
                    col_data[col_idx] = [row[0] if row else '' for row in vals]

            # ── Шаг 6: строим карту существующих строк по ключу ───────────────
            # num_data_rows = максимальная длина любого из прочитанных столбцов.
            # total_existing_rows = реальная последняя занятая строка в листе
            # (используем максимум между ключевым столбцом и всеми критическими).
            num_data_rows = max((len(v) for v in col_data.values()), default=0)
            total_existing_rows = max(total_existing_rows, num_data_rows + 1)  # +1 за заголовок
            logger.info(f"Реальных занятых строк в {sheet_name} (включая заголовок): {total_existing_rows}")

            existing_data_dict = {}
            for row_offset in range(num_data_rows):
                id_val = col_data.get(id_col_idx, [])
                id_val = id_val[row_offset] if row_offset < len(id_val) else ''
                if not id_val:
                    continue
                sheet_row = row_offset + 2  # 1-based, +1 за заголовок

                if is_groups and name_col_idx != -1:
                    nm_vals = col_data.get(name_col_idx, [])
                    nm_val = nm_vals[row_offset] if row_offset < len(nm_vals) else ''
                    row_key = (id_val, nm_val)
                else:
                    row_key = id_val

                # data: dict col_idx -> значение (для сравнения с новыми данными)
                row_cell_data = {ci: (col_data[ci][row_offset]
                                      if ci in col_data and row_offset < len(col_data[ci])
                                      else '')
                                 for ci in critical_indices}
                existing_data_dict[row_key] = {
                    'row_idx': sheet_row,
                    'data': row_cell_data,
                }

            # ── Шаг 7: сравниваем df с существующими данными ─────────────────
            if is_groups and name_column:
                df_unique = df.drop_duplicates(subset=[id_column, name_column], keep='last')
            else:
                df_unique = df.drop_duplicates(subset=[id_column], keep='last')

            new_rows_list = []
            batch_value_updates = []
            deleted_rows = set(existing_data_dict.keys())

            for _, row in df_unique.iterrows():
                row_id = str(row[id_column])
                if is_groups and name_column:
                    row_name = str(row[name_column]) if pd.notna(row.get(name_column, '')) else ''
                    row_key = (row_id, row_name)
                else:
                    row_key = row_id

                if row_key in existing_data_dict:
                    deleted_rows.discard(row_key)
                    existing_row_cells = existing_data_dict[row_key]['data']
                    row_idx = existing_data_dict[row_key]['row_idx']

                    # Собираем только изменившиеся ячейки
                    cell_updates = []
                    for col_idx in critical_indices:
                        col_name = col_idx_to_df_col.get(col_idx, headers[col_idx] if col_idx < len(headers) else '')
                        new_value = str(row[col_name]) if col_name in row and pd.notna(row.get(col_name)) else ''
                        old_value = existing_row_cells.get(col_idx, '')
                        if new_value != old_value:
                            cell_updates.append({
                                'range': f"{sheet_name}!{_col_letter(col_idx)}{row_idx}",
                                'values': [[new_value]]
                            })

                    batch_value_updates.extend(cell_updates)
                else:
                    new_rows_list.append(row)

            # ── Шаг 8: отправляем обновления батчами ─────────────────────────
            if batch_value_updates:
                logger.info(f"Отправляем {len(batch_value_updates)} обновлений ячеек в {sheet_name}")
                BATCH_UPDATE_CHUNK = 5000
                for chunk_start in range(0, len(batch_value_updates), BATCH_UPDATE_CHUNK):
                    chunk = batch_value_updates[chunk_start:chunk_start + BATCH_UPDATE_CHUNK]
                    self._execute_with_retry(
                        self.service.spreadsheets().values().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body={'valueInputOption': 'RAW', 'data': chunk}
                        )
                    )
                    logger.info(f"  Обновления: чанк {chunk_start + 1}–{chunk_start + len(chunk)}")

            # ── Шаг 9: добавляем новые строки ────────────────────────────────
            if new_rows_list:
                new_rows_df = pd.DataFrame(new_rows_list, columns=df.columns)
                # start_row = строка ПОСЛЕ последней занятой (total_existing_rows уже включает заголовок)
                start_row = total_existing_rows + 1
                needed = start_row + len(new_rows_df) - 1
                logger.info(f"Новые строки: {len(new_rows_df)} шт., запись начиная со строки {start_row}")
                self._ensure_sheet_capacity(spreadsheet_id, sheet_name, needed)

                for i in range(0, len(new_rows_df), CONFIG["BATCH_SIZE"]):
                    batch = new_rows_df.iloc[i:i + CONFIG["BATCH_SIZE"]]
                    values = batch.fillna('').values.tolist()
                    self._execute_with_retry(
                        self.service.spreadsheets().values().update(
                            spreadsheetId=spreadsheet_id,
                            range=f"{sheet_name}!A{start_row}",
                            valueInputOption='RAW',
                            body={'values': values}
                        )
                    )
                    start_row += len(values)
                logger.info(f"Добавлено {len(new_rows_df)} новых строк в {sheet_name}")

            # ── Шаг 10: удаляем строки отсутствующих позиций ─────────────────
            if deleted_rows:
                rows_to_delete = sorted(
                    [existing_data_dict[k]['row_idx'] for k in deleted_rows], reverse=True)
                logger.info(f"Удаление {len(deleted_rows)} позиций, отсутствующих у поставщика")
                sheet_id = self._get_sheet_id(sheet_name)
                delete_requests = [
                    {'deleteDimension': {'range': {
                        'sheetId': sheet_id, 'dimension': 'ROWS',
                        'startIndex': r - 1, 'endIndex': r
                    }}}
                    for r in rows_to_delete
                ]
                for chunk_start in range(0, len(delete_requests), CONFIG["DELETE_BATCH_SIZE"]):
                    chunk = delete_requests[chunk_start:chunk_start + CONFIG["DELETE_BATCH_SIZE"]]
                    self._execute_with_retry(
                        self.service.spreadsheets().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body={'requests': chunk}
                        )
                    )
                logger.info(f"Удалено {len(rows_to_delete)} строк")

            # ── Шаг 11: удаляем дубликаты ────────────────────────────────────
            sheet_type = "products" if "Export Products Sheet" in sheet_name else "groups"
            self._remove_duplicates(spreadsheet_id, sheet_name, sheet_type)

            logger.info(
                f"Готово [{sheet_name}]: обновлено ячеек={len(batch_value_updates)}, "
                f"добавлено строк={len(new_rows_list)}, удалено строк={len(deleted_rows)}")
            return True, None

        except Exception as e:
            logger.error(f"Ошибка записи в {sheet_name}: {e}", exc_info=True)
            return False, str(e)

    def _get_sheet_id(self, sheet_name):
        try:
            cached = self._sheet_id_cache.get(sheet_name)
            if cached is not None:
                return cached[0] if isinstance(cached, tuple) else cached

            self._wait_if_needed()
            sheets_metadata = self.service.spreadsheets().get(
                spreadsheetId=CONFIG["OUTPUT_SPREADSHEET_ID"],
                fields='sheets.properties'
            ).execute()

            for sheet in sheets_metadata.get('sheets', []):
                title = sheet['properties']['title']
                sid = sheet['properties']['sheetId']
                rc = sheet['properties'].get('gridProperties', {}).get('rowCount', 0)
                self._sheet_id_cache[title] = (sid, rc)

            cached = self._sheet_id_cache.get(sheet_name)
            if cached is not None:
                return cached[0] if isinstance(cached, tuple) else cached

            logger.error(f"Лист {sheet_name} не найден")
            raise ValueError(f"Лист {sheet_name} не найден")

        except Exception as e:
            logger.error(f"Ошибка получения ID листа: {e}")
            raise

    def _column_index_to_letter(self, index):
        result = ""
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result if result else "A"

    def _remove_duplicates(self, spreadsheet_id, sheet_name, sheet_type):
        try:
            self._wait_if_needed()
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=sheet_name
            ).execute()

            values = result.get('values', [])
            if not values or len(values) <= 1:
                logger.info(f"Нет данных для удаления дубликатов в {sheet_name}")
                return

            headers = values[0]

            if sheet_type == "products":
                id_col = "Ідентифікатор_товару" if "Ідентифікатор_товару" in headers else "Код_товару"
                if id_col not in headers:
                    logger.warning(f"Столбец {id_col} не найден в заголовках")
                    return

                id_idx = headers.index(id_col)
                critical_fields = {
                    "Наявність": headers.index("Наявність") if "Наявність" in headers else None,
                    "Ціна": headers.index("Ціна") if "Ціна" in headers else None,
                    "Кількість": headers.index("Кількість") if "Кількість" in headers else None,
                    "Знижка": headers.index("Знижка") if "Знижка" in headers else None
                }

                seen_ids = {}
                rows_to_delete = []
                rows_to_update = {}

                for i in range(1, len(values)):
                    row = values[i]
                    if len(row) > id_idx:
                        current_id = row[id_idx].strip()
                        if current_id in seen_ids:
                            original_idx = seen_ids[current_id]
                            if original_idx not in rows_to_update:
                                rows_to_update[original_idx] = values[original_idx].copy()

                            for field, idx in critical_fields.items():
                                if idx is not None and idx < len(row):
                                    if row[idx].strip():
                                        if idx >= len(rows_to_update[original_idx]):
                                            rows_to_update[original_idx].extend(
                                                [''] * (idx - len(rows_to_update[original_idx]) + 1))
                                        rows_to_update[original_idx][idx] = row[idx]

                            rows_to_delete.append(i + 1)
                        else:
                            seen_ids[current_id] = i

                # ----------------------------------------------------------------
                # ИСПРАВЛЕНИЕ #2: обновляем оригинальные строки одним батч-запросом
                # вместо N отдельных values().update() в цикле.
                # ----------------------------------------------------------------
                if rows_to_update:
                    batch_updates = [
                        {
                            'range': f"{sheet_name}!A{idx + 1}",
                            'values': [updated_row]
                        }
                        for idx, updated_row in rows_to_update.items()
                    ]
                    for chunk_start in range(0, len(batch_updates), CONFIG["BATCH_SIZE"]):
                        chunk = batch_updates[chunk_start:chunk_start + CONFIG["BATCH_SIZE"]]
                        self._wait_if_needed()
                        self.service.spreadsheets().values().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body={
                                'valueInputOption': 'RAW',
                                'data': chunk
                            }
                        ).execute()
                    logger.info(
                        f"Обновлено {len(rows_to_update)} оригинальных строк (слияние дубликатов) одним батчем.")

                # ----------------------------------------------------------------
                # ИСПРАВЛЕНИЕ #3: удаляем дубликаты одним batchUpdate
                # вместо N отдельных batchUpdate с одним запросом каждый.
                # ----------------------------------------------------------------
                if rows_to_delete:
                    sheet_id = self._get_sheet_id(sheet_name)
                    delete_requests = [
                        {
                            'deleteDimension': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'dimension': 'ROWS',
                                    'startIndex': row_idx - 1,
                                    'endIndex': row_idx
                                }
                            }
                        }
                        for row_idx in sorted(rows_to_delete, reverse=True)
                    ]

                    # Разбиваем на чанки, если дубликатов очень много
                    for chunk_start in range(0, len(delete_requests), CONFIG["DELETE_BATCH_SIZE"]):
                        chunk = delete_requests[chunk_start:chunk_start + CONFIG["DELETE_BATCH_SIZE"]]
                        self._wait_if_needed()
                        self.service.spreadsheets().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body={'requests': chunk}
                        ).execute()

                    logger.info(f"Удалено {len(rows_to_delete)} дубликатов в {sheet_name} (батч-запрос)")

            else:
                # --- Группы ---
                group_num_col = "Ідентифікатор_групи"
                group_name_col = "Назва_групи"

                if group_num_col not in headers or group_name_col not in headers:
                    logger.warning("Не найдены необходимые столбцы для групп")
                    return

                num_idx = headers.index(group_num_col)
                name_idx = headers.index(group_name_col)

                seen_groups = {}
                rows_to_delete = []

                for i in range(1, len(values)):
                    row = values[i]
                    if len(row) > max(num_idx, name_idx):
                        group_key = (row[num_idx].strip(), row[name_idx].strip())
                        if group_key in seen_groups:
                            rows_to_delete.append(i + 1)
                        else:
                            seen_groups[group_key] = i

                # ----------------------------------------------------------------
                # ИСПРАВЛЕНИЕ #4: то же для групп — один batchUpdate вместо цикла
                # ----------------------------------------------------------------
                if rows_to_delete:
                    sheet_id = self._get_sheet_id(sheet_name)
                    delete_requests = [
                        {
                            'deleteDimension': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'dimension': 'ROWS',
                                    'startIndex': row_idx - 1,
                                    'endIndex': row_idx
                                }
                            }
                        }
                        for row_idx in sorted(rows_to_delete, reverse=True)
                    ]

                    for chunk_start in range(0, len(delete_requests), CONFIG["DELETE_BATCH_SIZE"]):
                        chunk = delete_requests[chunk_start:chunk_start + CONFIG["DELETE_BATCH_SIZE"]]
                        self._wait_if_needed()
                        self.service.spreadsheets().batchUpdate(
                            spreadsheetId=spreadsheet_id,
                            body={'requests': chunk}
                        ).execute()

                    logger.info(f"Удалено {len(rows_to_delete)} дубликатов в {sheet_name} (батч-запрос)")

        except Exception as e:
            logger.error(f"Ошибка при удалении дубликатов в {sheet_name}: {e}")

    def _clear_sheet(self, spreadsheet_id, sheet_name):
        self._wait_if_needed()
        self.service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            body={}
        ).execute()

    def _write_batch(self, spreadsheet_id, sheet_name, batch_df, is_first_batch):
        cleaned_df = batch_df.copy()

        for col in cleaned_df.columns:
            if cleaned_df[col].dtype == 'object':
                cleaned_df[col] = cleaned_df[col].apply(
                    lambda x: '' if pd.isna(x) else str(x) if isinstance(x, (pd.Series, list, dict)) else x)

        if is_first_batch:
            values = [cleaned_df.columns.tolist()]
        else:
            values = []

        for _, row in cleaned_df.iterrows():
            row_list = []
            for col in cleaned_df.columns:
                value = row[col]
                if pd.isna(value):
                    row_list.append('')
                elif isinstance(value, (pd.Series, list, dict)):
                    row_list.append(str(value))
                else:
                    row_list.append(value)
            values.append(row_list)

        if is_first_batch:
            range_name = f"{sheet_name}!A1"
        else:
            start_row = self._get_next_empty_row(spreadsheet_id, sheet_name)
            range_name = f"{sheet_name}!A{start_row}"

        body = {'values': values}

        for attempt in range(CONFIG["MAX_RETRIES"]):
            try:
                self._wait_if_needed()
                logger.info(f"Записываем пакет данных в {sheet_name} по диапазону {range_name}: {len(values)} строк")

                self.service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption='RAW',
                    body=body
                ).execute()
                return
            except Exception as e:
                logger.warning(f"Попытка {attempt + 1} записи не удалась: {str(e)}")
                if attempt == CONFIG["MAX_RETRIES"] - 1:
                    logger.error(f"Все попытки записи не удались. Последняя ошибка: {str(e)}")
                    raise Exception(f"Ошибка API при записи данных: {str(e)}")
                time.sleep(2 ** attempt)

    def _get_next_empty_row(self, spreadsheet_id, sheet_name):
        self._wait_if_needed()
        result = self.service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A:A"
        ).execute()
        values = result.get('values', [])
        return len(values) + 1

    def _ensure_sheet_capacity(self, spreadsheet_id, sheet_name, required_rows):
        """Расширяет лист если строк меньше required_rows.
        Использует кэш _sheet_id_cache для хранения {sheet_name: (sheetId, rowCount)},
        чтобы не делать лишний API-запрос если размер уже известен и достаточен.
        """
        try:
            cached = self._sheet_id_cache.get(sheet_name)
            # cached хранит либо int (только sheetId — старый формат), либо tuple (sheetId, rowCount)
            if isinstance(cached, tuple):
                sheet_id, cached_rows = cached
                if cached_rows >= required_rows:
                    logger.info(
                        f"Лист {sheet_name}: кэш показывает {cached_rows} строк >= {required_rows}, расширение не нужно.")
                    return False
            # Нужно узнать актуальный размер
            self._wait_if_needed()
            sheets_info = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields='sheets.properties'
            ).execute()

            # Обновляем весь кэш попутно
            for sheet in sheets_info.get('sheets', []):
                t = sheet['properties']['title']
                sid = sheet['properties']['sheetId']
                rc = sheet['properties'].get('gridProperties', {}).get('rowCount', 0)
                self._sheet_id_cache[t] = (sid, rc)

            entry = self._sheet_id_cache.get(sheet_name)
            if not entry:
                logger.error(f"Лист {sheet_name} не найден в таблице.")
                return False

            sheet_id, current_rows = entry
            logger.info(f"Лист {sheet_name} имеет {current_rows} строк, требуется {required_rows}")

            if current_rows < required_rows:
                new_row_count = required_rows + 1000  # запас
                logger.info(f"Расширяем лист {sheet_name} до {new_row_count} строк")

                request = {
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': sheet_id,
                            'gridProperties': {'rowCount': new_row_count}
                        },
                        'fields': 'gridProperties.rowCount'
                    }
                }
                self._wait_if_needed()
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': [request]}
                ).execute()

                # Обновляем кэш
                self._sheet_id_cache[sheet_name] = (sheet_id, new_row_count)
                logger.info(f"Лист {sheet_name} успешно расширен до {new_row_count} строк")
                return True

            return False
        except Exception as e:
            logger.error(f"Ошибка при расширении листа {sheet_name}: {e}")
            return False

    def _ensure_min_empty_rows(self, sheet_name, min_empty_rows=1000):
        try:
            self._wait_if_needed()
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A:A"
            ).execute()

            values = result.get('values', [])
            last_used_row_index = len(values)

            self._wait_if_needed()
            sheets_info = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id,
                fields='sheets.properties'
            ).execute()

            target_sheet = None
            for sheet in sheets_info.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    target_sheet = sheet
                    break

            if not target_sheet:
                logger.error(f"Лист {sheet_name} не найден для проверки пустых строк.")
                return

            current_total_rows = target_sheet['properties'].get('gridProperties', {}).get('rowCount', 0)
            empty_rows = current_total_rows - last_used_row_index

            logger.info(
                f"Лист '{sheet_name}': Всего строк: {current_total_rows}, Занятых строк: {last_used_row_index}, Пустых строк: {empty_rows}")

            if empty_rows < min_empty_rows:
                rows_to_add = min_empty_rows - empty_rows
                new_row_count = current_total_rows + rows_to_add

                logger.info(
                    f"Лист '{sheet_name}': Недостаточно пустых строк ({empty_rows} < {min_empty_rows}). Добавляем {rows_to_add} строк для достижения {new_row_count} строк.")

                request = {
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': self._get_sheet_id(sheet_name),
                            'gridProperties': {
                                'rowCount': new_row_count
                            }
                        },
                        'fields': 'gridProperties.rowCount'
                    }
                }

                self._wait_if_needed()
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': [request]}
                ).execute()

                logger.info(f"Лист '{sheet_name}' успешно расширен до {new_row_count} строк.")
            else:
                logger.info(
                    f"Лист '{sheet_name}': Количество пустых строк ({empty_rows}) соответствует или превышает минимум ({min_empty_rows}). Изменений не требуется.")

        except Exception as e:
            logger.error(f"Ошибка при проверке/добавлении пустых строк в {sheet_name}: {e}")


def load_mapping():
    try:
        with open(CONFIG["MAPPING_FILE"], 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Ошибка загрузки маппинга: {e}")
        raise


def get_template_columns():
    products_columns = [
        "Код_товару", "Назва_позиції", "Назва_позиції_укр", "Пошукові_запити", "Пошукові_запити_укр",
        "Опис", "Опис_укр", "Тип_товару", "Ціна", "Валюта", "Одиниця_виміру", "Мінімальний_обсяг_замовлення",
        "Оптова_ціна", "Мінімальне_замовлення_опт", "Посилання_зображення", "Наявність", "Кількість",
        "Номер_групи", "Назва_групи", "Посилання_підрозділу", "Можливість_поставки", "Термін_поставки",
        "Спосіб_пакування", "Спосіб_пакування_укр", "Унікальний_ідентифікатор", "Ідентифікатор_товару",
        "Ідентифікатор_підрозділу", "Ідентифікатор_групи", "Виробник", "Країна_виробник", "Знижка",
        "ID_групи_різновидів", "Особисті_нотатки", "Продукт_на_сайті", "Термін_дії_знижки_від",
        "Термін_дії_знижки_до", "Ціна_від", "Ярлик", "HTML_заголовок", "HTML_заголовок_укр",
        "HTML_опис", "HTML_опис_укр", "Код_маркування_(GTIN)", "Номер_пристрою_(MPN)",
        "Вага,кг", "Ширина,см", "Висота,см", "Довжина,см", "Де_знаходиться_товар",
        "Товар_на_видалення", "Причина_видалення_товару"
    ]

    for i in range(1, 25):
        products_columns.extend([
            f"Назва_Характеристики_{i}" if i > 1 else "Назва_Характеристики",
            f"Одиниця_виміру_Характеристики_{i}" if i > 1 else "Одиниця_виміру_Характеристики",
            f"Значення_Характеристики_{i}" if i > 1 else "Значення_Характеристики"
        ])

    groups_columns = [
        "Номер_групи", "Назва_групи", "Назва_групи_укр", "Ідентифікатор_групи",
        "Номер_батьківської_групи", "Ідентифікатор_батьківської_групи", "HTML_заголовок_групи",
        "HTML_заголовок_групи_укр", "HTML_опис_групи", "HTML_опис_групи_укр",
        "Опис_групи_до_списку_товарних_позицій", "Опис_групи_до_списку_товарних_позицій_укр",
        "Опис_групи_після_списку_товарних_позицій", "Опис_групи_після_списку_товарних_позицій_укр",
        "Посилання_зображення_групи"
    ]

    return products_columns, groups_columns


# ══════════════════════════════════════════════════════════════════════════════
# ЭТАП 1 И ЭТАП 2: Нормализация и перевод характеристик под стандарты Prom.ua
# ══════════════════════════════════════════════════════════════════════════════

UKR_PARAM_NAMES = {
    # Одежда и обувь
    "Размер": "Розмір",
    "Розмір": "Розмір",
    "Размер женской обуви": "Розмір жіночого взуття",
    "Размер мужской обуви": "Розмір чоловічого взуття",
    "Сезон": "Сезон",
    "Материал верха": "Матеріал верху",
    "Материал": "Матеріал",
    "Матеріал": "Матеріал",
    "Цвет": "Колір",
    "Колір": "Колір",
    "Группа цвета": "Група кольору",
    "Материал подкладки": "Матеріал підкладки",
    "Материал подошвы": "Матеріал підошви",
    "Материал стельки": "Матеріал устілки",
    "Пол": "Стать",
    "Стать": "Стать",
    "Бренд": "Бренд",
    "Производитель": "Виробник",
    "Виробник": "Виробник",
    "Застежка": "Застібка",
    "Стиль": "Стиль",
    "Состояние": "Стан",
    "Состав": "Склад",
    "Размерная группа": "Розмірна група",
    "Страна производитель": "Країна-виробник",
    "Страна регистрации бренда": "Країна реєстрації бренду",
    "Страна-производитель товара": "Країна-виробник товару",
    "Особенности модели": "Особливості моделі",
    "Декорирование": "Декорування",
    "Тип ткани": "Тип тканини",
    "Событие": "Подія",
    "Ценовой сегмент": "Ціновий сегмент",
    "Материал наполнителя": "Матеріал наповнювача",
    "Комментарий к замерам": "Коментар до вимірів",
    "Вид обуви": "Вид взуття",
    "Каблук/Подошва": "Каблук/Підошва",
    "Высота голенища/задника": "Висота халяви/задника",
    "Мембрана": "Мембрана",
    "Узоры и принты": "Візерунки і принти",
    "Форма мыска/носка": "Форма миска/носка",
    "Размер женской одежды": "Розмір жіночого одягу",
    "Фасон выреза горловины": "Фасон вирізу горловини",
    "Фасон рукава": "Фасон рукава",
    "Тип носка": "Тип носка",
    "Стирка": "Прання",

    # Сумки, кошельки, ремни, галантерея (KIRS / BAGSROOM)
    "Тип сумки": "Тип сумки",
    "Форма": "Форма",
    "Жесткость": "Жорсткість",
    "Жорсткість": "Жорсткість",
    "Отделения": "Відділення",
    "Відділення": "Відділення",
    "Ручки": "Ручки",
    "Плечевой ремень": "Плечовий ремінь",
    "Плечовий ремінь": "Плечовий ремінь",
    "Длина": "Довжина",
    "Довжина": "Довжина",
    "Ширина": "Ширина",
    "Высота": "Висота",
    "Высота ": "Висота",
    "Висота": "Висота",
    "Глубина": "Глибина",
    "Глубина ": "Глибина",
    "Глибина": "Глибина",
    "Толщина": "Товщина",
    "Товщина": "Товщина",
    "Вид кожи": "Вид шкіри",
    "Вид кожи по способу отделки": "Вид шкіри за способом обробки",
    "Тип застежки": "Тип застібки",
    "Подкладка": "Підкладка",
    "Количество отделений": "Кількість відділень",
    "Количество отделений для кредитных карт": "Кількість відділень для кредитних карт",
    "Количество отделений для купюр": "Кількість відділень для купюр",
    "Количество отделений для монет": "Кількість відділень для монет",
    "Количество отделений для документов": "Кількість відділень для документів",
    "Количество отделений на молнии": "Кількість відділень на блискавці",
    "Количество прозрачных отделений": "Кількість прозорих відділень",
    "Количество вмещаемых карт": "Кількість вміщуваних карт",
    "Расположение монетницы": "Розташування монетниці",
    "Длина ручек": "Довжина ручок",
    "Длина плечевого ремня": "Довжина плечового ременя",
    "Длина ремня": "Довжина ременя",
    "Замок": "Замок",
    "Карманы": "Кишені",
    "Материал пряжки": "Матеріал пряжки",
    "Международный размер": "Міжнародний розмір",
    "Модель": "Модель",
    "Мoдель": "Модель",
    "Цвет фурнитуры": "Колір фурнітури",
    "Цвет пряжки": "Колір пряжки",
    "Цвет застежки": "Колір застібки",
    "Цвет меха": "Колір хутра",
    "Вид брелка": "Вид брелока",
    "Вид изделия": "Вид виробу",
    "Двусторонний ремень": "Двосторонній ремінь",
    "Размер женских ремней": "Розмір жіночих ременів",
    "Размер мужских ремней": "Розмір чоловічих ременів",
    "Размер ремня": "Розмір ременя",
    "Тип ремня": "Тип ременя",
    "Ширина ремня": "Ширина ременя",
    "Укорачивается": "Вкорочується",
    "Страховочный ремешок": "Страхувальний ремінець",
    "Телескопическая ручка": "Телескопічна ручка",
    "Крепление на тележку": "Кріплення на візок",
    "Максимальный обхват талии": "Максимальний обхват талії",

    # Спорт и мячи (FOOTBALLERS)
    "Классификация бутс": "Класифікація бутс",
    "Полнота обуви": "Повнота взуття",
    "Шнуровка": "Шнурівка",
    "Съемные шипы": "Знімні шипи",
    "Вращающиеся шипы": "Обертові шипи",
    "Специальные покрытия бутс": "Спеціальні покриття бутс",
    "Защитный язык шнуровки": "Захисний язичок шнурівки",
    "Способ сшивки мяча": "Спосіб зшивання м'яча",
    "Вес": "Вага",
    "Возрастная группа": "Вікова група",
    "Вид стельки": "Вид устілки",
    "Супинатор": "Супінатор",
    "Материал шипов": "Матеріал шипів",
    "Вид спорта": "Вид спорту",
    "Тип мяча": "Тип м'яча",
    "Тип соединения панелей": "Тип з'єднання панелей",

    # Электроника, техника, инструменты (SPECULANT / BAGSROOM)
    "Беспроводные возможности": "Бездротові можливості",
    "Вид оборудования": "Вид обладнання",
    "Входная вилка": "Вхідна вилка",
    "Выключатель": "Вимикач",
    "Выходная мощность": "Вихідна потужність",
    "Выходная розетка": "Вихідна розетка",
    "Выходное напряжение (output)": "Вихідна напруга (output)",
    "Диаметр": "Діаметр",
    "Длина кабеля": "Довжина кабелю",
    "Длина шнура": "Довжина шнура",
    "Дополнительные функции": "Додаткові функції",
    "Емкость аккумулятора": "Ємність акумулятора",
    "Заземление": "Заземлення",
    "Защита от короткого замыкания": "Захист від короткого замикання",
    "Индикатор работы": "Індикатор роботи",
    "Индикация времени": "Індикація часу",
    "Интерфейс подключения": "Інтерфейс підключення",
    "Интерфейсы": "Інтерфейси",
    "Керамическое покрытие": "Керамічне покриття",
    "Количество выходных розеток": "Кількість вихідних розеток",
    "Количество насадок": "Кількість насадок",
    "Количество отверстий для крепления": "Кількість отворів для кріплення",
    "Количество светодиодов": "Кількість світлодіодів",
    "Комплектация инструментами/аксесуарами": "Комплектація інструментами/аксесуарами",
    "Комплектация инструментами/аксессуарами": "Комплектація інструментами/аксесуарами",
    "Корпус": "Корпус",
    "Максимальное время установки таймера": "Максимальний час встановлення таймера",
    "Максимальный вес": "Максимальна вага",
    "Марка": "Марка",
    "Материал колбы": "Матеріал колби",
    "Материал корпуса": "Матеріал корпусу",
    "Материал корпуса термоса": "Матеріал корпусу термоса",
    "Материал платформы": "Матеріал платформи",
    "Материал подошвы утюга": "Матеріал підошви праски",
    "Мощность лампы для сушки ногтей": "Потужність лампи для сушіння нігтів",
    "Набор переходников": "Набір перехідників",
    "Наличие колес": "Наявність коліс",
    "Наличие магнитов": "Наявність магнітів",
    "Напряжение": "Напруга",
    "Напряжение питания": "Напруга живлення",
    "Напряжение сети": "Напруга мережі",
    "Номинал": "Номінал",
    "Номинальное напряжение": "Номінальна напруга",
    "Объем": "Об'єм",
    "Объем емкости для молотого кофе": "Об'єм ємності для меленої кави",
    "Петелька для подвешивания": "Петелька для підвішування",
    "Поддерживаемые карты памяти": "Підтримувані карти пам'яті",
    "Подключение устройства": "Підключення пристрою",
    "Постоянная подача пара": "Постійна подача пари",
    "Потребляемая мощность": "Споживана потужність",
    "Радио": "Радіо",
    "Размеры": "Розміри",
    "Регулировка подачи пара": "Регулювання подачі пари",
    "Регулировка размера": "Регулювання розміру",
    "Сечение жил": "Перетин жил",
    "Сила выходного тока": "Сила вихідного струму",
    "Совместимость с маркой": "Сумісність з маркою",
    "Суммарная мощность звука": "Сумарна потужність звуку",
    "Терморегулятор": "Терморегулятор",
    "Тип": "Тип",
    "Тип весов": "Тип вагів",
    "Тип зажигалки": "Тип запальнички",
    "Тип кабеля": "Тип кабелю",
    "Тип кофемолки": "Тип кавомолки",
    "Тип лампы": "Тип лампи",
    "Тип механизма": "Тип механізму",
    "Тип миксера": "Тип міксера",
    "Тип питания": "Тип живлення",
    "Тип привода": "Тип приводу",
    "Тип разъема": "Тип роз'єму",
    "Тип термоса": "Тип термоса",
    "Тип удлинителя": "Тип подовжувача",
    "Тип часов": "Тип годинника",
    "Форм-фактор": "Форм-фактор",
    "Формат": "Формат",
    "Формат акустики": "Формат акустики",
    "Функция автоотключения": "Функція автовідключення",
    "Цвет корпуса": "Колір корпусу",
    "Цвет свечения": "Колір світіння",
    "Цветовая температура": "Колірна температура",
    "Цоколь": "Цоколь",
    "Диагональ экрана ноутбука": "Діагональ екрана ноутбука",

    # Мебель, текстиль, дом (BAGSROOM)
    "Вес стола": "Вага столу",
    "Вид стола по своему назначению": "Вид столу за своїм призначенням",
    "Высота от пола": "Висота від підлоги",
    "Высота подушки": "Висота подушки",
    "Высота сиденья стульчика": "Висота сидіння стільчика",
    "Высота стеллажа": "Висота стелажа",
    "Высота стола": "Висота столу",
    "Высота стульчика с учетом спинки": "Висота стільчика з урахуванням спинки",
    "Гарантийный срок": "Гарантійний термін",
    "Глубина стеллажа": "Глибина стелажа",
    "Глубина стола": "Глибина столу",
    "Глубина столешницы": "Глибина стільниці",
    "Глубина стульчика": "Глибина стільчика",
    "Декоративные элементы": "Декоративні елементи",
    "Длина подушки": "Довжина подушки",
    "Длина стола": "Довжина столу",
    "Длина столешницы": "Довжина стільниці",
    "Зеркало в комплекте": "Дзеркало в комплекті",
    "Количество выдвижных ящиков": "Кількість висувних ящиків",
    "Количество мест": "Кількість місць",
    "Количество столов в группе": "Кількість столів у групі",
    "Количество тумб": "Кількість тумб",
    "Комплектация подлокотниками": "Комплектація підлокітниками",
    "Материал стола": "Матеріал столу",
    "Наличие надстройки": "Наявність надбудови",
    "Наличие спинки": "Наявність спинки",
    "Наполнитель": "Наповнювач",
    "Ножки": "Ніжки",
    "Опоры": "Опори",
    "Ортопедические свойства": "Ортопедичні властивості",
    "Основной материал наволочки": "Основний матеріал наволочки",
    "Полки": "Полиці",
    "Ростовая группа": "Ростова група",
    "Сиденье": "Сидіння",
    "Спинка": "Спинка",
    "Столешница": "Стільниця",
    "Стульчик": "Стільчик",
    "Тумба в комплекте": "Тумба в комплекті",
    "Фасады": "Фасади",
    "Форма стеллажа": "Форма стелажа",
    "Форма фасада": "Форма фасаду",
    "Форма шкафа": "Форма шафи",
    "Цвет (оттенок) мебели": "Колір (відтінок) меблів",
    "Ширина подушки": "Ширина подушки",
    "Ширина стеллажа": "Ширина стелажа",
    "Ширина стульчика": "Ширина стільчика",

    # Прочее / игры / уход
    "Вид творчества": "Вид творчості",
    "Возраст": "Вік",
    "Жанр": "Жанр",
    "Кий": "Кий",
    "Количество деталей": "Кількість деталей",
    "Количество предметов в наборе": "Кількість предметів в наборі",
    "Количество цветов в наборе": "Кількість кольорів у наборі",
    "Количество шаров в комплекте": "Кількість куль у комплекті",
    "Мем": "Мем",
    "Отделка и украшения": "Оздоблення та прикраси",
    "Персонажи": "Персонажі",
    "Расположение": "Розташування",
    "Тематика": "Тематика",
    "Тематика костюма": "Тематика костюма",
    "Треугольник": "Трикутник",
    "Упаковка": "Упаковка",
    "Упаковка средства": "Упаковка засобу",
    "Характеристика": "Характеристика",
    "Цвет игрового поля": "Колір ігрового поля",
    "Ширина упаковки": "Ширина упаковки",
    "Высота упаковки": "Висота упаковки",
    "Длина упаковки": "Довжина упаковки",
    "Количество в упаковке": "Кількість в упаковці",
    "Количество составных элементов": "Кількість складових елементів",
    "Назначение": "Призначення",
    "Особенности": "Особливості",
    "Тип назначения": "Тип призначення"
}

COLOR_MAP = {
    "черный": "Чорний", "чёрный": "Чорний", "чорний": "Чорний",
    "белый": "Білий", "білий": "Білий",
    "красный": "Червоний", "червоний": "Червоний",
    "синий": "Синій", "синій": "Синій",
    "голубой": "Блакитний", "блакитний": "Блакитний",
    "зеленый": "Зелений", "зелёный": "Зелений", "зелений": "Зелений",
    "желтый": "Жовтий", "жёлтый": "Жовтий", "жовтий": "Жовтий",
    "серый": "Сірий", "сірий": "Сірий",
    "коричневый": "Коричневий", "коричневий": "Коричневий",
    "бежевый": "Бежевий", "бежевий": "Бежевий",
    "оранжевый": "Помаранчевий", "помаранчевий": "Помаранчевий",
    "фиолетовый": "Фіолетовий", "фіолетовий": "Фіолетовий",
    "розовый": "Рожевий", "рожевий": "Рожевий",
    "бордовый": "Бордовий", "бордовий": "Бордовий",
    "золотистый": "Золотистий", "золотой": "Золотистий", "золотий": "Золотистий",
    "серебристый": "Сріблястий", "серебряный": "Сріблястий", "сріблястий": "Сріблястий",
    "хаки": "Хакі", "хакі": "Хакі",
    "бирюзовый": "Бірюзовий", "бірюзовий": "Бірюзовий",
    "разноцветный": "Різнокольоровий", "різнокольоровий": "Різнокольоровий"
}


def normalize_characteristic_value(char_name: str, value: str) -> str:
    """
    Нормализует значения характеристик под эталонные словари Prom.ua
    для попадания в основные фильтры каталога (а не в пользовательские).
    """
    if not value or not isinstance(value, str):
        return value

    val_clean = value.strip()
    val_lower = val_clean.lower()

    # Пол / Стать
    if char_name in ["Стать", "Пол"]:
        if any(w in val_lower for w in ["чоловіч", "мужск", "муж."]) and any(w in val_lower for w in ["жіноч", "женск", "жен."]):
            return "Унісекс"
        if any(w in val_lower for w in ["муж", "чоловіч"]):
            return "Чоловічий"
        if any(w in val_lower for w in ["жен", "жіноч"]):
            return "Жіночий"
        if any(w in val_lower for w in ["унисекс", "унісекс"]):
            return "Унісекс"
        if any(w in val_lower for w in ["детск", "дитяч"]):
            return "Дитячий"

    # Сезон
    elif char_name in ["Сезон"]:
        if any(w in val_lower for w in ["демисезон", "демісезон", "весна", "осень", "осінь"]):
            return "Демісезон"
        if any(w in val_lower for w in ["зим", "зима"]):
            return "Зима"
        if any(w in val_lower for w in ["лет", "літо", "літн"]):
            return "Літо"
        if any(w in val_lower for w in ["всесезон", "круглый"]):
            return "Всесезонний"

    # Состояние / Стан
    elif char_name in ["Стан", "Состояние"]:
        if any(w in val_lower for w in ["нов", "новое", "новый", "новий"]):
            return "Новий"
        if any(w in val_lower for w in ["б/у", "б/в", "вживан"]):
            return "Б/в"

    # Цвет / Колір
    elif char_name in ["Колір", "Група кольору", "Цвет", "Группа цвета"]:
        if val_lower in COLOR_MAP:
            return COLOR_MAP[val_lower]
        for delim in [', ', '; ', ' / ', '/', ',']:
            if delim in val_clean:
                parts = [p.strip() for p in val_clean.split(delim)]
                mapped = [COLOR_MAP.get(p.lower(), p.capitalize()) for p in parts]
                sep = ', ' if delim in [', ', ','] else (' / ' if delim in ['/', ' / '] else delim)
                return sep.join(mapped)

    # Застежка / Застібка
    elif char_name in ["Застібка", "Застежка", "Тип застежки"]:
        if "молни" in val_lower or "блискав" in val_lower:
            return "Блискавка"
        if "пуговиц" in val_lower or "ґудзик" in val_lower:
            return "Ґудзики"
        if "шнуров" in val_lower:
            return "Шнурівка"
        if "липучк" in val_lower:
            return "Липучка"
        if "кнопк" in val_lower:
            return "Кнопки"
        if "без застеж" in val_lower or "без застіб" in val_lower:
            return "Без застібки"

    return val_clean


def get_paired_insole_val(pair_str: str) -> str:
    """Возвращает длину по стельке в см для спаренных размеров обуви/тапочек."""
    insole_map = {
        '35-36': '22,5 см',
        '36-37': '23,5 см',
        '37-38': '24 см',
        '38-39': '24,5 см',
        '39-40': '25 см',
        '40-41': '25,5 см',
        '41-42': '26,5 см',
        '42-43': '27,5 см',
        '43-44': '28 см',
        '44-45': '28,5 см',
        '45-46': '29 см',
        '46-47': '29,5 см',
    }
    pair_clean = pair_str.strip()
    if pair_clean in insole_map:
        return insole_map[pair_clean]
    try:
        parts = [int(p) for p in pair_clean.split('-')]
        avg_s = sum(parts) / len(parts)
        val = 23.5 + (avg_s - 36.5) * 0.75
        val = round(val * 2) / 2.0
        return f"{str(val).replace('.', ',')} см"
    except Exception:
        return ""


def clean_counterfeit_words(text: str) -> str:
    """
    Удаляет слова 'реплика', 'репліка', 'копия', 'люкс реплика', '1:1',
    заменяя их на нейтральные термины ('модель'), чтобы описания
    не блокировались автоматическим фильтром контрафакта Prom.ua.
    """
    if not text:
        return ""
    text = re.sub(r'(?i)\s*\((?:дуже\s+)?(?:високо)?якісна\s+реплік[аи]\)', '', text)
    text = re.sub(r'(?i)\s*\((?:люкс\s+)?реплік[аи]\)', '', text)
    text = re.sub(r'(?i)\s*\((?:люкс\s+)?реплик[аи]\)', '', text)
    text = re.sub(r'(?i)\s*\(1:1\)', '', text)
    text = re.sub(r'(?i)\bвисокоякісна\s+репліка\s+(?:на\s+)?', 'Високоякісна модель ', text)
    text = re.sub(r'(?i)\bдуже\s+якісна\s+репліка\s+(?:на\s+)?', 'Якісна модель ', text)
    text = re.sub(r'(?i)\bякісна\s+репліка\s+(?:на\s+)?', 'Якісна модель ', text)
    text = re.sub(r'(?i)\bлюкс\s+репліка\s+(?:на\s+)?', 'Люкс модель ', text)
    text = re.sub(r'(?i)\bреплік[аиуе]\s+(?:на\s+)?', 'модель ', text)
    text = re.sub(r'(?i)\bреплик[аиуе]\s+(?:на\s+)?', 'модель ', text)
    text = re.sub(r'(?i)\bкопі[яєїю]\s+(?:на\s+)?', 'модель ', text)
    text = re.sub(r'(?i)\bкопи[яеию]\s+(?:на\s+)?', 'модель ', text)
    return text


def repair_and_format_size_table(html_content: str, is_ukr: bool = False) -> str:
    """
    Форматирует HTML-таблицы замеров для аккуратного отображения на Prom.ua
    (границы ячеек, отступы, центрирование).
    Также устраняет системную ошибку поставщика AGER для тапочек/обуви со сдвоенными размерами
    (когда спаренные размеры 36-37 ошибочно разбивались на 36 в размер и 37 в стельку),
    восстанавливая строку спаренных размеров и строку реальных замеров по стельке в см.
    """
    if not html_content or '<table' not in html_content:
        return html_content

    def fix_table_match(match):
        table_html = match.group(0)

        header_title = "Заміри (фактичні заміри можуть відрізнятися +- 2 см)" if is_ukr else "Замеры (фактические замеры могут отличаться +- 2 см)"
        m_head = re.search(r"""<td[^>]*class=["\']?thead_td["\']?[^>]*>(.*?)</td>""", table_html, re.DOTALL | re.IGNORECASE)
        if not m_head:
            m_head = re.search(r"""<th[^>]*>(.*?)</th>""", table_html, re.DOTALL | re.IGNORECASE)
        if m_head:
            header_title = re.sub(r'<[^>]+>', '', m_head.group(1)).strip()

        rows = re.findall(r"""<tr[^>]*class=["\']?tbody_tr["\']?[^>]*>(.*?)</tr>""", table_html, re.DOTALL | re.IGNORECASE)
        if not rows:
            rows = re.findall(r"""<tr[^>]*>(.*?)</tr>""", table_html, re.DOTALL | re.IGNORECASE)
            if rows and header_title in rows[0]:
                rows = rows[1:]

        parsed_rows = []
        for r in rows:
            first_m = re.search(r"""<td[^>]*class=["\']?tbody_first-td["\']?[^>]*>(.*?)</td>""", r, re.DOTALL | re.IGNORECASE)
            if first_m:
                row_label = re.sub(r'<[^>]+>', '', first_m.group(1)).strip()
            else:
                cells = re.findall(r"""<td[^>]*>(.*?)</td>""", r, re.DOTALL | re.IGNORECASE)
                row_label = re.sub(r'<[^>]+>', '', cells[0]).strip() if cells else ''

            tds = re.findall(r"""<td[^>]*class=["\']?tbody_td["\']?[^>]*>(.*?)</td>""", r, re.DOTALL | re.IGNORECASE)
            if not tds:
                all_cells = re.findall(r"""<td[^>]*>(.*?)</td>""", r, re.DOTALL | re.IGNORECASE)
                tds = [re.sub(r'<[^>]+>', '', c).strip() for c in all_cells[1:]]
            else:
                tds = [re.sub(r'<[^>]+>', '', c).strip() for c in tds]

            if row_label or tds:
                parsed_rows.append((row_label, tds))

        # Случай 1: Сырой баг AGER (Строка 1: [36, 38, 40], Строка 2: [37, 39, 41])
        if len(parsed_rows) == 2:
            r1_label, r1_vals = parsed_rows[0]
            r2_label, r2_vals = parsed_rows[1]
            if len(r1_vals) == len(r2_vals) and len(r1_vals) > 0:
                try:
                    n1 = [int(v.replace(',', '.').split()[0]) for v in r1_vals]
                    n2 = [int(v.replace(',', '.').split()[0]) for v in r2_vals]
                    diffs = [b - a for a, b in zip(n1, n2)]
                    if all(d == 1 for d in diffs):
                        paired_sizes = [f"{a}-{b}" for a, b in zip(n1, n2)]
                        size_label = "Розмір" if is_ukr else "Размер"
                        insole_label = "Довжина по устілці (взуття)" if is_ukr else "Длина по стельке (обувь)"
                        insole_vals = [get_paired_insole_val(ps) for ps in paired_sizes]
                        parsed_rows = [
                            (size_label, paired_sizes),
                            (insole_label, insole_vals)
                        ]
                except (ValueError, IndexError):
                    pass

        # Случай 2: Уже свернутая таблица из 1 строки со спаренными размерами (36-37, 41-42) - добавляем строку стельки
        elif len(parsed_rows) == 1:
            r1_label, r1_vals = parsed_rows[0]
            if any('-' in str(v) for v in r1_vals):
                insole_label = "Довжина по устілці (взуття)" if is_ukr else "Длина по стельке (обувь)"
                insole_vals = [get_paired_insole_val(str(ps)) for ps in r1_vals]
                parsed_rows.append((insole_label, insole_vals))

        if not parsed_rows:
            return table_html

        col_count = max(len(r[1]) for r in parsed_rows) + 1

        res = [
            f'<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #dcdcdc; text-align: center; margin: 12px 0;">',
            f'  <thead>',
            f'    <tr style="background-color: #f7f7f7; font-weight: bold;">',
            f'      <th colspan="{col_count}" style="padding: 8px 12px; border: 1px solid #dcdcdc; text-align: center; font-size: 13px;">{header_title}</th>',
            f'    </tr>',
            f'  </thead>',
            f'  <tbody>'
        ]

        for r_label, r_vals in parsed_rows:
            res.append('    <tr>')
            res.append(f'      <td style="padding: 7px 10px; border: 1px solid #dcdcdc; background-color: #fafafa; font-weight: 600; text-align: left; font-size: 13px;">{r_label}</td>')
            for val in r_vals:
                res.append(f'      <td style="padding: 7px 10px; border: 1px solid #dcdcdc; text-align: center; font-size: 13px;">{val}</td>')
            res.append('    </tr>')

        res.append('  </tbody>')
        res.append('</table>')

        return '\n'.join(res)

    cleaned = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r'<table[^>]*>.*?</table>', fix_table_match, cleaned, flags=re.DOTALL | re.IGNORECASE)
    return cleaned.strip()


def sanitize_description(text: str, is_ukr: bool = False) -> str:
    """
    Санитизирует описание товаров под требования модерации Prom.ua:
    1. Удаляет опасные теги (<script>, <style>, <iframe>, <object>, <embed>).
    2. Удаляет внешние ссылки (http, https, ссылки на поставщиков, телеграм, инстаграм).
    3. Раскрывает теги <a> в чистый текст.
    4. Удаляет контактные номера телефонов.
    5. Форматирует и нормализует таблицы замеров.
    6. Очищает запрещенные слова контрафакта ('реплика', 'копия' и др.).
    """
    if not text or pd.isna(text):
        return ""
    text = str(text)

    # 0. Очистка слов контрафакта/реплик для прохождения модерации Prom.ua
    text = clean_counterfeit_words(text)
    # 1. Удаление тегов скриптов и стилей
    text = re.sub(r'<(script|style|iframe|object|embed)[^>]*>.*?</\1>', '', text, flags=re.DOTALL | re.IGNORECASE)

    # 2. Форматирование таблиц замеров
    if '<table' in text.lower():
        text = repair_and_format_size_table(text, is_ukr=is_ukr)

    # 3. Раскрытие тегов <a> в обычный текст
    text = re.sub(r'<a\s+[^>]*>(.*?)</a>', r'\1', text, flags=re.DOTALL | re.IGNORECASE)

    # 4. Удаление прямых ссылок
    url_pattern = r'https?://[^\s<>"\'\]]+|www\.[^\s<>"\'\]]+'
    text = re.sub(url_pattern, '', text, flags=re.IGNORECASE)

    # 5. Удаление упоминаний доменов поставщиков и мессенджеров
    domains = [
        r'ager\.ua[^\s<>"\'\]]*', r'mydrop\.com\.ua[^\s<>"\'\]]*', r'spekylant\.com\.ua[^\s<>"\'\]]*',
        r'bagsroom\.com\.ua[^\s<>"\'\]]*', r'kirs\.com\.ua[^\s<>"\'\]]*', r'easydrop\.one[^\s<>"\'\]]*',
        r't\.me/[^\s<>"\'\]]+', r'instagram\.com/[^\s<>"\'\]]+', r'facebook\.com/[^\s<>"\'\]]+'
    ]
    for d in domains:
        text = re.sub(d, '', text, flags=re.IGNORECASE)

    # 6. Удаление номеров телефонов (+380..., (067)...)
    phone_pattern = r'(?:\+?38[\s\-]?)?\(?0\d{2}\)?[\s\-\.]*\d{3}[\s\-\.]*\d{2}[\s\-\.]*\d{2}'
    text = re.sub(phone_pattern, '', text)

    # 7. Очистка лишних пустых строк и тегов переноса
    text = re.sub(r'(<br\s*/?>\s*){3,}', '<br><br>', text, flags=re.IGNORECASE)
    text = re.sub(r'(<p>\s*(&nbsp;|\s)*</p>\s*)+', '', text, flags=re.IGNORECASE)

    return text.strip()


def unify_variation_groups(df: pd.DataFrame) -> pd.DataFrame:
    """
    Синхронизирует параметры группы магазина и категорию Prom.ua для всех позиций,
    имеющих одинаковый ID_групи_різновидів.
    Правило Prom.ua: у основного товара и всех его разновидностей обязаны быть 100% одинаковыми:
      - Номер_групи
      - Назва_групи
      - Ідентифікатор_групи
      - Ідентифікатор_підрозділу
    """
    if df.empty or 'ID_групи_різновидів' not in df.columns:
        return df

    has_var = df['ID_групи_різновидів'].notna() & (df['ID_групи_різновидів'].astype(str).str.strip() != '') & (df['ID_групи_різновидів'].astype(str).str.strip() != '0')
    if not has_var.any():
        return df

    logger.info(f"Синхронизация групп разновидностей для {has_var.sum()} позиций...")

    var_groups = df[has_var].groupby(df.loc[has_var, 'ID_групи_різновидів'].astype(str).str.strip())

    canonical_data = {}
    for gid, group in var_groups:
        exact_match = group[group['Ідентифікатор_товару'].astype(str).str.strip() == gid]
        if not exact_match.empty:
            parent_row = exact_match.iloc[0]
        else:
            with_cat = group[group['Ідентифікатор_підрозділу'].notna() & (group['Ідентифікатор_підрозділу'].astype(str).str.strip() != '')]
            if not with_cat.empty:
                parent_row = with_cat.iloc[0]
            else:
                parent_row = group.iloc[0]

        c_num = str(parent_row.get('Номер_групи') or '').strip()
        c_name = str(parent_row.get('Назва_групи') or '').strip()
        c_gid = str(parent_row.get('Ідентифікатор_групи') or c_num).strip()
        c_sub = str(parent_row.get('Ідентифікатор_підрозділу') or '').strip()

        if not c_num:
            for _, r in group.iterrows():
                val = str(r.get('Номер_групи') or '').strip()
                if val:
                    c_num = val
                    c_name = str(r.get('Назва_групи') or '').strip()
                    c_gid = str(r.get('Ідентифікатор_групи') or val).strip()
                    break

        if not c_sub:
            for _, r in group.iterrows():
                val = str(r.get('Ідентифікатор_підрозділу') or '').strip()
                if val:
                    c_sub = val
                    break

        canonical_data[gid] = {
            'Номер_групи': c_num,
            'Назва_групи': c_name,
            'Ідентифікатор_групи': c_gid,
            'Ідентифікатор_підрозділу': c_sub
        }

    var_ids_series = df['ID_групи_різновидів'].astype(str).str.strip()

    num_updates = 0
    for col in ['Номер_групи', 'Назва_групи', 'Ідентифікатор_групи', 'Ідентифікатор_підрозділу']:
        col_map = {gid: data[col] for gid, data in canonical_data.items() if data[col]}
        mapped = var_ids_series.map(col_map)
        mask = has_var & mapped.notna() & (mapped != '')
        changes = (df.loc[mask, col].astype(str).str.strip() != mapped[mask]).sum()
        if changes > 0:
            logger.info(f"Колонка {col}: скорректировано {changes} разновидностей для совпадения с основным товаром.")
            df.loc[mask, col] = mapped[mask]
            num_updates += changes

    logger.info(f"Синхронизация групп разновидностей завершена. Всего исправлений: {num_updates}")
    return df



# ══════════════════════════════════════════════════════════════════════════════
# ЭТАП 3: Жесткий маппинг категорий поставщиков на портальные категории Prom.ua
# для гарантированного попадания в целевые разделы и исключения лишней комиссии
# ══════════════════════════════════════════════════════════════════════════════

CATEGORY_PORTAL_MAP = {
    # ── AGER (18 категорий обуви) ─────────────────────────────────────────────
    "145": "35809",        # Мужская обувь -> чоловіче взуття, загальне
    "320": "32299",        # Женская обувь -> жіноче взуття, загальне
    "574": "35702",        # Женские босоножки -> жіночі босоніжки та сандалі
    "578": "32205",        # Мужские вьетнамки, шлепанцы -> чоловічі сандалі і шльопанці
    "579": "3220713",      # Мужские кроссовки -> кросівки та кеди
    "580": "3220713",      # Мужские кеды -> кросівки та кеди
    "581": "3220713",      # Женские кеды -> кросівки та кеди
    "582": "3220713",      # Женские кроссовки -> кросівки та кеди
    "584": "34304",        # Женские вьетнамки, сланцы -> взуття для пляжу та басейну
    "649": "3220404",      # Женские ботинки, ботильоны -> жіночі ботильйони, черевики
    "697": "32206",        # Женские тапочки -> Жіночі домашні капці
    "705": "3220416",      # Угги женские -> унти, уггі, сноубутси
    "710": "35804",        # Мужские слипоны -> чоловічі мокасини
    "744": "3220413",      # Мужские ботинки -> чоловічі черевики
    "751": "35805",        # Мужчкие тапочки -> чоловічі домашні капці
    "755": "35703",        # Женские лоферы -> жіночі мокасини і сліпони
    "756": "35702",        # Женские сабо, вьетнамки, шлепанцы -> жіночі босоніжки та сандалі
    "757": "35702",        # Женские шлепанцы -> жіночі босоніжки та сандалі

    # ── FOOTBALLERS (4 категории спорта) ──────────────────────────────────────
    "1740596412029": "3220707", # ФУТБОЛЬНІ БУТСИ -> футбольне взуття
    "139085133": "2023",         # ФУТБОЛЬНІ МЯЧІ -> спортивні ігрові м'ячі
    "156096408": "2023",         # Футбольні мʼячі (Не бренд) -> спортивні ігрові м'ячі
    "ФУТБОЛЬНІ ГЕТРИ ТА ТРИМАЧІ": "203303", # ФУТБОЛЬНІ ГЕТРИ -> футбольні гетри

    # ── IZIDROP (29 категорий обуви + защита от ловушек SALE / РОЗПРОДАЖ) ─────
    "174059646768": "3220713",  # NIKE чоловічі -> кросівки та кеди
    "174059647649": "3220713",  # ASICS чоловічі -> кросівки та кеди
    "174059647650": "3220713",  # NIKE жіночі -> кросівки та кеди
    "174059647652": "3220713",  # Зимові кросівки -> кросівки та кеди
    "174059647698": "3220713",  # ADIDAS жіночі -> кросівки та кеди
    "174059647699": "3220713",  # ADIDAS чоловічі -> кросівки та кеди
    "174059647730": "3220713",  # New Balance жіночі -> кросівки та кеди
    "174059647731": "3220713",  # New Balance чоловічі -> кросівки та кеди
    "174059648026": "3220713",  # REEBOK чоловічі -> кросівки та кеди
    "174059648027": "3220713",  # REEBOK жіночі -> кросівки та кеди
    "174059648028": "3220713",  # PUMA чоловічі -> кросівки та кеди
    "174059648053": "3220713",  # PUMA жіночі -> кросівки та кеди
    "174059648140": "3220713",  # Salomon чоловічі -> кросівки та кеди
    "174059648528": "3220713",  # ASICS жіночі -> кросівки та кеди
    "174059648878": "3220713",  # COLUMBIA чоловічі -> кросівки та кеди
    "174059649253": "3220713",  # Без бренду -> кросівки та кеди
    "1740596410588": "3220713", # SALE чоловічі (розпродаж) -> кросівки та кеди
    "1740596410589": "3220713", # SALE жіночі (розпродаж) -> кросівки та кеди
    "1740596412985": "3220713", # HOKA чоловічі -> кросівки та кеди
    "1740596412986": "3220713", # HOKA жіночі -> кросівки та кеди
    "1740596413319": "3220713", # Balenciaga жіночі -> кросівки та кеди
    "1740596413344": "34304",   # CROCS -> взуття для пляжу та басейну
    "1740596413789": "3220713", # Salomon жіночі -> кросівки та кеди
    "1740596414307": "3220713", # LACOSTE чоловічі -> кросівки та кеди
    "1740596415336": "3220713", # ECCO -> кросівки та кеди
    "1740596415477": "3220413", # SALE (розпродаж) Зимові (winter) -> чоловічі черевики
    "1740596416263": "3220713", # РОЗПРОДАЖ 2026 року (SALE 2026) -> кросівки та кеди
    "1740596416718": "3220713", # UNDER ARMOUR -> кросівки та кеди
    "1740596416802": "3220713", # ON Cloudtilt -> кросівки та кеди

    # ── KIRS (21 категория — галантерея, сумки, брелоки, детские рюкзаки) ─────
    "21050159": "380601",  # Рюкзаки городские женские -> міські та спортивні рюкзаки
    "21050946": "380630",  # Спортивные сумки -> спортивні сумки
    "21050128": "380630",  # Спортивные сумки и рюкзаки -> спортивні сумки
    "21509797": "380650",  # Бананки - сумки на пояс -> поясні сумки
    "21001911": "152411",  # Мужские сумки и барсетки -> чоловічі сумки
    "14709295": "152405",  # Кошельки и портмоне -> гаманці та портмоне
    "14709554": "152411",  # Сумки, рюкзаки, портфели -> чоловічі сумки
    "22449458": "380501",  # Сумки и рюкзаки детские (игрушки) -> дитячі сумки і рюкзаки
    "14628441": "33911",   # Галантерея и аксессуары -> галантерея та аксесуари, загальне
    "97968928": "380110",  # Сумки и рюкзаки для ноутбуков -> сумки та рюкзаки для ноутбуків
    "22449719": "152403",  # Женские сумочки и клатчи -> жіночі сумочки і клатчі
    "32214469": "380601",  # Рюкзаки спортивные -> міські та спортивні рюкзаки
    "76824950": "380650",  # Мужские сумки бананки на грудь (Кросс боди) -> поясні сумки
    "43986659": "380601",  # Мужские городские рюкзаки -> міські та спортивні рюкзаки
    "103087291": "32801",  # Женские ремни -> ремені та пояси
    "103087295": "32801",  # Мужские ремни -> ремені та пояси
    "25276002": "152402",  # Портфели деловые -> портфелі ділові
    "21509203": "32801",   # Поясные ремни мужские и женские -> ремені та пояси
    "25111628": "170702",  # Брелки меховые -> брелоки
    "102220419": "33911",  # Галантерея и аксессуары -> галантерея та аксесуари, загальне
    "107622359": "152405", # Кошельки и портмоне -> гаманці та портмоне

    # ── MOYDROP (13 категорий обуви) ──────────────────────────────────────────
    "68072": "3220713",  # ЧОЛОВІЧЕ ВЗУТТЯ (ДЕМІСЕЗОН) -> кросівки та кеди
    "68074": "3220713",  # Жіноче Демісезон -> кросівки та кеди
    "80847": "3221214",  # Дитяче Взуття Демісезон -> дитячі та підліткові кросівки, кеди
    "93016": "3220413",  # ЗИМА ЧОЛОВІЧЕ -> чоловічі черевики
    "93017": "3220413",  # ЗИМА -> чоловічі черевики
    "169399": "3220416", # ЗИМОВІ ЖІНОЧІ УГГ/ СНОУБУТСИ -> унти, уггі, сноубутси
    "169401": "3220404", # ЖІНОЧІ ЗИМОВІ ЧЕРЕВИКИ -> жіночі ботильйони, черевики
    "169402": "3220713", # ЖІНОЧІ ЗИМОВІ КРОСІВКИ/ КЕДИ/ЛОФЕРИ -> кросівки та кеди
    "169403": "35702",   # ЖІНОЧІ ЗИМОВІ КЛОГИ / САБО -> жіночі босоніжки та сандалі
    "169410": "34304",   # Тапки, босоніжки, крокси -> взуття для пляжу та басейну
    "169411": "3220713", # Кросівки -> кросівки та кеди
    "169412": "3220713", # Кеди -> кросівки та кеди
    "169413": "35703",   # Лофери /Мокасини -> жіночі мокасини і сліпони

    # ── BAGSROOM (21 категория сумок, аксессуаров и мебели) ───────────────────
    "127986225": "152405", # Складные -> гаманці та портмоне
    "124805235": "152411", # Месенджеры, барсетки -> чоловічі сумки
    "127986217": "32801",  # Ремни -> ремені та пояси
    "127986221": "32801",  # Женские ремни -> ремені та пояси
    "127986226": "152403", # Клатчи -> жіночі сумочки і клатчі
    "127986212": "152405", # Мужские кошельки -> гаманці та портмоне
    "127986218": "40505",  # Мягкое стекло -> скатертини, серветки сервірувальні
    "127986223": "15030325",# Подушки для стульчиков -> Меблеві подушки
    "127986224": "152405", # Зажимы -> гаманці та портмоне
    "124805236": "380601", # Рюкзаки -> міські та спортивні рюкзаки
    "127986215": "152413", # Обкладинки -> обкладинки для документів
    "127986216": "380310", # Ключницы -> чохли для ключів
    "124805233": "380650", # Слинги -> поясні сумки
    "124805229": "380601", # Женские кожаные рюкзаки -> міські та спортивні рюкзаки
    "124805232": "380650", # Бананки -> поясні сумки
    "124805234": "380110", # Портфели для ноутбуков и документов -> сумки та рюкзаки для ноутбуків
    "147863556": "371305", # Детские стульчики -> дитячі стільчики і табуретки
    "147863557": "371303", # Детские столики -> дитячі та підліткові столи
    "147862750": "370243", # Комплекты стульчики и столики -> Комплекти шкільних меблів
    "127986219": "371308", # Полочки для игрушек -> дитячі шафи та стелажі
    "127986222": "370243", # Детские столики и стульчики -> Комплекти шкільних меблів

    # ── BAYRAKTAR (48 категорий тактики, обуви и туризма) ─────────────────────
    "134414": "240410",  # Туристичне спорядження -> товари для туризму та подорожей, загальне
    "143323": "240423",  # Кемпінгова мебель -> туристичні складані меблі, загальне
    "143320": "240424",  # Кухня та туристичний посуд -> туристичний посуд, загальне
    "87276": "150410",   # Ліхтарі ручні та налобні -> ручні та налобні ліхтарі
    "134415": "240409",  # Намети та аксесуари -> туристичні намети і тенти
    "143324": "240439",  # Рушники швидкосохнучі -> туристичні рушники
    "153055": "152401",  # Рюкзаки та сумки -> Армійські спецсумки і рюкзаки
    "139627": "240407",  # Спальні мішки -> спальні мішки
    "143321": "1531",    # Термоси та термокружки -> термоси
    "143322": "240405",  # Туристичні коврики та каремати -> каремати та сидушки
    "151094": "240435",  # Чохли -> аксесуари для спальних мішків, карематів та наметів
    "134420": "240435",  # Аксесуари для палаток -> аксесуари для спальних мішків, карематів та наметів
    "134416": "240409",  # Експедиційні палатки -> туристичні намети і тенти
    "134418": "240409",  # Кемпінгові палатки -> туристичні намети і тенти
    "134417": "240409",  # Універсальні палатки -> туристичні намети і тенти
    "134419": "240409",  # Шатри, тенти, туристичні туалети -> туристичні намети і тенти
    "147181": "240423",  # Розкладушки та стільці -> туристичні складані меблі, загальне
    "152371": "240423",  # Туристичні розкладні ліжка -> туристичні складані меблі, загальне
    "147182": "240423",  # Туристичні розкладні столи -> туристичні складані меблі, загальне
    "33626": "690102",   # Берці -> Тактичне взуття
    "33625": "690102",   # Кросівки -> Тактичне взуття
    "33623": "690102",   # Напівчеревики -> Тактичне взуття
    "33622": "690102",   # Черевики -> Тактичне взуття
    "34086": "240449",   # Гідратори та питні системи -> питні системи, туристичні гідратори
    "33919": "152401",   # Рюкзаки та баули -> Армійські спецсумки і рюкзаки
    "33920": "380650",   # Сумки/органайзери -> поясні сумки
    "135079": "690102",  # GARMONT -> Тактичне взуття
    "109320": "690102",  # Берці Belleville USA -> Тактичне взуття
    "152456": "690102",  # Кросівки Belleville USA -> Тактичне взуття
    "88865": "690102",   # Тактичні берці LOWA -> Тактичне взуття
    "88868": "690102",   # Тактичні берці SALOMON -> Тактичне взуття
    "139170": "690102",  # Тактичні кросівки AKU -> Тактичне взуття
    "88866": "690102",   # Тактичні кросівки LOWA -> Тактичне взуття
    "123529": "690102",  # Тактичні кросівки MERRELL -> Тактичне взуття
    "88869": "690102",   # Тактичні кросівки SALOMON -> Тактичне взуття
    "139169": "690102",  # Тактичні черевики AKU -> Тактичне взуття
    "88864": "690102",   # Тактичні черевики LOWA -> Тактичне взуття
    "88867": "690102",   # Тактичні черевики SALOMON -> Тактичне взуття
    "125078": "690102",  # Черевики тактичні MERRELL -> Тактичне взуття
    "33630": "152401",   # Рюкзаки та сумки -> Армійські спецсумки і рюкзаки
    "82947": "152401",   # Рюкзаки, сумки та баули -> Армійські спецсумки і рюкзаки
    "81777": "380650",   # Сумки, бананки, органайзери -> поясні сумки
    "33621": "690102",   # Тактичне взуття -> Тактичне взуття
    "88863": "690102",   # Тактичне взуття LOWA, SALOMON, MERRELL -> Тактичне взуття
    "81797": "32205",    # Тапочки та вʼєтнамки -> чоловічі сандалі і шльопанці
    "139160": "3220704", # Трекінгове та туристичне взуття -> Туристичне взуття
    "81779": "152401"    # Штурмові рюкзаки Mil-Tec -> Армійські спецсумки і рюкзаки
}


def resolve_portal_category(supp_cat_id: str, title: str, supplier_name: str) -> str:
    """
    Принудительно сопоставляет категорию поставщика с точным ID категории Prom.ua (Ідентифікатор_підрозділу).
    Включает эвристическую защиту от аномалий названий (SALE / РОЗПРОДАЖ, плюшевые брелоки и т.д.).
    """
    cat_key = str(supp_cat_id).strip()
    if cat_key in CATEGORY_PORTAL_MAP:
        return CATEGORY_PORTAL_MAP[cat_key]

    title_lower = str(title).lower().strip()

    # Специфические эвристики для IZIDROP (только обувь)
    if supplier_name == "IZIDROP":
        if any(w in title_lower for w in ["зимов", "winter", "угг", "ботинк", "черевик"]):
            return "3220413"  # чоловічі черевики
        if any(w in title_lower for w in ["crocs", "крокс", "сланц", "шльоп", "сабо"]):
            return "34304"    # взуття для пляжу та басейну
        return "3220713"      # кросівки та кеди (дефолт для всего ассортимента IZIDROP)

    # Специфические эвристики для KIRS (галантерея, сумки, игрушки)
    if supplier_name == "KIRS":
        if any(w in title_lower for w in ["брелок", "брелоч"]):
            return "170702"  # брелоки
        if any(w in title_lower for w in ["іграшк", "игрушк", "плюшев"]):
            return "2604"    # м'які іграшки
        if any(w in title_lower for w in ["дитяч", "детск"]):
            return "380501"  # дитячі сумки і рюкзаки
        if any(w in title_lower for w in ["гаманець", "кошелек", "портмоне", "зажим", "клатч"]):
            return "152405"  # гаманці та портмоне
        if any(w in title_lower for w in ["бананка", "на пояс", "крос-боді", "кросс-боди"]):
            return "380650"  # поясні сумки
        if any(w in title_lower for w in ["рюкзак"]):
            return "380601"  # міські та спортивні рюкзаки
        if any(w in title_lower for w in ["ремінь", "ремень", "пояс"]):
            return "32801"   # ремені та пояси
        if any(w in title_lower for w in ["сумка", "барсетка", "мессенджер", "портфель"]):
            return "152411"  # чоловічі сумки

    # Специфические эвристики для BAGSROOM (сумки, рюкзаки, чемоданы)
    if supplier_name == "BAGSROOM":
        if any(w in title_lower for w in ["рюкзак"]):
            return "380601"
        if any(w in title_lower for w in ["гаманець", "кошелек", "портмоне"]):
            return "152405"
        if any(w in title_lower for w in ["бананка", "на пояс"]):
            return "380650"
        if any(w in title_lower for w in ["чемодан", "валіз"]):
            return "152404"  # дорожні сумки та валізи
        return "152411"

    # Специфические эвристики для BAYRAKTAR (тактическая обувь и амуниция)
    if supplier_name == "BAYRAKTAR":
        if any(w in title_lower for w in ["берці", "берцы"]):
            return "690102"
        if any(w in title_lower for w in ["кросівки", "кроссовки"]):
            return "3220713"
        if any(w in title_lower for w in ["ботинк", "черевик"]):
            return "3220413"
        if any(w in title_lower for w in ["рюкзак"]):
            return "380601"
        return "690102"

    # Специфические эвристики для FOOTBALLERS (футбольные товары)
    if supplier_name == "FOOTBALLERS":
        if any(w in title_lower for w in ["м'яч", "мяч"]):
            return "2023"     # футбольні м'ячі
        if any(w in title_lower for w in ["бутси", "бутсы", "сороконожк", "сороконіжк", "футзал"]):
            return "3220707"  # футбольне взуття
        if any(w in title_lower for w in ["кросівки", "кроссовки", "кеди", "кеды"]):
            return "3220713"  # кросівки та кеди
        if any(w in title_lower for w in ["щитки", "щиток"]):
            return "203302"   # футбольні щитки
        if any(w in title_lower for w in ["рукавиці", "рукавички", "перчатки"]):
            return "203108"   # футбольні рукавички
        if any(w in title_lower for w in ["форма", "футболка"]):
            return "30124"    # спортивна форма
        if any(w in title_lower for w in ["гетри", "гетры"]):
            return "203303"   # футбольні гетри
        return "2023"

    # Общие эвристики по ключевым словам
    if any(w in title_lower for w in ["кросівки", "кроссовки", "кеди", "кеды"]):
        return "3220713"
    if any(w in title_lower for w in ["бутси", "бутсы", "сороконожк", "сороконіжк", "футзал"]):
        return "3220707"
    if any(w in title_lower for w in ["берці", "берцы", "тактичн"]):
        return "690102"
    if any(w in title_lower for w in ["м'яч", "мяч"]):
        return "2023"
    if any(w in title_lower for w in ["термос", "термокружк"]):
        return "1531"
    if any(w in title_lower for w in ["намет", "палатк"]):
        return "240409"
    if any(w in title_lower for w in ["спальник", "спальний мішок"]):
        return "240407"

    return ""


def transform_supplier_data(supplier_df, supplier_name, sheet_type, mapping):
    logger.info(f"Преобразование данных {supplier_name} для {sheet_type}")

    products_cols, groups_cols = get_template_columns()
    template_cols = products_cols if sheet_type == "products" else groups_cols

    mapping_key = "products_mapping" if sheet_type == "products" else "groups_mapping"
    supplier_mapping = mapping.get(mapping_key, {}).get(supplier_name, {})

    if not supplier_mapping:
        logger.warning(f"Нет маппинга для {supplier_name} ({sheet_type})")
        return pd.DataFrame(columns=template_cols)

    data_dict = {col: [] for col in template_cols}

    for _, row in supplier_df.iterrows():
        row_dict = row.to_dict() if isinstance(row, pd.Series) else row
        row_data = {col: '' for col in template_cols}

        # Специальная обработка для IZIDROP если это группы
        if supplier_name == "IZIDROP" and sheet_type == "groups":
            if 'parentId' in row_dict:
                row_dict['parentId'] = '' if pd.isna(row_dict['parentId']) else row_dict['parentId']

        # Добавляем обработку новых полей для каждого поставщика
        if supplier_name == "AGER":
            row_data['Ідентифікатор_товару'] = row_dict.get('id', '')
            row_data['Код_товару'] = row_dict.get('vendorCode', '')
        elif supplier_name == "IZIDROP":
            row_data['Ідентифікатор_товару'] = row_dict.get('id', '')
            row_data['Код_товару'] = row_dict.get('barcode', '')
        elif supplier_name in ["MOYDROP", "SPECULANT", "KIRS", "BAGSROOM", "BAYRAKTAR"]:
            row_data['Ідентифікатор_товару'] = row_dict.get('id', '')
            row_data['Код_товару'] = row_dict.get('vendorCode', '')
        # Для FOOTBALLERS идентификаторы маппятся через основной цикл (column_mapping.json)

        # Основной цикл маппинга на основе column_mapping.json
        for supplier_col, target_col in supplier_mapping.items():
            if supplier_col not in row_dict:
                continue

            value = row_dict[supplier_col]
            if pd.isna(value) or value == '':
                continue

            if isinstance(target_col, list) and supplier_col.startswith("param_"):
                char_name = target_col[2]
                norm_char_name = UKR_PARAM_NAMES.get(char_name, char_name)

                for idx in range(1, 25):
                    name_col_target = f"Назва_Характеристики_{idx}" if idx > 1 else "Назва_Характеристики"
                    value_col_target = f"Значення_Характеристики_{idx}" if idx > 1 else "Значення_Характеристики"

                    if not row_data[name_col_target]:
                        row_data[name_col_target] = norm_char_name
                        str_value = str(value)
                        if str_value.startswith("'") and not str_value.startswith("''"):
                            str_value = str_value[1:]
                        row_data[value_col_target] = normalize_characteristic_value(norm_char_name, str_value)
                        break
            else:
                if target_col in row_data:
                    str_value = str(value)
                    if str_value.startswith("'") and not str_value.startswith("''"):
                        str_value = str_value[1:]
                    row_data[target_col] = str_value

        # Специальная обработка повторяющихся характеристик FOOTBALLERS
        if supplier_name == "FOOTBALLERS" and sheet_type == "products":
            # Нормализация первой характеристики (слот 1), если она была заполнена через основной маппинг
            if row_data.get('Назва_Характеристики'):
                raw_slot1 = str(row_data['Назва_Характеристики']).strip()
                norm_slot1 = UKR_PARAM_NAMES.get(raw_slot1, raw_slot1)
                row_data['Назва_Характеристики'] = norm_slot1
                if row_data.get('Значення_Характеристики'):
                    row_data['Значення_Характеристики'] = normalize_characteristic_value(
                        norm_slot1, str(row_data['Значення_Характеристики'])
                    )

            current_char_slot_index = 2

            for i in range(1, 24):
                source_name_key = f"Назва_Характеристики.{i}"
                source_unit_key = f"Одиниця_виміру_Характеристики.{i}"
                source_value_key = f"Значення_Характеристики.{i}"

                target_name_col = f"Назва_Характеристики_{current_char_slot_index}"
                target_unit_col = f"Одиниця_виміру_Характеристики_{current_char_slot_index}"
                target_value_col = f"Значення_Характеристики_{current_char_slot_index}"

                if source_name_key in row_dict and pd.notna(row_dict[source_name_key]) and str(
                        row_dict[source_name_key]).strip() != '':
                    if current_char_slot_index <= 24:
                        raw_char_name = str(row_dict[source_name_key]).strip()
                        norm_char_name = UKR_PARAM_NAMES.get(raw_char_name, raw_char_name)
                        row_data[target_name_col] = norm_char_name

                        if source_value_key in row_dict and pd.notna(row_dict[source_value_key]):
                            raw_val = str(row_dict[source_value_key])
                            row_data[target_value_col] = normalize_characteristic_value(norm_char_name, raw_val)
                        else:
                            row_data[target_value_col] = ''

                        if source_unit_key in row_dict and pd.notna(row_dict[source_unit_key]):
                            row_data[target_unit_col] = str(row_dict[source_unit_key])
                        else:
                            row_data[target_unit_col] = ''

                        current_char_slot_index += 1
                    else:
                        logger.warning(
                            f"Достигнут лимит в 24 характеристики для товара {row_dict.get('Код_товару', '')} от FOOTBALLERS.")
                        break
                else:
                    break

        # Определение точной портальной категории Prom.ua (Ідентифікатор_підрозділу)
        # Выполняется ПОСЛЕ основного маппинга, чтобы использовать уже нормализованные названия и группы
        if sheet_type == "products":
            raw_cat_id = str(
                row_dict.get('categoryId', '')
                or row_dict.get('group_id', '')
                or row_dict.get('Номер_групи', '')
                or row_data.get('Номер_групи', '')
                or row_data.get('Ідентифікатор_групи', '')
                or row_dict.get('Ідентифікатор_підрозділу', '')
                or ''
            ).strip()
            product_title = str(
                row_data.get('Назва_позиції_укр', '')
                or row_data.get('Назва_позиції', '')
                or row_dict.get('name_ua', '')
                or row_dict.get('name', '')
                or ''
            ).strip()
            portal_cat_id = resolve_portal_category(raw_cat_id, product_title, supplier_name)
            if portal_cat_id:
                row_data['Ідентифікатор_підрозділу'] = portal_cat_id

            # Нормализация внутренней группы магазина и логистики (Номер_групи, Назва_групи)
            norm_group_id, norm_group_name = resolve_product_group(
                supplier_name=supplier_name,
                raw_cat_id=raw_cat_id,
                product_title=product_title
            )
            if norm_group_id:
                row_data['Номер_групи'] = norm_group_id
                row_data['Назва_групи'] = norm_group_name
                row_data['Ідентифікатор_групи'] = norm_group_id

        # Значения по умолчанию для продуктов
        if sheet_type == "products":
            if not row_data['Валюта']:
                row_data['Валюта'] = 'UAH'
            if not row_data['Одиниця_виміру']:
                row_data['Одиниця_виміру'] = 'шт.'

            if 'available' in row_dict and not row_data['Наявність']:
                available_value = str(row_dict['available']).lower().strip()
                if available_value in ['true', '+', '1', 'yes', 'да', 'в наличии', 'є в наявності']:
                    row_data['Наявність'] = '+'
                elif available_value in ['false', '-', '0', 'no', 'нет', 'немає в наявності']:
                    row_data['Наявність'] = '-'

            if not row_data['Наявність'] and 'Кількість' in row_data and row_data['Кількість'] != '':
                try:
                    quantity = int(float(str(row_data['Кількість']).replace(',', '.')))
                    if quantity > 0:
                        row_data['Наявність'] = '+'
                    else:
                        row_data['Наявність'] = '-'
                except ValueError:
                    logger.warning(
                        f"Не удалось преобразовать Кількість '{row_data['Кількість']}' в число для товара {row_data.get('Код_товару', '')}")

            row_data['Особисті_нотатки'] = supplier_name

            # Санитизация описаний и форматирование таблиц замеров (для исключения блокировок модерации Prom.ua)
            if row_data.get('Опис'):
                row_data['Опис'] = sanitize_description(row_data['Опис'], is_ukr=False)
            if row_data.get('Опис_укр'):
                row_data['Опис_укр'] = sanitize_description(row_data['Опис_укр'], is_ukr=True)

        for col in template_cols:
            data_dict[col].append(row_data.get(col, ''))

    result_df = pd.DataFrame(data_dict, columns=template_cols)

    if sheet_type == "groups":
        result_df['Джерело_даних'] = supplier_name

    logger.info(f"Данные {supplier_name} ({sheet_type}) успешно преобразованы, получено строк: {len(result_df)}")
    return result_df


def process_supplier(gsheets, spreadsheet_id, supplier_name, mapping):
    logger.info(f"Обработка поставщика {supplier_name}")

    try:
        gsheets._wait_if_needed()
        spreadsheet = gsheets.service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields='sheets.properties'
        ).execute()

        sheet_names = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
        logger.info(f"Найденные листы для {supplier_name}: {sheet_names}")

        products_df = None
        groups_df = None

        for sheet_name in sheet_names:
            sheet_lower = sheet_name.lower()

            if products_df is None and any(x in sheet_lower for x in ["product", "товар", "export"]):
                df = gsheets.get_sheet_data(spreadsheet_id, sheet_name)
                if not df.empty:
                    products_df = df
                    logger.info(f"Найден лист товаров: {sheet_name} ({len(df)} строк)")

            if groups_df is None:
                if supplier_name == "IZIDROP":
                    if sheet_name == "Export Groups Sheet":
                        raw_data = gsheets.service.spreadsheets().values().get(
                            spreadsheetId=spreadsheet_id,
                            range=sheet_name
                        ).execute()

                        values = raw_data.get('values', [])
                        if values:
                            data = []
                            headers = values[0] if len(values) > 0 else []

                            id_idx = headers.index('id') if 'id' in headers else 0
                            name_idx = headers.index('name') if 'name' in headers else 1
                            parent_idx = headers.index('parentId') if 'parentId' in headers else -1

                            for row in values[1:]:
                                item = {
                                    'id': row[id_idx] if len(row) > id_idx else '',
                                    'name': row[name_idx] if len(row) > name_idx else ''
                                }
                                if parent_idx != -1 and len(row) > parent_idx:
                                    item['parentId'] = row[parent_idx]
                                else:
                                    item['parentId'] = ''
                                data.append(item)

                            groups_df = pd.DataFrame(data)
                            logger.info(f"Обработан лист групп IZIDROP: {sheet_name} ({len(groups_df)} строк)")
                else:
                    if any(x in sheet_lower for x in ["group", "груп", "category"]):
                        df = gsheets.get_sheet_data(spreadsheet_id, sheet_name)
                        if not df.empty:
                            groups_df = df
                            logger.info(f"Найден лист групп: {sheet_name} ({len(df)} строк)")

        results = {}

        if products_df is not None:
            results["products"] = transform_supplier_data(
                products_df, supplier_name, "products", mapping)

            # --- ЗАЩИТА: Фильтрация товаров с нулевой, пустой или некорректной ценой ---
            if 'Ціна' in results["products"].columns and not results["products"].empty:
                orig_len = len(results["products"])
                price_num = pd.to_numeric(
                    results["products"]['Ціна'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.'),
                    errors='coerce'
                ).fillna(0)
                results["products"] = results["products"][price_num > 0].copy()
                filtered_count = orig_len - len(results["products"])
                if filtered_count > 0:
                    logger.warning(
                        f"Поставщик {supplier_name}: отфильтровано {filtered_count} товаров с ценой <= 0 или пустой ценой (не допущены в шаблон импорта)"
                    )
        else:
            logger.warning(f"Нет данных товаров для {supplier_name}")
            results["products"] = pd.DataFrame()

        if groups_df is not None:
            if supplier_name == "IZIDROP":
                if 'parentId' not in groups_df.columns:
                    groups_df['parentId'] = ''
                groups_df['parentId'] = groups_df['parentId'].fillna('')

            results["groups"] = transform_supplier_data(
                groups_df, supplier_name, "groups", mapping)
            logger.debug(f"Пример преобразованных групп:\n{results['groups'].head()}")
        else:
            logger.warning(f"Нет данных групп для {supplier_name}")
            results["groups"] = pd.DataFrame()

        logger.info(f"Итоговые данные: товаров - {len(results['products'])}, групп - {len(results['groups'])}")
        return results

    except Exception as e:
        logger.error(f"Критическая ошибка при обработке {supplier_name}: {str(e)}", exc_info=True)
        return {"products": pd.DataFrame(), "groups": pd.DataFrame()}


def validate_dataframe(df, sheet_type):
    logger.info(f"Validating {sheet_type} DataFrame")

    rows, cols = df.shape
    logger.info(f"DataFrame size: {rows} rows x {cols} columns")

    for col in df.columns:
        max_len = df[col].astype(str).map(len).max()
        if max_len > 1000:
            logger.warning(f"Column '{col}' has very long values (max length: {max_len})")

        if df[col].dtype == 'object':
            unusual_chars = df[col].astype(str).str.contains(
                r'[^\w\s.,;:()\[\]{}"\'<>?!@#$%^&*+=\-/\\]', regex=True)
            if unusual_chars.any():
                unusual_count = unusual_chars.sum()
                logger.warning(f"Column '{col}' has {unusual_count} cells with unusual characters")

    if len(df.columns) != len(set(df.columns)):
        logger.warning("DataFrame has duplicate column names!")
        for col in df.columns:
            if list(df.columns).count(col) > 1:
                logger.warning(f"Column '{col}' appears multiple times")

    return True


# --- ДОБАВЛЕНА АДАПТИРОВАННАЯ ФУНКЦИЯ apply_price_discount ---
def apply_price_discount(gsheets, spreadsheet_id, sheet_name='Export Products Sheet'):
    """
    Функция пересчета цен и скидок, адаптированная под googleapiclient.
    Использует внутренние методы GoogleSheetsManager (включая ожидание и ретраи),
    считывает финальный лист после записи и применяет нужные патчи.
    """
    logger.info(f"--- ЗАПУСК ОБРАБОТКИ ЦЕН/СКИДОК для листа '{sheet_name}' ---")
    try:
        # 1. Читаем все данные целевого листа
        gsheets._wait_if_needed()
        result = gsheets._execute_with_retry(
            gsheets.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=sheet_name,
                valueRenderOption='FORMATTED_VALUE'
            )
        )

        values = result.get('values', [])
        if not values or len(values) < 2:
            logger.warning(f"Лист '{sheet_name}' пуст или не содержит данных. Обработка цен/скидок пропущена.")
            return

        headers = [str(h).strip() for h in values[0]]

        def find_column_index(headers_list, col_name):
            try:
                return headers_list.index(col_name)
            except ValueError:
                return -1

        price_idx = find_column_index(headers, 'Ціна')
        price_from_idx = find_column_index(headers, 'Ціна_від')
        discount_idx = find_column_index(headers, 'Знижка')

        if price_idx == -1 or price_from_idx == -1:
            logger.warning(f"Столбцы 'Ціна' и/или 'Ціна_від' не найдены на листе. Обработка пропущена.")
            return

        # Локальная функция для преобразования индекса колонки в буквы (A, B, C...)
        def get_column_letter(col_idx):
            letter = ''
            n = col_idx + 1
            while n > 0:
                n, rem = divmod(n - 1, 26)
                letter = chr(65 + rem) + letter
            return letter

        def parse_price(val):
            val = re.sub(r'\s+', '', val)
            val = val.replace(',', '.')
            return float(val)

        price_updates = []
        rows_updated = 0

        # 2. Идем по строкам и считаем скидки
        for row_idx_0based, row in enumerate(values[1:]):
            row_num = row_idx_0based + 2  # +1 за заголовок, +1 так как нумерация строк в Google начинается с 1

            # Дополняем строку пустыми значениями, если она обрывается до нужных индексов
            if len(row) < len(headers):
                row.extend([''] * (len(headers) - len(row)))

            current_price = row[price_idx].strip()
            current_price_from = row[price_from_idx].strip()
            current_discount = row[discount_idx].strip() if discount_idx != -1 else ''

            if not current_price:
                continue

            # По умолчанию (если скидки нет или отменена)
            val_price = current_price
            val_price_from = current_price
            discount_str = ""

            # Проверяем, есть ли цена со скидкой
            if current_price_from:
                try:
                    p_base = parse_price(current_price)
                    p_from = parse_price(current_price_from)

                    if p_base != p_from:
                        if p_base > p_from:
                            val_price = current_price
                            val_price_from = current_price_from
                            discount = p_base - p_from
                        else:
                            val_price = current_price_from
                            val_price_from = current_price
                            discount = p_from - p_base

                        discount_str = str(round(discount, 2))
                        if discount_str.endswith('.0'):
                            discount_str = discount_str[:-2]
                except ValueError:
                    # Если попался текст вроде "Немає в наявності", игнорируем ошибку конвертации
                    continue

            needs_update = False

            # Сравниваем расчетные значения с текущими на листе
            if current_price != val_price:
                col_letter = get_column_letter(price_idx)
                price_updates.append({
                    'range': f"{sheet_name}!{col_letter}{row_num}",
                    'values': [[val_price]]
                })
                needs_update = True

            if current_price_from != val_price_from:
                col_letter = get_column_letter(price_from_idx)
                price_updates.append({
                    'range': f"{sheet_name}!{col_letter}{row_num}",
                    'values': [[val_price_from]]
                })
                needs_update = True

            if discount_idx != -1 and current_discount != discount_str:
                col_letter = get_column_letter(discount_idx)
                price_updates.append({
                    'range': f"{sheet_name}!{col_letter}{row_num}",
                    'values': [[discount_str]]
                })
                needs_update = True

            if needs_update:
                rows_updated += 1

        logger.info(f"Найдено {rows_updated} товаров, требующих обновления цен/скидок. "
                    f"Подготовлено {len(price_updates)} операций обновления ячеек.")

        if not price_updates:
            logger.info("Нет операций для обновления цен/скидок. Данные актуальны.")
            logger.info(f"--- Обработка цен/скидок для листа '{sheet_name}' завершена ---")
            return

        # 3. Отправляем батчи
        PRICE_CHUNK_SIZE = 10000
        total_ops = len(price_updates)
        total_chunks = (total_ops + PRICE_CHUNK_SIZE - 1) // PRICE_CHUNK_SIZE
        logger.info(
            f"Отправка обновлений цен/скидок: {total_ops} операций, {total_chunks} чанков по {PRICE_CHUNK_SIZE}...")

        for chunk_idx in range(total_chunks):
            chunk_start = chunk_idx * PRICE_CHUNK_SIZE
            chunk_end = min(chunk_start + PRICE_CHUNK_SIZE, total_ops)
            chunk_data = price_updates[chunk_start:chunk_end]

            logger.info(f"  Цены: чанк {chunk_idx + 1}/{total_chunks}, операции {chunk_start + 1}-{chunk_end}...")

            gsheets._wait_if_needed()
            gsheets._execute_with_retry(
                gsheets.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={
                        'valueInputOption': 'USER_ENTERED',
                        'data': chunk_data
                    }
                )
            )

            if chunk_idx < total_chunks - 1:
                time.sleep(1)

        logger.info(f"Обработка цен/скидок завершена успешно: {rows_updated} строк, {total_ops} операций.")

    except Exception as e:
        logger.error(f"Неожиданная ошибка в apply_price_discount: {e}", exc_info=True)

    logger.info(f"--- Обработка цен/скидок для листа '{sheet_name}' завершена ---")


# -------------------------------------------------------------

def main():
    try:
        logger.info("=== Начало работы скрипта ===")

        gsheets = GoogleSheetsManager()
        mapping = load_mapping()

        all_products = []
        all_groups = []

        for supplier_name, spreadsheet_id in CONFIG["SUPPLIERS"].items():
            logger.info(f"Обработка поставщика {supplier_name}")
            supplier_data = process_supplier(gsheets, spreadsheet_id, supplier_name, mapping)

            if supplier_data.get("products") is not None:
                products_df = supplier_data["products"]
                logger.info(f"Получено продуктов от {supplier_name}: {len(products_df)}")

                if validate_dataframe(products_df, "products"):
                    all_products.append(products_df)
                else:
                    logger.error(f"Ошибка валидации данных продуктов для {supplier_name}")

            if supplier_data.get("groups") is not None:
                groups_df = supplier_data["groups"]
                logger.info(f"Получено групп от {supplier_name}: {len(groups_df)}")

                if validate_dataframe(groups_df, "groups"):
                    groups_with_source = groups_df.assign(Джерело_даних=supplier_name)
                    all_groups.append(groups_with_source)
                else:
                    logger.error(f"Ошибка валидации данных групп для {supplier_name}")

        products_cols, groups_cols = get_template_columns()

        merged_products = pd.concat(all_products, ignore_index=True)
        if 'Ціна' in merged_products.columns and not merged_products.empty:
            orig_total = len(merged_products)
            price_num_final = pd.to_numeric(
                merged_products['Ціна'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.'),
                errors='coerce'
            ).fillna(0)
            merged_products = merged_products[price_num_final > 0].copy()
            if orig_total > len(merged_products):
                logger.warning(f"Итоговая таблица: удалено {orig_total - len(merged_products)} позиций с некорректной/нулевой ценой!")

        # Унификация групп для всех разновидностей одного товара (ID_групи_різновидів)
        merged_products = unify_variation_groups(merged_products)

        # Эталонное 3-уровневое дерево групп Prom.ua (62 группы: 5 корневых, 23 подраздела, 34 конечные с логистикой)
        merged_groups = get_canonical_groups_df()

        logger.info(f"Объединено продуктов: {len(merged_products)}")
        logger.info(f"Объединено групп: {len(merged_groups)}")

        required_product_columns = ["Код_товару", "Ідентифікатор_товару", "Назва_позиції", "Ціна", "Наявність"]
        for column in required_product_columns:
            if column not in merged_products.columns:
                logger.error(f"В объединенных продуктах отсутствует обязательный столбец: {column}")
                raise ValueError(f"Отсутствует обязательный столбец: {column}")

        required_group_columns = ["Номер_групи", "Назва_групи"]
        for column in required_group_columns:
            if column not in merged_groups.columns:
                logger.error(f"В объединенных группах отсутствует обязательный столбец: {column}")
                raise ValueError(f"Отсутствует обязательный столбец: {column}")

        logger.info("Запись данных продуктов...")
        success, error = gsheets.write_sheet_data(
            CONFIG["OUTPUT_SPREADSHEET_ID"],
            "Export Products Sheet",
            merged_products
        )
        if not success:
            raise Exception(f"Ошибка записи продуктов: {error}")

        logger.info("Запись данных групп...")
        success, error = gsheets.write_sheet_data(
            CONFIG["OUTPUT_SPREADSHEET_ID"],
            "Export Groups Sheet",
            merged_groups
        )
        if not success:
            raise Exception(f"Ошибка записи групп: {error}")

        # --- ВНЕДРЕННАЯ ФУНКЦИЯ ВЫЗЫВАЕТСЯ ЗДЕСЬ ---
        apply_price_discount(gsheets, CONFIG["OUTPUT_SPREADSHEET_ID"], "Export Products Sheet")
        # -------------------------------------------

        logger.info("Проверка и обеспечение минимума пустых строк (1000)...")
        gsheets._ensure_min_empty_rows("Export Products Sheet", 1000)
        gsheets._ensure_min_empty_rows("Export Groups Sheet", 1000)

        logger.info("=== Скрипт успешно завершен ===")

    except Exception as e:
        logger.error(f"КРИТИЧЕСКАЯ ОШИБКА: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
