import pandas as pd
import json
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time
from datetime import datetime
import os  # <--- ВОТ ЭТА СТРОКА ОБЯЗАТЕЛЬНА
import re  # ДОБАВЛЕНО: нужно для обработки строк с ценами

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
        "AGER": os.environ.get("PARS_AGER_SPREADSHEET_ID", "PARS_AGER_SPREADSHEET_ID"),
        "FOOTBALLERS": os.environ.get("PARS_FOOTBALLERS_SPREADSHEET_ID", "PARS_FOOTBALLERS_SPREADSHEET_ID"),
        "IZIDROP": os.environ.get("PARS_IZIDROP_SPREADSHEET_ID", "PARS_IZIDROP_SPREADSHEET_ID"),
        "MOYDROP": os.environ.get("PARS_MOYDROP_SPREADSHEET_ID", "PARS_MOYDROP_SPREADSHEET_ID"),
        "BAYRAKTAR": os.environ.get("PARS_BAYRAKTAR_SPREADSHEET_ID", "PARS_BAYRAKTAR_SPREADSHEET_ID"),
        "KIRS": os.environ.get("PARS_KIRS_SPREADSHEET_ID", "PARS_KIRS_SPREADSHEET_ID"),
        "BAGSROOM": os.environ.get("PARS_BAGSROOM_SPREADSHEET_ID", "PARS_BAGSROOM_SPREADSHEET_ID")
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

            # ── Шаг 2: определяем ключевые столбцы ────────────────────────────
            is_groups = "Export Groups Sheet" in sheet_name
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

            def _col_letter(idx):
                letter, n = '', idx + 1
                while n > 0:
                    n, rem = divmod(n - 1, 26)
                    letter = chr(65 + rem) + letter
                return letter

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
                critical_fields = ["Назва_групи", "Назва_групи_укр", "Ідентифікатор_групи"]
                for f in critical_fields:
                    if f in headers and f in df.columns:
                        idx = headers.index(f)
                        col_idx_to_df_col[idx] = f
            else:
                base_critical_fields = [
                    "Назва_позиції", "Назва_позиції_укр", "Опис", "Опис_укр",
                    "Наявність", "Ціна", "Ціна_від", "Кількість", "Знижка",
                    "Ярлик", "Де_знаходиться_товар", "Тип_товару",
                ]
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
                for chunk_start in range(0, len(batch_value_updates), CONFIG["BATCH_SIZE"]):
                    chunk = batch_value_updates[chunk_start:chunk_start + CONFIG["BATCH_SIZE"]]
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
        # 1) Сначала собираем характеристики в структурированный список
        collected_params = []
        
        # Словарь нормализации/перевода характеристик на украинский язык
        UKR_PARAM_NAMES = {
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
            # Дополнительные параметры для сумок (KIRS / BAGSROOM) и аксессуаров
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
            "Висота": "Висота",
            "Глубина": "Глибина",
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
            # Спортивные параметры и бутсы (FOOTBALLERS)
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
            "Вид изделия": "Вид виробу",
            # Дополнительные характеристики сумок и аксессуаров (KIRS)
            "Тип назначения": "Тип призначення",
            "Упаковка средства": "Упаковка засобу",
            "Назначение": "Призначення",
            "Количество в упаковке": "Кількість в упаковці",
            "Диагональ экрана ноутбука": "Діагональ екрана ноутбука",
            "Страховочный ремешок": "Страхувальний ремінець",
            "Цвет фурнитуры": "Колір фурнітури",
            "Высота упаковки": "Висота упаковки",
            "Длина упаковки": "Довжина упаковки",
            "Размер женской одежды": "Розмір жіночого одягу",
            "Количество составных элементов": "Кількість складових елементів",
            "Фасон выреза горловины": "Фасон вирізу горловини",
            "Тип лампы": "Тип лампи",
            # Дополнительные параметры BAGSROOM
            "Декоративные элементы": "Декоративні елементи",
            "Объем": "Об'єм",
            "Высота подушки": "Висота подушки",
            "Максимальный обхват талии": "Максимальний обхват талії",
            "Особенности": "Особливості",
            "Высота стеллажа": "Висота стелажа",
            # Мебельные параметры (BAGSROOM)
            "Опоры": "Опори",
            "Фасады": "Фасади",
            "Высота от пола": "Висота від підлоги",
            "Глубина столешницы": "Глибина стільниці",
            "Длина столешницы": "Довжина стільниці",
            "Высота стола": "Висота столу",
            "Высота сиденья стульчика": "Висота сидіння стільчика",
            "Высота стульчика с учетом спинки": "Висота стільчика з урахуванням спинки",
            "Количество выдвижных ящиков": "Кількість висувних ящиків",
            "Гарантийный срок": "Гарантійний термін",
            "Вид стола по своему назначению": "Вид столу за своїм призначенням"
        }

        # Приоритет главных фильтров Prom.ua для первых слотов
        PRIORITY_PARAMS = [
            "Розмір",
            "Розмір жіночого взуття",
            "Розмір чоловічого взуття",
            "Сезон",
            "Матеріал верху",
            "Матеріал",
            "Колір",
            "Група кольору",
            "Матеріал підкладки",
            "Матеріал підошви",
            "Матеріал устілки",
            "Стать",
            "Бренд",
            "Виробник",
            "Стиль",
            "Склад",
            "Тип сумки",
            "Форма",
            "Жорсткість",
            "Застібка"
        ]

        handled_param_cols = set()
        for supplier_col, target_col in supplier_mapping.items():
            if supplier_col not in row_dict:
                continue

            value = row_dict[supplier_col]
            if pd.isna(value) or value == '':
                continue

            if isinstance(target_col, list) and supplier_col.startswith("param_"):
                handled_param_cols.add(supplier_col)
                char_name = target_col[2] if len(target_col) > 2 else supplier_col.replace("param_", "")
                # Переводим на украинский язык
                char_name_ukr = UKR_PARAM_NAMES.get(char_name, UKR_PARAM_NAMES.get(supplier_col.replace("param_", ""), char_name))
                str_value = str(value)
                if str_value.startswith("'") and not str_value.startswith("''"):
                    str_value = str_value[1:]
                collected_params.append((char_name_ukr, str_value))
            else:
                if target_col in row_data:
                    str_value = str(value)
                    if str_value.startswith("'") and not str_value.startswith("''"):
                        str_value = str_value[1:]
                    row_data[target_col] = str_value

        # Универсальный динамический сбор всех несмаппленных параметров (param_*)
        for k, v in row_dict.items():
            if k.startswith("param_") and k not in handled_param_cols:
                if pd.isna(v) or v == '':
                    continue
                raw_name = k.replace("param_", "").strip()
                ukr_name = UKR_PARAM_NAMES.get(raw_name, raw_name)
                str_val = str(v)
                if str_val.startswith("'") and not str_val.startswith("''"):
                    str_val = str_val[1:]
                collected_params.append((ukr_name, str_val))

        # Сортируем собранные характеристики по приоритету Prom.ua
        def get_param_priority(item):
            name = item[0]
            if name in PRIORITY_PARAMS:
                return PRIORITY_PARAMS.index(name)
            return len(PRIORITY_PARAMS) + 100

        collected_params.sort(key=get_param_priority)

        # Раскладываем по слотам 1..24
        slot_idx = 1
        for char_name, char_val in collected_params:
            if slot_idx > 24:
                break
            name_col_target = f"Назва_Характеристики_{slot_idx}" if slot_idx > 1 else "Назва_Характеристики"
            value_col_target = f"Значення_Характеристики_{slot_idx}" if slot_idx > 1 else "Значення_Характеристики"
            row_data[name_col_target] = char_name
            row_data[value_col_target] = char_val
            slot_idx += 1

        # Специальная обработка повторяющихся характеристик FOOTBALLERS
        if supplier_name == "FOOTBALLERS" and sheet_type == "products":
            current_char_slot_index = 1

            for i in range(1, 24):
                source_name_key = f"Назва_Характеристики.{i}"
                source_unit_key = f"Одиниця_виміру_Характеристики.{i}"
                source_value_key = f"Значення_Характеристики.{i}"

                target_name_col = f"Назва_Характеристики_{current_char_slot_index}" if current_char_slot_index > 1 else "Назва_Характеристики"
                target_unit_col = f"Одиниця_виміру_Характеристики_{current_char_slot_index}" if current_char_slot_index > 1 else "Одиниця_виміру_Характеристики"
                target_value_col = f"Значення_Характеристики_{current_char_slot_index}" if current_char_slot_index > 1 else "Значення_Характеристики"

                if source_name_key in row_dict and pd.notna(row_dict[source_name_key]) and str(
                        row_dict[source_name_key]).strip() != '':
                    if current_char_slot_index <= 24:
                        raw_char_name = str(row_dict[source_name_key]).strip()
                        norm_char_name = UKR_PARAM_NAMES.get(raw_char_name, raw_char_name)
                        row_data[target_name_col] = norm_char_name

                        if source_value_key in row_dict and pd.notna(row_dict[source_value_key]):
                            row_data[target_value_col] = str(row_dict[source_value_key])
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

        merged_products = pd.concat(all_products, ignore_index=True) if all_products \
            else pd.DataFrame(columns=products_cols)
        merged_groups = pd.concat(all_groups, ignore_index=True) if all_groups \
            else pd.DataFrame(columns=groups_cols)

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
