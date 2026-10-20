#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import time
from datetime import datetime
import logging
from typing import List, Dict, Any, Tuple, Set
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("feed_processor.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("feed_processor")

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'key_sheet.json'
PRODUCT_SHEET_NAME = 'Export Products Sheet'
CATEGORY_SHEET_NAME = 'Export Groups Sheet'
BATCH_SIZE = 30  # Уменьшили размер пакета для безопасности
API_DELAY = 1.1  # Для соблюдения лимита 60 запросов в минуту
MAX_COLUMN = 'CY'


class FeedProcessor:
    def __init__(self, target_spreadsheet_id: str, source_spreadsheet_id: str):
        self.target_spreadsheet_id = target_spreadsheet_id
        self.source_spreadsheet_id = source_spreadsheet_id
        self.service = self._get_sheets_service()
        self.source_products = []
        self.source_categories = []
        self.target_products = []
        self.target_categories = []
        self.column_order = {}  # Для сохранения порядка столбцов
        self.request_count = 0
        self.last_request_time = time.time()

    def _get_sheets_service(self):
        """Create and return Google Sheets API service."""
        try:
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            return build('sheets', 'v4', credentials=creds).spreadsheets()
        except Exception as e:
            logger.error(f"Failed to create Google Sheets service: {e}")
            raise

    def _rate_limit(self):
        """Manage API request rate limiting."""
        self.request_count += 1
        current_time = time.time()
        elapsed = current_time - self.last_request_time

        # Если не прошло API_DELAY секунд с последнего запроса
        if elapsed < API_DELAY:
            time.sleep(API_DELAY - elapsed)

        # Если достигли BATCH_SIZE запросов, делаем дополнительную паузу
        if self.request_count >= BATCH_SIZE:
            time.sleep(API_DELAY * 2)
            self.request_count = 0

        self.last_request_time = time.time()

    def _fetch_sheet_data(self, spreadsheet_id: str, sheet_name: str) -> tuple:
        """Fetch data from a sheet preserving original column order."""
        try:
            self._rate_limit()
            # Получаем все данные как есть
            result = self.service.values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:{MAX_COLUMN}",
                valueRenderOption="UNFORMATTED_VALUE"
            ).execute()

            values = result.get('values', [])
            if not values:
                return [], []

            headers = values[0]
            data_rows = values[1:]

            # Сохраняем порядок столбцов для этого листа
            self.column_order[sheet_name] = headers

            # Нормализуем данные (все строки должны иметь одинаковое количество столбцов)
            normalized_data = []
            for row in data_rows:
                normalized_row = row + [''] * (len(headers) - len(row))
                normalized_data.append(normalized_row[:len(headers)])  # Обрезаем лишние столбцы, если они есть

            return headers, normalized_data

        except HttpError as e:
            logger.error(f"Error fetching data from {sheet_name}: {e}")
            return [], []

    def _format_headers_bold(self, sheet_name: str):
        """Format headers as bold in the target spreadsheet."""
        try:
            self._rate_limit()

            # Получаем ID листа для форматирования
            sheet_metadata = self.service.get(spreadsheetId=self.target_spreadsheet_id).execute()
            sheet_id = None

            for sheet in sheet_metadata['sheets']:
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break

            if sheet_id is None:
                logger.error(f"Sheet {sheet_name} not found in target spreadsheet")
                return

            # Создаем запрос на форматирование
            requests = [{
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat.bold"
                }
            }]

            # Применяем форматирование
            self._rate_limit()
            self.service.batchUpdate(
                spreadsheetId=self.target_spreadsheet_id,
                body={"requests": requests}
            ).execute()

            logger.info(f"Headers in {sheet_name} formatted as bold")

        except HttpError as e:
            logger.error(f"Error formatting headers in {sheet_name}: {e}")

    def _create_row_key(self, row: List, headers: List[str], id_column_index: int = 0) -> str:
        """Create a unique key for a row based on ID column (default is first column)."""
        if id_column_index < len(row):
            return str(row[id_column_index])
        return ""

    def _find_differences(self, source_headers: List[str], source_data: List[List],
                          target_headers: List[str], target_data: List[List]) -> Tuple[
        List[List], List[List], List[str], Dict]:
        """Find rows to add, update, and delete with strict comparison."""
        source_dict = {}
        target_dict = {}
        id_column_index = 0  # ID в первом столбце

        # Собираем данные источника (последняя версия каждого товара)
        for row in reversed(source_data):  # берем последнюю версию дубликатов
            key = str(row[id_column_index])
            if key and key not in source_dict:
                source_dict[key] = row

        # Собираем данные цели (только уникальные товары)
        for i, row in enumerate(target_data):
            key = str(row[id_column_index])
            if key:
                target_dict[key] = (row, i + 1)  # позиция в таблице (+1 для заголовка)

        # Определяем изменения
        rows_to_add = []
        rows_to_update = []
        update_positions = {}

        for key, source_row in source_dict.items():
            if key not in target_dict:
                rows_to_add.append(source_row)  # новый товар
            else:
                target_row, pos = target_dict[key]
                # Сравниваем ВСЕ поля, кроме служебных (например, 'last_updated')
                for i, (src_val, tgt_val) in enumerate(zip(source_row, target_row)):
                    if i >= len(target_headers):
                        break
                    if str(src_val) != str(tgt_val):
                        rows_to_update.append(source_row)
                        update_positions[key] = pos
                        break

        # Товары для удаления (если их нет в источнике)
        rows_to_delete = [k for k in target_dict if k not in source_dict]

        return rows_to_add, rows_to_update, rows_to_delete, update_positions

    def _remove_duplicates(self, headers: List[str], data: List[List]) -> List[List]:
        """Удаляет все дубликаты, оставляя только первую версию."""
        seen = set()
        unique = []
        for row in data:
            if not row:
                continue
            row_id = str(row[0])  # ID в первом столбце
            if row_id not in seen:
                seen.add(row_id)
                unique.append(row)
        return unique

    def _update_specific_rows(self, spreadsheet_id: str, sheet_name: str,
                              headers: List[str], rows_to_update: List[List],
                              update_positions: Dict):
        """Update specific rows in the target spreadsheet."""
        if not rows_to_update:
            return

        try:
            for row in rows_to_update:
                key = self._create_row_key(row, headers)
                if key in update_positions:
                    position = update_positions[key]
                    # Формируем правильный диапазон строго в пределах длины строки
                    # Используем буквы от A до Z для первых 26 столбцов
                    max_col_index = min(len(row), len(headers)) - 1

                    # Создаем диапазон с учетом многобуквенных колонок (AA, AB и т.д.)
                    if max_col_index <= 25:  # Если индекс в пределах A-Z
                        end_col = chr(65 + max_col_index)
                    else:  # Для колонок после Z (AA, AB, ...)
                        first_char = chr(65 + (max_col_index // 26) - 1)
                        second_char = chr(65 + (max_col_index % 26))
                        end_col = first_char + second_char

                    range_name = f"{sheet_name}!A{position}:{end_col}{position}"

                    # Обрезаем данные до нужной длины
                    row_data = row[:max_col_index + 1]

                    self._rate_limit()
                    self.service.values().update(
                        spreadsheetId=spreadsheet_id,
                        range=range_name,
                        valueInputOption="RAW",
                        body={'values': [row_data]}
                    ).execute()

            logger.info(f"Updated {len(rows_to_update)} rows in {sheet_name}")

        except HttpError as e:
            logger.error(f"Error updating specific rows in {sheet_name}: {e}")
            # Добавляем больше информации для отладки
            if rows_to_update:
                sample_row = rows_to_update[0]
                sample_key = self._create_row_key(sample_row, headers)
                if sample_key in update_positions:
                    position = update_positions[sample_key]
                    logger.error(
                        f"Debug info - Row length: {len(sample_row)}, Headers length: {len(headers)}, Position: {position}")

    def _add_new_rows(self, spreadsheet_id: str, sheet_name: str, headers: List[str], rows_to_add: List[List]):
        """Add new rows to the target spreadsheet."""
        if not rows_to_add:
            return

        try:
            # Определяем диапазон для добавления (после последней существующей строки)
            self._rate_limit()
            current_data = self.service.values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:{MAX_COLUMN}"
            ).execute().get('values', [])

            start_row = len(current_data) + 1

            # Добавляем данные батчами
            for i in range(0, len(rows_to_add), BATCH_SIZE):
                batch = rows_to_add[i:i + BATCH_SIZE]

                # Обрезаем данные до количества заголовков
                normalized_batch = [row[:len(headers)] for row in batch]

                range_name = f"{sheet_name}!A{start_row}"

                self._rate_limit()
                self.service.values().update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption="RAW",
                    body={'values': normalized_batch}
                ).execute()

                start_row += len(batch)

            logger.info(f"Added {len(rows_to_add)} new rows to {sheet_name}")

        except HttpError as e:
            logger.error(f"Error adding new rows to {sheet_name}: {e}")

    def _delete_rows(self, spreadsheet_id: str, sheet_name: str,
                     headers: List[str], target_data: List[List], rows_to_delete: List[str]):
        """Delete rows that no longer exist in the source."""
        if not rows_to_delete:
            return

        try:
            # Получаем ID листа для операции удаления
            self._rate_limit()
            sheet_metadata = self.service.get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None

            for sheet in sheet_metadata['sheets']:
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break

            if sheet_id is None:
                logger.error(f"Sheet {sheet_name} not found in target spreadsheet")
                return

            # Собираем индексы строк для удаления
            rows_to_delete_indices = []
            id_column_index = 0  # Индекс колонки с ID

            for i, row in enumerate(target_data):
                key = self._create_row_key(row, headers, id_column_index)
                if key in rows_to_delete:
                    # +1 для учета индекса отсчитываемого от нуля и +1 для заголовка = +2
                    rows_to_delete_indices.append(i + 2)

            # Сортируем индексы в обратном порядке, чтобы удаление не меняло позиции оставшихся строк
            rows_to_delete_indices.sort(reverse=True)

            # Удаляем строки по одной, чтобы избежать проблем с индексами
            for row_index in rows_to_delete_indices:
                delete_request = {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": row_index - 1,  # API использует индексацию с 0
                            "endIndex": row_index
                        }
                    }
                }

                self._rate_limit()
                self.service.batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={
                        "requests": [delete_request]
                    }
                ).execute()

            logger.info(f"Deleted {len(rows_to_delete)} rows from {sheet_name}")

        except HttpError as e:
            logger.error(f"Error deleting rows from {sheet_name}: {e}")

    def _sync_sheet_data(self, sheet_name: str, source_headers: List[str], source_data: List[List]):
        """Полная синхронизация с данными поставщика (с сохранением порядка)"""
        logger.info(f"Синхронизация {sheet_name}...")

        # Получаем очищенные от дубликатов данные поставщика (в оригинальном порядке)
        unique_source_data = []
        seen_ids = set()
        for row in source_data:
            row_id = str(row[0]) if row else None
            if row_id and row_id not in seen_ids:
                seen_ids.add(row_id)
                unique_source_data.append(row)

        # Полностью перезаписываем целевую таблицу
        self._update_sheet_with_data(
            self.target_spreadsheet_id,
            sheet_name,
            source_headers,
            unique_source_data
        )

        self._format_headers_bold(sheet_name)
        logger.info(f"Синхронизация {sheet_name} завершена. Обновлено {len(unique_source_data)} строк")

    def _update_sheet_with_data(self, spreadsheet_id: str, sheet_name: str, headers: List[str], data: List[List]):
        """Полное обновление листа с сохранением порядка и без пустых строк"""
        if not headers:
            return

        try:
            # 1. Очищаем весь лист
            self._rate_limit()
            self.service.values().clear(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:{MAX_COLUMN}"
            ).execute()

            # 2. Записываем заголовки
            self._rate_limit()
            self.service.values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A1",
                valueInputOption="RAW",
                body={'values': [headers]}
            ).execute()

            # 3. Записываем данные (пачками по BATCH_SIZE)
            if data:
                for i in range(0, len(data), BATCH_SIZE):
                    batch = data[i:i + BATCH_SIZE]
                    start_row = i + 2  # +2 т.к. строка 1 - заголовки

                    self._rate_limit()
                    self.service.values().update(
                        spreadsheetId=spreadsheet_id,
                        range=f"{sheet_name}!A{start_row}",
                        valueInputOption="RAW",
                        body={'values': batch}
                    ).execute()

            logger.info(f"Лист {sheet_name} полностью обновлен")

        except HttpError as e:
            logger.error(f"Ошибка при обновлении {sheet_name}: {e}")
            raise

    def fetch_source_data(self):
        """Fetch data from source spreadsheet."""
        logger.info("Fetching data from source spreadsheet...")

        # Получаем данные продуктов
        product_headers, product_data = self._fetch_sheet_data(self.source_spreadsheet_id, PRODUCT_SHEET_NAME)
        if product_headers and product_data:
            self.source_products = (product_headers, product_data)
            logger.info(f"Fetched {len(product_data)} products with {len(product_headers)} columns from source")

        # Получаем данные категорий
        category_headers, category_data = self._fetch_sheet_data(self.source_spreadsheet_id, CATEGORY_SHEET_NAME)
        if category_headers and category_data:
            self.source_categories = (category_headers, category_data)
            logger.info(f"Fetched {len(category_data)} categories with {len(category_headers)} columns from source")

    def sync_data(self):
        """Простая и жесткая синхронизация данных"""
        # Продукты (работает)
        if self.source_products:
            headers, data = self.source_products
            self._update_sheet_with_data(
                self.target_spreadsheet_id,
                PRODUCT_SHEET_NAME,
                headers,
                data
            )
            logger.info(f"Продукты обновлены: {len(data)} строк")

        # Категории (исправляем проблему)
        if self.source_categories:
            headers, data = self.source_categories
            # 1. Полностью очищаем лист
            self._rate_limit()
            self.service.values().clear(
                spreadsheetId=self.target_spreadsheet_id,
                range=f"{CATEGORY_SHEET_NAME}!A:Z"
            ).execute()

            # 2. Записываем новые данные
            if data:
                self._rate_limit()
                self.service.values().update(
                    spreadsheetId=self.target_spreadsheet_id,
                    range=f"{CATEGORY_SHEET_NAME}!A1",
                    valueInputOption="RAW",
                    body={'values': [headers] + data}
                ).execute()

            logger.info(f"Категории перезаписаны: {len(data)} строк")
            self._format_headers_bold(CATEGORY_SHEET_NAME)

    def run(self):
        """Run the complete process."""
        start_time = datetime.now()
        logger.info(f"=== Process started at {start_time} ===")

        try:
            self.fetch_source_data()
            self.sync_data()

            duration = (datetime.now() - start_time).total_seconds()
            logger.info(f"=== Process completed in {duration:.2f} seconds ===")

        except Exception as e:
            logger.error(f"Process failed: {e}")
            raise


if __name__ == "__main__":
    # Получаем ID целевой таблицы (куда копируем)
    # ИСПРАВЛЕНО: Теперь ищем секрет PARS_FOOTBALLERS_SPREADSHEET_ID
    TARGET_SPREADSHEET_ID = os.environ.get("PARS_FOOTBALLERS_SPREADSHEET_ID", "ВАШ_ID_ТАБЛИЦЫ")
    
    # Получаем ID исходной таблицы (откуда копируем)
    # ИСПРАВЛЕНО: Теперь ищем секрет FOOTBALLERS_FEED_SPREADSHEET_ID
    SOURCE_SPREADSHEET_ID = os.environ.get("FOOTBALLERS_FEED_SPREADSHEET_ID", "ВАШ_ID_ТАБЛИЦЫ_ПОСТАВЩИКА")

    processor = FeedProcessor(
        target_spreadsheet_id=TARGET_SPREADSHEET_ID,
        source_spreadsheet_id=SOURCE_SPREADSHEET_ID
    )
    processor.run()
