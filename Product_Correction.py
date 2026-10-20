import os
import gspread

# ----------------------------------------------------
# Настройки: получаем ID из переменной окружения
# ----------------------------------------------------
SPREADSHEET_ID = os.environ.get('IMPORT_PROM_UA_SPREADSHEET_ID', '1WKlCqKReIp7ZByP6O8ctmjJ3Nm7eFsrcVv77y-9XKd4')
WORKSHEET_NAME = 'Export Products Sheet'
CREDENTIALS_FILE = 'key_sheet.json'

# Названия столбцов (должны совпадать с названиями в вашей таблице)
PRODUCT_CODE_COL_NAME = 'Код_товару'
PRODUCT_TYPE_COL_NAME = 'Тип_товару'
MIN_ORDER_QTY_COL_NAME = 'Мінімальний_обсяг_замовлення'
PRICE_COL_NAME = 'Ціна'

# ----------------------------------------------------
# Список товаров для обновления
# Заполните этот список кодами товаров и новыми значениями
# ----------------------------------------------------
PRODUCTS_TO_UPDATE = [
    {
        'product_code': '1434',
        'new_product_type': 'w',
        'new_min_order_qty': '50'
    },
    {
        'product_code': '1788',
        'new_product_type': 'w',
        'new_min_order_qty': '50'
    },
    {
        'product_code': '2341',
        'new_product_type': 'w',
        'new_min_order_qty': '20'
    },
    {
        'product_code': '1622ОПТ',          # <-- Замените на реальный код товара
        'new_product_type': 'w',
        'new_min_order_qty': '10',
        'new_price': '97'               # <-- Дополнительный столбец "Ціна" только для этого товара
    },
    # Шаблон для добавления новых товаров (без поля цены):
    # {
    #     'product_code': 'КОД_ТОВАРА',
    #     'new_product_type': 'НОВОЕ_ЗНАЧЕНИЕ_ТИПА',
    #     'new_min_order_qty': 'НОВОЕ_ЗНАЧЕНИЕ_МИНИМУМА'
    # },
    #
    # Шаблон для товара с обновлением цены:
    # {
    #     'product_code': 'КОД_ТОВАРА',
    #     'new_product_type': 'НОВОЕ_ЗНАЧЕНИЕ_ТИПА',
    #     'new_min_order_qty': 'НОВОЕ_ЗНАЧЕНИЕ_МИНИМУМА',
    #     'new_price': 'НОВАЯ_ЦЕНА'
    # },
]

# ----------------------------------------------------
# Основная логика скрипта
# ----------------------------------------------------

def update_multiple_products():
    """
    Ищет несколько товаров по коду и обновляет указанные столбцы.
    Для большинства товаров обновляются два столбца: Тип_товару и Мінімальний_обсяг_замовлення.
    Если в данных товара присутствует поле 'new_price' — дополнительно обновляется столбец Ціна.
    """
    try:
        # Авторизация и открытие таблицы
        gc = gspread.service_account(filename=CREDENTIALS_FILE)
        sh = gc.open_by_key(SPREADSHEET_ID)
        worksheet = sh.worksheet(WORKSHEET_NAME)

        print(f"Подключено к листу: {WORKSHEET_NAME}")

        # Получение заголовков
        headers = worksheet.row_values(1)

        try:
            # Находим индексы нужных столбцов
            product_code_col_index = headers.index(PRODUCT_CODE_COL_NAME) + 1
            product_type_col_index = headers.index(PRODUCT_TYPE_COL_NAME) + 1
            min_order_qty_col_index = headers.index(MIN_ORDER_QTY_COL_NAME) + 1
            price_col_index = headers.index(PRICE_COL_NAME) + 1
        except ValueError as e:
            print(f"Ошибка: Не удалось найти столбец с именем {e}. Проверьте названия столбцов.")
            return

        # Находим все значения в столбце "Код_товару"
        product_codes_list = worksheet.col_values(product_code_col_index)
        all_updates = []

        # Итерируемся по списку товаров, которые нужно обновить
        for product_data in PRODUCTS_TO_UPDATE:
            product_code = product_data['product_code']
            new_product_type = product_data['new_product_type']
            new_min_order_qty = product_data['new_min_order_qty']
            new_price = product_data.get('new_price')  # None, если поле не указано

            try:
                # Находим индекс строки для текущего товара
                row_index_to_update = product_codes_list.index(product_code) + 1
                print(f"Найден товар с кодом '{product_code}' в строке {row_index_to_update}.")

                # Формируем адреса ячеек
                update_range_type = gspread.utils.rowcol_to_a1(row_index_to_update, product_type_col_index)
                update_range_qty = gspread.utils.rowcol_to_a1(row_index_to_update, min_order_qty_col_index)

                all_updates.append({'range': update_range_type, 'values': [[new_product_type]]})
                all_updates.append({'range': update_range_qty, 'values': [[new_min_order_qty]]})

                # Обновляем столбец "Ціна" только если поле задано для этого товара
                if new_price is not None:
                    update_range_price = gspread.utils.rowcol_to_a1(row_index_to_update, price_col_index)
                    all_updates.append({'range': update_range_price, 'values': [[new_price]]})
                    print(f"  → Для товара '{product_code}' будет обновлена цена: {new_price}")

            except ValueError:
                print(f"Предупреждение: Товар с кодом '{product_code}' не найден. Пропускаем его.")
                continue

        if all_updates:
            # Отправляем все обновления одним пакетом
            worksheet.batch_update(all_updates)
            print("Все данные успешно обновлены.")
        else:
            print("Не найдено товаров для обновления. Никаких изменений не внесено.")

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Ошибка: Таблица с ID '{SPREADSHEET_ID}' не найдена. Проверьте ID.")
    except gspread.exceptions.WorksheetNotFound:
        print(f"Ошибка: Лист с названием '{WORKSHEET_NAME}' не найден. Проверьте название.")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

if __name__ == '__main__':
    update_multiple_products()
