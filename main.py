import os
import sys
import time
import requests
import logging
import pandas as pd
import subprocess
import gspread  # ДОБАВЛЕНО: для работы с Google Sheets через сервисный аккаунт

# Изменен импорт для совместимости с разными версиями
try:
    from io import StringIO
except ImportError:
    from StringIO import StringIO
from datetime import datetime, timedelta

# ==================== КОНФИГУРАЦИЯ ====================
# Ключ таблицы теперь берется из секретов GitHub (или используется старый как запасной)
SPREADSHEET_KEY = os.environ.get("SUBSCRIBERS_SERVICES_SPREADSHEET_ID", "ВАШ_ID_ТАБЛИЦЫ_ПОДПИСКИ")

# Ссылка для скачивания (используется как резервная, если нет сервисного аккаунта)
SUBSCRIBERS_SERVICES_URL = os.environ.get(
    "SUBSCRIBERS_SERVICES_URL", 
    f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_KEY}/gviz/tq?tqx=out:csv&sheet=Users"
)

TARGET_COLUMN = "Import_Subscribers"  # Название столбца для поиска email
# Словарь: скрипт → пауза в секундах ПОСЛЕ его выполнения (перед следующим скриптом).
# Большая пауза перед import_script_0.py достигается через паузу после yml_parser_AGER.py (последний парсер).
SCRIPTS_TO_RUN = {
    #"yml_parser_FOOTBALLERS.py": 60,
    #"yml_parser_SPECULANT.py": 60,
    #"yml_parser_KIRS.py": 60,
    #"yml_parser_MOYDROP.py": 60,
    #"yml_parser_BAYRAKTAR.py": 60,
    #"yml_parser_IZIDROP.py": 60,
    #"yml_parser_BAGSROOM.py": 60,
    #"yml_parser_AGER.py": 70,
    "import_script_0.py": 90,
    "IMPORT_PROM_UA.py": 60,
    #"Product_Correction.py": 10,
    #"SYNC_HEAVY_PROM_UA.py": 10,
    #"API_imports_on_prom.ua.py": 10,
}
LOG_FILE_TO_CHECK = "feed_processor.log"  # Лог-файл для проверки
MAX_LOG_LINES = 200  # Максимальное количество строк в логе
API_DELAY = 5  # Задержка между запросами к API
MAX_RETRIES = 7  # Максимальное количество попыток при ошибке 429

# URL для вызова вашего Google Apps Script (подтягивается из секретов GitHub)
GOOGLE_APPS_SCRIPT_URL = os.environ.get(
    "GOOGLE_APPS_SCRIPT_URL", 
    "ВАШ_GOOGLE_APPS_SCRIPT_URL"
)

# --- TELEGRAM CONFIG ---
# Скрипт сначала ищет их в переменных окружения (для GitHub),
# если не находит - использует значения по умолчанию (для локального запуска)
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "ВАШ_ТОКЕН_БОТА")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "ВАШ_CHAT_ID")
# =======================================================

# Настройка логгера (с дублированием в стандартный вывод для GitHub Actions)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('launcher.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)


def send_telegram_message(text):
    """Отправляет сообщение в Telegram"""
    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == "ВАШ_ТОКЕН_БОТА":
        logging.warning("Telegram токен не настроен. Пропуск отправки.")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        response = requests.post(url, json={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": text,
            "parse_mode": "HTML"
        }, timeout=10)
        response.raise_for_status()
        logging.info("Уведомление в Telegram отправлено.")
    except Exception as e:
        logging.error(f"Ошибка отправки в Telegram: {e}")


def clean_email(email_str):
    """Очистка email от пробелов, кавычек, невидимых символов и приведение к нижнему регистру"""
    if not email_str:
        return ""
    return str(email_str).strip().strip("'\"").replace('\xa0', '').strip().lower()


def check_user_file():
    """Проверка наличия и содержимого email пользователя"""
    env_email = os.environ.get("USER_EMAIL")
    if env_email:
        return clean_email(env_email)

    if not os.path.exists("user.txt"):
        logging.error("Файл user.txt не найден")
        print("\nФайл user.txt не найден! Пожалуйста, создайте его и добавьте ваш Email")
        time.sleep(5)
        exit()

    with open("user.txt", "r", encoding="utf-8") as f:
        user_email = clean_email(f.read())

    if not user_email:
        logging.error("Файл user.txt пуст")
        print("\nФайл user.txt пуст. Пожалуйста, добавьте ваш Email в файл и перезапустите скрипт")
        time.sleep(5)
        exit()

    return user_email


def safe_api_request(url, retries=MAX_RETRIES):
    """Безопасный запрос к API с повторными попытками при ошибках"""
    for attempt in range(retries):
        try:
            time.sleep(API_DELAY)
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            return response
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 429:
                wait_time = (2 ** attempt) * 5
                logging.warning(f"Превышен лимит запросов к API. Попытка {attempt + 1}/{retries}. Ждем {wait_time} секунд...")
                time.sleep(wait_time)
                continue
            logging.error(f"HTTP ошибка при обращении к {url}: {e}")
            raise
        except requests.RequestException as e:
            logging.error(f"Сетевая ошибка при обращении к {url}: {e}")
            raise
    logging.error(f"Не удалось выполнить запрос к {url} после {retries} попыток")
    raise Exception(f"Не удалось получить данные после {retries} попыток")


def check_subscription():
    """Проверка подписки пользователя в Google Таблице"""
    if os.environ.get("SKIP_SUBSCRIPTION_CHECK", "").lower() in ("true", "1", "yes"):
        logging.info("SKIP_SUBSCRIPTION_CHECK включен, проверка подписки пропущена.")
        return True

    logging.info("Начинаем проверку подписки...")
    user_email = check_user_file()
    logging.info(f"Email пользователя для проверки (длина {len(user_email)}): '{user_email}'")

    try:
        # Вариант 1: Через сервисный аккаунт (GitHub Actions и локальный запуск с ключом)
        if os.path.exists("key_sheet.json"):
            logging.info("Используем key_sheet.json. Подключаемся через сервисный аккаунт...")
            gc = gspread.service_account(filename="key_sheet.json")
            sh = gc.open_by_key(SPREADSHEET_KEY)
            worksheet = sh.worksheet("Users") # лист "Users", где хранятся подписчики

            records = worksheet.get_all_records()
            df = pd.DataFrame(records)

        # Вариант 2: Через публичный CSV (если ключ недоступен)
        else:
            logging.info("Файл key_sheet.json не найден. Пробуем получить данные через публичный CSV...")
            response = safe_api_request(SUBSCRIBERS_SERVICES_URL)
            csv_data = StringIO(response.text)
            df = pd.read_csv(csv_data)

        if TARGET_COLUMN not in df.columns:
            logging.error(f"Столбец '{TARGET_COLUMN}' не найден в таблице Users. Доступные столбцы: {list(df.columns)}")
            return False

        raw_emails = df[TARGET_COLUMN].dropna().tolist()
        emails = [clean_email(e) for e in raw_emails if clean_email(e)]
        logging.info(f"Получено {len(emails)} emails из листа '{TARGET_COLUMN}': {emails}")

        is_valid = user_email in emails
        if not is_valid:
            logging.error(f"Email '{user_email}' не найден среди разрешенных в таблице: {emails}")
        else:
            logging.info(f"Email '{user_email}' успешно авторизован!")

        return is_valid

    except Exception as e:
        logging.error(f"Ошибка при проверке подписки: {e}")
        return False



def check_and_clear_log():
    """Проверяет и очищает лог-файл при необходимости"""
    if not os.path.exists(LOG_FILE_TO_CHECK):
        logging.warning(f"Лог-файл {LOG_FILE_TO_CHECK} не найден")
        return

    try:
        with open(LOG_FILE_TO_CHECK, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        if len(lines) > MAX_LOG_LINES:
            logging.info(f"Лог-файл превысил {MAX_LOG_LINES} строк. Очистка...")
            with open(LOG_FILE_TO_CHECK, 'w', encoding='utf-8') as f:
                f.write(f"Лог очищен {datetime.now()}\n")
            logging.info("Лог-файл успешно очищен")
    except Exception as e:
        logging.error(f"Ошибка при работе с лог-файлом: {e}")


def run_script_with_retries(script_path, retries=3):
    """Запускает скрипт с несколькими попытками при ошибках"""
    for attempt in range(1, retries + 1):
        try:
            logging.info(f"Попытка {attempt}/{retries} запуска скрипта: {script_path}")
            result = subprocess.run(
                ["python", script_path],
                check=True,
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=1200
            )
            logging.info(f"Скрипт {script_path} успешно выполнен (попытка {attempt}/{retries}).")
            return result
        except subprocess.TimeoutExpired:
            logging.error(f"Скрипт {script_path} превысил время выполнения (попытка {attempt}/{retries})")
            if attempt == retries:
                logging.error(f"Скрипт {script_path} не выполнен после {retries} попыток из-за таймаута.")
                raise
            time.sleep(60 * attempt)
        except subprocess.CalledProcessError as e:
            logging.error(f"Ошибка в скрипте {script_path} (попытка {attempt}/{retries}). Код возврата: {e.returncode}")
            logging.error(f"Stdout: {e.stdout}")
            logging.error(f"Stderr: {e.stderr}")
            if attempt == retries:
                logging.error(f"Скрипт {script_path} не выполнен после {retries} попыток из-за ошибки.")
                raise
            time.sleep(60 * attempt)
        except Exception as e:
            logging.error(f"Неожиданная ошибка при запуске {script_path} (попытка {attempt}/{retries}): {e}")
            if attempt == retries:
                raise
            time.sleep(60 * attempt)
    return None


def trigger_google_apps_script(url, retries=3):
    """Отправляет запрос для запуска Google Apps Script с повторными попытками."""
    if not url or url == "ВАШ_URL_GOOGLE_APPS_SCRIPT_ЗДЕСЬ" or url == "ВАШ_GOOGLE_APPS_SCRIPT_URL":
        logging.warning("URL Google Apps Script не настроен. Пропуск запуска.")
        print("URL Google Apps Script не настроен. Пропуск запуска.")
        return False

    for attempt in range(1, retries + 1):
        logging.info(f"Запуск Google Apps Script (попытка {attempt}/{retries}): {url}")
        print(f"Запуск Google Apps Script (попытка {attempt}/{retries})...")
        try:
            response = requests.get(url, timeout=1800)
            response.raise_for_status()
            logging.info(f"Google Apps Script успешно запущен. Ответ сервера: {response.status_code}")
            print(f"Google Apps Script успешно запущен. Ответ сервера: {response.status_code}")
            return True
        except Exception as e:
            logging.error(f"Ошибка при вызове Google Apps Script (попытка {attempt}/{retries}): {e}")
            print(f"Ошибка при вызове Google Apps Script (попытка {attempt}/{retries}): {e}")
            if attempt < retries:
                wait = 30 * attempt
                logging.info(f"Ожидание {wait} сек перед следующей попыткой...")
                time.sleep(wait)

    logging.error(f"Google Apps Script не удалось запустить после {retries} попыток.")
    print(f"Google Apps Script не удалось запустить после {retries} попыток.")
    return False


def dump_script_log(log_file: str, script_name: str, last_n_lines: int = 80):
    """Читает лог-файл скрипта и выводит последние строки в консоль и в launcher.log.
    Вызывается при ошибке, чтобы причина была видна в GitHub Actions."""
    if not os.path.exists(log_file):
        print(f"[{script_name}] Лог-файл '{log_file}' не найден.")
        logging.warning(f"[{script_name}] Лог-файл '{log_file}' не найден.")
        return
    try:
        with open(log_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        tail = lines[-last_n_lines:] if len(lines) > last_n_lines else lines
        sep = "=" * 20
        sep2 = "=" * 60
        header = f"\n{sep} ЛОГ {script_name} (последние {len(tail)} строк) {sep}\n"
        footer = f"{sep2}\n"
        output = header + "".join(tail) + footer
        print(output)
        logging.info(output)
    except Exception as e:
        print(f"[{script_name}] Не удалось прочитать лог '{log_file}': {e}")
        logging.error(f"[{script_name}] Не удалось прочитать лог '{log_file}': {e}")

def get_products_count(log_file: str = "google_sheets_import.log") -> str:
    """Читает лог import_script_0.py и возвращает количество обработанных товаров."""
    if not os.path.exists(log_file):
        return "неизвестно (лог не найден)"
    try:
        with open(log_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        # Ищем последнюю строку с "Объединено продуктов:"
        for line in reversed(lines):
            if "Объединено продуктов:" in line:
                # Извлекаем число из строки вида "... - INFO - Объединено продуктов: 19709"
                return line.strip().split("Объединено продуктов:")[-1].strip()
        return "не найдено в логе"
    except Exception as e:
        logging.error(f"Ошибка при чтении количества товаров: {e}")
        return "ошибка чтения"

# Маппинг: скрипт → его лог-файл (добавьте сюда другие скрипты если нужно)
SCRIPT_LOG_FILES = {
    "import_script_0.py": "google_sheets_import.log",
    "yml_parser_FOOTBALLERS.py": "feed_processor.log",
    "yml_parser_KIRS.py": "feed_processor.log",
    "yml_parser_MOYDROP.py": "feed_processor.log",
    "yml_parser_BAYRAKTAR.py": "feed_processor.log",
    "yml_parser_IZIDROP.py": "feed_processor.log",
    "yml_parser_BAGSROOM.py": "feed_processor.log",
    "yml_parser_AGER.py": "feed_processor.log",
    "IMPORT_PROM_UA.py": "feed_processor.log",
    "Product_Correction.py": "feed_processor.log",
}


def run_scripts_sequentially():
    """Запускает скрипты последовательно с контролем квот"""
    total_start_time = time.time()
    execution_times = {}
    report_text = ""  # Будущий текст для Telegram

    for script_name, delay_after in SCRIPTS_TO_RUN.items():
        if not os.path.exists(script_name):
            logging.error(f"Скрипт {script_name} не найден!")
            continue

        script_start_time = time.time()
        logging.info(f"\n======== ЗАПУСК СКРИПТА: {script_name} ========")
        print(f"\nЗапуск скрипта: {script_name}")

        script_successfully_completed = False
        try:
            result = run_script_with_retries(script_name)

            if result:
                if result.stdout:
                    print(result.stdout)
                    logging.info(f"Вывод {script_name}:\n{result.stdout}")
                if result.stderr:
                    print(result.stderr)
                    logging.warning(f"Сообщения в stderr от {script_name}:\n{result.stderr}")

                if "429" in result.stderr or "Quota exceeded" in result.stderr:
                    logging.warning("Обнаружена ошибка квоты в выводе скрипта!")
                    time.sleep(30)

                if result.returncode == 0:
                    script_successfully_completed = True
                else:
                    logging.error(f"Скрипт {script_name} завершился с кодом ошибки: {result.returncode}")

        except Exception as e:
            logging.error(f"Фатальная ошибка при выполнении {script_name}: {e}")

        # Если скрипт завершился с ошибкой — сразу выводим его лог в консоль GitHub Actions
        if not script_successfully_completed and script_name in SCRIPT_LOG_FILES:
            dump_script_log(SCRIPT_LOG_FILES[script_name], script_name)

        exec_time = time.time() - script_start_time
        execution_times[script_name] = exec_time

        print(f"Время выполнения {script_name}: {exec_time:.2f} секунд")
        logging.info(f"Время выполнения {script_name}: {exec_time:.2f} секунд")

        if script_name == "import_script_0.py":
            if script_successfully_completed:
                logging.info(f"Скрипт {script_name} успешно завершен. Запускаем Google Apps Script для перевода...")
            else:
                logging.warning(
                    f"Скрипт {script_name} завершился с ошибкой, но Google Apps Script будет запущен в любом случае.")
                print(f"ВНИМАНИЕ: {script_name} завершился с ошибкой. Google Apps Script запускается принудительно...")
            trigger_google_apps_script(GOOGLE_APPS_SCRIPT_URL)

        if delay_after > 0:
            logging.info(f"Ожидание {delay_after} сек перед следующим скриптом...")
            print(f"Ожидание {delay_after} сек перед следующим скриптом...")
            time.sleep(delay_after)

    # Формирование финальной статистики
    total_time = time.time() - total_start_time
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S,%f')[:-3]

    # Собираем блок текста в том формате, который вы просили (как в логах)
    stats_header = "<code>====== Статистика выполнения ======</code>"
    stats_total = f"{now_str} - INFO - Общее время выполнения: {total_time:.2f} секунд"

    lines = []
    lines.append(stats_header)
    lines.append(f"<code>{stats_total}")

    for script, t in execution_times.items():
        line = f"{now_str} - INFO - {script}: {t:.2f} секунд"
        lines.append(line)
        logging.info(line)  # Дублируем в лог как обычно

    # ← ЗДЕСЬ, после цикла (без отступа)
    products_count = get_products_count()
    products_line = f"{now_str} - INFO - 📦 Всего товаров обработано: {products_count}"
    lines.append(f"<code>{products_line}</code>")
    logging.info(products_line)
    final_status = f"{now_str} - INFO - ======= ВСЕ СКРИПТЫ УСПЕШНО ВЫПОЛНЕНЫ (или обработаны согласно логике) =======</code>"
    lines.append(final_status)

    # Объединяем всё в одно сообщение
    full_report = "\n".join(lines)

    # Вывод в консоль
    print("\n" + full_report.replace("<code>", "").replace("</code>", ""))

    # Отправка в Telegram
    send_telegram_message(full_report)


def clear_log_files():
    """Очищает все лог файлы перед началом работы"""
    log_files = ['launcher.log', 'feed_processor.log', 'google_sheets_import.log']

    for log_file in log_files:
        try:
            if os.path.exists(log_file):
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write(f"Log cleared at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                logging.info(f"Лог-файл {log_file} очищен")
        except Exception as e:
            logging.error(f"Ошибка при очистке лога {log_file}: {e}")


def main():
    """Основная функция управления процессом"""
    logging.info("======= СТАРТ ПРОГРАММЫ =======")
    clear_log_files()

    if not check_subscription():
        logging.error("Доступ запрещен: Email не найден в списке подписчиков")
        print('Доступ запрещен. Email не найден в списке подписчиков.')
        send_telegram_message("❌ Доступ запрещен: Email не найден в списке.")
        time.sleep(5)
        exit()

    logging.info("Подписка подтверждена. Продолжение работы...")
    print('Подписка подтверждена. Продолжение работы...')

    check_and_clear_log()
    run_scripts_sequentially()

    logging.info("======= ВСЕ СКРИПТЫ УСПЕШНО ВЫПОЛНЕНЫ =======")


if __name__ == '__main__':
    main()
