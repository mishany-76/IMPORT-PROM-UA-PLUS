# IMPORT-PROM-UA-PLUS

Двухэшелонный конвейер импорта и синхронизации товаров на маркетплейс Prom.ua с полной изоляцией данных, динамическим управлением через секреты GitHub Actions и поддержкой 8 поставщиков (включая нового поставщика Bayraktar Military).

---

## 🎯 Архитектура двухэшелонного импорта

```
[Поставщики YML / Sheets]
  ├── AGER, FOOTBALLERS, IZIDROP, MOYDROP, BAYRAKTAR, KIRS, BAGSROOM, SPECULANT
  └── (Скрипты yml_parser_*.py параллельно наполняют PARS-таблицы)
          │
          ▼
[Шаблон_імпорту_prom.ua_PLUS] (Промежуточная таблица-агрегатор)
  └── import_script_0.py нормализует и объединяет данные всех поставщиков
          │
          ▼
[IMPORT-PROM-UA-PLUS] (Таблица легкого эшелона)
  ├── IMPORT_PROM_UA.py выполняет smart diff, обновляет остатки, цены, категории
  └── Product_Correction.py применяет точечные бизнес-правила (опт, кратность)
          │
          ▼
[IMPORT-PROM-UA-PLUS-HEAVY] (Таблица тяжелого эшелона)
  └── SYNC_HEAVY_PROM_UA.py синхронизирует изменения из легкого эшелона,
      сохраняя полные описания, ссылки на изображения и все атрибуты
          │
          ▼
[Cloudflare R2 + Prom.ua API]
  └── API_imports_on_prom.ua.py экспортирует XLSX в бакет Cloudflare R2
      и инициирует импорт по ссылке через Prom.ua API для магазина VaillantVibe
```

---

## 📦 Список поставщиков и их параметры

| Поставщик | Тип источника | Город отправки | Скрипт парсера |
| :--- | :--- | :--- | :--- |
| **AGER** | YML Feed | Одеса | `yml_parser_AGER.py` |
| **FOOTBALLERS** | Google Sheet | Одеса | `yml_parser_FOOTBALLERS.py` |
| **IZIDROP** | YML Feed | Одеса | `yml_parser_IZIDROP.py` |
| **MOYDROP** | YML Feed | Одеса | `yml_parser_MOYDROP.py` |
| **BAYRAKTAR** | YML Feed (MyDrop API) | с.Шаровечка, Хмельницька область | `yml_parser_BAYRAKTAR.py` |
| **KIRS** | YML Feed | Дніпро | `yml_parser_KIRS.py` |
| **BAGSROOM** | YML Feed | Київ | `yml_parser_BAGSROOM.py` |
| **SPECULANT** | YML Feed | Полтава | `yml_parser_SPECULANT.py` |

---

## 🔐 Необходимые секреты GitHub Actions

Все чувствительные данные вынесены в **Settings -> Secrets and variables -> Actions**:

### 1. Системные и сервисные секреты
- `KEY_SHEET_JSON`: Полное содержимое сервисного аккаунта Google Cloud (`client_email`, `private_key` и др.)
- `USER_EMAIL`: Email владельца Google Sheets таблиц для предоставления прав доступа
- `TELEGRAM_BOT_TOKEN`: Токен Telegram-бота для отправки отчетов и алертов
- `TELEGRAM_CHAT_ID`: ID чата Telegram для уведомлений
- `GOOGLE_APPS_SCRIPT_URL`: Webhook Google Apps Script (опционально)

### 2. Ссылки на фиды поставщиков
- `AGER_FEED_URL`
- `AGER_USER_AGENT`
- `BAGSROOM_FEED_URL`
- `BAYRAKTAR_FEED_URL`
- `FOOTBALLERS_FEED_URL`
- `IZIDROP_FEED_URL`
- `KIRS_FEED_URL`
- `MOYDROP_FEED_URL`
- `SPECULANT_FEED_URL`
- `FOOTBALLERS_FEED_SPREADSHEET_ID`
- `SUBSCRIBERS_SERVICES_SPREADSHEET_ID`
- `SUBSCRIBERS_SERVICES_URL`

### 3. ID парсинг-таблиц поставщиков (`PARS_*`)
- `PARS_AGER_SPREADSHEET_ID`
- `PARS_BAGSROOM_SPREADSHEET_ID`
- `PARS_BAYRAKTAR_SPREADSHEET_ID`
- `PARS_FOOTBALLERS_SPREADSHEET_ID`
- `PARS_IZIDROP_SPREADSHEET_ID`
- `PARS_KIRS_SPREADSHEET_ID`
- `PARS_MOYDROP_SPREADSHEET_ID`
- `PARS_SPECULANT_SPREADSHEET_ID`

### 4. ID целевых таблиц двухэшелонного конвейера
- `IMPORT_TEMPLATE_PROM_SPREADSHEET_ID`: ID промежуточной таблицы `Шаблон_імпорту_prom.ua_PLUS`
- `IMPORT_PROM_UA_SPREADSHEET_ID`: ID легкой таблицы `IMPORT-PROM-UA-PLUS`
- `IMPORT_PROM_UA_HEAVY_SPREADSHEET_ID`: ID тяжелой таблицы `IMPORT-PROM-UA-PLUS-HEAVY`

### 5. Интеграция Cloudflare R2 и Prom.ua API
- `R2_ACCESS_KEY`: Access Key ID для Cloudflare R2
- `R2_SECRET_KEY`: Secret Access Key для Cloudflare R2
- `R2_ACCOUNT_ID`: Account ID Cloudflare
- `R2_BUCKET_NAME`: Имя бакета R2
- `R2_PUBLIC_DOMAIN`: Публичный домен R2 (например, `https://pub-xxxx.r2.dev`)
- `PROM_TOKEN_CITY`: Токен API магазина Prom.ua `CityStyle` (Флагман, ID `c3973756`)
- `PROM_TOKEN_BEST`: Токен API магазина Prom.ua `BestSeller` (Сателлит, ID `c3972457`)
- `PROM_TOKEN_VAILLANT`: Токен API магазина Prom.ua `VaillantVibe` (Сателлит, ID `c4082383`)

---

## 🚀 Порядок локального запуска

1. Установить зависимости:
   ```bash
   pip install -r requirements.txt
   ```
2. Положить в корень `key_sheet.json` и `user.txt`.
3. Запустить полный цикл:
   ```bash
   python main.py
   ```

---

## 🛡️ Архитектурные стандарты и защита данных

1. **Каноническое дерево категорий (62 группы):**
   - 3-уровневая структура каталога (5 корневых, 23 подраздела, 34 конечные группы).
   - Логистическая изоляция складов по суффиксам групп: `(НП / УП)` для AGER, `(НП / МР)` для KIRS, группы без суффикса (только Нова Пошта) для остальных 6 поставщиков.
2. **Защита от модерационных блокировок Prom.ua:**
   - Автоматическая фильтрация стоп-слов контрафакта (`репліка`, `люкс копія` и др.) с заменой на нейтральные термины (`модель`).
   - Унификация групп разновидностей (`unify_variation_groups`) для исключения ошибки несовпадения ID группы основного товара и его вариаций.
3. **Универсальный ремонт размерных сеток AGER:**
   - Двухуровневая архитектура: эталонный справочник фабричных замеров 27 артикулов обуви AGER + универсальный fallback-парсер с автоматической коррекцией высоты платформы (3 см) и спаренных размеров.
   - Монолитная HTML5-сборка таблиц без паразитных текстовых узлов и foster parenting.
