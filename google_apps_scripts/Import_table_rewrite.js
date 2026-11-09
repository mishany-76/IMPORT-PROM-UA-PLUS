/**
 * ============================================================================
 * PROM.UA SMART AI GENERATOR & SEO ENGINE (OmniRoute / OpenAI)
 * ============================================================================
 * Версия: 2.7 (Universal Multi-Model Engine: Deep Extractor, Expanded Prompting & Deadlock Protection)
 * Автор: Mikhail & Orchestrator
 * Дата: 28 октября 2026 г.
 *
 * ЧТО УЛУЧШЕНО В ВЕРСИИ 2.7:
 * 1. Совместимость с GPT-5.4-mini и GPT-4.1-mini:
 *    - Промпт переработан под особенности современных моделей: требование к объему
 *      задается через 50-65 уникальных поисковых фраз (100-130+ слов), с полным охватом
 *      всех 8 семантических категорий, что гарантирует плотность 900-1015 символов.
 *    - Температура адаптирована до 0.5-0.6 для богатства лексики без галлюцинаций.
 * 2. Всеядный полиморфный экстрактор (flattenObjectTree + findDeepFieldValue):
 *    - Поддержка любой степени вложенности (keywords.ru, tasks[0].keywords_uk, data.translations).
 *    - Поддержка любых языковых кодов (uk, ua, ukr, ukrainian, ru, russian).
 *    - Поддержка любых нотаций: snake_case, camelCase, PascalCase, тире, пробелы.
 *    - Устойчивое сопоставление строк: row, row_number, id, а также позиционный фоллбек.
 * 3. Информативное логирование и сохранение строгого порога качества:
 *    - Сохранен жесткий порог длины (>= 500 символов) для недопущения куцых тегов в Prom.ua.
 *    - В логи Cloud Console выводятся точные длины: RU len: X, UA len: Y.
 * 4. Защита от зависаний (Anti-Deadlock):
 *    - При недостаточности длины выполняется 1 автоматический повтор с усиленным требованием.
 *    - Если батч не проходит порог даже после повтора, указатель STRING LOG продвигается,
 *      исключая зацикливание скрипта на одной строке и не блокируя каталог из 21+ тыс. товаров.
 * ============================================================================
 */

const ENGINE_CONFIG = {
  BATCH_SIZE: 3,        // 3 товара в батче (оптимально для 2 языков по 1000 символов)
  ITERATIONS: 2,        // 2 итерации за запуск (6 товаров за 5 минут)
  SYSTEMIC_SHEET: "Systemic",
  DATA_SHEET: "Export Products Sheet"
};

function main() {
  Logger.log("Smart Prom Engine V2.7 started.");

  const props = PropertiesService.getScriptProperties();
  props.setProperty("SCRIPT_STATUS", "RUNNING");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sysSheet = ss.getSheetByName(ENGINE_CONFIG.SYSTEMIC_SHEET);
    const dataSheet = ss.getSheetByName(ENGINE_CONFIG.DATA_SHEET);

    if (!sysSheet || !dataSheet) {
      Logger.log("ОШИБКА: Не найден лист 'Systemic' или 'Export Products Sheet'.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    const sysHeaders = sysSheet.getRange(1, 1, 1, sysSheet.getLastColumn()).getValues()[0];
    const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];

    // 1. Загружаем системные настройки и правила
    const settings = loadSystemicSettings(sysSheet, sysHeaders);
    if (!settings.apiKey) {
      Logger.log("ОШИБКА: OMNIROUTE_API_KEY не заполнен на листе Systemic!");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    const taskConfigs = loadTaskConfigurations(sysSheet, sysHeaders);
    if (taskConfigs.length === 0) {
      Logger.log("ОШИБКА: Не найдены конфигурации задач на листе Systemic.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    const supplierColIdx = findColumnIndex(dataHeaders, ["Особисті_нотатки"]);
    if (supplierColIdx === -1) {
      Logger.log("ОШИБКА: В Export Products Sheet не найден столбец 'Особисті_нотатки'.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    const totalRows = dataSheet.getLastRow();
    let currentRow = settings.stringLog + 1;
    Logger.log(`Старт с строки ${currentRow}. Всего строк в таблице: ${totalRows}`);

    let completedIterations = 0;

    while (currentRow <= totalRows && completedIterations < ENGINE_CONFIG.ITERATIONS) {
      // 2. Сканируем строки и собираем батч товаров, требующих обработки
      const batch = [];
      let scanRow = currentRow;

      while (scanRow <= totalRows && batch.length < ENGINE_CONFIG.BATCH_SIZE) {
        const rowData = dataSheet.getRange(scanRow, 1, 1, dataSheet.getLastColumn()).getValues()[0];
        const supplier = String(rowData[supplierColIdx] || "").trim().toUpperCase();

        if (supplier) {
          const matchingTasks = taskConfigs.filter(cfg => cfg.supplierName === supplier);
          const pendingTasks = [];

          for (const task of matchingTasks) {
            // Проверяем статус задачи по rowData прямо из памяти
            if (!isTaskCompletedFast(rowData, task.outputIndices, task.minOutputLength)) {
              pendingTasks.push(task);
            }
          }

          if (pendingTasks.length > 0) {
            batch.push({
              rowNumber: scanRow,
              rowData: rowData,
              supplier: supplier,
              tasks: pendingTasks
            });
          }
        }
        scanRow++;
      }

      if (batch.length === 0) {
        Logger.log("Все последующие строки проверены и заполнены. Завершение.");
        sysSheet.getRange(settings.stringLogCell).setValue(totalRows);
        deleteTimeTrigger();
        props.setProperty("SCRIPT_STATUS", "IDLE");
        return;
      }

      Logger.log(`Сформирован батч из ${batch.length} товаров (строки ${batch[0].rowNumber} - ${batch[batch.length - 1].rowNumber}).`);

      // 3. Выполняем AI генерацию через OmniRoute с авто-повтором и защитой от зависаний
      const result = executeBatchAI(batch, dataSheet, settings);

      if (!result.success) {
        if (result.isFatalApiError) {
          Logger.log("Сбой при выполнении батча через API (сеть / лимиты). Остановка.");
          break;
        } else {
          // Защита от мертвого цикла: продвигаем указатель вперед, чтобы не блокировать каталог
          const lastBatchRow = batch[batch.length - 1].rowNumber;
          Logger.log(`Внимание: товары батча (строки ${batch[0].rowNumber} - ${lastBatchRow}) не прошли валидацию качества (длина < 500). Сдвигаем указатель STRING LOG на строку ${lastBatchRow}, чтобы не блокировать конвейер каталога.`);
          sysSheet.getRange(settings.stringLogCell).setValue(lastBatchRow);
          currentRow = lastBatchRow + 1;
          completedIterations++;
          continue;
        }
      }

      completedIterations++;
      const nextLogVal = result.maxSavedRow > 0 ? result.maxSavedRow : batch[batch.length - 1].rowNumber;
      sysSheet.getRange(settings.stringLogCell).setValue(nextLogVal);
      currentRow = nextLogVal + 1;

      if (completedIterations < ENGINE_CONFIG.ITERATIONS && currentRow <= totalRows) {
        Utilities.sleep(settings.queryTimeFrame * 1000);
      }
    }

    // 4. Управление триггером
    if (currentRow <= totalRows) {
      setTimeTrigger();
      Logger.log(`Итерация завершена. Указатель на строке ${currentRow}. Установлен триггер на 5 минут.`);
    } else {
      deleteTimeTrigger();
      Logger.log("Обработка всех строк завершена! Триггер удален.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
    }

  } catch (err) {
    Logger.log("Ошибка в main: " + err.message);
    props.setProperty("SCRIPT_STATUS", "IDLE");
  }
}

/**
 * Выполнение батча через OpenAI Endpoint OmniRoute с адаптивным промптингом,
 * авто-повтором и всеядным извлечением
 */
function executeBatchAI(batch, dataSheet, settings) {
  const itemsPayload = batch.map(item => {
    const itemTasks = item.tasks.map(t => {
      const inputText = extractInputData(item.rowData, t.inputIndices, t.inputSeparator);
      return {
        task_type: t.taskType,
        input_data: inputText,
        output_columns_count: t.outputIndices.length
      };
    });

    return {
      row: item.rowNumber,
      supplier: item.supplier,
      tasks: itemTasks
    };
  });

  const baseSystemInstruction = 
    "Ты — ведущий SEO-эксперт и локализатор маркетплейса Prom.ua.\n" +
    "Твоя задача — пакетно обработать товары строго по правилам их задач для максимального охвата в поиске Prom.ua.\n\n" +
    "ОБЯЗАТЕЛЬНЫЕ ТРЕБОВАНИЯ ДЛЯ ЗАДАЧ:\n\n" +
    "1. ДЛЯ ЗАДАЧИ 'Keywords RU+UA':\n" +
    "   - Ты ОБЯЗАН вернуть ДВА независимых поля: 'keywords_ru' (на русском) и 'keywords_ua' (на украинском).\n" +
    "   - ОБЪЕМ И ПЛОТНОСТЬ: В каждом поле сгенерируй НЕ МЕНЕЕ 50-65 уникальных поисковых фраз (от 100 до 130 слов суммарно), чтобы максимально плотно занять лимит поля Prom.ua (строго от 900 до 1015 символов!).\n" +
    "   - КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО выдавать краткий ответ (менее 500 символов или менее 45 фраз)! Ответы короче 500 символов отбраковываются валидацией.\n" +
    "   - ФОРМАТ: Все запросы через запятую с пробелом, строго строчными буквами, без кавычек и точек.\n" +
    "   - СТОП-СЛОВА: Запрещено использовать слова 'купить' (в RU) и 'купити' (в UA).\n" +
    "   - СЕМАНТИЧЕСКИЙ ОХВАТ (обязательно включи все 8 групп):\n" +
    "     1) Прямые названия товара, точные синонимы, народные названия, профессиональный и разговорный сленг.\n" +
    "     2) Характеристики: цвет, оттенки, материал, ткань, состав, сезон, тип застежки, форма, длина, размер.\n" +
    "     3) Стиль и назначение: повседневный, деловой, молодежный, спортивный, праздничный, базовый, тренд.\n" +
    "     4) Сценарии использования: на работу, в школу, на прогулку, для тренировок, для отдыха, на подарок.\n" +
    "     5) Фразы покупателей (long-tail): 'стильный качественный недорогой красивый модный удобный практичный'.\n" +
    "     6) Целевая аудитория: мужской, женский, подростковый, для парней, для девушек (в зависимости от товара).\n" +
    "     7) Смежные и ассоциативные запросы, аналоги.\n" +
    "     8) Грамматические варианты: единственное и множественное число, разные падежи.\n\n" +
    "2. ДЛЯ ЗАДАЧИ 'Translate UA->RU':\n" +
    "   - 'title_ru': точный перевод названия (до 110 символов).\n" +
    "   - 'description_ru': качественный перевод описания со 100% сохранением всей HTML-разметки, тегов, стилей и таблиц.\n\n" +
    "3. ДЛЯ ЗАДАЧИ 'Translate RU->UA':\n" +
    "   - 'title_ua': точный перевод названия на украинский язык (до 110 символов).\n" +
    "   - 'description_ua': качественный перевод описания на украинский язык со 100% сохранением всей HTML-разметки и таблиц.\n\n" +
    "СТРОГО ВЕРНИ JSON объект вида: { \"results\": [ ... ] }\n" +
    "Пример элемента results:\n" +
    "{\n" +
    "  \"row\": 26,\n" +
    "  \"keywords_ru\": \"запрос1, запрос2, запрос3, ... (всего 50-65 фраз, длина от 900 до 1015 символов)\",\n" +
    "  \"keywords_ua\": \"запит1, запит2, запит3, ... (всего 50-65 фраз, длина от 900 до 1015 символов)\"\n" +
    "}";

  // Допускаем до 2 попыток генерации (основная + усиленный повтор при коротких ответах)
  for (let attempt = 1; attempt <= 2; attempt++) {
    let currentSystemInstruction = baseSystemInstruction;
    let temp = 0.5;

    if (attempt === 2) {
      Logger.log("Повторная попытка генерации батча с усиленным требованием к длине тегов...");
      currentSystemInstruction += "\n\nКРИТИЧЕСКИ ВАЖНО: Предыдущий ответ был отбракован из-за недостаточной длины (<500 символов)! Сгенерируй минимум 65 развернутых поисковых фраз (не менее 1000 символов в каждом поле)!";
      temp = 0.6;
    }

    const payload = {
      model: settings.modelName || "openai/gpt-4.1-mini",
      messages: [
        { role: "system", content: currentSystemInstruction },
        { role: "user", content: JSON.stringify(itemsPayload) }
      ],
      temperature: temp,
      max_tokens: 8192,
      response_format: { type: "json_object" }
    };

    const options = {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + settings.apiKey,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    let resCode = 0;
    let resText = "";

    try {
      const response = UrlFetchApp.fetch(settings.omnirouteUrl, options);
      resCode = response.getResponseCode();
      resText = response.getContentText();
    } catch (fetchErr) {
      Logger.log("Сетевая ошибка UrlFetchApp: " + fetchErr.message);
      return { success: false, maxSavedRow: 0, isFatalApiError: true };
    }

    if (resCode !== 200) {
      Logger.log(`HTTP Ошибка OmniRoute ${resCode}: ${resText}`);
      return { success: false, maxSavedRow: 0, isFatalApiError: true };
    }

    const parsed = smartUniversalParser(resText);
    if (!parsed || !parsed.results || !Array.isArray(parsed.results)) {
      Logger.log("Ошибка: парсер не смог восстановить валидный список results из ответа AI.");
      if (attempt === 2) {
        return { success: false, maxSavedRow: 0, isFatalApiError: false };
      }
      continue;
    }

    const savedRows = [];

    for (let rIdx = 0; rIdx < parsed.results.length; rIdx++) {
      const resItem = parsed.results[rIdx];
      const itemRow = resItem.row || resItem.row_number || resItem.rowNumber || resItem.id || resItem.productId || resItem.row_id;
      
      let match = batch.find(b => String(b.rowNumber) === String(itemRow));
      // Позиционное сопоставление (fallback), если модель забыла указать row
      if (!match && batch.length === parsed.results.length) {
        match = batch[rIdx];
      }
      if (!match) continue;

      const extractedTaskOutputs = extractOutputsForMatchedItem(match, resItem);
      let rowFullyValid = true;

      // ⭐️ АТОМАРНАЯ ВАЛИДАЦИЯ И ИНФОРМАТИВНОЕ ЛОГИРОВАНИЕ
      for (let tIdx = 0; tIdx < match.tasks.length; tIdx++) {
        const taskConfig = match.tasks[tIdx];
        const taskVals = extractedTaskOutputs[tIdx] || [];
        const isKeywords = taskConfig.taskType && taskConfig.taskType.toLowerCase().includes("keywords");

        for (let oIdx = 0; oIdx < taskConfig.outputIndices.length; oIdx++) {
          const val = String(taskVals[oIdx] || "").trim();
          if (isKeywords) {
            // Сохраняем строгий порог качества не менее 500 символов
            if (val.length < 500) {
              rowFullyValid = false;
            }
          } else if (val.length === 0) {
            rowFullyValid = false;
          }
        }

        // Вывод точных метрик в лог
        if (isKeywords) {
          const ruLen = String(taskVals[0] || "").trim().length;
          const uaLen = String(taskVals[1] || "").trim().length;
          if (rowFullyValid) {
            Logger.log(`Строка ${match.rowNumber}: валидация пройдена (RU len: ${ruLen}, UA len: ${uaLen}).`);
          } else {
            Logger.log(`Строка ${match.rowNumber}: ответ отбракован (RU len: ${ruLen}, UA len: ${uaLen}. Требуется >= 500 по обоим языкам).`);
          }
        }
      }

      if (!rowFullyValid) {
        continue;
      }

      // Записываем оба проверенных языка в таблицу
      for (let tIdx = 0; tIdx < match.tasks.length; tIdx++) {
        const taskConfig = match.tasks[tIdx];
        const taskVals = extractedTaskOutputs[tIdx] || [];

        for (let oIdx = 0; oIdx < taskConfig.outputIndices.length; oIdx++) {
          const col = taskConfig.outputIndices[oIdx];
          let val = String(taskVals[oIdx] || "").trim();

          if (taskConfig.taskType && taskConfig.taskType.toLowerCase().includes("keywords")) {
            val = formatPromKeywords(val);
          }

          if (col !== -1 && val) {
            dataSheet.getRange(match.rowNumber, col + 1).setValue(val);
          }
        }
      }

      savedRows.push(match.rowNumber);
    }

    Logger.log(`Батч обработан (попытка ${attempt}): успешно записано товаров: ${savedRows.length} из ${batch.length}.`);

    if (savedRows.length > 0) {
      const maxSaved = Math.max(...savedRows);
      return { success: true, maxSavedRow: maxSaved, isFatalApiError: false };
    }

    // Если ни одна строка не сохранена и это первая попытка — цикл повторит с усиленным промптом
  }

  Logger.log("Внимание: Ни одна строка батча не сохранена после 2 попыток.");
  return { success: false, maxSavedRow: 0, isFatalApiError: false };
}

/**
 * Всеядный полиморфный экстрактор выходов с поддержкой любой вложенности и любых вариантов именования
 */
function extractOutputsForMatchedItem(matchItem, resItem) {
  const resultByTask = [];
  // Рекурсивно сплющиваем весь объект любой структуры и вложенности
  const flatTreeMap = flattenObjectTree(resItem);

  for (let t = 0; t < matchItem.tasks.length; t++) {
    const task = matchItem.tasks[t];
    const taskTypeLower = task.taskType.toLowerCase();
    const taskVals = [];

    // 1. ЗАДАЧА КЛЮЧЕВЫХ СЛОВ (Keywords RU+UA)
    if (taskTypeLower.includes("keywords")) {
      const ruVal = findDeepFieldValue(flatTreeMap, resItem, ["ru", "russian"], ["keyword", "query", "search", "запрос", "тег", "ключ"]);
      const uaVal = findDeepFieldValue(flatTreeMap, resItem, ["ua", "uk", "ukr", "ukrainian"], ["keyword", "query", "search", "запрос", "тег", "ключ"]);

      if (ruVal && uaVal) {
        taskVals.push(ruVal);
        taskVals.push(uaVal);
        resultByTask.push(taskVals);
        continue;
      }
    }

    // 2. ЗАДАЧА ПЕРЕВОДА (Translate UA->RU или RU->UA)
    if (taskTypeLower.includes("translate")) {
      const isTargetRu = taskTypeLower.includes("->ru") || taskTypeLower.includes("to_ru") || taskTypeLower.includes("toru");
      const targetLangs = isTargetRu ? ["ru", "russian"] : ["ua", "uk", "ukr", "ukrainian"];

      const title = findDeepFieldValue(flatTreeMap, resItem, targetLangs, ["title", "name", "название", "заголовок"]) ||
                    resItem.title || resItem.name;
      const desc = findDeepFieldValue(flatTreeMap, resItem, targetLangs, ["description", "desc", "text", "html", "описание"]) ||
                   resItem.description || resItem.desc;

      if (title && desc) {
        taskVals.push(title);
        taskVals.push(desc);
        resultByTask.push(taskVals);
        continue;
      }
    }

    // 3. Если модель отдала массив outputs: [ [val1, val2] ] или [ val1, val2 ]
    if (resItem.outputs !== undefined && Array.isArray(resItem.outputs)) {
      if (Array.isArray(resItem.outputs[t]) && resItem.outputs[t].length === task.outputIndices.length) {
        resultByTask.push(resItem.outputs[t].map(v => String(v || "").trim()));
        continue;
      } else if (resItem.outputs.length === task.outputIndices.length && typeof resItem.outputs[0] === "string") {
        resultByTask.push(resItem.outputs.map(v => String(v || "").trim()));
        continue;
      }
    }

    // 4. Резервный сбор всех строковых значений подходящей длины из плоского пула
    const fallbackVals = [];
    const needed = task.outputIndices.length;
    const allStrings = Object.values(flatTreeMap).filter(v => v.length > 5);

    for (let o = 0; o < needed; o++) {
      fallbackVals.push(o < allStrings.length ? allStrings[o] : "");
    }
    resultByTask.push(fallbackVals);
  }

  return resultByTask;
}

/**
 * Рекурсивное сплющивание любого объекта/массива произвольной степени вложенности
 * Превращает { a: { b: { c: "val" } } } в { "a_b_c": "val" }
 */
function flattenObjectTree(obj, prefix, map) {
  if (!map) map = {};
  if (obj === null || obj === undefined) return map;

  if (typeof obj === "string" || typeof obj === "number" || typeof obj === "boolean") {
    const cleanPrefix = (prefix || "").toLowerCase().replace(/[^a-z0-9_]/g, "_");
    map[cleanPrefix] = String(obj).trim();
    return map;
  }

  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      flattenObjectTree(obj[i], prefix ? `${prefix}_${i}` : `${i}`, map);
    }
    return map;
  }

  if (typeof obj === "object") {
    const keys = Object.keys(obj);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const nextPrefix = prefix ? `${prefix}_${k}` : k;
      flattenObjectTree(obj[k], nextPrefix, map);
    }
  }

  return map;
}

/**
 * Поиск значения в объекте/дереве по языковым меткам и смысловым токенам
 */
function findDeepFieldValue(flatTreeMap, directObj, langTokens, typeTokens) {
  // 1. Проверяем прямые свойства объекта нижнего регистра
  const directKeys = Object.keys(directObj || {});
  for (let i = 0; i < directKeys.length; i++) {
    const rawKey = directKeys[i];
    const k = rawKey.toLowerCase().replace(/[^a-z0-9_]/g, "_");
    const val = directObj[rawKey];
    if (typeof val === "string" && val.trim().length > 0) {
      const hasLang = langTokens.some(l => k.includes(`_${l}`) || k.includes(`${l}_`) || k.endsWith(l) || k.startsWith(l));
      const hasType = typeTokens.some(t => k.includes(t));
      if (hasLang && hasType) {
        return val.trim();
      }
    }
  }

  // 2. Ищем по всем путям произвольной вложенности в flatTreeMap
  const paths = Object.keys(flatTreeMap || {});
  
  // Приоритет 1: путь содержит и языковой токен, и смысловой токен
  for (let i = 0; i < paths.length; i++) {
    const p = paths[i];
    const val = flatTreeMap[p];
    if (!val || val.length === 0) continue;

    const parts = p.split("_");
    const hasLang = langTokens.some(l => parts.includes(l) || p.includes(`_${l}`) || p.includes(`${l}_`) || p.endsWith(`_${l}`));
    const hasType = typeTokens.some(t => parts.includes(t) || p.includes(t));

    if (hasLang && hasType) {
      return val;
    }
  }

  // Приоритет 2: путь содержит языковой токен и общий контекст ключевых слов/тегов
  for (let i = 0; i < paths.length; i++) {
    const p = paths[i];
    const val = flatTreeMap[p];
    if (!val || val.length === 0) continue;

    const parts = p.split("_");
    const hasLang = langTokens.some(l => parts.includes(l));
    const hasKeywordContext = p.includes("keyword") || p.includes("запрос") || p.includes("tag") || p.includes("seo");

    if (hasLang && hasKeywordContext) {
      return val;
    }
  }

  return "";
}

/**
 * Быстрая проверка завершенности задачи по уже загруженным данным rowData
 */
function isTaskCompletedFast(rowData, outputIndices, minLength) {
  for (let i = 0; i < outputIndices.length; i++) {
    const idx = outputIndices[i];
    if (idx === -1) continue;
    const val = String(rowData[idx] || "").trim();
    if (val.length < (minLength || 1)) return false;
  }
  return true;
}

/**
 * Форматирование ключевых слов для Prom.ua:
 * - Все через запятую в нижнем регистре
 * - Исключение стоп-слов "купить" и "купити"
 * - Лимит до 1020 символов с аккуратным завершением по последней запятой
 */
function formatPromKeywords(text) {
  if (!text) return "";
  let clean = String(text)
    .replace(/["'{}[\]]/g, "")
    .split(/[,;\n]+/)
    .map(s => s.trim().toLowerCase())
    .filter(s => s.length > 0 && s.indexOf("купить") === -1 && s.indexOf("купити") === -1)
    .join(", ");

  if (clean.length > 1020) {
    const sub = clean.substring(0, 1020);
    const lastComma = sub.lastIndexOf(",");
    if (lastComma > 500) {
      clean = sub.substring(0, lastComma).trim();
    } else {
      clean = sub.trim();
    }
  }
  return clean;
}

/**
 * Самовосстанавливающийся всеядный парсер JSON с извлечением из любого текста/markdown
 */
function smartUniversalParser(rawText) {
  try {
    let outer;
    try {
      outer = JSON.parse(rawText);
    } catch (e) {
      outer = null;
    }

    let content = (outer && outer.choices?.[0]?.message?.content) ? outer.choices[0].message.content : rawText;

    if (typeof content === "string") {
      // Извлекаем блок кода markdown, если он есть
      const jsonBlockMatch = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
      let jsonStr = jsonBlockMatch ? jsonBlockMatch[1].trim() : content.trim();

      // Очистка от случайного префиксного текста до первой скобки { или [
      const firstBrace = jsonStr.indexOf("{");
      const firstBracket = jsonStr.indexOf("[");
      let startIdx = -1;
      if (firstBrace !== -1 && firstBracket !== -1) {
        startIdx = Math.min(firstBrace, firstBracket);
      } else {
        startIdx = Math.max(firstBrace, firstBracket);
      }

      if (startIdx > 0) {
        jsonStr = jsonStr.substring(startIdx);
      }

      // Очистка от случайного суффиксного текста после последней скобки } или ]
      const lastBrace = jsonStr.lastIndexOf("}");
      const lastBracket = jsonStr.lastIndexOf("]");
      const endIdx = Math.max(lastBrace, lastBracket);
      if (endIdx !== -1 && endIdx < jsonStr.length - 1) {
        jsonStr = jsonStr.substring(0, endIdx + 1);
      }

      try {
        content = JSON.parse(jsonStr);
      } catch (jsonErr) {
        content = repairIncompleteJson(jsonStr);
      }
    }

    if (!content) return null;
    if (content.results && Array.isArray(content.results)) return content;
    if (content.results && typeof content.results === "object") return { results: [content.results] };
    if (Array.isArray(content)) return { results: content };
    if (content.items && Array.isArray(content.items)) return { results: content.items };
    if (content.products && Array.isArray(content.products)) return { results: content.products };
    if (content.data && Array.isArray(content.data)) return { results: content.data };
    return { results: [content] };
  } catch (e) {
    Logger.log("Ошибка smartUniversalParser: " + e.message);
    return null;
  }
}

/**
 * Авто-ремонт незавершенного или обрезанного JSON
 */
function repairIncompleteJson(str) {
  let s = str.trim();
  const quoteCount = (s.match(/"/g) || []).length;
  if (quoteCount % 2 !== 0) s += '"';
  let openBrackets = (s.match(/\[/g) || []).length - (s.match(/\]/g) || []).length;
  let openBraces = (s.match(/\{/g) || []).length - (s.match(/\}/g) || []).length;

  while (openBrackets > 0) { s += "]"; openBrackets--; }
  while (openBraces > 0) { s += "}"; openBraces--; }

  try {
    return JSON.parse(s);
  } catch (err) {
    return null;
  }
}

/**
 * Сборка входного текста из колонок строки
 */
function extractInputData(rowData, indices, separator) {
  const parts = [];
  for (let i = 0; i < indices.length; i++) {
    const idx = indices[i];
    if (idx !== -1 && rowData[idx] !== undefined && rowData[idx] !== null) {
      const v = String(rowData[idx]).trim();
      if (v) parts.push(v);
    }
  }
  return parts.join(separator || ". ");
}

/**
 * Загрузка конфигураций задач поставщиков с листа Systemic
 */
function loadTaskConfigurations(sysSheet, headers) {
  const configs = [];
  const lastRow = sysSheet.getLastRow();
  if (lastRow < 2) return configs;

  const colSupplier = findColumnIndex(headers, ["Supplier Name"]);
  const colTaskType = findColumnIndex(headers, ["Task Type"]);
  const colInputs = findColumnIndex(headers, ["Input Columns (Letters)"]);
  const colOutputs = findColumnIndex(headers, ["Output Columns (Letters)"]);
  const colPrompt = findColumnIndex(headers, ["Prompt Cell"]);
  const colSysInst = findColumnIndex(headers, ["System Instruction Cell"]);
  const colMinLen = findColumnIndex(headers, ["Min Output Length"]);
  const colSep = findColumnIndex(headers, ["Input Separator"]);

  const numCols = sysSheet.getLastColumn();
  const data = sysSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const supplier = String(row[colSupplier] || "").trim().toUpperCase();
    if (!supplier) continue;

    configs.push({
      supplierName: supplier,
      taskType: String(row[colTaskType] || "").trim(),
      inputIndices: lettersToIndices(String(row[colInputs] || "")),
      outputIndices: lettersToIndices(String(row[colOutputs] || "")),
      promptCell: String(row[colPrompt] || "").trim(),
      systemInstructionCell: String(row[colSysInst] || "").trim(),
      minOutputLength: parseInt(row[colMinLen], 10) || 0,
      inputSeparator: String(row[colSep] || "\n")
    });
  }

  return configs;
}

/**
 * Загрузка глобальных параметров с листа Systemic
 */
function loadSystemicSettings(sysSheet, headers) {
  const stringLogIdx = findColumnIndex(headers, ["STRING LOG"]);
  const timeFrameIdx = findColumnIndex(headers, ["QUERY_TIME_FRAME"]);
  const omniKeyIdx = findColumnIndex(headers, ["OMNIROUTE_API_KEY"]);
  const omniUrlIdx = findColumnIndex(headers, ["OMNIROUTE_URL"]);
  const modelIdx = findColumnIndex(headers, ["OPENAI_MODEL"]);

  const stringLog = stringLogIdx !== -1 ? parseInt(sysSheet.getRange(2, stringLogIdx + 1).getValue(), 10) || 1 : 1;
  const queryTimeFrame = timeFrameIdx !== -1 ? parseInt(sysSheet.getRange(2, timeFrameIdx + 1).getValue(), 10) || 10 : 10;
  
  const apiKey = omniKeyIdx !== -1 ? String(sysSheet.getRange(2, omniKeyIdx + 1).getValue()).trim() : "";
  const omnirouteUrl = omniUrlIdx !== -1 ? String(sysSheet.getRange(2, omniUrlIdx + 1).getValue()).trim() : "https://omniroute.citystyleshop.online/v1/chat/completions";
  const modelName = modelIdx !== -1 ? String(sysSheet.getRange(2, modelIdx + 1).getValue()).trim() : "openai/gpt-4.1-mini";

  return {
    stringLog: stringLog,
    stringLogCell: `A2`,
    queryTimeFrame: queryTimeFrame,
    apiKey: apiKey,
    omnirouteUrl: omnirouteUrl,
    modelName: modelName
  };
}

/**
 * Преобразование буквенных колонок в числовые индексы (0-based)
 */
function lettersToIndices(letters) {
  if (!letters) return [];
  return letters.split(",").map(letStr => {
    const l = letStr.trim().toUpperCase();
    if (!l) return -1;
    let idx = 0;
    for (let i = 0; i < l.length; i++) {
      idx = idx * 26 + (l.charCodeAt(i) - 65 + 1);
    }
    return idx - 1;
  });
}

/**
 * Поиск индекса колонки по списку возможных заголовков
 */
function findColumnIndex(headers, names) {
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const idx = headers.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Установка триггера на перезапуск через 5 минут
 */
function setTimeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === "main")) {
    ScriptApp.newTrigger("main")
      .timeBased()
      .everyMinutes(5)
      .create();
  }
}

/**
 * Удаление триггера перезапуска
 */
function deleteTimeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "main") {
      ScriptApp.deleteTrigger(t);
    }
  });
}
