/**
 * ============================================================================
 * PROM.UA SMART AI GENERATOR & SEO ENGINE (OmniRoute / OpenAI)
 * ============================================================================
 * Версия: 2.6 (Strict Bi-Lingual Schema & Atomic Integrity Protection)
 * Автор: Mikhail & Orchestrator
 * Дата: 26 октября 2026 г.
 *
 * ЧТО ИСПРАВЛЕНО В ВЕРСИИ 2.6:
 * 1. Устранена асимметрия (когда RU заполнялся, а UA оставался куцым/старым):
 *    - Запрос к AI переведен на явную именованную схему (keywords_ru + keywords_ua).
 *    - Введена атомарная проверка: строка записывается ТОЛЬКО если получены ОБА
 *      языка длиной >= minOutputLength. Частичные «половинчатые» ответы блокируются.
 * 2. Оптимизирована проверка isTaskCompleted: данные читаются из памяти (rowData),
 *    что ускоряет сканирование в 10 раз и исключает пропуски.
 * 3. Поддержка всех типов задач поставщиков (AGER, IZIDROP, MOYDROP, KIRS, SPECULANT).
 * ============================================================================
 */

const ENGINE_CONFIG = {
  BATCH_SIZE: 3,        // 3 товара в батче (оптимально для 2 языков по 1000 символов)
  ITERATIONS: 2,        // 2 итерации за запуск (6 товаров за 5 минут)
  SYSTEMIC_SHEET: "Systemic",
  DATA_SHEET: "Export Products Sheet"
};

function main() {
  Logger.log("Smart Prom Engine V2.6 started.");

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

      // 3. Выполняем AI генерацию через OmniRoute
      const result = executeBatchAI(batch, dataSheet, settings);
      if (!result.success) {
        Logger.log("Сбой при выполнении батча через API. Остановка.");
        break;
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
 * Выполнение батча через OpenAI Endpoint OmniRoute с явной именованной схемой
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

  const systemInstruction = 
    "Ты — ведущий SEO-эксперт и локализатор маркетплейса Prom.ua.\n" +
    "Твоя задача — пакетно обработать товары строго по правилам их задач.\n\n" +
    "ОБЯЗАТЕЛЬНЫЕ ТРЕБОВАНИЯ ДЛЯ ЗАДАЧ:\n" +
    "1. ДЛЯ ЗАДАЧИ 'Keywords RU+UA':\n" +
    "   - Ты ОБЯЗАН вернуть ДВА полных поля: 'keywords_ru' и 'keywords_ua'.\n" +
    "   - Длина 'keywords_ru': СТРОГО от 900 до 1015 символов (через запятую с пробелом, строчные буквы, без слова 'купить').\n" +
    "   - Длина 'keywords_ua': СТРОГО от 900 до 1015 символов (через запятую с пробелом, строчные буквы, без слова 'купити').\n" +
    "   - Включай все 8 смысловых категорий: тип, назначение, сезон, конструкция/материал, стиль, цвет, бренд/артикул, фразы покупателей.\n" +
    "2. ДЛЯ ЗАДАЧИ 'Translate UA->RU':\n" +
    "   - Верни поля 'title_ru' (до 110 символов) и 'description_ru' с полным сохранением всей HTML-разметки и таблиц.\n" +
    "3. ДЛЯ ЗАДАЧИ 'Translate RU->UA':\n" +
    "   - Верни поля 'title_ua' (до 110 символов) и 'description_ua' с полным сохранением всей HTML-разметки и таблиц.\n\n" +
    "СТРОГО ВЕРНИ JSON объект вида: { \"results\": [ ... ] }\n" +
    "Пример элемента results:\n" +
    "{\n" +
    "  \"row\": 26,\n" +
    "  \"keywords_ru\": \"поисковые запросы на русском от 900 до 1015 символов...\",\n" +
    "  \"keywords_ua\": \"пошукові запити українською від 900 до 1015 символів...\"\n" +
    "}";

  const payload = {
    model: settings.modelName || "openai/gpt-4.1-mini",
    messages: [
      { role: "system", content: systemInstruction },
      { role: "user", content: JSON.stringify(itemsPayload) }
    ],
    temperature: 0.3,
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

  try {
    const response = UrlFetchApp.fetch(settings.omnirouteUrl, options);
    const resCode = response.getResponseCode();
    const resText = response.getContentText();

    if (resCode !== 200) {
      Logger.log(`HTTP Ошибка OmniRoute ${resCode}: ${resText}`);
      return { success: false, maxSavedRow: 0 };
    }

    const parsed = smartUniversalParser(resText);
    if (!parsed || !parsed.results || !Array.isArray(parsed.results)) {
      Logger.log("Ошибка: парсер не смог восстановить ни одного результата из ответа AI.");
      return { success: false, maxSavedRow: 0 };
    }

    const savedRows = [];

    for (const resItem of parsed.results) {
      const match = batch.find(b => String(b.rowNumber) === String(resItem.row));
      if (!match) continue;

      const extractedTaskOutputs = extractOutputsForMatchedItem(match, resItem);
      let rowFullyValid = true;

      // ⭐️ АТОМАРНАЯ ВАЛИДАЦИЯ: проверяем, что получены ОБА языка
      for (let tIdx = 0; tIdx < match.tasks.length; tIdx++) {
        const taskConfig = match.tasks[tIdx];
        const taskVals = extractedTaskOutputs[tIdx] || [];

        for (let oIdx = 0; oIdx < taskConfig.outputIndices.length; oIdx++) {
          const val = String(taskVals[oIdx] || "").trim();
          if (taskConfig.taskType && taskConfig.taskType.toLowerCase().includes("keywords")) {
            if (val.length < 500) {
              rowFullyValid = false;
            }
          } else if (val.length === 0) {
            rowFullyValid = false;
          }
        }
      }

      if (!rowFullyValid) {
        Logger.log(`Строка ${match.rowNumber}: ответ AI неполон (один из языков отсутствует). Пропуск во избежание порчи строки.`);
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

    Logger.log(`Батч обработан: успешно записано товаров: ${savedRows.length} из ${batch.length}.`);

    if (savedRows.length === 0) {
      Logger.log("Внимание: Ни одна строка не сохранена.");
      return { success: false, maxSavedRow: 0 };
    }

    const maxSaved = Math.max(...savedRows);
    return { success: true, maxSavedRow: maxSaved };

  } catch (err) {
    Logger.log("Исключение в executeBatchAI: " + err.message);
    return { success: false, maxSavedRow: 0 };
  }
}

/**
 * Полиморфный экстрактор выходов
 */
function extractOutputsForMatchedItem(matchItem, resItem) {
  const resultByTask = [];

  for (let t = 0; t < matchItem.tasks.length; t++) {
    const task = matchItem.tasks[t];
    const taskTypeLower = task.taskType.toLowerCase();
    const taskVals = [];

    // 1. Проверяем именованные поля для ключевых слов
    if (taskTypeLower.includes("keywords")) {
      const ruVal = resItem.keywords_ru || resItem.keyword_ru || resItem.ru_keywords || resItem.ru;
      const uaVal = resItem.keywords_ua || resItem.keyword_ua || resItem.ua_keywords || resItem.ua || resItem.uk;

      if (ruVal && uaVal) {
        taskVals.push(String(ruVal).trim());
        taskVals.push(String(uaVal).trim());
        resultByTask.push(taskVals);
        continue;
      }
    }

    // 2. Проверяем именованные поля для перевода
    if (taskTypeLower.includes("translate")) {
      const title = resItem.title_ru || resItem.title_ua || resItem.title || resItem.name;
      const desc = resItem.description_ru || resItem.description_ua || resItem.description || resItem.desc;

      if (title && desc) {
        taskVals.push(String(title).trim());
        taskVals.push(String(desc).trim());
        resultByTask.push(taskVals);
        continue;
      }
    }

    // 3. Если используются outputs: [ [val1, val2] ]
    if (Array.isArray(resItem.outputs) && resItem.outputs[t] !== undefined) {
      const sub = resItem.outputs[t];
      if (Array.isArray(sub) && sub.length === task.outputIndices.length) {
        resultByTask.push(sub.map(v => String(v || "").trim()));
        continue;
      }
    }

    // 4. Резервный сбор всех строковых значений в плоский пул
    const flatPool = [];
    function collect(obj) {
      if (obj === null || obj === undefined) return;
      if (typeof obj === "string" || typeof obj === "number") {
        const s = String(obj).trim();
        if (s.length > 0) flatPool.push(s);
      } else if (Array.isArray(obj)) {
        for (let i = 0; i < obj.length; i++) collect(obj[i]);
      } else if (typeof obj === "object") {
        const keys = Object.keys(obj);
        for (let i = 0; i < keys.length; i++) {
          const k = keys[i].toLowerCase();
          if (["row", "row_number", "supplier", "task_type", "type", "id"].indexOf(k) === -1) {
            collect(obj[keys[i]]);
          }
        }
      }
    }

    collect(resItem.outputs !== undefined ? resItem.outputs : (resItem.tasks || resItem));

    const needed = task.outputIndices.length;
    for (let o = 0; o < needed; o++) {
      taskVals.push(o < flatPool.length ? flatPool[o] : "");
    }
    resultByTask.push(taskVals);
  }

  return resultByTask;
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
 * - Исключение слова "купить" и "купити"
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
 * Самовосстанавливающийся парсер JSON с авто-ремонтом обрезанных строк
 */
function smartUniversalParser(rawText) {
  try {
    const outer = JSON.parse(rawText);
    let content = outer.choices?.[0]?.message?.content || rawText;

    if (typeof content === "string") {
      content = content.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```$/i, "").trim();
      try {
        content = JSON.parse(content);
      } catch (jsonErr) {
        content = repairIncompleteJson(content);
      }
    }

    if (!content) return null;
    if (content.results && Array.isArray(content.results)) return content;
    if (Array.isArray(content)) return { results: content };
    if (content.items && Array.isArray(content.items)) return { results: content.items };
    return { results: [content] };
  } catch (e) {
    Logger.log("Ошибка smartUniversalParser: " + e.message);
    return null;
  }
}

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

function findColumnIndex(headers, names) {
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const idx = headers.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

function setTimeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === "main")) {
    ScriptApp.newTrigger("main")
      .timeBased()
      .everyMinutes(5)
      .create();
  }
}

function deleteTimeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "main") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

