/**
 * ============================================================================
 * PROM.UA SMART ORCHESTRATOR & WATCHDOG (V2.1)
 * ============================================================================
 * Версия: 2.1 (Auto-Wakeup & Debounce for Smart Prom AI Engine)
 * Автор: Mikhail & Orchestrator
 * Дата: 15 октября 2026 г.
 *
 * НАЗНАЧЕНИЕ:
 * 1. Слушает изменения таблицы от внешних скриптов (GitHub Actions / Python).
 * 2. Выжидает 10 минут (debounce), чтобы импорт успел полностью записаться.
 * 3. Сканирует таблицу от предыдущего STRING LOG (или снизу вверх), проверяя
 *    все правила поставщиков с листа Systemic.
 * 4. Находит ПЕРВУЮ строку, требующую заполнения, ставит указатель STRING LOG
 *    и запускает main() для пакетной обработки.
 * 5. Если все строки заполнены — засыпает и переводит статус в IDLE.
 * ============================================================================
 */

/**
 * 1. ТРИГГЕР НА ИЗМЕНЕНИЕ (Устанавливается как триггер типа "При изменении" / "On change")
 */
function onChangeTrigger(e) {
  const props = PropertiesService.getScriptProperties();
  const scriptStatus = props.getProperty("SCRIPT_STATUS");

  // Если скрипт уже активно работает (RUNNING) или ожидает запуска (WAITING), выходим
  if (scriptStatus === "RUNNING" || scriptStatus === "WAITING") {
    return;
  }

  Logger.log("Зафиксировано внешнее обновление таблицы. Планируем проверку через 10 минут...");

  // Фиксируем статус ожидания
  props.setProperty("SCRIPT_STATUS", "WAITING");

  // Удаляем старые триггеры оркестратора, если они остались
  deleteSpecificTrigger("superMasterOrchestrator");

  // Создаем разовый отложенный запуск через 10 минут (600 000 мс)
  ScriptApp.newTrigger("superMasterOrchestrator")
    .timeBased()
    .after(10 * 60 * 1000)
    .create();
}

/**
 * 2. ГЛАВНАЯ ФУНКЦИЯ ОРКЕСТРАТОРА (СТАРТЕР)
 */
function superMasterOrchestrator() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("SCRIPT_STATUS", "RUNNING");

  Logger.log("Оркестратор проснулся. Начинаем поиск строк, требующих работы...");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const systemicSheet = ss.getSheetByName("Systemic");
    const dataSheet = ss.getSheetByName("Export Products Sheet");

    if (!systemicSheet || !dataSheet) {
      Logger.log("ОШИБКА: Не найден лист 'Systemic' или 'Export Products Sheet'.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    // Удаляем старые зависшие триггеры функции main перед сканированием
    deleteSpecificTrigger("main");

    const sysHeaders = systemicSheet.getRange(1, 1, 1, systemicSheet.getLastColumn()).getValues()[0];
    const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];

    // Загружаем правила конфигурации поставщиков
    const taskConfigs = loadTaskConfigurationsForMaster(systemicSheet, sysHeaders);
    const supplierColIdx = findColumnIndexInMaster(dataHeaders, ["Особисті_нотатки"]);
    const stringLogColIdx = findColumnIndexInMaster(sysHeaders, ["STRING LOG"]);

    if (supplierColIdx === -1 || stringLogColIdx === -1) {
      Logger.log("ОШИБКА: Не найден столбец 'Особисті_нотатки' или 'STRING LOG'.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("В таблице нет данных.");
      props.setProperty("SCRIPT_STATUS", "IDLE");
      return;
    }

    // Считываем текущий STRING LOG (чтобы не сканировать с самого начала таблицы)
    let currentLog = parseInt(systemicSheet.getRange(2, stringLogColIdx + 1).getValue(), 10) || 1;
    if (currentLog >= lastRow) {
      currentLog = 1; // Если лог уперся в конец, проверим таблицу с начала на случай пропущенных новинок
    }

    let firstUnprocessedRow = -1;

    // Сканируем строки начиная с currentLog
    for (let r = currentLog + 1; r <= lastRow; r++) {
      const rowValues = dataSheet.getRange(r, 1, 1, dataSheet.getLastColumn()).getValues()[0];
      const supplierValue = String(rowValues[supplierColIdx] || "").trim().toUpperCase();

      if (!supplierValue) continue;

      const matchingTasks = taskConfigs.filter(cfg => cfg.supplierName === supplierValue);
      if (matchingTasks.length === 0) continue;

      // Проверяем задачи для этого поставщика
      let rowNeedsWork = false;
      for (const task of matchingTasks) {
        if (!isTaskDoneForMaster(rowValues, task.outputIndices, task.minOutputLength, dataSheet, r)) {
          rowNeedsWork = true;
          break;
        }
      }

      if (rowNeedsWork) {
        firstUnprocessedRow = r;
        break;
      }
    }

    if (firstUnprocessedRow !== -1) {
      // Ставим указатель на строку ПЕРЕД первой необработанной
      systemicSheet.getRange(2, stringLogColIdx + 1).setValue(firstUnprocessedRow - 1);
      Logger.log(`Найдена строка, требующая обработки: ${firstUnprocessedRow}. Запускаем умный main()...`);
      
      // Запускаем рабочий движок
      main();
    } else {
      Logger.log("Все строки полностью заполнены. Работы нет, Оркестратор засыпает.");
      systemicSheet.getRange(2, stringLogColIdx + 1).setValue(lastRow);
      props.setProperty("SCRIPT_STATUS", "IDLE");
    }

  } catch (err) {
    Logger.log("Исключение в superMasterOrchestrator: " + err.message);
    props.setProperty("SCRIPT_STATUS", "IDLE");
  } finally {
    deleteSpecificTrigger("superMasterOrchestrator");
  }
}

/**
 * Вспомогательные функции Оркестратора
 */
function deleteSpecificTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function isTaskDoneForMaster(rowValues, outputIndices, minLength, dataSheet, rowNum) {
  for (const idx of outputIndices) {
    if (idx === -1) continue;
    const val = String(rowValues[idx] || "").trim();
    if (val.length < (minLength || 1)) {
      return false;
    }
  }
  return true;
}

function loadTaskConfigurationsForMaster(sysSheet, headers) {
  const configs = [];
  const lastRow = sysSheet.getLastRow();
  if (lastRow < 2) return configs;

  const colSupplier = findColumnIndexInMaster(headers, ["Supplier Name"]);
  const colTaskType = findColumnIndexInMaster(headers, ["Task Type"]);
  const colOutputs = findColumnIndexInMaster(headers, ["Output Columns (Letters)"]);
  const colMinLen = findColumnIndexInMaster(headers, ["Min Output Length"]);

  const numCols = sysSheet.getLastColumn();
  const data = sysSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const supplier = String(row[colSupplier] || "").trim().toUpperCase();
    if (!supplier) continue;

    configs.push({
      supplierName: supplier,
      taskType: String(row[colTaskType] || "").trim(),
      outputIndices: lettersToIndicesInMaster(String(row[colOutputs] || "")),
      minOutputLength: parseInt(row[colMinLen], 10) || 0
    });
  }

  return configs;
}

function lettersToIndicesInMaster(letters) {
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

function findColumnIndexInMaster(headers, names) {
  for (const name of names) {
    const idx = headers.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

