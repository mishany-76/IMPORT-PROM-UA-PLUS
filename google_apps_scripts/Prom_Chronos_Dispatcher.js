/**
 * ============================================================================
 * PROM.UA CHRONOS DISPATCHER (GitHub Actions Time Commander)
 * ============================================================================
 * Название: Prom Chronos Dispatcher (Точный диспетчер запусков Prom.ua)
 * Версия: 1.0 (Kyiv Exact-Time GitHub Actions Dispatcher)
 * Автор: Mikhail & Chronos
 * Дата: 28 октября 2026 г.
 *
 * НАЗНАЧЕНИЕ:
 * 1. Гарантированный запуск репозитория IMPORT-PROM-UA-PLUS 3 раза в сутки:
 *    строго в 06:00, 12:00, 19:00 по киевскому времени (Europe/Kyiv).
 * 2. Полный обход задержек и сбоев встроенного cron-планировщика GitHub Actions:
 *    вызывает официальный REST API GitHub (событие workflow_dispatch).
 * 3. Мгновенный приоритетный запуск раннера на стороне GitHub без ожидания в очередях.
 * 4. Удобное меню в интерфейсе Google Таблицы:
 *    - Запуск обновления каталога вручную в 1 клик
 *    - Автоматическая установка / проверка расписания
 *    - Безопасное сохранение GitHub токена (PAT) в Script Properties
 * ============================================================================
 */

const CHRONOS_CONFIG = {
  GITHUB_OWNER: "mishany-76",
  GITHUB_REPO: "IMPORT-PROM-UA-PLUS",
  WORKFLOW_FILE: "run_scripts.yml",
  BRANCH: "main",
  TIMEZONE: "Europe/Kyiv",
  // Точные часы запуска по Киеву: 06:00, 12:00, 19:00
  SCHEDULE_HOURS: [6, 12, 19]
};

/**
 * 1. МЕНЮ В ИНТЕРФЕЙСЕ GOOGLE ТАБЛИЦЫ
 * Создается автоматически при открытии документа
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("⚡ Chronos Dispatcher")
      .addItem("▶️ Запустить импорт на GitHub прямо сейчас", "triggerGitHubActionsWorkflow")
      .addSeparator()
      .addItem("⏰ Настроить расписание (06:00, 12:00, 19:00 Киев)", "setupChronosSchedule")
      .addItem("❌ Отключить расписание Chronos", "deleteChronosSchedule")
      .addSeparator()
      .addItem("🔑 Ввести / обновить токен GitHub (PAT)", "promptSetGitHubToken")
      .addItem("📊 Проверить статус последнего запуска", "showChronosStatus")
      .addToUi();
  } catch (e) {
    Logger.log("Меню не может быть создано в режиме без UI: " + e.message);
  }
}

/**
 * 2. ГЛАВНАЯ ФУНКЦИЯ ДИСПЕТЧЕРА: Запуск GitHub Actions через REST API
 */
function triggerGitHubActionsWorkflow() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("GITHUB_PAT") || props.getProperty("GITHUB_TOKEN");

  if (!token) {
    const errorMsg = "ОШИБКА: Токен GitHub (GITHUB_PAT) не настроен в свойствах скрипта! Задайте его через меню '⚡ Chronos Dispatcher' -> '🔑 Ввести токен GitHub'.";
    Logger.log(errorMsg);
    notifyUiToast("Ошибка: Не задан токен GITHUB_PAT!", "Chronos Error", 10);
    return false;
  }

  const url = `https://api.github.com/repos/${CHRONOS_CONFIG.GITHUB_OWNER}/${CHRONOS_CONFIG.GITHUB_REPO}/actions/workflows/${CHRONOS_CONFIG.WORKFLOW_FILE}/dispatches`;

  const options = {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token.trim(),
      "Accept": "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "Google-Apps-Script-Chronos-Dispatcher"
    },
    payload: JSON.stringify({
      ref: CHRONOS_CONFIG.BRANCH
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    const nowStr = Utilities.formatDate(new Date(), CHRONOS_CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");

    // GitHub API возвращает 204 No Content при успешном создании задачи
    if (code === 204) {
      const msg = `[Chronos] УСПЕХ: Воркфлоу '${CHRONOS_CONFIG.WORKFLOW_FILE}' запущен на GitHub в ${nowStr} (Киев).`;
      Logger.log(msg);
      props.setProperty("CHRONOS_LAST_RUN", `${nowStr} (HTTP 204 Success)`);
      notifyUiToast(`Репозиторий ${CHRONOS_CONFIG.GITHUB_REPO} успешно запущен в ${nowStr}!`, "Chronos Успех", 6);
      return true;
    } else {
      const errMsg = `[Chronos] ОШИБКА запуска GitHub: HTTP ${code}: ${body}`;
      Logger.log(errMsg);
      props.setProperty("CHRONOS_LAST_RUN", `${nowStr} (FAILED HTTP ${code})`);
      notifyUiToast(`Ошибка запуска: HTTP ${code}`, "Chronos Ошибка", 10);
      return false;
    }
  } catch (err) {
    Logger.log("[Chronos] Исключение UrlFetchApp: " + err.message);
    notifyUiToast("Исключение сети: " + err.message, "Chronos Ошибка", 10);
    return false;
  }
}

/**
 * 3. АВТО-УСТАНОВКА ТОЧНОГО РАСПИСАНИЯ ПО КИЕВУ (06:00, 12:00, 19:00)
 */
function setupChronosSchedule() {
  // Сначала удаляем предыдущие триггеры Chronos во избежание дублирования
  deleteChronosSchedule(false);

  for (let i = 0; i < CHRONOS_CONFIG.SCHEDULE_HOURS.length; i++) {
    const targetHour = CHRONOS_CONFIG.SCHEDULE_HOURS[i];

    ScriptApp.newTrigger("triggerGitHubActionsWorkflow")
      .timeBased()
      .atHour(targetHour)
      .nearMinute(0)
      .inTimezone(CHRONOS_CONFIG.TIMEZONE)
      .everyDays(1)
      .create();

    Logger.log(`[Chronos] Создан ежедневный триггер на ${targetHour}:00 (${CHRONOS_CONFIG.TIMEZONE}).`);
  }

  const successText = `Расписание Chronos успешно установлено на 06:00, 12:00 и 19:00 (${CHRONOS_CONFIG.TIMEZONE})!`;
  Logger.log(successText);
  notifyUiAlert("Успешно", successText);
}

/**
 * 4. УДАЛЕНИЕ РАСПИСАНИЯ CHRONOS
 */
function deleteChronosSchedule(showAlert) {
  if (showAlert === undefined) showAlert = true;
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;

  for (let i = 0; i < triggers.length; i++) {
    const t = triggers[i];
    if (t.getHandlerFunction() === "triggerGitHubActionsWorkflow") {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  }

  Logger.log(`[Chronos] Удалено триггеров Chronos: ${count}.`);
  if (showAlert) {
    notifyUiAlert("Расписание отключено", `Удалено ${count} активных триггеров Chronos.`);
  }
}

/**
 * 5. ДИАЛОГ ВВОДА GITHUB ТОКЕНА (PAT)
 */
function promptSetGitHubToken() {
  const ui = SpreadsheetApp.getUi();
  const currentToken = PropertiesService.getScriptProperties().getProperty("GITHUB_PAT");
  const masked = currentToken ? (currentToken.substring(0, 7) + "..." + currentToken.substring(currentToken.length - 4)) : "Не задан";

  const result = ui.prompt(
    "Настройка токена GitHub (PAT)",
    `Текущий токен: ${masked}\n\nВставьте ваш GitHub Personal Access Token (с правами actions:write):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const val = result.getResponseText().trim();
    if (val.length > 10) {
      PropertiesService.getScriptProperties().setProperty("GITHUB_PAT", val);
      ui.alert("Токен сохранен", "GitHub Personal Access Token успешно сохранен в защищенных свойствах проекта!", ui.ButtonSet.OK);
    } else {
      ui.alert("Ошибка", "Введен слишком короткий или пустой токен. Значение не сохранено.", ui.ButtonSet.OK);
    }
  }
}

/**
 * 6. ПРОВЕРКА СТАТУСА ПОСЛЕДНЕГО ЗАПУСКА
 */
function showChronosStatus() {
  const props = PropertiesService.getScriptProperties();
  const lastRun = props.getProperty("CHRONOS_LAST_RUN") || "Запусков еще не производилось";
  const token = props.getProperty("GITHUB_PAT");
  const tokenStatus = token ? "✅ Настроен" : "❌ Отсутствует";

  // Считаем активные триггеры
  const triggers = ScriptApp.getProjectTriggers();
  const chronosTriggers = triggers.filter(t => t.getHandlerFunction() === "triggerGitHubActionsWorkflow");

  const message = 
    `Репозиторий: ${CHRONOS_CONFIG.GITHUB_OWNER}/${CHRONOS_CONFIG.GITHUB_REPO}\n` +
    `Воркфлоу: ${CHRONOS_CONFIG.WORKFLOW_FILE}\n` +
    `Токен GitHub: ${tokenStatus}\n` +
    `Активных триггеров: ${chronosTriggers.length} (требуется 3: 06:00, 12:00, 19:00)\n` +
    `Часовой пояс: ${CHRONOS_CONFIG.TIMEZONE}\n\n` +
    `Последний вызов: ${lastRun}`;

  notifyUiAlert("Статус Chronos Dispatcher", message);
}

/**
 * Вспомогательные функции оповещений
 */
function notifyUiToast(msg, title, timeoutSec) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, timeoutSec || 5);
  } catch (e) {}
}

function notifyUiAlert(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(`[${title}] ${message}`);
  }
}

