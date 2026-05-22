/**
 * ЕДИНЫЙ ПАКЕТНЫЙ ЗАПРОС ДЛЯ HTML-ПАНЕЛИ (Smart Polling API)
 * Забирает абсолютно все данные за один вызов.
 */
function getGameState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const currentTurn = ss.getRangeByName("чей_ход").getValue();
    const remainingActions = ss.getRangeByName("ост.д").getValue();
    const actionNotification = ss.getRangeByName("ед.д_оповещение").getValue();
    
    // Безопасное чтение последней строки лога
    let lastLog = "Бой начался.";
    const logRange = ss.getRangeByName("журнал_боя");
    const logValues = logRange.getValues().flat().filter(String);
    if (logValues.length > 1) {
      lastLog = logValues[0]; // Предполагаем, что новые записи пишем наверх
    }

    return {
      success: true,
      currentTurn: currentTurn,
      remainingActions: remainingActions,
      statusMessage: actionNotification,
      lastLogEntry: lastLog,
      timestamp: new Date().getTime()
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Логика проведения атаки
 */
function executeAttackCommand() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Проверка лимита действий перед атакой
  const status = ss.getRangeByName("ед.д_оповещение").getValue();
  if (status !== "можно ходить") {
    throw new Error("У вас не хватает действий в этом ходе!");
  }
  
  const currentAttacker = ss.getRangeByName("чей_ход").getValue();
  const targetsRange = ss.getRangeByName("все_цели");
  const targets = targetsRange.getValues().flat().filter(String);
  
  if (targets.length === 0) {
    throw new Error("Не выбраны цели для атаки!");
  }
  
  // Рассчитываем износ оружия на основе количества целей
  const targetCount = targets.length;
  const wearAmount = applyWeaponWear(currentAttacker, targetCount);
  
  // Формируем батл-лог
  const descriptionText = `[${currentAttacker}] 💥 Атака оружием по целям: [${targets.join(", ")}]. Износ: -${wearAmount} ед.`;
  ss.getRangeByName("описание").setValue(descriptionText);
  
  // Запись в журнал боя
  writeToBattleLog(descriptionText, targets);
  
  // Списание стоимости ОД (происходит через вычитание в ResourceEngine)
  deductTurnActions();
}

/**
 * Запись строк поцелево в журнал_боя
 */
function writeToBattleLog(mainText, targets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logRange = ss.getRangeByName("журнал_боя");
  const currentLogs = logRange.getValues();
  
  // Защищенная строка заголовков — это индекс 0. Сдвигаем данные вниз.
  // Для простоты примера: вставляем новую строку в массив логов сразу после заголовка.
  let newLogs = [];
  newLogs.push(currentLogs[0]); // Сохраняем заголовки
  
  // Первая строка хода получает базовую инфу
  newLogs.push([mainText, targets[0] || ""]);
  
  // Остальные цели выстраиваются строго в столбец G (индекс 1 в двумерном массиве лога целей)
  for (let i = 1; i < targets.length; i++) {
    newLogs.push(["", targets[i]]);
  }
  
  // Дописываем старые логи, если есть место
  for (let j = 1; j < currentLogs.length; j++) {
    if (newLogs.length < currentLogs.length) {
      newLogs.push(currentLogs[j]);
    }
  }
  
  logRange.setValues(newLogs);
}
