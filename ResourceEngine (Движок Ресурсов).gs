/**
 * Списание стоимости хода из накопителя ост.д
 */
function deductTurnActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costRange = ss.getRangeByName("ед.д");
  const remainingRange = ss.getRangeByName("ост.д");
  
  const cost = costRange.getValue();
  const remaining = remainingRange.getValue();
  
  const nextActions = remaining - cost;
  remainingRange.setValue(nextActions < 0 ? 0 : nextActions);
}

/**
 * Расчет и применение износа оружия к карте персонажа
 */
function applyWeaponWear(playerName, targetCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Находим именную карту персонажа (лист с именем игрока)
  const charSheet = ss.getSheetByName(playerName);
  if (!charSheet) return 0;
  
  // Допустим, тип атаки проверяется формулой на листе, но нам нужно списать прочность:
  // Координаты прочности оружия на карте (пример: строка 21, колонка 10)
  const durabilityRange = charSheet.getRange(21, 10);
  const currentDurability = durabilityRange.getValue();
  
  // Износ равен количеству целей (поцелевое списание)
  const wear = targetCount;
  const newDurability = currentDurability - wear;
  
  durabilityRange.setValue(newDurability < 0 ? 0 : newDurability);
  return wear;
}

/**
 * Обнуление остатков ОД персонажа (вызывается при Смене Хода)
 */
function clearRemainingActionsPool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const remainingRange = ss.getRangeByName("ост.д");
  remainingRange.setValue(0);
}
