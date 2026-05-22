/**
 * Пересчет и чистка списков участников боя.
 * Убирает мертвых, обновляет диапазоны для "чей_ход" и "все_цели".
 */
function rebuildBattlefieldDropDowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Считываем "участники_боя" (где логика листа отфильтровала только живых)
  const participantsRange = ss.getRangeByName("участники_боя");
  const aliveUnits = participantsRange.getValues().flat().filter(String);
  
  const mainSheet = ss.getSheetByName("Поле_Боя"); // Имя вашего боевого листа
  if (!mainSheet) return;
  
  // Динамически перестраиваем валидацию данных (выпадающий список) для ячейки "чей_ход"
  const turnCell = ss.getRangeByName("чей_ход");
  if (aliveUnits.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(aliveUnits, true)
      .setAllowInvalid(false)
      .setHelpText("Выберите активного персонажа из списка живых.")
      .build();
    turnCell.setDataValidation(rule);
  } else {
    turnCell.clearDataValidation();
  }
}
