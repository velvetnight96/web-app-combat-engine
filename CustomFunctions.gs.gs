/**
 * Автоматически находит имя персонажа в карточке на листе БАЗА 
 * и забирает значение указанной ячейки с его личного листа.
 * 
 * @param {string} targetCell Адрес ячейки, которую нужно забрать с листа персонажа (например, "AK21").
 * @param {any} triggerCell Необязательный аргумент для принудительного обновления (указывайте диапазон "обновить").
 * @return {any} Значение запрашиваемой ячейки.
 * @customfunction
 */
function GET_CHARACTER_RESOURCE(targetCell, triggerCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (!targetCell) return "Укажите целевую ячейку";

  // Скрипт сам определяет ячейку, в которой написана эта формула
  const cell = ss.getActiveRange();
  const cellRow = cell.getRow();
  const cellCol = cell.getColumn();
  
  // Ищем, в какой именованный диапазон "Карта_..." входит эта ячейка
  const namedRanges = ss.getNamedRanges();
  let characterName = "";
  
  for (let i = 0; i < namedRanges.length; i++) {
    let range = namedRanges[i].getRange();
    let name = namedRanges[i].getName();
    
    if (name.startsWith("Карта_") && range.getSheet().getName() === sheet.getName() &&
        cellRow >= range.getRow() && cellRow <= range.getLastRow() &&
        cellCol >= range.getColumn() && cellCol <= range.getLastColumn()) {
      
      // Находим ячейку 3 строки и 3 столбца ВНУТРИ этого диапазона
      characterName = range.getCell(3, 3).getValue().toString().trim();
      break;
    }
  }
  
  if (characterName === "") return "Карта не найдена";
  
  try {
    const targetSheet = ss.getSheetByName(characterName);
    if (!targetSheet) return "Лист '" + characterName + "' не найден";
    
    return targetSheet.getRange(targetCell).getValue();
  } catch(e) {
    return "Ошибка чтения ячейки " + targetCell;
  }
}
