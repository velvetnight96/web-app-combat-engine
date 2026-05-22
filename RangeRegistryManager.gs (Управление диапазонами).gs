/**
 * =========================================================================
 * МОДУЛЬ ДВУСТОРОННЕЙ СИНХРОНИЗАЦИИ ИМЕНОВАННЫХ ДИАПАЗОНОВ (COMBAT ENGINE)
 * =========================================================================
 */

const REGISTRY_SHEET_NAME = "_Ranges_DB";

/**
 * ФУНКЦИЯ 1: ПРЯМАЯ СИНХРОНИЗАЦИЯ (Ввод "1" в "диапазон_ген" или кнопка)
 * Из таблицы реестра -> В физические диапазоны Google Таблицы.
 */
function generateAndSyncNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  
  if (!dbSheet) {
    throw new Error(`Лист "${REGISTRY_SHEET_NAME}" не найден!`);
  }
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(`Реестр диапазонов на листе "${REGISTRY_SHEET_NAME}" пуст!`);
  }
  
  // Читаем первые 4 столбца: А (Имя), B (Тип), C (Лист), D (Координаты)
  const registryData = dbSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  // Собираем чистые имена из первого столбца (индекс 0)
  const managedNames = registryData.map(row => row[0] ? row[0].toString().trim() : "").filter(String);
  
  // --- ЭТАП 1: ОЧИСТКА СТАРЫХ ДУБЛЕЙ ---
  const currentNamedRanges = ss.getNamedRanges();
  for (let i = 0; i < currentNamedRanges.length; i++) {
    const rangeName = currentNamedRanges[i].getName();
    if (managedNames.indexOf(rangeName) !== -1) {
      currentNamedRanges[i].remove();
    }
  }
  
  // --- ЭТАП 2: ГЕНЕРАЦИЯ ДИАПАЗОНОВ ---
  let createdCount = 0;
  let errors = [];
  
  for (let j = 0; j < registryData.length; j++) {
    // ОБНОВЛЕННЫЕ ИНДЕКСЫ С УЧЕТОМ СТОЛБЦА ТИП (0 - Имя, 2 - Лист, 3 - Координаты)
    const name = registryData[j][0] ? registryData[j][0].toString().trim() : "";
    const sheetName = registryData[j][2] ? registryData[j][2].toString().trim() : "";
    const a1Notation = registryData[j][3] ? registryData[j][3].toString().trim() : "";
    
    if (!name || !sheetName || !a1Notation) continue; 
    
    // Защита системных управляющих ячеек
    if (name === "диапазон_ген" || name === "диапазон_синх" || name === "сообщение" || name === "алфавит_дб" || name === "тип_дб") continue;
    
    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      errors.push(`Лист "${sheetName}" не существует (для диапазона "${name}")`);
      continue;
    }
    
    try {
      const targetRange = targetSheet.getRange(a1Notation);
      ss.setNamedRange(name, targetRange);
      createdCount++;
    } catch (err) {
      errors.push(`Ошибка адреса "${a1Notation}" для "${name}": ${err.message}`);
    }
  }
  
  SpreadsheetApp.flush();
  
  if (errors.length > 0) {
    throw new Error("Ошибки генерации:\n" + errors.join("\n"));
  }
}

/**
 * ФУНКЦИЯ 2: ОБРАТНОЕ СКАНИРОВАНИЕ (Ввод "1" в "диапазон_синх" или кнопка)
 * Из физических диапазонов Таблицы -> Обратно в текст реестра.
 */
function reverseScanAndReadNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  
  if (!dbSheet) {
    throw new Error(`Лист "${REGISTRY_SHEET_NAME}" не найден!`);
  }
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(`Реестр пуст. Нечего сканировать!`);
  }
  
  const namesList = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => n.toString().trim());
  let updatedRowsData = [];
  
  for (let i = 0; i < namesList.length; i++) {
    const currentName = namesList[i];
    if (!currentName) {
      updatedRowsData.push(["", ""]);
      continue;
    }
    
    const systemRange = ss.getRangeByName(currentName);
    if (systemRange) {
      const actualSheetName = systemRange.getSheet().getName();
      const actualA1Notation = systemRange.getA1Notation();
      updatedRowsData.push([actualSheetName, actualA1Notation]);
    } else {
      updatedRowsData.push(["⚠️ УДАЛЕН НА ЛИСТАХ", "⚠️ ТРЕБУЕТСЯ СИНХРОНИЗАЦИЯ"]);
    }
  }
  
  // ОБНОВЛЕННЫЙ АДРЕС ЗАПИСИ: Теперь пишем пачкой строго в Столбцы C и D (3 и 4 столбцы)
  const targetWriteRange = dbSheet.getRange(2, 3, lastRow - 1, 2);
  targetWriteRange.setValues(updatedRowsData);
  
  SpreadsheetApp.flush();
}

/**
 * ФУНКЦИЯ 3: СОРТИРОВКА РЕЕСТРА ПО АЛФАВИТУ (Ввод "1" в "алфавит_дб" или кнопка)
 * Выделяет строго столбцы A-E и сортирует по Столбцу A.
 */
function sortRegistryByAlphabet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  if (!dbSheet) return;
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return;
  
  // ЖЕСТКОЕ ОГРАНИЧЕНИЕ: Вместо всех столбцов берем ровно 5 (столбцы A, B, C, D, E)
  const dataRange = dbSheet.getRange(2, 1, lastRow - 1, 5);
  
  dataRange.sort({column: 1, ascending: true}); // Сортировка по Столбцу A
  SpreadsheetApp.flush();
}

/**
 * ФУНКЦИЯ 4: МНОГОУРОВНЕВАЯ СОРТИРОВКА (Ввод "1" в "тип_дб" или кнопка)
 * Выделяет строго столбцы A-E. Группирует по Типу (Столбец B), затем по Алфавиту (Столбец A).
 */
function sortRegistryByTypeAndAlphabet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  
  if (!dbSheet) {
    throw new Error(`Лист "${REGISTRY_SHEET_NAME}" не найден!`);
  }
  
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return;
  
  // ЖЕСТКОЕ ОГРАНИЧЕНИЕ: Выделяем область шириной ровно в 5 столбцов (A-E)
  const dataRange = dbSheet.getRange(2, 1, lastRow - 1, 5);
  
  // Каскадная сортировка внутри защищенного диапазона A-E
  dataRange.sort([
    {column: 2, ascending: true}, 
    {column: 1, ascending: true}
  ]);
  
  SpreadsheetApp.flush();
}
