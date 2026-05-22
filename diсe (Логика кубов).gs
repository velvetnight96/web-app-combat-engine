function mainFunction() {
  animateDice(); 
  checkValue(); 
}

// 1. АНИМАЦИЯ (ТЕПЕРЬ ДЛЯ ОДНОГО ИЛИ ДВУХ КУБИКОВ)
function animateDice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cellOutput1 = ss.getRangeByName("Result_dice");
  const cellOutput2 = ss.getRangeByName("Result_dice_2");
  
  // Проверяем статус преимущества/помехи
  const advRange = ss.getRangeByName("преимущество");
  const advStatus = advRange ? advRange.getValue().toString().trim() : "Нет";
  const isDoubleRoll = (advStatus === "Преимущество" || advStatus === "Помеха");

  const valRange = ss.getRangeByName("Tipe_dice");
  const val = valRange ? valRange.getValue().toString().toUpperCase() : "D20";
  
  if (!cellOutput1) return;

  // Сброс стилей перед броском для первого кубика
  cellOutput1.setFontSize(10).setFontWeight("normal").setBackground(null).setFontColor("#000000");
  
  // Сброс стилей и очистка для второго кубика (если он есть)
  if (cellOutput2) {
    cellOutput2.setFontSize(10).setFontWeight("normal").setBackground(null).setFontColor("#000000");
    if (!isDoubleRoll) cellOutput2.clearContent(); // Если бросок одиночный, стираем старый второй кубик
  }

  let maxSides = 20;
  const matches = val.match(/\d+/);
  if (matches) maxSides = parseInt(matches, 10);

  const frames = 8; 
  for (let i = 0; i < frames; i++) {
    let tempNum1 = Math.floor(Math.random() * maxSides) + 1;
    cellOutput1.setValue(tempNum1);
    
    if (isDoubleRoll && cellOutput2) {
      let tempNum2 = Math.floor(Math.random() * maxSides) + 1;
      cellOutput2.setValue(tempNum2);
    }
    
    SpreadsheetApp.flush(); 
    Utilities.sleep(100); 
  }
}

// 2. РАСЧЕТ И ФИНАЛЬНОЕ ОФОРМЛЕНИЕ (ОБНОВЛЕНО ПОД ДВА КУБИКА)
function checkValue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const valRange = ss.getRangeByName("Tipe_dice");
  const val = valRange ? valRange.getValue().toString().toUpperCase() : "D20";
  
  const cellOutput1 = ss.getRangeByName("Result_dice");
  const cellOutput2 = ss.getRangeByName("Result_dice_2");
  const statusCell = ss.getRangeByName("Info_dice");
  const messDiceCell = ss.getRangeByName("mess_dice");

  // Читаем статус преимущества
  const advRange = ss.getRangeByName("преимущество");
  const advStatus = advRange ? advRange.getValue().toString().trim() : "Нет";
  const isDoubleRoll = (advStatus === "Преимущество" || advStatus === "Помеха");

  if (!cellOutput1) return;

  let maxSides = 20; 
  const matches = val.match(/\d+/);
  if (matches) maxSides = parseInt(matches, 10);
  
  // Бросаем финальные значения
  const roll1 = Math.floor(Math.random() * maxSides) + 1;
  cellOutput1.setValue(roll1);
  cellOutput1.setFontSize(18).setFontWeight("bold");
  styleSingleDice(cellOutput1, roll1, maxSides); // Стилизуем первый кубик

  let roll2 = 0;
  if (isDoubleRoll && cellOutput2) {
    roll2 = Math.floor(Math.random() * maxSides) + 1;
    cellOutput2.setValue(roll2);
    cellOutput2.setFontSize(18).setFontWeight("bold");
    styleSingleDice(cellOutput2, roll2, maxSides); // Стилизуем второй кубик
  }

  // Обновляем статус инфо
  if (statusCell) {
    statusCell.setValue("брошен " + val + (isDoubleRoll ? " (" + advStatus + ")" : ""));
  }

  // Формируем финальное сообщение в mess_dice
  if (messDiceCell) {
    let message = "";
    if (!isDoubleRoll) {
      message = getDiceMessage(roll1, maxSides);
    } else {
      // Логика подсказки для преимущества / помехи
      if (advStatus === "Преимущество") {
        const bestRoll = Math.max(roll1, roll2);
        message = "Преимущество! Выбирай: " + bestRoll + " (" + getDiceMessage(bestRoll, maxSides) + ")";
      } else if (advStatus === "Помеха") {
        const worstRoll = Math.min(roll1, roll2);
        message = "Помеха... Твой результат: " + worstRoll + " (" + getDiceMessage(worstRoll, maxSides) + ")";
      }
    }
    messDiceCell.setValue(message);
  }
}

// Вспомогательная функция для окрашивания отдельного кубика
function styleSingleDice(cell, result, maxSides) {
  if (result === 1) {
    cell.setBackground("#ff4d4d").setFontColor("#ffffff");
  } else if (result === maxSides) {
    cell.setBackground("#ffd700").setFontColor("#000000");
  } else if (maxSides === 100) {
    if (result < 40) cell.setBackground("#fff2cc").setFontColor("#000000");
    else cell.setBackground("#d1e7dd").setFontColor("#000000");
  } else if (maxSides >= 15) {
    if (result < 10) cell.setBackground("#fff2cc").setFontColor("#000000");
    else cell.setBackground("#d1e7dd").setFontColor("#000000");
  } else {
    cell.setBackground("#d1e7dd").setFontColor("#000000");
  }
}

// Вспомогательная функция для получения текста оценки
function getDiceMessage(result, maxSides) {
  if (result === 1) return "полный крах";
  if (result === maxSides) return "вау, да ты жжешь!";
  if (maxSides === 100) return (result < 40) ? "могло быть и хуже..." : "хороший результат";
  if (maxSides >= 15) return (result < 10) ? "могло быть и хуже..." : "хороший результат";
  return "нормальный результат";
}
