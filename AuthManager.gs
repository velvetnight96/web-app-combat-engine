/**
 * Главная функция запуска сайта. 
 * Указывает Google, какой именно HTML-файл открывать по ссылке.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('auth-screen')
      .setTitle('Кто ты, Герой?')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Сверяет логин/пароль с базой данных на листе _Users_DB.
 */
function authenticateUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('_Users_DB');
  
  if (!userSheet) {
    return { success: false, message: "Ошибка: Лист _Users_DB не найден!" };
  }
  
  const userData = userSheet.getDataRange().getValues();
  let isAuthenticated = false;
  
  for (let i = 1; i < userData.length; i++) {
    const dbUser = userData[i][0].toString().trim();
    const dbPass = userData[i][1].toString().trim();
    
    if (dbUser.toLowerCase() === username.toString().trim().toLowerCase() && dbPass === password.toString()) {
      isAuthenticated = true;
      break;
    }
  }
  
  if (!isAuthenticated) {
    return { success: false, message: "Неверный ник или пароль!" };
  }
  
  return {
    success: true,
    username: username,
    characters: getPlayerCharactersList(username)
  };
}

/**
 * Быстрый вход без пароля (для localStorage в браузере).
 */
function authenticateUserWithoutPassword(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('_Users_DB');
  
  if (!userSheet) return { success: false };
  
  const userData = userSheet.getDataRange().getValues();
  let userExists = false;
  
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0].toString().trim().toLowerCase() === username.toString().trim().toLowerCase()) {
      userExists = true;
      break;
    }
  }
  
  if (!userExists) return { success: false };
  
  return {
    success: true,
    username: username,
    characters: getPlayerCharactersList(username)
  };
}

/**
 * Внутренняя функция сбора персонажей игрока (с защитой от пустых ячеек).
 */
function getPlayerCharactersList(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const matchedCharacters = [];
  
  sheets.forEach(sheet => {
    const k1Range = sheet.getRange("K1");
    if (!k1Range.isBlank() && k1Range.getValue().toString().trim() === "Инфо") {
      const ownerVal = sheet.getRange("L2").getValue();
      const owner = ownerVal ? ownerVal.toString().trim() : "";
      
      if (owner.toLowerCase() === username.toString().trim().toLowerCase()) {
        // Добавляем защиту: проверяем, есть ли вообще данные в ячейках
        const charNameVal = sheet.getRange("L6").getValue();
        const levelVal = sheet.getRange("L7").getValue();
        
        matchedCharacters.push({
          sheetName: sheet.getName(),
          charName: charNameVal ? charNameVal.toString() : "Безымянный",
          level: levelVal ? levelVal.toString() : "1"
        });
      }
    }
  });
  
  return matchedCharacters;
}

/**
 * Считывает все данные с конкретного листа для Web App.
 */
function getCharacterFullProfile(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { success: false, message: "Лист не найден" };
  }
  
  const mainBlock = sheet.getRange("K1:L45").getValues();
  
  const profile = {
    success: true,
    name: mainBlock[5][1],
    level: mainBlock[6][1],
    soulType: mainBlock[7][1],
    race: mainBlock[2][1],
    hpCurrent: 85, // Временная статика, пока не подключим формулы
    hpMax: 100,
    apCurrent: 6,
    apMax: 6,
    equipment: {},
    resists: [],
    consumables: [],
    weapons: []
  };

  // Экипировка (Строки 13-26)
  for (let i = 12; i <= 25; i++) {
    let slotName = mainBlock[i][0].toString().replace(":", "").trim();
    let slotValue = mainBlock[i][1].toString().trim();
    if (slotName) {
      profile.equipment[slotName] = slotValue || "—";
    }
  }

  // Сопротивления (Строки 28-42)
  for (let i = 27; i <= 41; i++) {
    let resName = mainBlock[i][0].toString().trim();
    let resValue = mainBlock[i][1];
    if (resName) {
      profile.resists.push({
        name: resName,
        value: typeof resValue === 'number' ? (resValue * 100).toFixed(0) + "%" : resValue || "0%"
      });
    }
  }

  // Расходники (V2:W15)
  const invBlock = sheet.getRange("V2:W15").getValues();
  invBlock.forEach(row => {
    let itemName = row[0].toString().trim();
    let itemCount = row[1];
    if (itemName) {
      profile.consumables.push({ name: itemName, count: itemCount + " шт." });
    }
  });

  // Запасное оружие (M13:N17)
  const weaponBlock = sheet.getRange("M13:N17").getValues();
  weaponBlock.forEach(row => {
    let wName = row[0].toString().trim();
    let wDurability = row[1];
    if (wName) {
      profile.weapons.push({ name: wName, count: wDurability + " Пр." });
    }
  });



  /**
 * Возвращает HTML-код игрового экрана по запросу фронтенда.
 */
function loadGameScreen() {
  return HtmlService.createHtmlOutputFromFile('game-screen').getContent();
}


  return profile;
}
