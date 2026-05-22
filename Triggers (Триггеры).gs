/**
 * =========================================================================
 * ГЛАВНЫЙ МОДУЛЬ ТРИГГЕРОВ И ПЕРЕХВАТА КОМАНД (COMBAT ENGINE)
 * =========================================================================
 * 
 * ВНИМАНИЕ: Данная функция должна быть привязана как УСТАНАВЛИВАЕМЫЙ ТРИГГЕР
 * в панели управления Google Скриптов (On Edit / При изменении).
 */
function installedOnEdit(e) {
  // Защита от пустых вызовов (например, программных или системных сбоев)
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const value = e.value;
  
  // Ловим ввод "1" или числовое значение 1 в ячейках-командах
  if (value === "1" || value === 1) {
    try {
      
      // -------------------------------------------------------------------
      // КОМАНДА 1: КОНЕЦ ХОДА (Передача хода новому персонажу)
      // -------------------------------------------------------------------
      if (isCellNamed(e.range, "кн_конец_хода")) {
        e.range.clearContent(); // Мгновенная очистка единицы
        showFlashMessage(sheet, "⏳ ОБРАБОТКА СМЕНЫ ХОДА ДВИЖКОМ...", "yellow");
        executeEndTurn();       // Вызов фазы из TurnEngine.gs
        showFlashMessage(sheet, "✅ ХОД УСПЕШНО ПЕРЕДАН", "green");
        return;
      }
      
      // -------------------------------------------------------------------
      // КОМАНДА 2: ГЕНЕРАЦИЯ ДИАПАЗОНОВ (Из текста реестра -> в систему)
      // -------------------------------------------------------------------
      if (isCellNamed(e.range, "диапазон_ген")) {
        e.range.clearContent();
        showFlashMessage(sheet, "⏳ ГЕНЕРАЦИЯ ДИАПАЗОНОВ ИЗ РЕЕСТРА...", "yellow");
        generateAndSyncNamedRanges(); // Вызов из RangeRegistryManager.gs
        showFlashMessage(sheet, "✅ ДИАПАЗОНЫ УСПЕШНО СОЗДАНЫ", "green");
        return;
      }
      
      // -------------------------------------------------------------------
      // КОМАНДА 3: СКАНИРОВАНИЕ ДИАПАЗОНОВ (Из системы -> обратно в реестр)
      // -------------------------------------------------------------------
      if (isCellNamed(e.range, "диапазон_синх")) {
        e.range.clearContent();
        showFlashMessage(sheet, "⏳ СКАНИРОВАНИЕ ИМЕН И ОБНОВЛЕНИЕ РЕЕСТРА...", "yellow");
        reverseScanAndReadNamedRanges(); // Вызов из RangeRegistryManager.gs
        showFlashMessage(sheet, "✅ РЕЕСТР ДАННЫХ ОБНОВЛЕН", "green");
        return;
      }
      
      // -------------------------------------------------------------------
      // КОМАНДА 4: СОРТИРОВКА ПО АЛФАВИТУ (Чистый алфавит по Столбцу A)
      // -------------------------------------------------------------------
      if (isCellNamed(e.range, "алфавит_дб")) {
        e.range.clearContent();
        showFlashMessage(sheet, "⏳ СОРТИРОВКА РЕЕСТРА ПО АЛФАВИТУ...", "yellow");
        sortRegistryByAlphabet(); // Вызов из RangeRegistryManager.gs
        showFlashMessage(sheet, "✅ СОРТИРОВКА УСПЕШНО ЗАВЕРШЕНА", "green");
        return;
      }

      // -------------------------------------------------------------------
      // КОМАНДА 5: ГРУППИРОВКА ПО ТИПАМ (Тип [B] -> затем Алфавит [A])
      // -------------------------------------------------------------------
      if (isCellNamed(e.range, "тип_дб")) {
        e.range.clearContent();
        showFlashMessage(sheet, "⏳ ГРУППИРОВКА РЕЕСТРА ПО ТИПАМ...", "yellow");
        sortRegistryByTypeAndAlphabet(); // Вызов из RangeRegistryManager.gs
        showFlashMessage(sheet, "✅ ГРУППИРОВКА УСПЕШНО ЗАВЕРШЕНА", "green");
        return;
      }
      
    } catch (error) {
      // Жесткий перехват любых ошибок: зачищаем единицу и красим панель в красный
      e.range.clearContent();
      showFlashMessage(sheet, "❌ ОШИБКА: " + error.message, "red");
    }
  }
}

/**
 * ВСПОМОГАТЕЛЬНАЯ УТИЛИТА: Проверка принадлежности ячейки именованному диапазону
 * Защищает от ложных срабатываний, если ячейка с таким же адресом нажата на другом листе.
 */
function isCellNamed(range, name) {
  try {
    const namedRange = range.getSheet().getParent().getRangeByName(name);
    if (!namedRange) return false;
    
    // Сверяем имя листа, индекс строки и индекс столбца
    return (range.getSheet().getName() === namedRange.getSheet().getName() &&
            range.getRow() === namedRange.getRow() &&
            range.getColumn() === namedRange.getColumn());
  } catch(e) {
    return false; // Если диапазон с таким именем вообще не создан в таблице
  }
}

/**
 * СИСТЕМНАЯ АНИМАЦИЯ: Вывод лога («flashMessage») в ячейку "сообщение"
 * Плавно переключает цвета фонов в зависимости от стадии выполнения.
 */
function showFlashMessage(sheet, text, type) {
  const ss = sheet.getParent();
  let msgRange;
  
  try {
    msgRange = ss.getRangeByName("сообщение");
  } catch(e) { 
    return; // Если ячейку "сообщение" забыли создать, скрипт не упадет, а просто проигнорирует анимацию
  }
  
  // Записываем системный текст
  msgRange.setValue(text);
  
  // Цветовая палитра под тактический дизайн интерфейса
  let color = "#FFFFFF";
  if (type === "green") color = "#D4EDDA";  // Успех (Мягкий зеленый)
  if (type === "red") color = "#F8D7DA";    // Критическая ошибка (Мягкий красный)
  if (type === "yellow") color = "#FFF3CD"; // Загрузка процесса (Мягкий желтый)
  
  msgRange.setBackground(color);
  
  // Форсируем немедленную отправку изменений на сервер Google Таблиц
  SpreadsheetApp.flush();
  
  // Если процесс завершился (успех или ошибка), даем игроку рассмотреть надпись 1.5 секунды
  // и плавно возвращаем нейтральный белый фон. Если это "загрузка", оставляем цвет гореть.
  if (type !== "yellow") {
    Utilities.sleep(1500);
    msgRange.setBackground("#FFFFFF");
    SpreadsheetApp.flush();
  }
}

/**
 * Создает пользовательское меню при открытии таблицы Google.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎮 Combat Engine')
    .addItem('📇 Открыть карточку персонажа', 'openCharacterCardSidebar')
    .addToUi();
}

/**
 * Открывает файл CharacterCard.html в боковой панели (Sidebar).
 */
function openCharacterCardSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('CharacterCard')
      .setTitle('Combat Engine — Авторизация')
      .setWidth(300); // Оптимальная ширина для боковой панели
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Создает меню при открытии таблицы.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎮 Combat Engine')
    .addItem('📱 Открыть Мобильный Движок (Web App)', 'openEngineLink')
    .addToUi();
}

/**
 * Показывает во всплывающем окне кликабельную ссылку на ваш Dev-сайт.
 */
function openEngineLink() {
  const devUrl = ScriptApp.getService().getUrl(); // Автоматически найдет вашу ссылку
  
  if (!devUrl) {
    SpreadsheetApp.getUi().alert("Сначала разверните проект как Веб-приложение!");
    return;
  }
  
  // Создаем красивое окошко со ссылкой, которая сама откроется в новой вкладке
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p style="font-family:sans-serif; color:#e2e8f0; text-align:center;">Нажмите для перехода на игровой сайт:</p>
     <div style="text-align:center; margin-top:15px;">
       <a href="${devUrl}" target="_blank" style="background:#6b46c1; color:white; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold; font-family:sans-serif;">ОТКРЫТЬ ДВИЖОК</a>
     </div>`
  )
  .setWidth(300)
  .setHeight(120);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Запуск Web App');
}


