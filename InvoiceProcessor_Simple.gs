/**
 * СИСТЕМА СВЕРКИ СЧЕТОВ С 1С - УПРОЩЕННАЯ ВЕРСИЯ
 * Работа только с отчетом "Список" из 1С
 */

// ============================================
// КОНФИГУРАЦИЯ
// ============================================

const CONFIG = {
  // ID вашей основной таблицы со счетами
  MAIN_SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  
  // Названия листов
  SHEETS: {
    PRELIMINARY: 'Упаковка',  // Лист с предварительными данными
    COMPARISON: 'Сверка',
    LOG: 'Лог обработки'
  },
  
  // Колонки в предварительной таблице (по индексу, начиная с 1)
  COLUMNS: {
    DATE: 1,           // A - Дата
    CLIENT: 2,         // B - Клиент
    TARIFF: 3,         // C - Тарифы
    SERVICE: 4,        // D - Услуга
    QUANTITY: 11,      // K - Количество
    AMOUNT: 13,        // M - Сумма
    COMMENT: 24,       // X - Комментарий с номером документа 1С
    INDEX: 23          // W - Индекс
  },
  
  // Регулярное выражение для поиска номеров документов 1С
  DOCUMENT_NUMBER_REGEX: /[А-ЯA-Z0-9]{2,4}УТ-\d{6}/g
};

// ============================================
// ЛОГИРОВАНИЕ
// ============================================

/**
 * Запись в лог
 */
function logMessage(message, level = 'INFO') {
  const timestamp = new Date();
  const logEntry = `[${timestamp.toLocaleString('ru-RU')}] [${level}] ${message}`;
  
  console.log(logEntry);
  
  // Записываем в лист логов
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.SHEETS.LOG);
      logSheet.getRange(1, 1, 1, 3).setValues([['Время', 'Уровень', 'Сообщение']]);
      logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 3).setValues([[timestamp, level, message]]);
    
    // Ограничиваем лог последними 1000 записями
    if (lastRow > 1000) {
      logSheet.deleteRows(2, lastRow - 1000);
    }
    
  } catch (error) {
    console.error('Ошибка записи в лог:', error.message);
  }
}

// ============================================
// НАСТРОЙКА СИСТЕМЫ
// ============================================

/**
 * Первичная настройка системы
 */
function setupSystem() {
  logMessage('=== НАСТРОЙКА СИСТЕМЫ ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Создаем необходимые листы
  [CONFIG.SHEETS.COMPARISON, CONFIG.SHEETS.LOG].forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
      logMessage(`Создан лист: ${sheetName}`);
    }
  });
  
  logMessage('Настройка завершена');
  logMessage('ВАЖНО: Укажите номера колонок в переменной CONFIG.COLUMNS');
}

/**
 * Проверка конфигурации
 */
function validateConfiguration() {
  logMessage('=== ПРОВЕРКА КОНФИГУРАЦИИ ===');
  
  const errors = [];
  
  // Проверка таблицы
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    logMessage('✓ Основная таблица найдена');
    
    // Проверка листов
    const prelimSheet = ss.getSheetByName(CONFIG.SHEETS.PRELIMINARY);
    if (prelimSheet) {
      logMessage('✓ Лист предварительной таблицы найден');
      
      // Проверка колонки с комментариями
      const testComment = prelimSheet.getRange(2, CONFIG.COLUMNS.COMMENT).getValue();
      logMessage(`Пример комментария из строки 2: "${testComment}"`);
      
      if (testComment) {
        const match = String(testComment).match(CONFIG.DOCUMENT_NUMBER_REGEX);
        if (match) {
          logMessage(`✓ Найден номер документа: ${match[0]}`);
        } else {
          logMessage('⚠ В комментарии не найден номер документа (формат: 02УТ-003392)', 'WARNING');
        }
      }
    } else {
      errors.push(`Лист "${CONFIG.SHEETS.PRELIMINARY}" не найден`);
    }
  } catch (error) {
    errors.push(`Ошибка доступа к таблице: ${error.message}`);
  }
  
  // Проверка регулярного выражения
  const testString = '02УТ-003392';
  const match = testString.match(CONFIG.DOCUMENT_NUMBER_REGEX);
  if (match) {
    logMessage('✓ Регулярное выражение работает корректно');
  } else {
    errors.push('Регулярное выражение не распознает номера документов');
  }
  
  if (errors.length > 0) {
    logMessage('\n❌ НАЙДЕНЫ ОШИБКИ:', 'ERROR');
    errors.forEach(error => logMessage(`  - ${error}`, 'ERROR'));
    return false;
  } else {
    logMessage('\n✓ Конфигурация валидна');
    return true;
  }
}

/**
 * Очистка всех обработанных данных
 */
function clearAllData() {
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COMPARISON);
  if (sheet) {
    sheet.clear();
    logMessage('Очищен лист: ' + CONFIG.SHEETS.COMPARISON);
  }
  
  logMessage('Все обработанные данные очищены');
}
