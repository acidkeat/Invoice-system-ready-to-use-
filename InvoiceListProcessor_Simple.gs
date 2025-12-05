/**
 * МОДУЛЬ РАБОТЫ С ОТЧЕТОМ "СПИСОК" ИЗ 1С
 * Упрощенная версия - только сверка по номерам документов
 */

// ============================================
// КОНФИГУРАЦИЯ ДЛЯ ОТЧЕТА "СПИСОК"
// ============================================

const LIST_CONFIG = {
  // ID папки Google Drive для отчета "Список"
  LIST_IMPORT_FOLDER_ID: '1BHeo0ZTZcjgVrwcRNSfzbVWVqxUG_pcV',  // Твоя папка
  
  // Название листа для импорта списка
  LIST_SHEET_NAME: 'Список из 1С',
  
  // Минимальная дата для обработки (игнорируем документы до этой даты)
  MIN_DATE: new Date('2025-09-01'),
  
  // Колонки в отчете "Список" (нумерация с 0!)
  LIST_COLUMNS: {
    PRIORITY: 0,      // Приоритет
    NUMBER: 1,        // Номер документа
    DATE: 2,          // Дата
    AMOUNT: 3,        // Сумма
    CLIENT: 4,        // Клиент
    STATUS: 5,        // Статус
    PAID_AMOUNT: 6,   // Сумма оплаты
    PAID_PERCENT: 7,  // % оплаты
    SHIP_PERCENT: 8,  // % отгрузки
    HAS_DIFF: 9,      // Есть расхождения
    CURRENCY: 10,     // Валюта
    OPERATION: 11,    // Операция
    AUTHOR: 12        // Автор
  }
};

// ============================================
// ОБРАБОТКА ОТЧЕТА "СПИСОК"
// ============================================

/**
 * Главная функция обработки отчета "Список"
 */
function processListReport() {
  const startTime = new Date();
  logMessage('=== ОБРАБОТКА ОТЧЕТА "СПИСОК" ===');
  
  try {
    // Получаем файлы из папки
    const files = getListFiles();
    
    if (files.length === 0) {
      logMessage('Файлов "Список" не найдено');
      return;
    }
    
    logMessage(`Найдено файлов: ${files.length}`);
    
    // Обрабатываем последний (самый свежий) файл
    const file = files[0];
    logMessage(`Обработка файла: ${file.getName()}`);
    
    const data = processListFile(file);
    
    if (data.length === 0) {
      logMessage('Данных для импорта не найдено', 'WARNING');
      return;
    }
    
    logMessage(`Извлечено записей: ${data.length}`);
    
    // Записываем в лист
    writeListData(data);
    
    // Обновляем сверку с учетом статусов
    updateComparisonWithStatuses();
    
    // Перемещаем в обработано
    moveToProcessedFolder(file);
    
    const duration = (new Date() - startTime) / 1000;
    logMessage(`Обработка завершена за ${duration.toFixed(2)} сек`);
    
  } catch (error) {
    logMessage(`ОШИБКА при обработке списка: ${error.message}`, 'ERROR');
    throw error;
  }
}

/**
 * Получение файлов "Список" из папки
 */
function getListFiles() {
  if (LIST_CONFIG.LIST_IMPORT_FOLDER_ID === 'YOUR_FOLDER_ID_HERE') {
    logMessage('Не настроен ID папки для отчета "Список"', 'ERROR');
    logMessage('Укажите ID в переменной LIST_CONFIG.LIST_IMPORT_FOLDER_ID', 'ERROR');
    return [];
  }
  
  try {
    const folder = DriveApp.getFolderById(LIST_CONFIG.LIST_IMPORT_FOLDER_ID);
    const files = [];
    
    // Ищем Excel файлы
    const iterator = folder.getFiles();
    while (iterator.hasNext()) {
      const file = iterator.next();
      const name = file.getName().toLowerCase();
      
      // Ищем файлы со словом "список" в названии
      if ((name.includes('список') || name.includes('list')) && 
          (name.endsWith('.xlsx') || name.endsWith('.xls'))) {
        files.push(file);
      }
    }
    
    if (files.length === 0) {
      logMessage('Файлы "Список" не найдены. Убедитесь что в названии есть слово "список" или "list"', 'WARNING');
    }
    
    // Сортируем по дате изменения (новые первые)
    files.sort((a, b) => b.getLastUpdated().getTime() - a.getLastUpdated().getTime());
    
    return files;
    
  } catch (error) {
    logMessage(`Ошибка доступа к папке: ${error.message}`, 'ERROR');
    return [];
  }
}

/**
 * Обработка файла "Список"
 */
function processListFile(file) {
  // Конвертируем Excel в Google Sheets
  const tempSheet = convertExcelToSheets(file);
  
  if (!tempSheet) {
    logMessage('Не удалось конвертировать файл списка', 'ERROR');
    return [];
  }
  
  try {
    const sheet = tempSheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    const result = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Начинаем со второй строки (пропускаем заголовки)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const docNumber = row[LIST_CONFIG.LIST_COLUMNS.NUMBER];
      const dateValue = row[LIST_CONFIG.LIST_COLUMNS.DATE];
      const status = row[LIST_CONFIG.LIST_COLUMNS.STATUS];
      
      if (!docNumber || !status) continue;
      
      // Обработка даты
      let docDate = null;
      
      if (typeof dateValue === 'object' && dateValue instanceof Date) {
        docDate = dateValue;
      } else if (typeof dateValue === 'string') {
        // Если это время (13:35, 12:32) - берем сегодняшнюю дату
        if (dateValue.match(/^\d{1,2}:\d{2}$/)) {
          docDate = today;
          logMessage(`Документ ${docNumber}: дата = время (${dateValue}), используем сегодня`);
        } else {
          // Пробуем распарсить дату
          const parts = dateValue.split('.');
          if (parts.length === 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1;
            const year = parseInt(parts[2]);
            docDate = new Date(year, month, day);
          }
        }
      }
      
      // Проверка на минимальную дату
      if (docDate && docDate < LIST_CONFIG.MIN_DATE) {
        continue; // Пропускаем старые документы
      }
      
      result.push({
        number: docNumber,
        date: docDate || today,
        amount: row[LIST_CONFIG.LIST_COLUMNS.AMOUNT] || 0,
        client: row[LIST_CONFIG.LIST_COLUMNS.CLIENT] || '',
        status: status,
        paidAmount: row[LIST_CONFIG.LIST_COLUMNS.PAID_AMOUNT] || null,
        paidPercent: row[LIST_CONFIG.LIST_COLUMNS.PAID_PERCENT] || null,
        shipPercent: row[LIST_CONFIG.LIST_COLUMNS.SHIP_PERCENT] || null,
        hasDiff: row[LIST_CONFIG.LIST_COLUMNS.HAS_DIFF] || 'Нет',
        author: row[LIST_CONFIG.LIST_COLUMNS.AUTHOR] || ''
      });
    }
    
    // Удаляем временный файл
    DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
    
    logMessage(`Обработано документов (после фильтрации): ${result.length}`);
    return result;
    
  } catch (error) {
    logMessage(`Ошибка обработки данных списка: ${error.message}`, 'ERROR');
    if (tempSheet) {
      DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
    }
    throw error;
  }
}

/**
 * Конвертация Excel в Google Sheets
 */
function convertExcelToSheets(file) {
  try {
    const blob = file.getBlob();
    const config = {
      title: `temp_${file.getName()}_${new Date().getTime()}`,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: LIST_CONFIG.LIST_IMPORT_FOLDER_ID}]
    };
    
    const resource = Drive.Files.insert(config, blob);
    const spreadsheet = SpreadsheetApp.openById(resource.id);
    
    return spreadsheet;
  } catch (error) {
    logMessage(`Ошибка конвертации файла: ${error.message}`, 'ERROR');
    return null;
  }
}

/**
 * Запись данных списка в Google Sheets
 */
function writeListData(data) {
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LIST_CONFIG.LIST_SHEET_NAME);
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = ss.insertSheet(LIST_CONFIG.LIST_SHEET_NAME);
    
    // Добавляем заголовки
    const headers = [
      'Дата импорта',
      'Номер документа',
      'Дата документа',
      'Клиент',
      'Сумма',
      'Статус',
      'Сумма оплаты',
      '% оплаты',
      '% отгрузки',
      'Расхождения',
      'Автор'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Очищаем старые данные (кроме заголовков)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
    }
  }
  
  const importDate = new Date();
  
  // Подготавливаем данные для записи
  const rows = data.map(item => [
    importDate,
    item.number,
    item.date,
    item.client,
    item.amount,
    item.status,
    item.paidAmount,
    item.paidPercent,
    item.shipPercent,
    item.hasDiff,
    item.author
  ]);
  
  // Записываем данные
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 11).setValues(rows);
    
    // Применяем форматирование
    applyListFormatting(sheet, rows.length);
    
    logMessage(`Записано строк в лист "${LIST_CONFIG.LIST_SHEET_NAME}": ${rows.length}`);
  }
}

/**
 * Применение форматирования к листу "Список из 1С"
 */
function applyListFormatting(sheet, dataRows) {
  const statusColumn = 6; // Колонка "Статус"
  
  for (let i = 2; i <= dataRows + 1; i++) {
    const statusCell = sheet.getRange(i, statusColumn);
    const status = statusCell.getValue();
    
    if (status === 'К отгрузке') {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#d9ead3'); // Зеленый
    } else if (status === 'На согласовании') {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#fff2cc'); // Желтый
    } else if (status === 'К выполнению / В резерве') {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#cfe2f3'); // Голубой
    }
  }
  
  // Автоширина колонок
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

// ============================================
// СВЕРКА С УЧЕТОМ СТАТУСОВ
// ============================================

/**
 * Обновление листа сверки с добавлением статусов из списка
 */
function updateComparisonWithStatuses() {
  logMessage('\n--- ОБНОВЛЕНИЕ СВЕРКИ СО СТАТУСАМИ ---');
  
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  
  // Получаем данные из предварительной таблицы
  const prelimSheet = ss.getSheetByName(CONFIG.SHEETS.PRELIMINARY);
  if (!prelimSheet) {
    logMessage('Лист предварительной таблицы не найден', 'ERROR');
    return;
  }
  
  // Получаем данные из списка
  const listSheet = ss.getSheetByName(LIST_CONFIG.LIST_SHEET_NAME);
  if (!listSheet) {
    logMessage('Лист "Список из 1С" не найден. Запустите processListReport() сначала.', 'WARNING');
    return;
  }
  
  const prelimData = prelimSheet.getDataRange().getValues();
  const listData = listSheet.getDataRange().getValues();
  
  // Создаем карту: номер документа → данные из списка
  const docMap = new Map();
  for (let i = 1; i < listData.length; i++) {
    const docNumber = listData[i][1]; // Колонка "Номер документа"
    if (docNumber) {
      docMap.set(docNumber, {
        date: listData[i][2],
        client: listData[i][3],
        amount: listData[i][4],
        status: listData[i][5],
        paidAmount: listData[i][6],
        paidPercent: listData[i][7],
        shipPercent: listData[i][8],
        hasDiff: listData[i][9],
        author: listData[i][10]
      });
    }
  }
  
  logMessage(`Загружено документов из списка: ${docMap.size}`);
  
  // Создаем или обновляем лист сверки
  let compSheet = ss.getSheetByName(CONFIG.SHEETS.COMPARISON);
  if (!compSheet) {
    compSheet = ss.insertSheet(CONFIG.SHEETS.COMPARISON);
  } else {
    compSheet.clear();
  }
  
  // Заголовки для сверки
  const headers = [
    'Дата',
    'Клиент',
    'Услуга',
    'Количество',
    'Сумма',
    'Номер документа',
    'Индекс',
    'Статус в системе',
    'Статус 1С',
    '% отгрузки',
    'Автор',
    'Комментарий'
  ];
  
  compSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  compSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  compSheet.setFrozenRows(1);
  
  const comparisonResults = [];
  let foundInSystem = 0;
  let notFoundInSystem = 0;
  
  // Анализируем каждую строку предварительной таблицы
  for (let i = 1; i < prelimData.length; i++) {
    const row = prelimData[i];
    const commentCell = row[CONFIG.COLUMNS.COMMENT - 1];
    
    if (!commentCell) continue;
    
    // Ищем номера документов в комментарии
    const docNumbers = String(commentCell).match(CONFIG.DOCUMENT_NUMBER_REGEX);
    
    if (!docNumbers || docNumbers.length === 0) {
      // Нет номера документа
      comparisonResults.push([
        row[CONFIG.COLUMNS.DATE - 1],
        row[CONFIG.COLUMNS.CLIENT - 1],
        row[CONFIG.COLUMNS.SERVICE - 1],
        row[CONFIG.COLUMNS.QUANTITY - 1],
        row[CONFIG.COLUMNS.AMOUNT - 1],
        '-',
        row[CONFIG.COLUMNS.INDEX - 1],
        'НЕ В СИСТЕМЕ',
        '-',
        '-',
        '-',
        'Документ не создан в 1С'
      ]);
      notFoundInSystem++;
    } else {
      // Проверяем документ в списке
      const docNumber = docNumbers[0];
      const docInfo = docMap.get(docNumber);
      
      if (docInfo) {
        // Документ найден
        comparisonResults.push([
          row[CONFIG.COLUMNS.DATE - 1],
          row[CONFIG.COLUMNS.CLIENT - 1],
          row[CONFIG.COLUMNS.SERVICE - 1],
          row[CONFIG.COLUMNS.QUANTITY - 1],
          row[CONFIG.COLUMNS.AMOUNT - 1],
          docNumber,
          row[CONFIG.COLUMNS.INDEX - 1],
          'В СИСТЕМЕ',
          docInfo.status,
          docInfo.shipPercent || '-',
          docInfo.author,
          `Статус: ${docInfo.status}`
        ]);
        foundInSystem++;
      } else {
        // Документ не найден в списке
        comparisonResults.push([
          row[CONFIG.COLUMNS.DATE - 1],
          row[CONFIG.COLUMNS.CLIENT - 1],
          row[CONFIG.COLUMNS.SERVICE - 1],
          row[CONFIG.COLUMNS.QUANTITY - 1],
          row[CONFIG.COLUMNS.AMOUNT - 1],
          docNumber,
          row[CONFIG.COLUMNS.INDEX - 1],
          'НЕ НАЙДЕН В СПИСКЕ',
          '-',
          '-',
          '-',
          'Не найдено в отчете "Список" или документ старше 01.09.2025'
        ]);
        notFoundInSystem++;
      }
    }
  }
  
  // Записываем результаты сверки
  if (comparisonResults.length > 0) {
    compSheet.getRange(2, 1, comparisonResults.length, headers.length).setValues(comparisonResults);
    
    // Применяем условное форматирование
    applyComparisonFormatting(compSheet, comparisonResults.length);
  }
  
  logMessage(`Найдено в системе: ${foundInSystem}`);
  logMessage(`Не найдено в системе: ${notFoundInSystem}`);
  logMessage('Сверка со статусами завершена');
}

/**
 * Форматирование листа сверки
 */
function applyComparisonFormatting(sheet, dataRows) {
  const systemStatusColumn = 8;  // Колонка "Статус в системе"
  const status1CColumn = 9;      // Колонка "Статус 1С"
  
  for (let i = 2; i <= dataRows + 1; i++) {
    const systemStatus = sheet.getRange(i, systemStatusColumn).getValue();
    const status1C = sheet.getRange(i, status1CColumn).getValue();
    
    // Определяем цвет по статусу
    if (systemStatus === 'В СИСТЕМЕ') {
      // Дополнительно смотрим на статус 1С
      if (status1C === 'К отгрузке') {
        sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#b6d7a8'); // Ярко-зеленый
      } else if (status1C === 'На согласовании') {
        sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#ffe599'); // Оранжевый
      } else if (status1C === 'К выполнению / В резерве') {
        sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#cfe2f3'); // Голубой
      } else {
        sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#d9ead3'); // Светло-зеленый
      }
    } else if (systemStatus === 'НЕ В СИСТЕМЕ') {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#fff2cc'); // Желтый
    } else {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#f4cccc'); // Красный
    }
  }
  
  // Автоширина колонок
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================

/**
 * Перемещение файла в папку "Обработано"
 */
function moveToProcessedFolder(file) {
  try {
    const folder = DriveApp.getFolderById(LIST_CONFIG.LIST_IMPORT_FOLDER_ID);
    
    // Создаем или получаем папку "Обработано"
    let processedFolder;
    const folders = folder.getFoldersByName('Обработано');
    
    if (folders.hasNext()) {
      processedFolder = folders.next();
    } else {
      processedFolder = folder.createFolder('Обработано');
    }
    
    // Перемещаем файл
    file.moveTo(processedFolder);
    logMessage(`Файл перемещен в "Обработано": ${file.getName()}`);
    
  } catch (error) {
    logMessage(`Ошибка перемещения файла: ${error.message}`, 'WARNING');
  }
}

/**
 * Получение статистики по статусам
 */
function getStatusStatistics() {
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  const compSheet = ss.getSheetByName(CONFIG.SHEETS.COMPARISON);
  
  if (!compSheet) {
    return { error: 'Лист сверки не найден' };
  }
  
  const data = compSheet.getDataRange().getValues();
  
  const stats = {
    total: data.length - 1,
    bySystemStatus: {},
    by1CStatus: {},
    byShipPercent: {}
  };
  
  for (let i = 1; i < data.length; i++) {
    const systemStatus = data[i][7];  // Статус в системе
    const status1C = data[i][8];      // Статус 1С
    const shipPercent = data[i][9];   // % отгрузки
    
    // По статусу в системе
    stats.bySystemStatus[systemStatus] = (stats.bySystemStatus[systemStatus] || 0) + 1;
    
    // По статусу 1С
    if (status1C && status1C !== '-') {
      stats.by1CStatus[status1C] = (stats.by1CStatus[status1C] || 0) + 1;
    }
    
    // По % отгрузки
    if (shipPercent && shipPercent !== '-') {
      const key = shipPercent === 100 ? '100%' : 'Частично';
      stats.byShipPercent[key] = (stats.byShipPercent[key] || 0) + 1;
    }
  }
  
  return stats;
}

/**
 * Вывод статистики по статусам
 */
function showStatusStatistics() {
  const stats = getStatusStatistics();
  
  if (stats.error) {
    Logger.log(stats.error);
    return;
  }
  
  Logger.log('=== СТАТИСТИКА ПО СТАТУСАМ ===');
  Logger.log(`Всего записей: ${stats.total}`);
  Logger.log('');
  
  Logger.log('По статусу в системе:');
  for (const [status, count] of Object.entries(stats.bySystemStatus)) {
    Logger.log(`  ${status}: ${count} (${(count / stats.total * 100).toFixed(1)}%)`);
  }
  Logger.log('');
  
  Logger.log('По статусу 1С:');
  for (const [status, count] of Object.entries(stats.by1CStatus)) {
    Logger.log(`  ${status}: ${count} (${(count / stats.total * 100).toFixed(1)}%)`);
  }
  Logger.log('');
  
  Logger.log('По % отгрузки:');
  for (const [status, count] of Object.entries(stats.byShipPercent)) {
    Logger.log(`  ${status}: ${count} (${(count / stats.total * 100).toFixed(1)}%)`);
  }
}

// ============================================
// АВТОМАТИЗАЦИЯ
// ============================================

/**
 * Создание автоматического триггера
 */
function createAutoTrigger() {
  // Удаляем существующие триггеры
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processListReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Создаем новый триггер (каждые 6 часов)
  ScriptApp.newTrigger('processListReport')
    .timeBased()
    .everyHours(6)
    .create();
    
  logMessage('Автоматический триггер создан (каждые 6 часов)');
}

/**
 * Удаление автоматического триггера
 */
function removeAutoTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processListReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  logMessage('Автоматический триггер удален');
}
