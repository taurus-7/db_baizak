// ========== КЕШ ДЛЯ БЫСТРОГО ПОИСКА ==========
let clientDataCache = null;
let lastCachedTime = 0;
const CACHE_TIMEOUT = 5 * 60 * 1000; // 5 минут

// ========== НАСТРОЙКИ ==========
const WORD_TEMPLATE1 = "шаблон_договора";
const WORD_TEMPLATE2 = "шаблон_допсоглашения";
const WORD_TEMPLATE3 = "шаблон_допка";
const EXCEL_TEMPLATE = "template_zayavlenie";
const OUTPUT_FOLDER = "документы";
const DB_SHEET_NAME = "база"; // Лист с базой клиентов
const TEMPLATES_FOLDER = "шаблоны"; // Имя папки с шаблонами

// Тарифы для культур (тенге/га)
const TARIFFS = {
  "картофель": 7067,
  "овощи(бюджет.орган)": 7067,
  "сах.свекла": 7133,
  "многолетние травы": 8400,
  "подсолнечник": 4800,
  "бахчевые": 3933,
  "кукуруза на зерно": 5733,
  "сады": 7733,
  "соя(масленичные)": 5067,
  "яровые зерновые": 3733,
  "озимая пшеница": 3133,
  "тал-терек": 7733
};

// КОНФИГУРАЦИЯ ДЛЯ РАЗНЫХ ЛИСТОВ EXCEL
const SHEET_CONFIG = {
  "өтінім": {
    BASIC_FIELDS: {
      "Номер договора": "D1",
      "Номер телефона": "C12",
      "Канал": "A16",
      "Наличие земель": "D13",
      "Орошаемые": "D14"
    },
    COMPOSITE_FIELDS: {
      "кх_фио": { cell: "C8", template: "к/х {{КХ}} {{ФИО чистый}}" },
      "full_adress": { cell: "C10", template: "Адрес: {{Село}}, {{Полный адрес}}" }
    },
    CULTURE_CELLS: {
      "сах.свекла": ["D22"],
      "овощи(бюджет.орган)": ["D23"],
      "многолетние травы": ["D24"],
      "подсолнечник": ["D25"],
      "бахчевые": ["D26"],
      "кукуруза на зерно": ["D27"],
      "сады": ["D28"],
      "картофель": ["D29"],
      "соя(масленичные)": ["D30"],
      "яровые зерновые": ["D31"],
      "озимая пшеница": ["D32"],
      "тал-терек": ["D33"]
    }
  },
  "жоспар": {
    SIMPLE_FIELDS: {
      "Номер договора": "B40"
    },
    COMPOSITE_FIELDS: {
      "кх_фио": { cell: "D40", template: "к/х {{КХ}} {{ФИО чистый}}" }
    }
  }
};

// Поля для преобразования в integer
const INTEGER_FIELDS = {
  "Номер": "номер",
  "Номер телефона": "телефон",
  "ИИН/БИН": "иин_бин",
  "Объем": "объем"
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Генератор документов')
    .addItem('Генерировать весь реестр', 'generateDocuments')
    .addItem('Найти клиента по ИИН', 'showForm')
    .addItem('Добавить нового клиента', 'showCreateForm')
    .addItem('Обновить кеш базы', 'resetCache')
    .addToUi();
}

// ========== ОСНОВНЫЕ ФУНКЦИИ ==========
function generateDocuments() {
  createBackup();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  
  const colIndex = {};
  headers.forEach((header, index) => {
    colIndex[header] = index;
  });
  
  const requiredColumns = ['Номер', 'ФИО чистый', 'КХ', 'Документы_созданы', 'Дата договора'];
  requiredColumns.forEach(col => {
    if (!(col in colIndex)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      colIndex[col] = sheet.getLastColumn() - 1;
    }
  });

  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const outputFolder = getOrCreateFolder(parentFolder, OUTPUT_FOLDER);
  
  const wordTemplate1 = getTemplateFile(parentFolder, WORD_TEMPLATE1);
  const wordTemplate2 = getTemplateFile(parentFolder, WORD_TEMPLATE2);
  const wordTemplate3 = getTemplateFile(parentFolder, WORD_TEMPLATE3);
  const excelTemplate = getTemplateFile(parentFolder, EXCEL_TEMPLATE);
  
  for (let i = 1; i < data.length && i < 1000; i++) {
    const row = data[i];
    
    if (!row[colIndex['Номер']] || !row[colIndex['ФИО чистый']]) continue;
    if (row[colIndex['Документы_созданы']]) continue;
    
    const clientFolderName = cleanFilename(`${row[colIndex['Номер']]}_${row[colIndex['ФИО чистый']]}_${row[colIndex['КХ']]}`);
    const clientFolder = getOrCreateFolder(outputFolder, clientFolderName);
    
    const context = createContext(row, colIndex);
    context['дата_договора'] = formatDate(row[colIndex['Дата договора']]);
    
    // Генерация договора
    const docName1 = `${cleanFilename(row[colIndex['Номер']])}_договор`;
    const newFile1 = wordTemplate1.makeCopy(docName1, clientFolder);
    const doc1 = DocumentApp.openById(newFile1.getId());
    replaceInDoc(doc1, context);
    doc1.saveAndClose();
    
    // Генерация допсоглашения
    const docName2 = `${cleanFilename(row[colIndex['Номер']])}_допсоглашение`;
    const newFile2 = wordTemplate2.makeCopy(docName2, clientFolder);
    const doc2 = DocumentApp.openById(newFile2.getId());
    replaceInDoc(doc2, context);
    doc2.saveAndClose();
    
    // Генерация допки
    const docName3 = `${cleanFilename(row[colIndex['Номер']])}_допка`;
    const newFile3 = wordTemplate3.makeCopy(docName3, clientFolder);
    const doc3 = DocumentApp.openById(newFile3.getId());
    replaceInDoc(doc3, context);
    doc3.saveAndClose();
    
    // Генерация заявления
    const appName = `${cleanFilename(row[colIndex['Номер']])}_заявление`;
    const newFile4 = excelTemplate.makeCopy(appName, clientFolder);
    const ssApp = SpreadsheetApp.openById(newFile4.getId());
    
    const sheets = ssApp.getSheets();
    sheets.forEach(sheetApp => {
      const sheetName = sheetApp.getName();
      const skipSheets = ["договор-1,3,5", "договор-2,4,6"];
      if (skipSheets.includes(sheetName)) return;
      
      const config = SHEET_CONFIG[sheetName];
      if (config) fillApplication(sheetApp, row, colIndex, config);
    });
    
    sheet.getRange(i + 1, colIndex['Документы_созданы'] + 1).setValue(new Date());
    SpreadsheetApp.flush();
  }
}

// ========== ФУНКЦИИ ДЛЯ РАБОТЫ С ФОРМОЙ ПОИСКА ==========
function showForm() {
  resetCache();
  
  const html = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Поиск клиента по ИИН');
}

function getClientDatabase() {
  const now = new Date().getTime();
  
  if (clientDataCache && (now - lastCachedTime) < CACHE_TIMEOUT) {
    return clientDataCache;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  
  if (!dbSheet) {
    console.error(`❌ Лист '${DB_SHEET_NAME}' не найден!`);
    return {};
  }
  
  // Получаем заголовки базы данных
  const headers = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];
  const dbColIndex = {};
  
  // Создаем индекс колонок для базы данных
  headers.forEach((header, index) => {
    dbColIndex[header] = index;
  });
  
  const lastRow = dbSheet.getLastRow();
  const data = dbSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  
  clientDataCache = {};
  
  data.forEach(row => {
    const iin = row[dbColIndex["ИИН/БИН"]]?.toString().trim() || '';
    if (iin && iin.length >= 5) {
      const client = {
        iin: iin,
        fullName: row[dbColIndex["ФИО чистый"]] || '',
        kx: row[dbColIndex["КХ"]] || '',
        village: row[dbColIndex["Село"]] || '',
        address: row[dbColIndex["Полный адрес"]] || '',
        phone: row[dbColIndex["Номер телефона"]] || '',
        l_availability: parseFloat(row[dbColIndex["Наличие земель"]]) || 0,
        l_irrigated: parseFloat(row[dbColIndex["Орошаемые"]]) || 0,
        volume_water: parseFloat(row[dbColIndex["Объем"]]) || 0,
        tarrif: parseFloat(row[dbColIndex["Тариф"]]) || 0,
        channel: row[dbColIndex["Канал"]] || '',
        channel_allocation: row[dbColIndex["Выдел"]] || '',
        cadastral: row[dbColIndex["Кадастровый номер"]] || '',
        sum_snds: parseFloat(row[dbColIndex["Суммандс"]]) || 0
      };
      
      // Добавляем данные по культурам
      Object.keys(TARIFFS).forEach(culture => {
        client[culture] = parseFloat(row[dbColIndex[culture]]) || 0;
      });
      
      clientDataCache[iin] = client;
    }
  });
  
  lastCachedTime = now;
  console.log(`🔄 Кеш базы обновлен. Записей: ${Object.keys(clientDataCache).length}`);
  return clientDataCache;
}

function findClientByIIN(iin) {
  const cleanIIN = iin.toString().trim();
  
  if (!/^\d{12}$/.test(cleanIIN)) {
    throw new Error("ИИН должен содержать 12 цифр");
  }
  
  const database = getClientDatabase();
  const client = database[cleanIIN];
  
  if (!client) {
    throw new Error("Клиент с ИИН " + cleanIIN + " не найден");
  }
  
  return client;
}

// ========== ФУНКЦИИ ДЛЯ СОЗДАНИЯ НОВЫХ КЛИЕНТОВ ==========
function showCreateForm() {
  const html = HtmlService.createHtmlOutputFromFile('CreateForm')
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Добавление нового клиента');
}

function generateDocumentsForClient(clientData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Проверяем обязательные поля
  if (!clientData.fullName || !clientData.kx) {
    throw new Error("Заполните обязательные поля: ФИО чистый и КХ");
  }
  
  const newRow = sheet.getLastRow() + 1;
  
  // Генерируем номер договора
  const contractNumber = clientData.contractNumber || generateContractNumber(newRow);
  
  // Формируем значение для столбца "ФИО"
  const fullNameWithKX = `${clientData.fullName || ''} к/х ${clientData.kx || ''}`;
  
  const mapping = {
    "Номер": newRow,  // Порядковый номер в реестре
    "ФИО": fullNameWithKX,
    "ИИН/БИН": clientData.iin,
    "ФИО чистый": clientData.fullName,
    "КХ": clientData.kx,
    "Село": clientData.village,
    "Полный адрес": clientData.address,
    "Номер телефона": clientData.phone,
    "Наличие земель": clientData.l_availability || 0,
    "Орошаемые": clientData.l_irrigated || 0,
    "Объем": clientData.volume_water || 0,
    "Тариф": clientData.tarrif || 0,
    "Канал": clientData.channel,
    "Выдел": clientData.channel_allocation,
    "Кадастровый номер": clientData.cadastral,
    "Суммандс": clientData.sum_snds || 0,
    "Номер договора": contractNumber,
    // Культуры
    "картофель": clientData["картофель"] || 0,
    "овощи(бюджет.орган)": clientData["овощи(бюджет.орган)"] || 0,
    "сах.свекла": clientData["сах.свекла"] || 0,
    "многолетние травы": clientData["многолетние травы"] || 0,
    "подсолнечник": clientData["подсолнечник"] || 0,
    "бахчевые": clientData["бахчевые"] || 0,
    "кукуруза на зерно": clientData["кукуруза на зерно"] || 0,
    "сады": clientData["сады"] || 0,
    "соя(масленичные)": clientData["соя(масленичные)"] || 0,
    "яровые зерновые": clientData["яровые зерновые"] || 0,
    "озимая пшеница": clientData["озимая пшеница"] || 0,
    "тал-терек": clientData["тал-терек"] || 0
  };
  
  // Создаем недостающие колонки
  for (const header in mapping) {
    if (!headers.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  }
  
  // Заполняем данные
  for (const [header, value] of Object.entries(mapping)) {
    const colIndex = headers.indexOf(header) + 1;
    if (colIndex > 0) {
      sheet.getRange(newRow, colIndex).setValue(value);
    }
  }
  
  // Получаем актуальные индексы колонок
  const colIndex = {};
  headers.forEach((header, index) => {
    colIndex[header] = index;
  });
  
  // Добавляем дату создания
  if (!colIndex['Документы_созданы']) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Документы_созданы');
    colIndex['Документы_созданы'] = sheet.getLastColumn() - 1;
  }
  
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const outputFolder = getOrCreateFolder(parentFolder, OUTPUT_FOLDER);
  
  // Формируем имя папки и файлов
  const folderName = cleanFilename(`${newRow}_${clientData.fullName}_${clientData.kx}`);
  const clientFolder = getOrCreateFolder(outputFolder, folderName);
  
  const row = sheet.getRange(newRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const context = createContext(row, colIndex);
  context['дата_договора'] = formatDate(new Date());
  context['номер_договора'] = contractNumber;
  
  // Генерация документов
  const filePrefix = cleanFilename(`${newRow}_${clientData.fullName}`);
  
  // Договор
  const docName1 = `${filePrefix}_договор`;
  const wordTemplate1 = getTemplateFile(parentFolder, WORD_TEMPLATE1);
  const newFile1 = wordTemplate1.makeCopy(docName1, clientFolder);
  const doc1 = DocumentApp.openById(newFile1.getId());
  replaceInDoc(doc1, context);
  doc1.saveAndClose();
  
  // Допсоглашение
  const docName2 = `${filePrefix}_допсоглашение`;
  const wordTemplate2 = getTemplateFile(parentFolder, WORD_TEMPLATE2);
  const newFile2 = wordTemplate2.makeCopy(docName2, clientFolder);
  const doc2 = DocumentApp.openById(newFile2.getId());
  replaceInDoc(doc2, context);
  doc2.saveAndClose();
  
  // Допка
  const docName3 = `${filePrefix}_допка`;
  const wordTemplate3 = getTemplateFile(parentFolder, WORD_TEMPLATE3);
  const newFile3 = wordTemplate3.makeCopy(docName3, clientFolder);
  const doc3 = DocumentApp.openById(newFile3.getId());
  replaceInDoc(doc3, context);
  doc3.saveAndClose();
  
  // Заявление
  const appName = `${filePrefix}_заявление`;
  const excelTemplate = getTemplateFile(parentFolder, EXCEL_TEMPLATE);
  const newFile4 = excelTemplate.makeCopy(appName, clientFolder);
  const ssApp = SpreadsheetApp.openById(newFile4.getId());
  
  const sheets = ssApp.getSheets();
  sheets.forEach(sheetApp => {
    const sheetName = sheetApp.getName();
    const skipSheets = ["договор-1,3,5", "договор-2,4,6"];
    if (skipSheets.includes(sheetName)) return;
    
    const config = SHEET_CONFIG[sheetName];
    if (config) fillApplication(sheetApp, row, colIndex, config);
  });
  
  sheet.getRange(newRow, colIndex['Документы_созданы'] + 1).setValue(new Date());
  SpreadsheetApp.flush();
  
  return "✅ 4 документа успешно созданы!";
}

function generateContractNumber(rowIndex) {
  const currentYear = new Date().getFullYear();
  const lastDigit = currentYear.toString().slice(-1);
  return `№ ${lastDigit}-${rowIndex} Ф/П`;
}

// ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
function resetCache() {
  clientDataCache = null;
  console.log("♻️ Кеш базы данных сброшен");
  return "Кеш обновлен!";
}

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const backupPrefix = sheet.getName() + "_backup_";
  
  deleteOldBackups(ss, backupPrefix);
  
  const timestamp = Utilities.formatDate(new Date(), "GMT+6", "yyyyMMdd_HHmmss");
  const backupName = backupPrefix + timestamp;
  sheet.copyTo(ss).setName(backupName);
}

function deleteOldBackups(ss, prefix) {
  const sheets = ss.getSheets();
  const backups = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(prefix)) {
      backups.push({
        name: name,
        date: name.replace(prefix, ""),
        sheet: sheet
      });
    }
  });
  
  backups.sort((a, b) => a.date.localeCompare(b.date));
  
  if (backups.length > 3) {
    const toDelete = backups.length - 3;
    for (let i = 0; i < toDelete; i++) {
      ss.deleteSheet(backups[i].sheet);
    }
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function getTemplateFile(parentFolder, fileName) {
  const templatesFolder = getOrCreateFolder(parentFolder, TEMPLATES_FOLDER);
  const files = templatesFolder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().startsWith(fileName)) {
      return file;
    }
  }
  throw new Error(`❌ Файл шаблона "${fileName}" не найден в папке "${TEMPLATES_FOLDER}"`);
}

function cleanFilename(text) {
  return String(text)
    .replace(/[\\\/:*?"<>|]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 100);
}

function createContext(row, colIndex) {
  const context = {};
  for (const header in colIndex) {
    const key = header
      .replace(/\s+/g, '_')
      .replace(/\//g, '_')
      .replace(/-/g, '_')
      .toLowerCase();
    let value = row[colIndex[header]];
    
    if (INTEGER_FIELDS[header]) {
      value = toInteger(value);
    }
    
    context[key] = value !== null && value !== undefined ? value.toString() : '';
  }
  return context;
}

function toInteger(value) {
  if (typeof value === 'string') {
    const cleaned = value.replace(/[^\d]/g, '');
    return cleaned ? parseInt(cleaned, 10) : 0;
  }
  return value !== null && value !== undefined ? parseInt(value, 10) : 0;
}

function replaceInDoc(doc, context) {
  const body = doc.getBody();
  for (const key in context) {
    body.replaceText(`{{${key}}}`, context[key] || '');
  }
}

function fillApplication(sheet, row, colIndex, config) {
  if (config.SIMPLE_FIELDS) {
    for (const field in config.SIMPLE_FIELDS) {
      const cell = config.SIMPLE_FIELDS[field];
      const value = row[colIndex[field]] || '';
      sheet.getRange(cell).setValue(value);
    }
  }
  
  if (config.COMPOSITE_FIELDS) {
    for (const field in config.COMPOSITE_FIELDS) {
      const { cell, template } = config.COMPOSITE_FIELDS[field];
      let result = template;
      
      const placeholders = template.match(/{{(.*?)}}/g) || [];
      placeholders.forEach(placeholder => {
        const key = placeholder.replace(/[{}]/g, '').trim();
        const value = row[colIndex[key]] || '';
        result = result.replace(placeholder, value);
      });
      
      sheet.getRange(cell).setValue(result);
    }
  }
  
  if (config.BASIC_FIELDS) {
    for (const field in config.BASIC_FIELDS) {
      const cell = config.BASIC_FIELDS[field];
      const value = row[colIndex[field]] || '';
      sheet.getRange(cell).setValue(value);
    }
  }
  
  if (config.CULTURE_CELLS) {
    for (const culture in config.CULTURE_CELLS) {
      const cells = config.CULTURE_CELLS[culture];
      const value = parseFloat(row[colIndex[culture]]) || 0;
      cells.forEach(cell => {
        sheet.getRange(cell).setValue(value);
      });
    }
  }
}

function formatDate(dateValue) {
  if (!dateValue) return "";
  try {
    const date = new Date(dateValue);
    return Utilities.formatDate(date, "GMT+6", "dd.MM.yyyy");
  } catch (e) {
    return dateValue.toString();
  }
}
