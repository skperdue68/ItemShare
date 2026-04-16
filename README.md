const APP_TITLE = 'ESO Item Share';
const EXPECTED_WINDOWS_PATH = String.raw`Documents\Elder Scrolls Online\live\SavedVariables\ItemShare.lua`;

const MASTER_SHEET_NAME = 'Master';

const HEADER_ROW = ['Name', 'Item Type', 'Quality', 'Trait', 'Date Added', 'Count', 'RequestedBy', 'SyncKey'];
const MASTER_HEADER_ROW = ['Name', 'Item Type', 'Quality', 'Trait', 'Date Added', 'Count', 'RequestedBy', 'Source Sheet', 'SyncKey'];

const COL_NAME = 1;
const COL_ITEM_TYPE = 2;
const COL_QUALITY = 3;
const COL_TRAIT = 4;
const COL_DATE_ADDED = 5;
const COL_COUNT = 6;
const COL_REQUESTED_BY = 7;
const COL_SYNC_KEY = 8;
const MASTER_COL_SOURCE_SHEET = 8;
const MASTER_COL_SYNC_KEY = 9;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Item Share')
    .addItem('Open Updater', 'showUploadDialog')
    .addItem('Reset/format active account sheet', 'prepareActiveSheetForImports')
    .addItem('Reapply Protections', 'reapplyProtectionsForActiveSheet')
    .addItem('Rebuild Master', 'rebuildMasterSheet')
    .addItem('Delete selected item safely', 'deleteSelectedItemSafely')
    .addToUi();
}

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UploadDialog')
    .setWidth(560)
    .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update ESO Item Share Sheet');
}

function getUploaderConfig() {
  return {
    appTitle: APP_TITLE,
    expectedWindowsPath: EXPECTED_WINDOWS_PATH
  };
}

function processUploadedSavedVariables(fileText) {
  if (!fileText || !String(fileText).trim()) {
    throw new Error('No file content was received.');
  }

  const parsedAccounts = parseAllSavedVariablesAccounts_(String(fileText));
  if (!parsedAccounts.length) {
    throw new Error('Could not find any ESO account sections with shared items in the saved variables file.');
  }

  let totalImportedItems = 0;
  let totalUpdatedRows = 0;
  let totalInsertedRows = 0;
  const updatedSheets = [];

  parsedAccounts.forEach(parsed => {
    const sheet = getOrCreateAccountSheet_(parsed.accountName);
    initializeAccountSheet_(sheet);

    const existingRows = readExistingRows_(sheet, false);
    const updates = mergeImportedRows_(existingRows, parsed.items);

    clearProtections_(sheet);
    writeRows_(sheet, updates.rows, false);
    sortAccountSheet_(sheet);
    applyProtections_(sheet, false);
    relockRequestedByCells_(sheet);

    totalImportedItems += parsed.items.length;
    totalUpdatedRows += updates.updatedRows;
    totalInsertedRows += updates.insertedRows;
    updatedSheets.push(sheet.getName());
  });

  rebuildMasterSheet_();
  activateMasterSheet_();
  SpreadsheetApp.flush();

  return {
    accountName: parsedAccounts.length === 1 ? parsedAccounts[0].accountName : parsedAccounts.length + ' accounts',
    importedItems: totalImportedItems,
    sheetName: updatedSheets.join(', '),
    updatedRows: totalUpdatedRows,
    insertedRows: totalInsertedRows
  };
}

function prepareActiveSheetForImports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === MASTER_SHEET_NAME) {
    initializeMasterSheet_(sheet);
    applyProtections_(sheet, true);
  } else {
    initializeAccountSheet_(sheet);
    applyProtections_(sheet, false);
  }
  SpreadsheetApp.getActive().toast('Active sheet prepared.', APP_TITLE, 5);
}

function reapplyProtectionsForActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const isMaster = sheet.getName() === MASTER_SHEET_NAME;
  applyProtections_(sheet, isMaster);
  relockRequestedByCells_(sheet);
  SpreadsheetApp.getActive().toast('Protections reapplied.', APP_TITLE, 5);
}

function rebuildMasterSheet() {
  rebuildMasterSheet_();
  activateMasterSheet_();
  SpreadsheetApp.getActive().toast('Master rebuilt.', APP_TITLE, 5);
}

function getOrCreateAccountSheet_(accountName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanName = sanitizeSheetName_(accountName);
  let sheet = ss.getSheetByName(cleanName);
  if (!sheet) {
    sheet = ss.insertSheet(cleanName);
  }
  return sheet;
}

function getOrCreateMasterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_SHEET_NAME, 1);
  }
  return sheet;
}

function activateMasterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = getOrCreateMasterSheet_();
  ss.setActiveSheet(master);
  ss.moveActiveSheet(1);
}

function sanitizeSheetName_(name) {
  let clean = String(name || '').trim() || 'Unknown Account';
  clean = clean.replace(/[\\\/\?\*\[\]\:]/g, '_');
  if (clean.length > 99) clean = clean.slice(0, 99);
  return clean;
}

function initializeAccountSheet_(sheet) {
  initializeSheetCommon_(sheet, HEADER_ROW);
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('@');
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('E:E').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F:F').setNumberFormat('0');
  sheet.getRange('G:G').setNumberFormat('@');
  sheet.getRange('H:H').setNumberFormat('@');

  sheet.setColumnWidths(1, 1, 280);
  sheet.setColumnWidths(2, 1, 130);
  sheet.setColumnWidths(3, 1, 110);
  sheet.setColumnWidths(4, 1, 150);
  sheet.setColumnWidths(5, 1, 130);
  sheet.setColumnWidths(6, 1, 80);
  sheet.setColumnWidths(7, 1, 160);
  safeHideColumn_(sheet, COL_SYNC_KEY);
}

function initializeMasterSheet_(sheet) {
  initializeSheetCommon_(sheet, MASTER_HEADER_ROW);
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('@');
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('E:E').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F:F').setNumberFormat('0');
  sheet.getRange('G:G').setNumberFormat('@');
  sheet.getRange('H:H').setNumberFormat('@');
  sheet.getRange('I:I').setNumberFormat('@');

  sheet.setColumnWidths(1, 1, 280);
  sheet.setColumnWidths(2, 1, 130);
  sheet.setColumnWidths(3, 1, 110);
  sheet.setColumnWidths(4, 1, 150);
  sheet.setColumnWidths(5, 1, 130);
  sheet.setColumnWidths(6, 1, 80);
  sheet.setColumnWidths(7, 1, 160);
  sheet.setColumnWidths(8, 1, 180);
  safeHideColumn_(sheet, MASTER_COL_SYNC_KEY);
}

function initializeSheetCommon_(sheet, headerRow) {
  sheet.clearConditionalFormatRules();

  if (sheet.getMaxRows() < 2) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }
  if (sheet.getMaxColumns() < headerRow.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headerRow.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

  const filterRows = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(1, 1, filterRows, headerRow.length).createFilter();
}

function readExistingRows_(sheet, isMaster) {
  const headers = isMaster ? MASTER_HEADER_ROW : HEADER_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  return values
    .filter(row => row[0])
    .map(row => ({
      name: String(row[0]).trim(),
      itemType: String(row[1] || '').trim(),
      quality: String(row[2] || '').trim(),
      trait: String(row[3] || '').trim(),
      dateAdded: row[4] instanceof Date ? row[4] : null,
      count: Number(row[5]) || 0,
      requestedBy: String(row[6] || '').trim(),
      sourceSheet: isMaster ? String(row[7] || '').trim() : sheet.getName(),
      syncKey: String(row[isMaster ? 8 : 7] || '').trim()
    }));
}

function mergeImportedRows_(existingRows, importedItems) {
  const rowMap = new Map();
  let updatedRows = 0;
  let insertedRows = 0;

  existingRows.forEach(row => {
    rowMap.set(buildRowKey_(row.name, row.itemType, row.quality, row.trait), {
      name: row.name,
      itemType: row.itemType,
      quality: row.quality,
      trait: row.trait,
      dateAdded: row.dateAdded,
      count: row.count,
      requestedBy: row.requestedBy,
      syncKey: row.syncKey || buildSyncKey_(row.name, row.itemType, row.quality, row.trait)
    });
  });

  importedItems.forEach(item => {
    const key = buildRowKey_(item.name, item.itemTypeName, item.qualityName, item.trait);

    if (rowMap.has(key)) {
      const existing = rowMap.get(key);
      existing.count = item.count;
      existing.itemType = item.itemTypeName;
      if (!existing.dateAdded) {
        existing.dateAdded = unixToDate_(item.firstDumpedAt);
      }
      existing.syncKey = existing.syncKey || buildSyncKey_(item.name, item.itemTypeName, item.qualityName, item.trait);
      updatedRows += 1;
    } else {
      rowMap.set(key, {
        name: item.name,
        itemType: item.itemTypeName,
        quality: item.qualityName,
        trait: item.trait,
        dateAdded: unixToDate_(item.firstDumpedAt),
        count: item.count,
        requestedBy: '',
        syncKey: buildSyncKey_(item.name, item.itemTypeName, item.qualityName, item.trait)
      });
      insertedRows += 1;
    }
  });

  const rows = Array.from(rowMap.values()).sort((a, b) => {
    return compareText_(a.itemType, b.itemType) ||
      compareText_(a.name, b.name) ||
      compareText_(a.quality, b.quality) ||
      compareText_(a.trait, b.trait);
  });

  return { rows, updatedRows, insertedRows };
}

function buildRowKey_(name, itemType, quality, trait) {
  return [
    String(name || '').trim().toLowerCase(),
    String(itemType || '').trim().toLowerCase(),
    String(quality || '').trim().toLowerCase(),
    String(trait || '').trim().toLowerCase()
  ].join('\u001f');
}

function buildSyncKey_(name, itemType, quality, trait) {
  return [name || '', itemType || '', quality || '', trait || ''].join('||');
}

function compareText_(a, b) {
  return String(a || '').localeCompare(String(b || ''), undefined, { sensitivity: 'base' });
}

function writeRows_(sheet, rows, isMaster) {
  const headers = isMaster ? MASTER_HEADER_ROW : HEADER_ROW;
  const dataRowCount = Math.max(sheet.getLastRow() - 1, 0);

  if (dataRowCount > 0) {
    sheet.getRange(2, 1, dataRowCount, headers.length).clearContent();
  }

  if (rows.length === 0) return;

  const totalRowsNeeded = rows.length + 1;
  if (sheet.getMaxRows() < totalRowsNeeded) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRowsNeeded - sheet.getMaxRows());
  }

  const values = rows.map(row => isMaster ? [
    row.name,
    row.itemType,
    row.quality,
    row.trait,
    row.dateAdded || '',
    row.count,
    row.requestedBy || '',
    row.sourceSheet || '',
    row.syncKey || ''
  ] : [
    row.name,
    row.itemType,
    row.quality,
    row.trait,
    row.dateAdded || '',
    row.count,
    row.requestedBy || '',
    row.syncKey || ''
  ]);

  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function sortAccountSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;

  sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).sort([
    { column: COL_ITEM_TYPE, ascending: true },
    { column: COL_NAME, ascending: true },
    { column: COL_QUALITY, ascending: true },
    { column: COL_TRAIT, ascending: true }
  ]);
}

function sortMasterSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;

  sheet.getRange(2, 1, lastRow - 1, MASTER_HEADER_ROW.length).sort([
    { column: COL_ITEM_TYPE, ascending: true },
    { column: COL_NAME, ascending: true },
    { column: COL_QUALITY, ascending: true },
    { column: COL_TRAIT, ascending: true },
    { column: MASTER_COL_SOURCE_SHEET, ascending: true }
  ]);
}

function rebuildMasterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = getOrCreateMasterSheet_();
  initializeMasterSheet_(master);

  const rows = [];
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === MASTER_SHEET_NAME) return;

    const existing = readExistingRows_(sheet, false);
    existing.forEach(row => {
      rows.push({
        name: row.name,
        itemType: row.itemType,
        quality: row.quality,
        trait: row.trait,
        dateAdded: row.dateAdded,
        count: row.count,
        requestedBy: row.requestedBy,
        sourceSheet: name,
        syncKey: row.syncKey || buildSyncKey_(row.name, row.itemType, row.quality, row.trait)
      });
    });
  });

  clearProtections_(master);
  writeRows_(master, rows, true);
  sortMasterSheet_(master);
  applyProtections_(master, true);
  relockRequestedByCells_(master);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(master);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);
}

function unixToDate_(epochSeconds) {
  const n = Number(epochSeconds);
  if (!n) return '';
  return new Date(n * 1000);
}

function parseAllSavedVariablesAccounts_(text) {
  const defaultIndex = text.indexOf('["Default"]');
  if (defaultIndex === -1) {
    throw new Error('Could not find the Default saved variables section.');
  }

  const defaultOpenBrace = text.indexOf('{', defaultIndex);
  if (defaultOpenBrace === -1) {
    throw new Error('Could not parse the Default saved variables section.');
  }

  const defaultCloseBrace = findMatchingBrace_(text, defaultOpenBrace);
  const defaultBlock = text.slice(defaultOpenBrace + 1, defaultCloseBrace);
  const accountEntries = parseTopLevelEntries_(defaultBlock);

  const accounts = [];
  accountEntries.forEach(entry => {
    if (!/^@/.test(entry.key)) return;

    const accountBlock = entry.tableText;
    const accountWideText = extractNamedTable_(accountBlock, '$AccountWide');
    const tableName = accountWideText.indexOf('["sharedItems"]') >= 0 ? 'sharedItems' : 'dumpedItems';
    const itemsText = extractNamedTable_(accountWideText, tableName);
    const itemEntries = parseTopLevelEntries_(itemsText);
    const items = itemEntries.map(itemEntry => parseDumpedItemEntry_(itemEntry.tableText)).filter(Boolean);

    accounts.push({
      accountName: entry.key,
      items: items
    });
  });

  return accounts;
}

function extractNamedTable_(text, tableName) {
  const marker = `["${tableName}"]`;
  const markerIndex = text.indexOf(marker);
  if (markerIndex === -1) {
    throw new Error(`Could not find the ${tableName} table in the uploaded file.`);
  }

  const braceStart = text.indexOf('{', markerIndex);
  if (braceStart === -1) {
    throw new Error(`Could not parse the ${tableName} table.`);
  }

  const braceEnd = findMatchingBrace_(text, braceStart);
  return text.slice(braceStart + 1, braceEnd);
}

function parseTopLevelEntries_(blockText) {
  const entries = [];
  let i = 0;

  while (i < blockText.length) {
    i = skipWhitespaceAndCommas_(blockText, i);
    if (i >= blockText.length) break;

    if (blockText[i] !== '[' || blockText[i + 1] !== '"') {
      i += 1;
      continue;
    }

    const keyStart = i + 2;
    const keyEnd = findStringTerminator_(blockText, keyStart);
    const key = unescapeLuaString_(blockText.slice(keyStart, keyEnd));

    i = keyEnd + 2;
    i = skipWhitespaceAndCommas_(blockText, i);

    if (blockText[i] !== '=') {
      throw new Error(`Malformed entry near key ${key}.`);
    }

    i += 1;
    i = skipWhitespaceAndCommas_(blockText, i);

    if (blockText[i] !== '{') {
      throw new Error(`Expected a table value for key ${key}.`);
    }

    const tableStart = i;
    const tableEnd = findMatchingBrace_(blockText, tableStart);
    const tableText = blockText.slice(tableStart + 1, tableEnd);

    entries.push({ key, tableText });
    i = tableEnd + 1;
  }

  return entries;
}

function parseDumpedItemEntry_(tableText) {
  const itemName = parseLuaStringField_(tableText, 'itemName');
  if (!itemName) return null;

  return {
    name: itemName,
    accountName: parseLuaStringField_(tableText, 'accountName') || '',
    itemType: parseLuaNumberField_(tableText, 'itemType') || 0,
    itemTypeName: parseLuaStringField_(tableText, 'itemTypeName') || String(parseLuaNumberField_(tableText, 'itemType') || ''),
    quality: parseLuaNumberField_(tableText, 'quality') || 0,
    qualityName: parseLuaStringField_(tableText, 'qualityName') || String(parseLuaNumberField_(tableText, 'quality') || ''),
    trait: parseLuaStringField_(tableText, 'trait') || '',
    count: parseLuaNumberField_(tableText, 'count') || 0,
    firstDumpedAt: parseLuaNumberField_(tableText, 'firstDumpedAt') || 0,
    lastDumpedAt: parseLuaNumberField_(tableText, 'lastDumpedAt') || 0
  };
}

function parseLuaStringField_(text, fieldName) {
  const re = new RegExp(`\\["${escapeRegExp_(fieldName)}"\\]\\s*=\\s*"((?:\\\\.|[^"\\\\])*)"`);
  const match = re.exec(text);
  return match ? unescapeLuaString_(match[1]) : '';
}

function parseLuaNumberField_(text, fieldName) {
  const re = new RegExp(`\\["${escapeRegExp_(fieldName)}"\\]\\s*=\\s*(-?\\d+(?:\\.\\d+)?)`);
  const match = re.exec(text);
  return match ? Number(match[1]) : 0;
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function unescapeLuaString_(value) {
  return String(value)
    .replace(/\\\\/g, '\\')
    .replace(/\\"/g, '"')
    .replace(/\\n/g, '\n')
    .replace(/\\r/g, '\r')
    .replace(/\\t/g, '\t');
}

function skipWhitespaceAndCommas_(text, index) {
  while (index < text.length && /[\s,]/.test(text[index])) {
    index += 1;
  }
  return index;
}

function findStringTerminator_(text, startIndex) {
  let i = startIndex;
  while (i < text.length) {
    if (text[i] === '"' && text[i - 1] !== '\\') {
      return i;
    }
    i += 1;
  }
  throw new Error('Unterminated string in uploaded file.');
}

function findMatchingBrace_(text, openIndex) {
  let depth = 0;
  let inString = false;

  for (let i = openIndex; i < text.length; i += 1) {
    const ch = text[i];
    const prev = i > 0 ? text[i - 1] : '';

    if (ch === '"' && prev !== '\\') {
      inString = !inString;
      continue;
    }

    if (inString) continue;

    if (ch === '{') depth += 1;
    if (ch === '}') {
      depth -= 1;
      if (depth === 0) return i;
    }
  }

  throw new Error('Could not match braces in uploaded file.');
}

function clearProtections_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(protection => {
    try {
      protection.remove();
    } catch (err) {}
  });
}

function applyProtections_(sheet, isMaster) {
  clearProtections_(sheet);

  const lastRow = Math.max(sheet.getMaxRows(), 2);
  const protectedColumns = 6;
  const range = sheet.getRange(1, 1, lastRow, protectedColumns);
  const protection = range.protect();
  protection.setDescription('Protected item data columns');

  const me = Session.getEffectiveUser();
  try {
    protection.addEditor(me);
    const editors = protection.getEditors();
    if (editors && editors.length) {
      protection.removeEditors(editors.filter(editor => editor.getEmail() !== me.getEmail()));
    }
  } catch (err) {}

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function relockRequestedByCells_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, COL_REQUESTED_BY, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() !== '') {
      protectRequestedByCell_(sheet.getRange(i + 2, COL_REQUESTED_BY));
    }
  }
}

function protectRequestedByCell_(range) {
  const sheet = range.getSheet();
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  for (const protection of protections) {
    try {
      const pRange = protection.getRange();
      if (
        pRange.getRow() === range.getRow() &&
        pRange.getColumn() === range.getColumn() &&
        pRange.getNumRows() === 1 &&
        pRange.getNumColumns() === 1
      ) {
        return;
      }
    } catch (err) {}
  }

  const protection = range.protect();
  protection.setDescription('Locked RequestedBy cell after entry');

  const me = Session.getEffectiveUser();
  try {
    protection.addEditor(me);
    const editors = protection.getEditors();
    if (editors && editors.length) {
      protection.removeEditors(editors.filter(editor => editor.getEmail() !== me.getEmail()));
    }
  } catch (err) {}

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }

  range.setBackground('#fce5cd');
  range.setNote('Locked after RequestedBy entry');
}

function clearCellProtection_(range) {
  const protections = range.getSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    try {
      const r = p.getRange();
      if (
        r.getRow() === range.getRow() &&
        r.getColumn() === range.getColumn() &&
        r.getNumRows() === 1 &&
        r.getNumColumns() === 1
      ) {
        p.remove();
      }
    } catch (err) {}
  });
}

function clearRequestedByVisuals_(range) {
  range.setBackground(null);
  range.clearNote();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (range.getColumn() !== COL_REQUESTED_BY || range.getRow() === 1) return;

  const value = String(range.getValue() || '').trim();

  // Sync even when blank so clearing propagates.
  syncRequestedByFromEdit_(sheet, range.getRow(), value);
}

function lockRequestedByOnEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (range.getColumn() !== COL_REQUESTED_BY || range.getRow() === 1) return;

  const value = String(range.getValue() || '').trim();

  // Sync even when blank so clearing propagates.
  syncRequestedByFromEdit_(sheet, range.getRow(), value);

  if (value !== '') {
    protectRequestedByCell_(sheet.getRange(range.getRow(), COL_REQUESTED_BY));
  } else {
    clearCellProtection_(sheet.getRange(range.getRow(), COL_REQUESTED_BY));
    clearRequestedByVisuals_(sheet.getRange(range.getRow(), COL_REQUESTED_BY));
  }
}

function rebuildMasterOnChange(e) {
  rebuildMasterSheet_();
  activateMasterSheet_();
}

function syncRequestedByFromEdit_(sheet, rowIndex, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isMaster = sheet.getName() === MASTER_SHEET_NAME;

  if (isMaster) {
    const sourceSheetName = String(sheet.getRange(rowIndex, MASTER_COL_SOURCE_SHEET).getValue() || '').trim();
    const syncKey = String(sheet.getRange(rowIndex, MASTER_COL_SYNC_KEY).getValue() || '').trim();
    if (!sourceSheetName || !syncKey) return;

    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) return;

    const sourceRow = findRowBySyncKey_(sourceSheet, syncKey, false);
    if (sourceRow > 0) {
      const targetRange = sourceSheet.getRange(sourceRow, COL_REQUESTED_BY);
      const currentValue = String(targetRange.getValue() || '').trim();

      if (currentValue !== value) {
        clearCellProtection_(targetRange);
        targetRange.setValue(value);

        if (value !== '') {
          protectRequestedByCell_(targetRange);
        } else {
          clearRequestedByVisuals_(targetRange);
        }
      }
    }
  } else {
    const syncKey = String(sheet.getRange(rowIndex, COL_SYNC_KEY).getValue() || '').trim();
    if (!syncKey) return;

    const master = getOrCreateMasterSheet_();
    const masterRow = findRowBySyncKey_(master, syncKey, true);
    if (masterRow > 0) {
      const targetRange = master.getRange(masterRow, COL_REQUESTED_BY);
      const currentValue = String(targetRange.getValue() || '').trim();

      if (currentValue !== value) {
        clearCellProtection_(targetRange);
        targetRange.setValue(value);

        if (value !== '') {
          protectRequestedByCell_(targetRange);
        } else {
          clearRequestedByVisuals_(targetRange);
        }
      }
    }
  }
}

function findRowBySyncKey_(sheet, syncKey, isMaster) {
  const col = isMaster ? MASTER_COL_SYNC_KEY : COL_SYNC_KEY;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === syncKey) {
      return i + 2;
    }
  }
  return -1;
}

function deleteSelectedItemSafely() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveRange() ? sheet.getActiveRange().getRow() : 0;

  if (row <= 1) {
    throw new Error('Select a data row first.');
  }

  if (sheet.getName() === MASTER_SHEET_NAME) {
    const sourceSheetName = String(sheet.getRange(row, MASTER_COL_SOURCE_SHEET).getValue() || '').trim();
    const syncKey = String(sheet.getRange(row, MASTER_COL_SYNC_KEY).getValue() || '').trim();

    if (sourceSheetName && syncKey) {
      const sourceSheet = ss.getSheetByName(sourceSheetName);
      if (sourceSheet) {
        const sourceRow = findRowBySyncKey_(sourceSheet, syncKey, false);
        if (sourceRow > 0) {
          sourceSheet.deleteRow(sourceRow);
        }
      }
    }

    sheet.deleteRow(row);
  } else {
    sheet.deleteRow(row);
  }

  rebuildMasterSheet_();
  activateMasterSheet_();
}

function safeHideColumn_(sheet, column) {
  try {
    sheet.hideColumns(column);
  } catch (err) {}
}
