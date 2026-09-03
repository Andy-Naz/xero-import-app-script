/**
 * Получает обязательные рабочие листы.
 */
function getRequiredSheets_(
  spreadsheet,
  templateKey
) {
  validateTemplateKey_(templateKey);

  return {
    template: getSheetOrThrow_(
      spreadsheet,
      templateKey
    ),

    history: getSheetOrThrow_(
      spreadsheet,
      `${templateKey}_XERO`
    ),

    transactions: getSheetOrThrow_(
      spreadsheet,
      `${templateKey}_transactions`
    ),

    contacts: getSheetOrThrow_(
      spreadsheet,
      CONFIG.SHEETS.CONTACTS
    )
  };
}


/**
 * Возвращает лист или выбрасывает понятную ошибку.
 */
function getSheetOrThrow_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Лист "${sheetName}" не найден.`);
  }

  return sheet;
}


/**
 * Читает текущие данные рабочего шаблона EUR.
 */
function readTemplateData_(templateSheet, rowCount) {
  return templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      1,
      rowCount,
      CONFIG.TEMPLATE.COLUMN_COUNT
    )
    .getValues();
}


/**
 * Добавляет текущие данные EUR в начало EUR_XERO.
 */
function archiveTemplateData_(
  templateSheet,
  historySheet,
  templateData
) {
  const rowCount = templateData.length;

  if (rowCount === 0) {
    return;
  }

  historySheet.insertRowsAfter(
    CONFIG.TEMPLATE.HEADER_ROW,
    rowCount
  );

  historySheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      1,
      rowCount,
      CONFIG.TEMPLATE.COLUMN_COUNT
    )
    .setValues(templateData);

  templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      1,
      rowCount,
      CONFIG.TEMPLATE.COLUMN_COUNT
    )
    .copyFormatToRange(
      historySheet,
      1,
      CONFIG.TEMPLATE.COLUMN_COUNT,
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      CONFIG.TEMPLATE.FIRST_DATA_ROW + rowCount - 1
    );
}


/**
 * Очищает старый EUR и записывает новые операции.
 */
function replaceTemplateData_(
  templateSheet,
  newTemplateData,
  previousRowCount
) {
  const newRowCount = newTemplateData.length;

  const rowsToClear = Math.max(
    previousRowCount,
    newRowCount
  );

  templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      1,
      rowsToClear,
      CONFIG.TEMPLATE.COLUMN_COUNT
    )
    .clearContent();

  if (newRowCount === 0) {
    return;
  }

  templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      1,
      newRowCount,
      CONFIG.TEMPLATE.COLUMN_COUNT
    )
    .setValues(newTemplateData);

  applyTemplateFormats_(templateSheet, newRowCount);
}


/**
 * Устанавливает форматы даты и суммы.
 */
function applyTemplateFormats_(templateSheet, rowCount) {
  templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      CONFIG.TEMPLATE.COLUMNS.DATE,
      rowCount,
      1
    )
    .setNumberFormat('dd/mm/yyyy');

  templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      CONFIG.TEMPLATE.COLUMNS.AMOUNT,
      rowCount,
      1
    )
    .setNumberFormat('0.00');
}


/**
 * Считает непрерывно заполненные строки
 * в выбранной колонке.
 */
function getContiguousDataRowCount_(
  sheet,
  firstDataRow,
  columnNumber
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < firstDataRow) {
    return 0;
  }

  const values = sheet
    .getRange(
      firstDataRow,
      columnNumber,
      lastRow - firstDataRow + 1,
      1
    )
    .getValues();

  let rowCount = 0;

  for (let index = 0; index < values.length; index++) {
    const value = values[index][0];

    if (value === '' || value === null) {
      break;
    }

    rowCount++;
  }

  return rowCount;
}

function getRequiredSheets_(spreadsheet) {
  return {
    template: getSheetOrThrow_(
      spreadsheet,
      CONFIG.SHEETS.TEMPLATE
    ),

    history: getSheetOrThrow_(
      spreadsheet,
      CONFIG.SHEETS.HISTORY
    ),

    transactions: getSheetOrThrow_(
      spreadsheet,
      CONFIG.SHEETS.TRANSACTIONS
    ),

    contacts: getSheetOrThrow_(
      spreadsheet,
      CONFIG.SHEETS.CONTACTS
    )
  };
}