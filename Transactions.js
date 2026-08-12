/**
 * Ищет в EUR_transactions строку,
 * соответствующую первой строке данных на EUR.
 */
function findMatchingTransactionRow_(
  spreadsheet,
  templateSheet,
  transactionsSheet
) {
  const templateValues = templateSheet
    .getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      1,
      1,
      CONFIG.TEMPLATE.COLUMN_COUNT
    )
    .getValues()[0];

  const templateDate =
    templateValues[CONFIG.TEMPLATE.COLUMNS.DATE - 1];

  const templateAmount =
    templateValues[CONFIG.TEMPLATE.COLUMNS.AMOUNT - 1];

  const templateDescription =
    templateValues[CONFIG.TEMPLATE.COLUMNS.DESCRIPTION - 1];

  validateTemplateValues_(
    templateDate,
    templateAmount,
    templateDescription
  );

  const lastTransactionRow = transactionsSheet.getLastRow();

  if (
    lastTransactionRow <
    CONFIG.TRANSACTIONS.FIRST_DATA_ROW
  ) {
    return null;
  }

  const transactionRows = transactionsSheet
    .getRange(
      CONFIG.TRANSACTIONS.FIRST_DATA_ROW,
      1,
      lastTransactionRow -
        CONFIG.TRANSACTIONS.FIRST_DATA_ROW +
        1,
      CONFIG.TRANSACTIONS.COLUMN_COUNT
    )
    .getValues();

  const timeZone = spreadsheet.getSpreadsheetTimeZone();

  for (let index = 0; index < transactionRows.length; index++) {
    const row = transactionRows[index];

    const valueDate =
      row[CONFIG.TRANSACTIONS.COLUMNS.VALUE_DATE - 1];

    const debit = toNumber_(
      row[CONFIG.TRANSACTIONS.COLUMNS.DEBIT - 1]
    );

    const credit = toNumber_(
      row[CONFIG.TRANSACTIONS.COLUMNS.CREDIT - 1]
    );

    const description3 =
      row[CONFIG.TRANSACTIONS.COLUMNS.DESCRIPTION_3 - 1];

    const transactionAmount = debit + credit;

    const isMatch =
      datesAreEqual_(templateDate, valueDate, timeZone) &&
      numbersAreEqual_(templateAmount, transactionAmount) &&
      normalizeText_(templateDescription) ===
        normalizeText_(description3);

    if (isMatch) {
      return CONFIG.TRANSACTIONS.FIRST_DATA_ROW + index;
    }
  }

  return null;
}


/**
 * Читает новые строки выписки.
 */
function readNewTransactionRows_(
  transactionsSheet,
  rowCount
) {
  return transactionsSheet
    .getRange(
      CONFIG.TRANSACTIONS.FIRST_DATA_ROW,
      1,
      rowCount,
      CONFIG.TRANSACTIONS.COLUMN_COUNT
    )
    .getValues();
}


/**
 * Преобразует строки UBS в формат шаблона Xero.
 */
function convertTransactionsToTemplate_(transactionRows) {
  return transactionRows.map(function(row) {
    const valueDate =
      row[CONFIG.TRANSACTIONS.COLUMNS.VALUE_DATE - 1];

    const debit = toNumber_(
      row[CONFIG.TRANSACTIONS.COLUMNS.DEBIT - 1]
    );

    const credit = toNumber_(
      row[CONFIG.TRANSACTIONS.COLUMNS.CREDIT - 1]
    );

    const description1 =
      row[CONFIG.TRANSACTIONS.COLUMNS.DESCRIPTION_1 - 1];

    const description3 =
      row[CONFIG.TRANSACTIONS.COLUMNS.DESCRIPTION_3 - 1];

    return [
      valueDate,          // A — *Date
      debit + credit,     // B — *Amount
      description1,       // C — Payee, пока исходный
      description3,       // D — Description
      '',                 // E — Reference
      ''                  // F — Check Number
    ];
  });
}