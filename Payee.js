/**
 * Загружает справочник SA Contacts.
 *
 * Колонка A — итоговое имя контакта в Xero.
 * Колонка B — банковское имя для поиска.
 *
 * Возвращает:
 * {
 *   xeroNames: Set,
 *   bankToXero: Map
 * }
 */
function getPayeeContactData_(contactsSheet) {
  const lastRow = contactsSheet.getLastRow();

  const contactData = {
    xeroNames: new Set(),
    bankToXero: new Map()
  };

  if (lastRow < CONFIG.CONTACTS.FIRST_DATA_ROW) {
    return contactData;
  }

  const rowCount =
    lastRow - CONFIG.CONTACTS.FIRST_DATA_ROW + 1;

  const values = contactsSheet
    .getRange(
      CONFIG.CONTACTS.FIRST_DATA_ROW,
      1,
      rowCount,
      CONFIG.CONTACTS.COLUMN_COUNT
    )
    .getValues();

  values.forEach(function(row) {
    const xeroContactName =
      row[CONFIG.CONTACTS.COLUMNS.XERO_NAME - 1];

    const bankPayeeName =
      row[CONFIG.CONTACTS.COLUMNS.BANK_NAME - 1];

    const normalizedXeroName =
      normalizePayeeKey_(xeroContactName);

    const normalizedBankName =
      normalizePayeeKey_(bankPayeeName);

    if (normalizedXeroName !== '') {
      contactData.xeroNames.add(normalizedXeroName);
    }

    if (
      normalizedBankName !== '' &&
      normalizeText_(xeroContactName) !== ''
    ) {
      contactData.bankToXero.set(
        normalizedBankName,
        String(xeroContactName).trim()
      );
    }
  });

  return contactData;
}


/**
 * Обрабатывает текущее значение Payee.
 *
 * 1. Если Payee уже находится в SA Contacts!A:A,
 *    значение считается обработанным и не меняется.
 *
 * 2. Иначе берётся текст до первого ";".
 *
 * 3. Сокращённое значение ищется в SA Contacts!B:B.
 *
 * 4. При совпадении возвращается значение из A:A.
 *
 * 5. При отсутствии совпадения возвращается исходный Payee.
 */
function resolveCurrentPayee_(currentPayee, contactData) {
  const originalValue = normalizeText_(currentPayee);

  if (originalValue === '') {
    return '';
  }

  const normalizedOriginal =
    normalizePayeeKey_(originalValue);

  /*
   * Значение уже является именем контакта Xero.
   * Повторно его не обрабатываем.
   */
  if (contactData.xeroNames.has(normalizedOriginal)) {
    return originalValue;
  }

  const bankName =
    getTextBeforeFirstSemicolon_(originalValue);

  const lookupKey =
    normalizePayeeKey_(bankName);

  if (contactData.bankToXero.has(lookupKey)) {
    return contactData.bankToXero.get(lookupKey);
  }

  return originalValue;
}


/**
 * Возвращает текст до первого ";".
 *
 * Если ";" нет, возвращает весь текст.
 */
function getTextBeforeFirstSemicolon_(value) {
  const text = normalizeText_(value);
  const semicolonIndex = text.indexOf(';');

  if (semicolonIndex === -1) {
    return text;
  }

  return text
    .substring(0, semicolonIndex)
    .trim();
}


/**
 * Нормализует значение для поиска:
 * - убирает лишние пробелы;
 * - игнорирует регистр.
 */
function normalizePayeeKey_(value) {
  return normalizeText_(value).toUpperCase();
}