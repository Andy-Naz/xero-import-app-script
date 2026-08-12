/**
 * Проверяет ключевые значения первой строки EUR.
 */
function validateTemplateValues_(
  date,
  amount,
  description
) {
  const missingFields = [];

  if (!(date instanceof Date) || isNaN(date.getTime())) {
    missingFields.push('EUR!A2 — *Date');
  }

  if (
    amount === '' ||
    amount === null ||
    isNaN(toNumber_(amount))
  ) {
    missingFields.push('EUR!B2 — *Amount');
  }

  if (normalizeText_(description) === '') {
    missingFields.push('EUR!D2 — Description');
  }

  if (missingFields.length > 0) {
    throw new Error(
      'Нельзя выполнить поиск. Проверьте:\n' +
      missingFields.join('\n')
    );
  }
}


/**
 * Сравнивает календарные даты без времени.
 */
function datesAreEqual_(
  firstDate,
  secondDate,
  timeZone
) {
  if (
    !(firstDate instanceof Date) ||
    !(secondDate instanceof Date) ||
    isNaN(firstDate.getTime()) ||
    isNaN(secondDate.getTime())
  ) {
    return false;
  }

  const firstFormatted = Utilities.formatDate(
    firstDate,
    timeZone,
    'yyyy-MM-dd'
  );

  const secondFormatted = Utilities.formatDate(
    secondDate,
    timeZone,
    'yyyy-MM-dd'
  );

  return firstFormatted === secondFormatted;
}


/**
 * Сравнивает денежные суммы с точностью до цента.
 */
function numbersAreEqual_(
  firstNumber,
  secondNumber
) {
  return (
    Math.round(toNumber_(firstNumber) * 100) ===
    Math.round(toNumber_(secondNumber) * 100)
  );
}


/**
 * Преобразует значение в число.
 */
function toNumber_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return 0;
  }

  if (typeof value === 'number') {
    return value;
  }

  const normalizedValue = String(value)
    .replace(/\s/g, '')
    .replace(',', '.');

  const number = Number(normalizedValue);

  return Number.isNaN(number) ? 0 : number;
}


/**
 * Нормализует текст перед сравнением.
 */
function normalizeText_(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Показывает понятную ошибку пользователю
 * и сохраняет полную ошибку в журнале.
 */
function showError_(error) {
  console.error(error);

  SpreadsheetApp.getUi().alert(
    'Ошибка',
    error instanceof Error
      ? error.message
      : String(error),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}