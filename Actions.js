/**
 * Находит строку банковской выписки,
 * соответствующую первой операции на листе EUR.
 *
 * Вызывается из меню:
 * MySky → Find Row
 */
function findLastImportedTransaction() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const sheets = getRequiredSheets_(spreadsheet);

    const matchedRow = findMatchingTransactionRow_(
      spreadsheet,
      sheets.template,
      sheets.transactions
    );

    if (matchedRow === null) {
      ui.alert(
        'Совпадение не найдено',
        'Строка 2 листа EUR не найдена в банковской выписке.',
        ui.ButtonSet.OK
      );

      return null;
    }

    ui.alert(
      'Операция найдена',
      `Совпавшая строка на листе EUR_transactions: ${matchedRow}`,
      ui.ButtonSet.OK
    );

    return matchedRow;
  } catch (error) {
    showError_(error);
    return null;
  }
}


/**
 * Архивирует текущие данные EUR в EUR_XERO
 * и переносит новые операции из EUR_transactions в EUR.
 *
 * Payee на этом этапе переносится без гармонизации:
 * EUR_transactions Description1 → EUR Payee.
 *
 * Гармонизация Payee запускается отдельно:
 * MySky → Update Payees
 *
 * Вызывается из меню:
 * MySky → Prepare EUR
 */
function prepareEURTemplate() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const sheets = getRequiredSheets_(spreadsheet);

    /*
     * 1. Ищем в банковской выписке операцию,
     * соответствующую первой строке текущего EUR.
     */
    const matchedTransactionRow = findMatchingTransactionRow_(
      spreadsheet,
      sheets.template,
      sheets.transactions
    );

    if (matchedTransactionRow === null) {
      ui.alert(
        'Операция не найдена',
        'Подготовка остановлена: EUR!2 не найдена в банковской выписке.',
        ui.ButtonSet.OK
      );

      return;
    }

    /*
     * Новые банковские операции находятся между:
     *
     * EUR_transactions!11
     * и строкой перед найденной операцией.
     */
    const newTransactionsCount =
      matchedTransactionRow -
      CONFIG.TRANSACTIONS.FIRST_DATA_ROW;

    if (newTransactionsCount <= 0) {
      ui.alert(
        'Новых операций нет',
        `Найденная операция находится в строке ${matchedTransactionRow}. ` +
        'Выше неё в выписке нет новых операций.',
        ui.ButtonSet.OK
      );

      return;
    }

    /*
     * 2. Считаем количество текущих строк EUR
     * по непрерывно заполненному столбцу A.
     */
    const currentTemplateRowCount =
      getContiguousDataRowCount_(
        sheets.template,
        CONFIG.TEMPLATE.FIRST_DATA_ROW,
        CONFIG.TEMPLATE.COLUMNS.DATE
      );

    if (currentTemplateRowCount === 0) {
      throw new Error(
        'На листе EUR нет данных для переноса в историю.'
      );
    }

    /*
     * 3. Читаем текущие данные EUR.
     */
    const currentTemplateData = readTemplateData_(
      sheets.template,
      currentTemplateRowCount
    );

    /*
     * 4. Читаем новые операции банковской выписки.
     */
    const newTransactionRows = readNewTransactionRows_(
      sheets.transactions,
      newTransactionsCount
    );

    /*
     * 5. Преобразуем строки UBS в формат Xero.
     *
     * Value date → *Date
     * Debit + Credit → *Amount
     * Description1 → Payee
     * Description3 → Description
     */
    const newTemplateData =
      convertTransactionsToTemplate_(
        newTransactionRows
      );

    /*
     * 6. Добавляем текущий EUR в начало EUR_XERO.
     */
    archiveTemplateData_(
      sheets.template,
      sheets.history,
      currentTemplateData
    );

    /*
     * 7. Очищаем текущий EUR и записываем новые операции.
     */
    replaceTemplateData_(
      sheets.template,
      newTemplateData,
      currentTemplateRowCount
    );

    spreadsheet.toast(
      `${currentTemplateRowCount} строк перенесено в EUR_XERO; ` +
      `${newTemplateData.length} строк загружено в EUR.`,
      'MySky',
      8
    );

    ui.alert(
      'Шаблон подготовлен',
      `Найденная строка выписки: ${matchedTransactionRow}\n\n` +
      `Перенесено в EUR_XERO: ${currentTemplateRowCount} строк\n` +
      `Загружено в EUR: ${newTemplateData.length} строк`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    showError_(error);
  }
}


/**
 * Обрабатывает текущие Payee на листе EUR
 * по актуальному справочнику SA Contacts.
 *
 * Логика:
 *
 * 1. Проверяем текущее значение EUR!C:C
 *    в SA Contacts!A:A.
 *
 * 2. Если значение найдено в A:A,
 *    оно уже гармонизировано и остаётся без изменений.
 *
 * 3. Если значение не найдено, берём текст
 *    до первого символа ";".
 *
 * 4. Ищем сокращённое значение в SA Contacts!B:B.
 *
 * 5. При совпадении записываем значение из SA Contacts!A:A.
 *
 * 6. При отсутствии совпадения сохраняем исходный Payee.
 *
 * Функцию можно безопасно запускать несколько раз.
 *
 * Вызывается из меню:
 * MySky → Update Payees
 */
function updateEURPayees() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const sheets = getRequiredSheets_(spreadsheet);

    /*
     * Считаем количество операций на EUR
     * по заполненным значениям в столбце A.
     */
    const rowCount = getContiguousDataRowCount_(
      sheets.template,
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      CONFIG.TEMPLATE.COLUMNS.DATE
    );

    if (rowCount === 0) {
      ui.alert(
        'Нет данных',
        'На листе EUR нет операций для обработки.',
        ui.ButtonSet.OK
      );

      return;
    }

    /*
     * Читаем текущие Payee из EUR!C2:C.
     */
    const payeeRange = sheets.template.getRange(
      CONFIG.TEMPLATE.FIRST_DATA_ROW,
      CONFIG.TEMPLATE.COLUMNS.PAYEE,
      rowCount,
      1
    );

    const currentPayees = payeeRange.getValues();

    /*
     * Загружаем актуальный справочник:
     *
     * xeroNames — значения SA Contacts!A:A;
     * bankToXero — соответствия B → A.
     */
    const contactData = getPayeeContactData_(
      sheets.contacts
    );

    let alreadyProcessedCount = 0;
    let replacedCount = 0;
    let notFoundCount = 0;
    let emptyCount = 0;

    /*
     * Обрабатываем каждое текущее значение Payee.
     */
    const updatedPayees = currentPayees.map(
      function(row) {
        const currentPayee = normalizeText_(row[0]);

        if (currentPayee === '') {
          emptyCount++;
          return [''];
        }

        const normalizedCurrentPayee =
          normalizePayeeKey_(currentPayee);

        /*
         * Текущее значение уже является
         * гармонизированным именем Xero.
         */
        if (
          contactData.xeroNames.has(
            normalizedCurrentPayee
          )
        ) {
          alreadyProcessedCount++;

          return [currentPayee];
        }

        /*
         * Значение пока не гармонизировано.
         * Обрезаем по ";" и ищем в SA Contacts!B:B.
         */
        const resolvedPayee = resolveCurrentPayee_(
          currentPayee,
          contactData
        );

        if (
          normalizePayeeKey_(resolvedPayee) !==
          normalizedCurrentPayee
        ) {
          replacedCount++;
        } else {
          notFoundCount++;
        }

        return [resolvedPayee];
      }
    );

    /*
     * Одной операцией обновляем EUR!C2:C.
     */
    payeeRange.setValues(updatedPayees);

    spreadsheet.toast(
      `Заменено: ${replacedCount}; ` +
      `уже обработано: ${alreadyProcessedCount}; ` +
      `не найдено: ${notFoundCount}.`,
      'MySky',
      8
    );

    ui.alert(
      'Payee обработаны',
      `Всего строк: ${rowCount}\n\n` +
      `Заменено по справочнику: ${replacedCount}\n` +
      `Уже было обработано: ${alreadyProcessedCount}\n` +
      `Не найдено в справочнике: ${notFoundCount}\n` +
      `Пустых Payee: ${emptyCount}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    showError_(error);
  }
}