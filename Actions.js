/**
 * Находит строку банковской выписки,
 * соответствующую первой операции на листе EUR.
 *
 * Вызывается из меню:
 * MySky → Find Row
 */
function prepareTemplate(templateKey) {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();

  try {
    const sheets = getRequiredSheets_(
      spreadsheet,
      templateKey
    );

    const matchedTransactionRow =
      findMatchingTransactionRow_(
        spreadsheet,
        sheets.template,
        sheets.transactions
      );

    if (matchedTransactionRow === null) {
      ui.alert(
        'Операция не найдена',
        `Подготовка ${templateKey} остановлена: ` +
        `первая строка листа "${templateKey}" ` +
        'не найдена в банковской выписке.',
        ui.ButtonSet.OK
      );

      return;
    }

    const newTransactionsCount =
      matchedTransactionRow -
      CONFIG.TRANSACTIONS.FIRST_DATA_ROW;

    if (newTransactionsCount <= 0) {
      ui.alert(
        'Новых операций нет',
        `Для ${templateKey} новых операций нет.`,
        ui.ButtonSet.OK
      );

      return;
    }

    const currentTemplateRowCount =
      getContiguousDataRowCount_(
        sheets.template,
        CONFIG.TEMPLATE.FIRST_DATA_ROW,
        CONFIG.TEMPLATE.COLUMNS.DATE
      );

    if (currentTemplateRowCount === 0) {
      throw new Error(
        `На листе "${templateKey}" нет данных ` +
        'для переноса в историю.'
      );
    }

    const currentTemplateData =
      readTemplateData_(
        sheets.template,
        currentTemplateRowCount
      );

    const newTransactionRows =
      readNewTransactionRows_(
        sheets.transactions,
        newTransactionsCount
      );

    const newTemplateData =
      convertTransactionsToTemplate_(
        newTransactionRows
      );

    archiveTemplateData_(
      sheets.template,
      sheets.history,
      currentTemplateData
    );

    replaceTemplateData_(
      sheets.template,
      newTemplateData,
      currentTemplateRowCount
    );

    spreadsheet.toast(
      `${templateKey}: ` +
      `${currentTemplateRowCount} строк архивировано; ` +
      `${newTemplateData.length} новых строк загружено.`,
      'MySky',
      8
    );

    ui.alert(
      `${templateKey} подготовлен`,
      `Найденная строка выписки: ${matchedTransactionRow}\n\n` +
      `Перенесено в ${templateKey}_XERO: ` +
      `${currentTemplateRowCount}\n` +
      `Загружено в ${templateKey}: ` +
      `${newTemplateData.length}`,
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
function updatePayees(templateKey) {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const ui = SpreadsheetApp.getUi();

  try {
    const sheets = getRequiredSheets_(
      spreadsheet,
      templateKey
    );

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
    `${templateKey}: Payee обработаны`,
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