function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MySky')
    .addItem(
      'Find Row',
      'findLastImportedTransaction'
    )
    .addSeparator()
    .addItem(
      'Prepare Template',
      'showPrepareTemplateDialog'
    )
    .addItem(
      'Update Payees',
      'showUpdatePayeesDialog'
    )
    .addToUi();
}