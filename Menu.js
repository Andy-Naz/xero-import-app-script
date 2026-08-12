function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MySky')
    .addItem(
      'Find Row',
      'findLastImportedTransaction'
    )
    .addSeparator()
    .addItem(
      'Prepare EUR',
      'prepareEURTemplate'
    )
    .addItem(
      'Update Payees',
      'updateEURPayees'
    )
    .addToUi();
}