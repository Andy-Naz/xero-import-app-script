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
      'prepareEURTemplate'
    )
    .addItem(
      'Update Payees',
      'updateEURPayees'
    )
    .addToUi();
}