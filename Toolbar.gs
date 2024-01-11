function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Form Duplicator')
    .addItem('Duplicate Form', 'duplicateForm')
    .addToUi();
}
