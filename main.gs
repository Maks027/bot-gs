START_COLUMN = 4;
START_ROW    = 1;
COLUMNS      = 4;

function openAddDialog() {
  var html = HtmlService.createTemplateFromFile('add').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function openRemoveDialog() {
  var html = HtmlService.createTemplateFromFile('rm').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function openEditDialog() {
  var html = HtmlService.createTemplateFromFile('edit').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function activeSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("catalog");
}

function extractA1Columns(notation) {
  return notation.split(':').map(s => s.match(/^[A-Z]+/)[0]).join(':')
}

function test() {
  addCategory("test");
  // var sheet = activeSheet();

  // range = sheet.getRange(2).getA1Notation();
  // Logger.log(sheet.getRange(2).getA1Notation())
  // Logger.log(extractA1Columns(range.getA1Notation()));

  // range.setBorder(null, null, null, null, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID);
}

// function setHeaderMetadata() {
//   var range = activeSheet().getRange('1:1');
//   range.addDeveloperMetadata('header');
// }

function addCategory(name) {
  var sheet       = activeSheet();
  var nextColomn  = nextCategoryLocation();

  sheet.insertColumns(nextColomn, COLUMNS);

  var firstHeaderRange  = sheet.getRange(START_ROW, nextColomn, 1, COLUMNS);
  var secondHeaderRange = sheet.getRange(START_ROW + 1, nextColomn, 1, COLUMNS);
  var dataRangeNotation = extractA1Columns(firstHeaderRange.getA1Notation())
  var dataRange         = sheet.getRange(dataRangeNotation);

  firstHeaderRange
    .mergeAcross()
    .setHorizontalAlignment("center")
    .setBackground('#0c343d')
    .setFontColor('#d9d9d9')
    .setFontFamily('Times New Roman')
    .setFontSize(14)
    .setFontStyle('bold')
    .setBorder(true, true, true, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setValue(name)
    .activate();

  secondHeaderRange
    .setFontFamily('Times New Roman')
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID)
    .setValues([['Артикул', 'Наименование', 'Ед. изм.', 'Изображение']]);

    dataRange
    .setBorder(null, null, null, null, true, null, '#666666', SpreadsheetApp.BorderStyle.SOLID)
    .setBorder(null, true, null, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);


  var fc = dataRangeNotation.split(':')[0]
  var firstColumnRange = sheet.getRange(`${fc}:${fc}`);

  firstColumnRange.addDeveloperMetadata('category', name);
}

function removeCategory(name) {
  var sheet       = activeSheet();
  var metadata    = sheet.createDeveloperMetadataFinder().withKey('category').withValue(name).find()[0];
  var firstColumn = metadata.getLocation().getColumn().getColumn();

  sheet.deleteColumns(firstColumn, COLUMNS);
  metadata.remove();
}

function getCategoriesMetadata() {
  return activeSheet().createDeveloperMetadataFinder().withKey('category').find();
}

function nextCategoryLocation() {
  var categoriesMetadata = getCategoriesMetadata();
  if (categoriesMetadata.length == 0) {
    return START_COLUMN;
  }
  var columns = [];
  categoriesMetadata.forEach(m => columns.push(m.getLocation().getColumn().getColumn()));
  return Math.max(...columns) + COLUMNS + 1;
}

function categoriesList() {
  var categories = [];
  getCategoriesMetadata().forEach(m => categories.push(m.getValue()));
  return categories;
}

function getMetadata(range) {
  var metadata = range.createDeveloperMetadataFinder().find();
}

function getAllMetadata(key = null) {
  var finder = activeSheet().createDeveloperMetadataFinder();
  if (key == null) {
    return finder.find();
  }
  return finder.withKey(key).find();
}

function printAllMetadata() {
  getAllMetadata().forEach(m => Logger.log(m.getKey()));
}

function removeAllMetadata() {
  getAllMetadata().forEach(m => m.remove());
  Logger.log(getAllMetadata())
}