SHEET_NAME   = "Catalog"
START_COLUMN = 4;
START_ROW    = 1;
COLUMNS      = 4;

function activeSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sheet == null) {
    sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1").setName(SHEET_NAME);
  }
  return sheet;
}

function generateControls() {
  var sheet = activeSheet();

  sheet.setColumnWidth(1, 30).setColumnWidth(2, 130).setColumnWidth(3, 30);

  var addImage = sheet.insertImage(getAddButtonImageUrl(), 2, 2, 4, 2);
  var rmImage = sheet.insertImage(getRemoveButtonImageUrl(), 2, 4, 4, 2);
  var editImage = sheet.insertImage(getEditButtonImageUrl(), 2, 6, 4, 2);

  ["B2:B3", "B4:B5", "B6:B7"].forEach(
    a1 => sheet.getRange(a1).mergeVertically().setBorder(true, true, true, true, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  );

  addImage.assignScript("openAddDialog");
  rmImage.assignScript("openRemoveDialog");
  editImage.assignScript("openEditDialog");

  setProtectedRanges();
}

function setProtectedRanges() {
  activeSheet().getRange('A:C').protect().setDescription('Controls').setWarningOnly(true);
  activeSheet().getRange('1:2').protect().setDescription('Controls').setWarningOnly(true);
}

function openAddDialog() {
  var html = HtmlService.createTemplateFromFile('add').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function openRemoveDialog() {
  var html = HtmlService.createTemplateFromFile('rm').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function openEditDialog() {
  var html = HtmlService.createTemplateFromFile('edit').evaluate().setWidth(400).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}


function extractA1Columns(notation) {
  return notation.split(':').map(s => s.match(/^[A-Z]+/)[0]).join(':')
}

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
    .setValue(name);

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

  sheet.getRange(2, dataRange.getLastColumn(), 1, 1).activate();

  var firstColumn = dataRangeNotation.split(':')[0];
  sheet.getRange(`${firstColumn}:${firstColumn}`).addDeveloperMetadata('category', name);
}

function removeCategory(name) {
  var metadata = getCategoryMetadata(name);
  activeSheet().deleteColumns(metadata.getLocation().getColumn().getColumn(), COLUMNS);
  metadata.remove();
}

function editCategory(name, newName) {
  var metadata = getCategoryMetadata(name);
  activeSheet().getRange(START_ROW, metadata.getLocation().getColumn().getColumn(), 1, COLUMNS).setValue(newName);

  metadata.setValue(newName);
}

function getCategoriesMetadata() {
  return activeSheet().createDeveloperMetadataFinder().withKey('category').find();
}

function getCategoryMetadata(name) {
  return activeSheet().createDeveloperMetadataFinder().withKey('category').withValue(name).find()[0];
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
