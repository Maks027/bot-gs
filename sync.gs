class Category {
  constructor(name) {
    this.name = name;
    /** @type {any[]} */
    this.items = [];
  }

  addItem(item) {
    this.items.push(item);
  }
}

class Item {
  constructor (internalId, name, unit, imageUrl) {
    this.internal_id = internalId;
    this.name = name;
    this.unit = unit;
    this.image_url = imageUrl;
  }
}

function parseCatalog () {
  var sheet = activeSheet();
  var categoriesMetadata = sheet.createDeveloperMetadataFinder().withKey('category').find();

  var categories = [];

  for (var i = 0; i < categoriesMetadata.length; i++) {
    var metadata = categoriesMetadata[i];
    var category = new Category(metadata.getValue());

    var columnRange = metadata.getLocation().getColumn();
    var lastRow     = columnRange.getLastRow();
    var range       = sheet.getRange(3, columnRange.getColumn(), lastRow, 4);
    var values      = range.getValues().filter(v => (v[0] != '') && (v[1] != ''));

    values.forEach(v => category.addItem(new Item(String(v[0]), String(v[1]), String(v[2]), String(v[3]))));

    categories.push(category);
  }

  // Logger.log(JSON.stringify(categories));

  return JSON.stringify(categories);
}