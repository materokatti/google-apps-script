function onEdit(e) {
  var ss = e.source;
  var activeRange = e.range;
  var activeSheet = activeRange.getSheet();
  var activeColumn = activeRange.getColumn();
  var selectedItem = activeRange.getValue();

  if (activeColumn === 6 && selectedItem.length > 0) {
    var otherSheet = ss.getSheetByName("sheet2");
    var items = otherSheet.getRange("B48:B" + otherSheet.getLastRow()).getValues();
    var valuesToCopy = [];

    for (var i = 0; i < items.length; i++) {
      if (items[i][0] === selectedItem) {
        var correspondingRow = i + 48;
        var correspondingValue = otherSheet.getRange(correspondingRow, 3).getValue();
        valuesToCopy.push([correspondingValue]);

        while (i+1 < items.length && (items[i+1][0] === "" || items[i+1][0] === undefined)) {
          i++;
          correspondingRow = i + 48;
          correspondingValue = otherSheet.getRange(correspondingRow, 3).getValue();
          valuesToCopy.push([correspondingValue]);
        }
        break;
      }
    }

    if (valuesToCopy.length > 0) {
      activeSheet.insertRows(activeRange.getRow() + 1, valuesToCopy.length);
      activeSheet.getRange(activeRange.getRow() + 1, 6, valuesToCopy.length, 1).setValues(valuesToCopy);
    }
  }
}
