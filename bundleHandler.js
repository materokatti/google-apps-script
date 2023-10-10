function onEdit(e) {
  // Get the spreadsheet that was edited
  var ss = e.source;

  // Get the range of cells that was edited
  var activeRange = e.range;

  // Get the sheet that contains the edited range
  var activeSheet = activeRange.getSheet();

  // Get the column number of the edited range
  var activeColumn = activeRange.getColumn();

  // Get the value in the edited cell
  var selectedItem = activeRange.getValue();

  // If the edited column is 6 and the value is not empty
  if (activeColumn === 6 && selectedItem.length > 0) {

    // Get the sheet named "tab2"
    var otherSheet = ss.getSheetByName("tab2");

    // Get the values in column B from row 48 to the last row
    var items = otherSheet.getRange("B48:B" + otherSheet.getLastRow()).getValues();

    // Array to store values that match the edited value
    var valuesToCopy = [];

    // Loop through the items in column B
    for (var i = 0; i < items.length; i++) {
      if (items[i][0] === selectedItem) {

        // Get the corresponding row number and value from column C
        var correspondingRow = i + 48;
        var correspondingValue = otherSheet.getRange(correspondingRow, 3).getValue();

        // Add the corresponding value to the valuesToCopy array
        valuesToCopy.push([correspondingValue]);

        // Loop through subsequent rows as long as they are empty or undefined
        while (i+1 < items.length && (items[i+1][0] === "" || items[i+1][0] === undefined)) {
          i++;
          correspondingRow = i + 48;
          correspondingValue = otherSheet.getRange(correspondingRow, 3).getValue();
          valuesToCopy.push([correspondingValue]);
        }
        break;
      }
    }

    // If there are any values to copy
    if (valuesToCopy.length > 0) {

      // Insert new rows below the edited cell for the values to copy
      activeSheet.insertRows(activeRange.getRow() + 1, valuesToCopy.length);

      // Set the values in the new rows to the values from the valuesToCopy array
      activeSheet.getRange(activeRange.getRow() + 1, 6, valuesToCopy.length, 1).setValues(valuesToCopy);
      
      // Set VLOOKUP formula for the new rows in column Y
      var vlookupRange = activeSheet.getRange(activeRange.getRow() + 1, 25, valuesToCopy.length, 1); // 25 corresponds to column Y
      var formula = "=VLOOKUP(F" + (activeRange.getRow() + 1).toString() + ", 'tab2'!$B$2:$D$92, 3, 0)";
      vlookupRange.setFormula(formula);
      
      // Delete the original edited row
      activeSheet.deleteRow(activeRange.getRow());
    }
  }
}
