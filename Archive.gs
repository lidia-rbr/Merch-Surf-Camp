/**
 * Update every month sheet with new values, and update settings data
 * @param {object} extraProductMapObj
 * @param {object} newExtraProductMap
 * @param {object} generalMap
 */
function oldupdateSheets(extraProductMapObj, newExtraProductMap, generalMap) {

  let oldItems = Object.keys(extraProductMapObj);
  let newItems = Object.keys(newExtraProductMap);
  let intersectionItems = oldItems.filter(x => newItems.includes(x));
  let deletedItems = oldItems.filter(x => !newItems.includes(x));
  let onlyNewItems = newItems.filter(x => !oldItems.includes(x));

  // Update settings sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsData = settingsSheet.getDataRange().getValues();
  const settingsDataHeaders = settingsData[0];
  const formattedSellersList = generalMap.sellers.split(",").map(x => [x]);
  const formattedPaimentTypesList = generalMap.paimentTypes.split(",").map(x => [x]);
  settingsSheet.getRange(2, settingsDataHeaders.indexOf("Paiment type") + 1, formattedPaimentTypesList.length, 1).setValues(formattedPaimentTypesList);
  settingsSheet.getRange(2, settingsDataHeaders.indexOf("Sellers") + 1, formattedSellersList.length, 1).setValues(formattedSellersList);

  // Build item array
  let formattedItemsArray = [];
  let nbOfProducts = 0;
  for (let i = 0; i < newItems.length; i++) {
    let row = [
      newItems[i],
      newExtraProductMap[newItems[i]]["colours"],
      newExtraProductMap[newItems[i]]["sizes"]
    ];
    formattedItemsArray.push(row);
    nbOfProducts = nbOfProducts + newExtraProductMap[newItems[i]]["colours"].split(",").length;
  }
  settingsSheet.getRange(2, 1, settingsSheet.getLastRow(), formattedItemsArray[0].length).clearContent();
  settingsSheet.getRange(2, 1, formattedItemsArray.length, formattedItemsArray[0].length).setValues(formattedItemsArray);

  // Update month sheets
  for (let i = 0; i < MONTHS_ARRAY.length; i++) {
    let thisSheet = ss.getSheetByName(MONTHS_ARRAY[i]);
    let thisSheetData = thisSheet.getDataRange().getValues();
    let thisSheetItemCol = thisSheetData.map(x => x[ITEM_COL]);
    let thisSheetColourCol = thisSheetData.map(x => x[COLOURS_COL]);

    // Delete items if necessary
    for (let j = thisSheetItemCol.length; j > 7; j--) {
      if (deletedItems.indexOf(thisSheetItemCol[j]) > -1) {
        let rangeToDelete = thisSheet.getRange(j + 1, 2, 1, NUMBER_OF_COL_TODELETE);
        rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
        thisSheetItemCol.splice(j, 1);
        thisSheetColourCol.splice(j, 1);
      }
    }
    let thisColourProduct;
    // Add item if necessary
    for (let j = 0; j < onlyNewItems.length; j++) {
      let newProductColoursArray = newExtraProductMap[onlyNewItems[j]]["colours"].split(",");
      for (let k = 0; k < newProductColoursArray.length; k++) {
        let thisColourProduct = (newProductColoursArray[k] == "") ? "-" : newProductColoursArray[k].trim();

        // Inventory
        let sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
        let destination = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
        let cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)"), 2, 1, NUMBER_OF_COL_TODELETE);
        cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
        thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)"), 2, 1, 2).setValues([
          [onlyNewItems[j], thisColourProduct]
        ]);
        thisSheetItemCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, onlyNewItems[j]);
        thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newProductColoursArray[k]);
        // sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

        // Current stock
        sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
        destination = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
        cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)"), 2, 1, NUMBER_OF_COL_TODELETE);
        cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
        thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)"), 2, 1, 2).setValues([
          [onlyNewItems[j], thisColourProduct]
        ]);
        thisSheetItemCol.splice(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 0, onlyNewItems[j]);
        thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newProductColoursArray[k]);
        sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

        // Total sold 
        sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("$End$") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
        destination = thisSheet.getRange(thisSheetItemCol.indexOf("$End$") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
        cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("$End$"), 2, 1, NUMBER_OF_COL_TODELETE);
        cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
        thisSheet.getRange(thisSheetItemCol.indexOf("$End$"), 2, 1, 2).setValues([
          [onlyNewItems[j], thisColourProduct]
        ]);
        thisSheetItemCol.splice(thisSheetItemCol.indexOf("$End$") - 1, 0, onlyNewItems[j]);
        thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newProductColoursArray[k]);
        sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      }
    }
    // Check if new or deleted colours in intersection products
    for (let j = 0; j < intersectionItems.length; j++) {
      let thisItemOldColoursTab = extraProductMapObj[intersectionItems[j]]["colours"].map(c => c.trim());
      let thisItemNewColoursTab = newExtraProductMap[intersectionItems[j]]["colours"].split(",").map(c => c.trim());

      if (JSON.stringify(thisItemOldColoursTab) != JSON.stringify(thisItemNewColoursTab)) {

        let newColours = thisItemNewColoursTab.filter(x => !thisItemOldColoursTab.includes(x));
        let deletedColours = thisItemOldColoursTab.filter(x => !thisItemNewColoursTab.includes(x));

        // Add new colours
        if (newColours.length > 0) {
          for (let k = 0; k < newColours.length; k++) {

            thisColourProduct = (newColours[k] == "") ? "-" : newColours[k];

            // Inventory
            // let sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
            // let destination = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
            // // Insert new colours just after product
            let cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf(intersectionItems[j]) + 2, 2, 1, NUMBER_OF_COL_TODELETE);
            cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
            thisSheet.getRange(thisSheetItemCol.indexOf(intersectionItems[j]) + 2, 2, 1, 2).setValues([
              [intersectionItems[j], thisColourProduct]
            ]);
            // thisSheetItemCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, onlyNewItems[j]);
            // thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newColours[k]);
            // sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

            // Current stock
            // sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
            // destination = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
            cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf(intersectionItems[j]) + 4 + nbOfProducts + k, 2, 1, NUMBER_OF_COL_TODELETE);
            cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
            thisSheet.getRange(thisSheetItemCol.indexOf(intersectionItems[j]) + 4 + nbOfProducts + k, 2, 1, 2).setValues([
              [intersectionItems[j], thisColourProduct]
            ]);
            // thisSheetItemCol.splice(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 0, onlyNewItems[j]);
            // thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newColours[k]);
            // sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

            // Total sold 
            // sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("$End$") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
            // destination = thisSheet.getRange(thisSheetItemCol.indexOf("$End$") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
            cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf(intersectionItems[j]) + 7 + 2 * (nbOfProducts + k), 2, 1, NUMBER_OF_COL_TODELETE);
            cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
            thisSheet.getRange(thisSheetItemCol.indexOf(intersectionItems[j]) + 7 + 2 * (nbOfProducts + k), 2, 1, 2).setValues([
              [intersectionItems[j], thisColourProduct]
            ]);
            // thisSheetItemCol.splice(thisSheetItemCol.indexOf("$End$") - 1, 0, onlyNewItems[j]);
            // thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newColours[k]);
            // sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
          }
        } else if (deletedColours.length > 0) {
          // Delete items if necessary
          for (let k = thisSheetItemCol.length; k > 7; k--) {
            if (deletedColours.indexOf(thisSheetColourCol[k]) > -1 && thisSheetItemCol[k] == intersectionItems[j]) {
              let rangeToDelete = thisSheet.getRange(k + 1, 2, 1, NUMBER_OF_COL_TODELETE);
              rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
              thisSheetItemCol.splice(k, 1);
              thisSheetColourCol.splice(k, 1);
            }
          }
        }
      }
    }
  }
  updateInventoryForNextMonth(ss.getId(), MONTHS_ARRAY.indexOf(ss.getActiveSheet().getName()));
}