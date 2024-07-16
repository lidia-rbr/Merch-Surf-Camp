/**
 * Display sales modal
 */
function displayAddSalesForm() {
  // Get ss data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsData = settingsSheet.getDataRange().getValues();
  const settingsDataHeaders = settingsData[0];
  const items = settingsData.map(x => x[settingsDataHeaders.indexOf("Items")]).filter(n => n);

  // Fill drop downs with ss data
  // Sellers
  let sellers = settingsData.map(x => x[settingsDataHeaders.indexOf("Sellers")]).filter(n => n);
  let htmlSellersOptions = "<option value=" + sellers[1] + " selected >" + sellers[1] + "</option > ";
  sellers.slice(2).forEach(function (row) {
    htmlSellersOptions += '<option value=' + row + '>' + row + "</option>";
  });

  const extraProductMap = {};
  // Get product info
  for (let i = 1; i < items.length; i++) {
    extraProductMap[items[i]] = {};
    extraProductMap[items[i]]["colours"] = settingsData[i][settingsDataHeaders.indexOf("Colour")].split(",").filter(n => n);
    extraProductMap[items[i]]["sizes"] = settingsData[i][settingsDataHeaders.indexOf("Size")].split(",").filter(n => n);
  }
  // Add cash card info
  extraProductMap["cashCard"] = settingsData.map(x => x[settingsDataHeaders.indexOf("Paiment type")]).filter(n => n).slice(1);

  let modal = HtmlService.createTemplateFromFile("Form");
  modal.sellersOptions = htmlSellersOptions;
  modal.extraProductMap = JSON.stringify(extraProductMap);

  modal = modal.evaluate();
  modal.setHeight(500).setWidth(850);
  SpreadsheetApp.getUi().showModalDialog(modal, "Add sales");
}

/**
 * Display sales modal
 */
function displayProductForm() {
  // Get ss data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsData = settingsSheet.getDataRange().getValues();
  const settingsDataHeaders = settingsData[0];
  const items = settingsData.map(x => x[settingsDataHeaders.indexOf("Items")]).filter(n => n);

  // Fill drop downs with ss data
  // Sellers
  let sellers = settingsData.map(x => x[settingsDataHeaders.indexOf("Sellers")]).filter(n => n);
  let sellersTextCommaSeparated = sellers[1]
  sellers.slice(2).forEach(function (s) {
    sellersTextCommaSeparated += ',' + s;
  });

  // Paiment types
  let paimentTypes = settingsData.map(x => x[settingsDataHeaders.indexOf("Paiment type")]).filter(n => n);
  let paimentTypesSeparated = paimentTypes[1]
  paimentTypes.slice(2).forEach(function (p) {
    paimentTypesSeparated += ',' + p;
  });

  const extraProductMap = {};
  // Get product info
  for (let i = 1; i < items.length; i++) {
    extraProductMap[items[i]] = {};
    extraProductMap[items[i]]["colours"] = settingsData[i][settingsDataHeaders.indexOf("Colour")].split(",").filter(n => n);
    extraProductMap[items[i]]["sizes"] = settingsData[i][settingsDataHeaders.indexOf("Size")].split(",").filter(n => n);
  }

  let modal = HtmlService.createTemplateFromFile("Product-form");
  modal.paimentTypesText = paimentTypesSeparated;
  modal.sellersText = sellersTextCommaSeparated;
  modal.extraProductMap = JSON.stringify(extraProductMap);

  modal = modal.evaluate();
  modal.setHeight(500).setWidth(850);
  SpreadsheetApp.getUi().showModalDialog(modal, "Manage data");
}

/**
 * Treat client side data and print inputs to sheets
 * @param {object} formData
 * @return {string} res
 */
function treatAndPrintClientSideData(formData) {

  // Get unique items list first
  let items = [];
  for (let i = 0; i < Object.keys(formData).length; i++) {
    let item = Object.keys(formData)[i].split("_")[0];
    if (item != "datePickerInput" &&
      item != "sellerId" &&
      item != "temp" &&
      items.indexOf(item) == -1) {
      items.push(item);
    }
  }
  // Build formatted sales array
  let salesArray = [];
  // Prepare timestamp
  const currentSheet = SpreadsheetApp.getActiveSheet();
  const currentSheetAllData = currentSheet.getDataRange().getValues();
  const userEmailAddress = Session.getEffectiveUser().getEmail();
  const timestamp = new Date();
  const userAndTimestamp = userEmailAddress + ", " + timestamp;
  const dateColInSales = currentSheetAllData.map(x => x[12]).filter(n => n); // +3
  for (let i = 0; i < items.length; i++) {
    let index = 1;
    console.log(items[i] + "_QtyId_" + index);
    console.log(formData[items[i] + "_QtyId_" + index]);
    while (formData[items[i] + "_QtyId_" + index]) {
      if (formData[items[i] + "_QtyId_" + index] != "") {
        // add row
        let row = [
          formData["datePickerInput"],
          formData["sellerId"],
          items[i].replaceAll("SPACE", " "),
          formData[items[i] + "_SizeId_" + index],
          formData[items[i] + "_ColourId_" + index],
          formData[items[i] + "_CashCardId_" + index],
          formData[items[i] + "_QtyId_" + index],
          formData[items[i] + "_StaffId_" + index],
          formData[items[i] + "_NotesId_" + index],
          userAndTimestamp
        ];
        salesArray.push(row);
      }
      index++;
    }
  }
  if (salesArray.length > 0) {
    currentSheet.getRange(dateColInSales.length + 2, 13, salesArray.length, salesArray[0].length).setValues(salesArray);
    // update timestamp
    currentSheet.getRange("J2").setValue(timestamp);
    currentSheet.getRange("J3").setValue(userEmailAddress);
  }
}

/**
 * to include css and js code in the HTML files
 * @param {string} filename
 * @returns {string} 
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

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
  console.log("nbOfProducts", nbOfProducts)
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

/**
 * Update every month sheet with new values, and update settings data
 * @param {object} extraProductMapObj
 * @param {object} newExtraProductMap
 * @param {object} generalMap
 */
function updateSheets(extraProductMapObj, newExtraProductMap, generalMap) {

  let oldItems = Object.keys(extraProductMapObj);
  let newItems = Object.keys(newExtraProductMap);

  let oldNumberOfProducts = 0;

  for (const product in extraProductMapObj) {
    if (extraProductMapObj.hasOwnProperty(product)) {
      const productDetails = extraProductMapObj[product];
      if (productDetails.colours && Array.isArray(productDetails.colours)) {
        oldNumberOfProducts += productDetails.colours.length;
      }
    }
  }

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
  console.log("nbOfProducts", nbOfProducts)
  settingsSheet.getRange(2, 1, settingsSheet.getLastRow(), formattedItemsArray[0].length).clearContent();
  settingsSheet.getRange(2, 1, formattedItemsArray.length, formattedItemsArray[0].length).setValues(formattedItemsArray);

  const newSettingsData = settingsSheet.getDataRange().getValues();

  // Update month sheets
  for (let i = 0; i < MONTHS_ARRAY.length; i++) {
    rebuildProductSections(oldNumberOfProducts, newSettingsData, MONTHS_ARRAY[i]);

  }
}

function rebuildProductSections(oldNumberOfProducts, settingsValues, sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const thisSheet = ss.getSheetByName(sheetName);
    let thisSheetData = thisSheet.getDataRange().getValues();
    let thisSheetItemCol = thisSheetData.map(x => x[ITEM_COL]);

    // const settingsValues = ss.getSheetByName("Settings").getDataRange().getValues();
    const settingsValuesHeaders = settingsValues[0];
    const items = settingsValues.map(x => x[settingsValuesHeaders.indexOf("Items")]);
    items.shift();
    const colours = settingsValues.map(x => x[settingsValuesHeaders.indexOf("Colour")]);
    const nbOfProducts = colours.flatMap(colour => colour.split(',')).length - 1;
    colours.shift();

    // const nbOfProducts = 8;

    // Get category indexes
    let productIndexes = thisSheetItemCol.reduce((indexes, item, index) => {
        if (item === 'Products') {
            indexes.push(index);
        }
        return indexes;
    }, []);

    const oldInventory = thisSheet.getRange(productIndexes[0] + 1, 2, oldNumberOfProducts + 1, 9).getValues();
    const oldInventoryHeaders = oldInventory.shift();
    const oldProductKeys = oldInventory.map(row => `${row[0]}-${row[1]}`);

    // Clear content for each category
    for (let i = 0; i < productIndexes.length; i++) {
        thisSheet.getRange(productIndexes[i] + 2, 2, oldNumberOfProducts, 10).clearContent();
    }

    // Delete rows, except first one to keep formatting
    let rangeToDelete;
    for (let i = productIndexes.length - 1; i >= 0; i--) {
        console.log(productIndexes[i] + 2);
        rangeToDelete = thisSheet.getRange(productIndexes[i] + 3, 2, oldNumberOfProducts - 1, 10);
        rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
    }

    // Build new data from settings tab
    let newData = [];
    let newDataWithOldInventory = [];
    for (let i = 0; i < items.length; i++) {
        let coloursTab = colours[i].split(",");
        for (let j = 0; j < coloursTab.length; j++) {
            let row = new Array(10);
            row[0] = items[i]
            row[1] = coloursTab[j].trim();
            let productKey = `${items[i].trim()}-${coloursTab[j].trim()}`;
            // Check if value existed before and retrieve inventory data if any
            // If no data, inject formula (except for Jan sheet)
            if (oldProductKeys.indexOf(productKey) > -1) {
                let oldRow = oldInventory[oldProductKeys.indexOf(productKey)];
                if (MONTHS_ARRAY[MONTHS_ARRAY.indexOf(sheetName) - 1]) {
                    let oldRowWithFormulas = oldRow.map((c, index) =>
                        c === ""
                            ? `=${MONTHS_ARRAY[MONTHS_ARRAY.indexOf(sheetName) - 1]}!${String.fromCharCode(index + 66)}${nbOfProducts + 12 + j + i}`
                            : c
                    );
                    oldRowWithFormulas.push("");
                    newData.push(row);
                    newDataWithOldInventory.push(oldRowWithFormulas);
                } else {
                    oldRow.push("");
                    newDataWithOldInventory.push(oldRow);
                    newData.push(row);
                }
            }
            else {
                newData.push(row);
                newDataWithOldInventory.push(row);
            }
        }
    }

    // Reinject data
    let totalSold = thisSheet.getRange(productIndexes[2] + 4 - oldNumberOfProducts * 2, 2, newData.length - 1, 10);
    totalSold.insertCells(SpreadsheetApp.Dimension.ROWS);
    thisSheet.getRange(productIndexes[2] + 4 - oldNumberOfProducts * 2, 2, newData.length, 10).setValues(newData);

    let currentStock = thisSheet.getRange(productIndexes[1] + 3 - oldNumberOfProducts, 2, newData.length - 1, 10);
    currentStock.insertCells(SpreadsheetApp.Dimension.ROWS);
    thisSheet.getRange(productIndexes[1] + 3 - oldNumberOfProducts, 2, newData.length, 10).setValues(newData);

    let inventory = thisSheet.getRange(productIndexes[0] + 2, 2, newData.length - 1, 10);
    inventory.insertCells(SpreadsheetApp.Dimension.ROWS);
    thisSheet.getRange(productIndexes[0] + 2, 2, newData.length, 10).setValues(newDataWithOldInventory);

    // Get new indexes
    thisSheetData = thisSheet.getDataRange().getValues();
    thisSheetItemCol = thisSheetData.map(x => x[ITEM_COL]);
    let newProductIndexes = thisSheetItemCol.reduce((indexes, item, index) => {
        if (item === 'Products') {
            indexes.push(index);
        }
        return indexes;
    }, []);

    // Reinject formulas
    // Total Sold (AUTOMATICALLY FILLED)
    // Fill first row
    let totalSoldFirstCellRange = thisSheet.getRange(newProductIndexes[2] + 2, 4);
    totalSoldFirstCellRange.setFormula(`=SUMIFS($S$4:$S, $O$4:$O, $B${newProductIndexes[2] + 2}, $P$4:$P, D$${newProductIndexes[2] + 1}, $Q$4:$Q, $C${newProductIndexes[2] + 2})`);
    let totalSoldFirstRowRange = thisSheet.getRange(newProductIndexes[2] + 2, 4, 1, NUMBER_OF_SIZES);
    totalSoldFirstCellRange.autoFill(totalSoldFirstRowRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    thisSheet.getRange(newProductIndexes[2] + 2, 11).setFormula(`=SUM(D${newProductIndexes[2] + 2}:J${newProductIndexes[2] + 2})`);
    // Expand first row
    thisSheet.getRange(newProductIndexes[2] + 2, 4, 1, NUMBER_OF_SIZES + 1).autoFill(
        thisSheet.getRange(newProductIndexes[2] + 2, 4, newData.length, NUMBER_OF_SIZES + 1),
        SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );

    // Total Sold (AUTOMATICALLY FILLED)
    // Fill first row =D9-D27
    let currentStockCellRange = thisSheet.getRange(newProductIndexes[1] + 2, 4);
    currentStockCellRange.setFormula(`=D${newProductIndexes[0] + 2}-D${newProductIndexes[2] + 2}`);
    let currentStockFirstRowRange = thisSheet.getRange(newProductIndexes[1] + 2, 4, 1, NUMBER_OF_SIZES);
    currentStockCellRange.autoFill(currentStockFirstRowRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    thisSheet.getRange(newProductIndexes[1] + 2, 11).setFormula(`=SUM(D${newProductIndexes[1] + 2}:J${newProductIndexes[1] + 2})`);
    // Expand first row
    thisSheet.getRange(newProductIndexes[1] + 2, 4, 1, NUMBER_OF_SIZES + 1).autoFill(
        thisSheet.getRange(newProductIndexes[1] + 2, 4, newData.length, NUMBER_OF_SIZES + 1),
        SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );

    // Inventory
    thisSheet.getRange(newProductIndexes[0] + 2, 11).setFormula(`=SUM(D${newProductIndexes[0] + 2}:J${newProductIndexes[0] + 2})`);
    // Expand first row
    thisSheet.getRange(newProductIndexes[0] + 2, 11).autoFill(
        thisSheet.getRange(newProductIndexes[0] + 2, 11, newData.length, 1),
        SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
}


/**
 * Adds menu in Sheets
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
    .addItem('Create a blank copy', 'createBlankCopyModal')
    .addToUi();
}

/**
 * Display loader while creating a copy of the file
 */
function createBlankCopyModal() {
  let modal = HtmlService.createTemplateFromFile("CreateCopy");
  modal = modal.evaluate();
  modal.setHeight(350).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(modal, "Create new file");
}

/**
 * Create a copy of existing file and duplicate template to rebuild months
 */
function createBlankCopy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSsName = "BLANK Merchandise spreadsheet";
  const newSs = ss.copy(newSsName);
  const newSsURL = newSs.getUrl();
  let res = {
    "url": newSsURL,
    "ssName": newSsName
  };
  // Format new spreadsheet
  const templateSheet = newSs.getSheetByName("Template");
  for (let i = 0; i < MONTHS_ARRAY.length - 1; i++) {
    let sheet = newSs.getSheetByName(MONTHS_ARRAY[i]);
    newSs.deleteSheet(sheet);
    templateSheet.copyTo(newSs).setName(MONTHS_ARRAY[i]).showSheet();
  }
  updateInventoryForNextMonth(newSs.getId(), 1);
  return res;
}

/**
 * Insert Stock formulas based on last month current stock values
 * @param {string} ssId
 */
function updateInventoryForNextMonth(ssId, monthIndex) {
  const ss = SpreadsheetApp.openById(ssId);
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsData = settingsSheet.getDataRange().getValues();
  const settingsDataHeaders = settingsData[0];
  const colorCol = settingsData.map(x => x[settingsDataHeaders.indexOf("Colour")]).filter(n => n).slice(1);
  let numberOfRowsInInventory = 0;
  colorCol.forEach((c) => {
    numberOfRowsInInventory = numberOfRowsInInventory + c.split(",").length;
  })
  const janSheet = ss.getSheetByName("Jan");
  const janSheetData = janSheet.getDataRange().getValues();
  const janSheetDataSecondCol = janSheetData.map(x => x[1]);
  let currentStockFirstCell = "D" + parseInt(janSheetDataSecondCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") + 3);
  let inventoryFormula = "=prevMonth!" + currentStockFirstCell;

  let safeMonthIndex = monthIndex == 1 ? 2 : monthIndex;
  console.log("safeMonthIndex", safeMonthIndex);
  for (let i = safeMonthIndex; i < MONTHS_ARRAY.length - 1; i++) {
    let thisSheet = ss.getSheetByName(MONTHS_ARRAY[i]);
    let formulaForThisSheet = inventoryFormula.replace("prevMonth", MONTHS_ARRAY[i - 1]);
    let sourceRange = thisSheet.getRange(FIRST_STOCK_ROW, 4);
    sourceRange.setFormula(formulaForThisSheet);
    let verticalDestinationRange = thisSheet.getRange(FIRST_STOCK_ROW, 4, numberOfRowsInInventory, 1);
    sourceRange.autoFill(verticalDestinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    let horizontalDestinationRange = thisSheet.getRange(FIRST_STOCK_ROW, 4, numberOfRowsInInventory, NUMBER_OF_SIZES);
    verticalDestinationRange.autoFill(horizontalDestinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
  return safeMonthIndex;
}


function testUpdate() {
  const index = 1;
  const ssId = "1rhhvymbRnv_oLc_FvsMKUqiw5usYcc5nDIQsbnqnDks";
  const monthIndexUsed = updateInventoryForNextMonth(ssId, index);
  console.log(monthIndexUsed)
}
