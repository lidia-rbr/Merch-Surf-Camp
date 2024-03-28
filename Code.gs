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
          items[i].replaceAll("-", " "),
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

// /**
//  * Display product modal
//  */
// function displayAddProductForm() {
//   // Get ss data
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const settingsSheet = ss.getSheetByName("Settings");
//   const settingsData = settingsSheet.getDataRange().getValues();
//   const settingsDataHeaders = settingsData[0];
//   const items = settingsData.map(x => x[settingsDataHeaders.indexOf("Items")]).filter(n => n);

//   // Fill drop downs with ss data
//   // Sellers
//   let sellers = settingsData.map(x => x[settingsDataHeaders.indexOf("Sellers")]).filter(n => n);
//   let htmlSellersOptions = "<option value=" + sellers[1] + " selected >" + sellers[1] + "</option > ";
//   sellers.slice(2).forEach(function (row) {
//     htmlSellersOptions += '<option value=' + row + '>' + row + "</option>";
//   });
//   // Hoodie colours
//   let hoodieColours = settingsData[settingsDataHeaders.indexOf("Colour")][items.indexOf("Hoodie")].split(",");
//   let htmlHoodieColoursOptions = "<option value=" + hoodieColours[0] + " selected >" + hoodieColours[0] + "</option > ";
//   hoodieColours.slice(1).forEach(function (val) {
//     htmlHoodieColoursOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // Hoodie sizes
//   let hoodieSizes = settingsData[items.indexOf("Hoodie")][settingsDataHeaders.indexOf("Size")].split(",");
//   let htmlHoodieSizesOptions = "<option value=" + hoodieSizes[0] + " selected >" + hoodieSizes[0] + "</option > ";
//   hoodieSizes.slice(1).forEach(function (val) {
//     htmlHoodieSizesOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // teeShirt colours
//   let teeShirtColours = settingsData[items.indexOf("T-Shirt")][settingsDataHeaders.indexOf("Colour")].split(",");
//   let htmlteeShirtColoursOptions = "<option value=" + teeShirtColours[0] + " selected >" + teeShirtColours[0] + "</option > ";
//   teeShirtColours.slice(1).forEach(function (val) {
//     htmlteeShirtColoursOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // teeShirt sizes
//   let teeShirtSizes = settingsData[items.indexOf("T-Shirt")][settingsDataHeaders.indexOf("Size")].split(",");
//   let htmlteeShirtSizesOptions = "<option value=" + teeShirtSizes[0] + " selected >" + teeShirtSizes[0] + "</option > ";
//   teeShirtSizes.slice(1).forEach(function (val) {
//     htmlteeShirtSizesOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // Paiment type
//   let paimentType = settingsData.map(x => x[settingsDataHeaders.indexOf("Paiment type")]).filter(n => n);
//   let htmlPaimentTypeOptions = "<option value=" + paimentType[1] + " selected >" + paimentType[1] + "</option > ";
//   paimentType.slice(2).forEach(function (row) {
//     htmlPaimentTypeOptions += '<option value=' + row + '>' + row + "</option>";
//   });
//   // Beanie size
//   let beanieSizes = settingsData[items.indexOf("Beanie")][settingsDataHeaders.indexOf("Size")].split(",");
//   let htmlBeanieSizesOptions = "<option value=" + beanieSizes[0] + " selected >" + beanieSizes[0] + "</option > ";
//   beanieSizes.slice(1).forEach(function (val) {
//     htmlBeanieSizesOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // beanie colours
//   let beanieColours = settingsData[items.indexOf("Beanie")][settingsDataHeaders.indexOf("Colour")].split(",");
//   let htmlBeanieColoursOptions = "<option value=" + beanieColours[0] + " selected >" + beanieColours[0] + "</option > ";
//   beanieColours.slice(1).forEach(function (val) {
//     htmlBeanieColoursOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // Sunhat size
//   let sunhatSizes = settingsData[items.indexOf("Sunhat")][settingsDataHeaders.indexOf("Size")].split(",");
//   let htmlsunhatSizesOptions = "<option value=" + sunhatSizes[0] + " selected >" + sunhatSizes[0] + "</option > ";
//   sunhatSizes.slice(1).forEach(function (val) {
//     htmlsunhatSizesOptions += '<option value=' + val + '>' + val + "</option>";
//   });
//   // Sunhat colours
//   let sunhatColours = settingsData[items.indexOf("Sunhat")][settingsDataHeaders.indexOf("Colour")].split(",");
//   let htmlSunhatColoursOptions = "<option value=" + sunhatColours[0] + " selected >" + sunhatColours[0] + "</option > ";
//   sunhatColours.slice(1).forEach(function (val) {
//     htmlSunhatColoursOptions += '<option value=' + val + '>' + val + "</option>";
//   });

//   let mapExample = {
//     "productName": "Test new Product",
//     "sizes": ["XS", "S"],
//     "colours": ["white", "blue"]
//   }

//   let modal = HtmlService.createTemplateFromFile("Form");
//   modal.sellersOptions = htmlSellersOptions;
//   modal.hoodieColourOptions = htmlHoodieColoursOptions;
//   modal.hoodieSizesOptions = htmlHoodieSizesOptions;
//   modal.teeShirtColourOptions = htmlteeShirtColoursOptions;
//   modal.teeShirtSizesOptions = htmlteeShirtSizesOptions;
//   modal.beanieSizesOptions = htmlBeanieSizesOptions;
//   modal.beanieColourOptions = htmlBeanieColoursOptions;
//   modal.sunhatSizesOptions = htmlsunhatSizesOptions;
//   modal.sunhatColourOptions = htmlSunhatColoursOptions
//   modal.paimentTypeOptions = htmlPaimentTypeOptions;
//   modal.extraProductMap = mapExample

//   modal = modal.evaluate();
//   modal.setHeight(500).setWidth(850);
//   SpreadsheetApp.getUi().showModalDialog(modal, "Add sales");
// }

/**
 * Update every month sheet with new values, and update settings data
 * @param {object} extraProductMapObj
 * @param {object} newExtraProductMap
 * @param {object} generalMap
 */
function updateSheets(extraProductMapObj, newExtraProductMap, generalMap) {
  // let extraProductMapObj = {
  //   'T-Shirt':
  //   {
  //     'colours': ['Black', ' White', ' Grey'],
  //     'sizes': ['XS', ' S', ' M', ' L', ' XL', ' XXL']
  //   },
  //   'Beanie': { sizes: ['XS'], colours: ['Grey'] },
  //   'Sunhat': { colours: ['-'], sizes: ['-'] },
  //   'Hoodie':
  //   {
  //     'sizes': ['XS', ' S', ' M', ' L', ' XL', ' XXL'],
  //     'colours': ['Black']
  //   }
  // }

  // let generalMap = {
  //   "sellers": "Diogo,Danny,Margo,Lidia",
  //   "paimentTypes": "free,cahs,card"
  // };

  // let newExtraProductMap = {
  //   'New product': { 'colours': 'Black,White', 'sizes': 'S,M' },
  //   'T-Shirt': { 'sizes': 'XS, S, M, L, XL, XXL', 'colours': 'Black, White, Grey' },
  //   'Beanie': { 'colours': 'Grey', 'sizes': 'XS' },
  //   'Hoodie': { 'sizes': 'XS, S, M, L, XL, XXL', 'colours': 'Black' }
  // }
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
  for (let i = 0; i < newItems.length; i++) {
    let row = [
      newItems[i],
      newExtraProductMap[newItems[i]]["colours"],
      newExtraProductMap[newItems[i]]["sizes"]
    ];
    formattedItemsArray.push(row);
  }
  settingsSheet.getRange(2, 1, settingsSheet.getLastRow(), formattedItemsArray[0].length).clearContent();
  settingsSheet.getRange(2, 1, formattedItemsArray.length, formattedItemsArray[0].length).setValues(formattedItemsArray);

  // Update month sheets
  // for (let i = 0; i < MONTHS_ARRAY.length; i++) {
  // let thisSheet = ss.getSheetByName(MONTHS_ARRAY[i]);
  let thisSheet = ss.getSheetByName("Copy of Template");
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

  // Add item if necessary
  for (let j = 0; j < onlyNewItems.length; j++) {
    let newProductColoursArray = newExtraProductMap[onlyNewItems[j]]["colours"].split(",");
    // for (let k = 0; k < newProductColoursArray.length; k++) {
    for (let k = 0; k < 2; k++) {
      // Inventory
      let sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
      let destination = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
      let cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)"), 2, 1, NUMBER_OF_COL_TODELETE);
      cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
      thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)"), 2, 1, 2).setValues([
        [onlyNewItems[j], newProductColoursArray[k]]
      ]);
      thisSheetItemCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, onlyNewItems[j]);
      thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newProductColoursArray[k]);
      sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

      // Current stock
      sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
      destination = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
      cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)"), 2, 1, NUMBER_OF_COL_TODELETE);
      cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
      thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)"), 2, 1, 2).setValues([
        [onlyNewItems[j], newProductColoursArray[k]]
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
        [onlyNewItems[j], newProductColoursArray[k]]
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

    console.log("old colors for item: " + intersectionItems[j] + " ", thisItemOldColoursTab)
    console.log(typeof thisItemOldColoursTab)
    console.log("new colors for item: " + intersectionItems[j] + " ", thisItemNewColoursTab)
    console.log(typeof thisItemNewColoursTab)

    if (JSON.stringify(thisItemOldColoursTab) != JSON.stringify(thisItemNewColoursTab)) {

      let newColours = thisItemNewColoursTab.filter(x => !thisItemOldColoursTab.includes(x));
      let deletedColours = thisItemOldColoursTab.filter(x => !thisItemNewColoursTab.includes(x));



      return

      // Add new colours
      if (newColours.length > 0) {
        for (let k = 0; k < newColours.length; k++) {
          // Inventory
          let sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
          let destination = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
          let cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)"), 2, 1, NUMBER_OF_COL_TODELETE);
          cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
          thisSheet.getRange(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)"), 2, 1, 2).setValues([
            [intersectionItems[j], newColours[k]]
          ]);
          thisSheetItemCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, onlyNewItems[j]);
          thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newProductColoursArray[k]);
          sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

          // Current stock
          sourceRange = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 1, NUMBER_OF_COL_TODELETE - 2);
          destination = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)") - 1, 4, 2, NUMBER_OF_COL_TODELETE - 2);
          cellsToAdd = thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)"), 2, 1, NUMBER_OF_COL_TODELETE);
          cellsToAdd.insertCells(SpreadsheetApp.Dimension.ROWS);
          thisSheet.getRange(thisSheetItemCol.indexOf("Total Sold (AUTOMATICALLY FILLED)"), 2, 1, 2).setValues([
            [intersectionItems[j], newColours[k]]
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
            [intersectionItems[j], newColours[k]]
          ]);
          thisSheetItemCol.splice(thisSheetItemCol.indexOf("$End$") - 1, 0, onlyNewItems[j]);
          thisSheetColourCol.splice(thisSheetItemCol.indexOf("Current Stock Levels (AUTOMATICALLY FILLED)") - 1, 0, newProductColoursArray[k]);
          sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
        }
      } else if (deletedColours.length > 0) {
        // Delete items if necessary
        for (let k = thisSheetItemCol.length; k > 7; k--) {
          if (deletedColours.indexOf(thisSheetColourCol[k]) > -1 && thisSheetItemCol[k] == intersectionItems[j]) {
            let rangeToDelete = thisSheet.getRange(k + 1, 2, 1, NUMBER_OF_COL_TODELETE);
            rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
            thisSheetItemCol.splice(k, 1);
            thisSheetColourCol.thisSheetItemCol.splice(k, 1);
          }
        }
      }
    }
  }


  // }


}


function motherfucker() {
  let array = [ 'Black', 'Barbecue', 'Grey' ];
  console.log(array.includes("Black"));
  
}
// const INVENTORY_CELL = "B7";
// const NUMBER_OF_COL_TODELETE = 10;
// const ITEM_COL = 2;
// const COLOURS_COL = 3;