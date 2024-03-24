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

function testObj() {
  let mapExample = {
    "Test new Product": {
      "sizes": ["XS", "S"],
      "colours": ["white", "blue"]
    },
    "second new product": {
      "sizes": ["XS", "S"],
      "colours": ["white", "blue"]
    }
  }
  for (let i = 0; i < Object.keys(mapExample).length; i++) {
    console.log(mapExample[Object.keys(mapExample)[i]]);
    console.log(Object.keys(mapExample)[i])
  }
}

/**
 * Treat client side data and print inputs to sheets
 * @param {object} formData
 * @return {string} res
 */
function treatAndPrintClientSideData(formData) {

  console.log(formData);
  return;




  const currentSheet = SpreadsheetApp.getActiveSheet();
  const currentSheetAllData = currentSheet.getDataRange().getValues();
  // Col 11
  console.log("formData", formData);
  const dateColInSales = currentSheetAllData.map(x => x[12]).filter(n => n); // +3
  let salesArray = [];
  // Date	Seller	Item	Size	Colour	Cash/card	Staff	Notes added by
  const userEmailAddress = Session.getEffectiveUser().getEmail();
  const timestamp = new Date();
  const userAndTimestamp = userEmailAddress + ", " + timestamp;
  // Get first row tee shirt
  if (formData["teeShirtQtyId"] != "") {
    let teeShirtFirstRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "T-Shirt",
      formData["teeShirtSizeId"],
      formData["teeShirtColourId"],
      formData["teeShirtCashCardId"],
      formData["teeShirtQtyId"],
      formData["teeShirtStaffId"],
      formData["teeShirtNotesId"],
      userAndTimestamp
    ];
    salesArray.push(teeShirtFirstRow);
  }
  let teeShirtIncrement = 2;
  let hoodieIncrement = 2;
  let beanieIncrement = 2;
  let sunhatIncrement = 2;
  // Get all extra tee shirts
  while (formData["teeShirtQtyId_" + teeShirtIncrement] && formData["teeShirtQtyId_" + teeShirtIncrement] != "") {
    // Add the row to result array
    let teeShirtRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "T-Shirt",
      formData["teeShirtSizeId_" + teeShirtIncrement],
      formData["teeShirtColourId_" + teeShirtIncrement],
      formData["teeShirtCashCardId_" + teeShirtIncrement],
      formData["teeShirtQtyId_" + teeShirtIncrement],
      formData["teeShirtStaffId_" + teeShirtIncrement],
      formData["teeShirtNotesId_" + teeShirtIncrement],
      userAndTimestamp
    ];
    salesArray.push(teeShirtRow);
    teeShirtIncrement++;
  }
  // Get first row hoodie
  if (formData["hoodieQtyId"] != "") {
    let hoodieFirstRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "Hoodie",
      formData["hoodieSizeId"],
      formData["hoodieColourId"],
      formData["hoodieCashCardId"],
      formData["hoodieQtyId"],
      formData["hoodieStaffId"],
      formData["hoodieNotesId"],
      userAndTimestamp
    ];
    salesArray.push(hoodieFirstRow);
  }
  // Get all extra hoodies
  while (formData["hoodieQtyId_" + hoodieIncrement] && formData["hoodieQtyId_" + hoodieIncrement] != "") {
    // Add the row to result array
    let hoodieRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "Hoodie",
      formData["hoodieSizeId_" + hoodieIncrement],
      formData["hoodieColourId_" + hoodieIncrement],
      formData["hoodieCashCardId_" + hoodieIncrement],
      formData["hoodieQtyId_" + hoodieIncrement],
      formData["hoodieStaffId_" + hoodieIncrement],
      formData["hoodieNotesId_" + hoodieIncrement],
      userAndTimestamp
    ];
    hoodieIncrement++;
    salesArray.push(hoodieRow);
  }
  // Get first row beanie
  if (formData["beanieQtyId"] != "") {
    let beanieFirstRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "Beanie",
      formData["beanieSizeId"],
      formData["beanieColourId"],
      formData["beanieCashCardId"],
      formData["beanieQtyId"],
      formData["beanieStaffId"],
      formData["beanieNotesId"],
      userAndTimestamp
    ];
    salesArray.push(beanieFirstRow);
  }
  // Get all extra beanies
  while (formData["beanieQtyId" + beanieIncrement] && formData["beanieQtyId" + beanieIncrement] != "") {
    // Add the row to result array
    let beanieRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "Beanie",
      formData["beanieSizeId_" + beanieIncrement],
      formData["beanieColourId_" + beanieIncrement],
      formData["beanieCashCardId_" + beanieIncrement],
      formData["beanieQtyId_" + beanieIncrement],
      formData["beanieStaffId_" + beanieIncrement],
      formData["beanieNotesId_" + beanieIncrement],
      userAndTimestamp
    ];
    beanieIncrement++;
    salesArray.push(beanieRow);
  }
  // Get first sunhat
  if (formData["sunhatQtyId"] != "") {
    let sunhatFirstRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "Sunhat",
      formData["sunhatSizeId"],
      formData["sunhatColourId"],
      formData["sunhatCashCardId"],
      formData["sunhatQtyId"],
      formData["sunhatStaffId"],
      formData["sunhatNotesId"],
      userAndTimestamp
    ];
    salesArray.push(sunhatFirstRow);
  }
  // Get all extra sunhats
  while (formData["sunhatQtyId" + hoodieIncrement] && formData["sunhatQtyId" + hoodieIncrement] != "") {
    // Add the row to result array
    let sunhatRow = [
      formData["datePickerInput"],
      formData["sellerId"],
      "Sunhat",
      formData["sunhatSizeId_" + sunhatIncrement],
      formData["sunhatColourId_" + sunhatIncrement],
      formData["sunhatCashCardId_" + sunhatIncrement],
      formData["sunhatQtyId_" + sunhatIncrement],
      formData["sunhatStaffId_" + sunhatIncrement],
      formData["sunhatNotesId_" + sunhatIncrement],
      userAndTimestamp
    ];
    sunhatIncrement++;
    salesArray.push(sunhatRow);
  }
  console.log("salesArray", salesArray);
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
 * Display product modal
 */
function displayAddProductForm() {
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
  // Hoodie colours
  let hoodieColours = settingsData[settingsDataHeaders.indexOf("Colour")][items.indexOf("Hoodie")].split(",");
  let htmlHoodieColoursOptions = "<option value=" + hoodieColours[0] + " selected >" + hoodieColours[0] + "</option > ";
  hoodieColours.slice(1).forEach(function (val) {
    htmlHoodieColoursOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // Hoodie sizes
  let hoodieSizes = settingsData[items.indexOf("Hoodie")][settingsDataHeaders.indexOf("Size")].split(",");
  let htmlHoodieSizesOptions = "<option value=" + hoodieSizes[0] + " selected >" + hoodieSizes[0] + "</option > ";
  hoodieSizes.slice(1).forEach(function (val) {
    htmlHoodieSizesOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // teeShirt colours
  let teeShirtColours = settingsData[items.indexOf("T-Shirt")][settingsDataHeaders.indexOf("Colour")].split(",");
  let htmlteeShirtColoursOptions = "<option value=" + teeShirtColours[0] + " selected >" + teeShirtColours[0] + "</option > ";
  teeShirtColours.slice(1).forEach(function (val) {
    htmlteeShirtColoursOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // teeShirt sizes
  let teeShirtSizes = settingsData[items.indexOf("T-Shirt")][settingsDataHeaders.indexOf("Size")].split(",");
  let htmlteeShirtSizesOptions = "<option value=" + teeShirtSizes[0] + " selected >" + teeShirtSizes[0] + "</option > ";
  teeShirtSizes.slice(1).forEach(function (val) {
    htmlteeShirtSizesOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // Paiment type
  let paimentType = settingsData.map(x => x[settingsDataHeaders.indexOf("Paiment type")]).filter(n => n);
  let htmlPaimentTypeOptions = "<option value=" + paimentType[1] + " selected >" + paimentType[1] + "</option > ";
  paimentType.slice(2).forEach(function (row) {
    htmlPaimentTypeOptions += '<option value=' + row + '>' + row + "</option>";
  });
  // Beanie size
  let beanieSizes = settingsData[items.indexOf("Beanie")][settingsDataHeaders.indexOf("Size")].split(",");
  let htmlBeanieSizesOptions = "<option value=" + beanieSizes[0] + " selected >" + beanieSizes[0] + "</option > ";
  beanieSizes.slice(1).forEach(function (val) {
    htmlBeanieSizesOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // beanie colours
  let beanieColours = settingsData[items.indexOf("Beanie")][settingsDataHeaders.indexOf("Colour")].split(",");
  let htmlBeanieColoursOptions = "<option value=" + beanieColours[0] + " selected >" + beanieColours[0] + "</option > ";
  beanieColours.slice(1).forEach(function (val) {
    htmlBeanieColoursOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // Sunhat size
  let sunhatSizes = settingsData[items.indexOf("Sunhat")][settingsDataHeaders.indexOf("Size")].split(",");
  let htmlsunhatSizesOptions = "<option value=" + sunhatSizes[0] + " selected >" + sunhatSizes[0] + "</option > ";
  sunhatSizes.slice(1).forEach(function (val) {
    htmlsunhatSizesOptions += '<option value=' + val + '>' + val + "</option>";
  });
  // Sunhat colours
  let sunhatColours = settingsData[items.indexOf("Sunhat")][settingsDataHeaders.indexOf("Colour")].split(",");
  let htmlSunhatColoursOptions = "<option value=" + sunhatColours[0] + " selected >" + sunhatColours[0] + "</option > ";
  sunhatColours.slice(1).forEach(function (val) {
    htmlSunhatColoursOptions += '<option value=' + val + '>' + val + "</option>";
  });

  let mapExample = {
    "productName": "Test new Product",
    "sizes": ["XS", "S"],
    "colours": ["white", "blue"]
  }

  let modal = HtmlService.createTemplateFromFile("Form");
  modal.sellersOptions = htmlSellersOptions;
  modal.hoodieColourOptions = htmlHoodieColoursOptions;
  modal.hoodieSizesOptions = htmlHoodieSizesOptions;
  modal.teeShirtColourOptions = htmlteeShirtColoursOptions;
  modal.teeShirtSizesOptions = htmlteeShirtSizesOptions;
  modal.beanieSizesOptions = htmlBeanieSizesOptions;
  modal.beanieColourOptions = htmlBeanieColoursOptions;
  modal.sunhatSizesOptions = htmlsunhatSizesOptions;
  modal.sunhatColourOptions = htmlSunhatColoursOptions
  modal.paimentTypeOptions = htmlPaimentTypeOptions;
  modal.extraProductMap = mapExample

  modal = modal.evaluate();
  modal.setHeight(500).setWidth(850);
  SpreadsheetApp.getUi().showModalDialog(modal, "Add sales");
}
