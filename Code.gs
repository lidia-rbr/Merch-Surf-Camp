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
 * Treat client side data and print inputs to sheets
 * @param {object} formData
 * @return {string} res
 */
function treatAndPrintClientSideData() {
  let formData = { 'T-Shirt-NotesId_1': '',
  'sunglasses-QtyId_1': '2',
  'Hoodie-QtyId_1': '2',
  'Hoodie-CashCardId_1': 'Cash',
  'temp-SizeId_1': '',
  'temp-StaffId_1': false,
  'Hoodie-NotesId_1': '',
  'sunglasses-ColourId_1': 'red',
  'sunglasses-CashCardId_1': 'Cash',
  'temp-CashCardId_1': '',
  'Beanie-QtyId_1': '2',
  'Hoodie-ColourId_1': 'Black',
  'sunglasses-SizeId_1': '',
  'T-Shirt-StaffId_1': false,
  'sellerId': 'Diogo',
  'sunglasses-StaffId_1': false,
  'Sunhat-SizeId_1': '-',
  'T-Shirt-CashCardId_1': 'Cash',
  'T-Shirt-QtyId_1': '2',
  'Beanie-NotesId_1': '',
  'Beanie-CashCardId_1': 'Cash',
  'Sunhat-CashCardId_1': 'Cash',
  'Beanie-ColourId_1': 'Grey',
  'Sunhat-NotesId_1': '',
  'Sunhat-StaffId_1': false,
  'T-Shirt-SizeId_1': 'XS',
  'temp-QtyId_1': '',
  'Hoodie-SizeId_1': 'XS',
  'datePickerInput': '2024-03-26',
  'temp-ColourId_1': '',
  'T-Shirt-ColourId_1': 'Black',
  'temp-NotesId_1': '',
  'Sunhat-QtyId_1': '2',
  'sunglasses-NotesId_1': '',
  'Hoodie-StaffId_1': false,
  'Beanie-SizeId_1': 'XS',
  'Beanie-StaffId_1': true,
  'Sunhat-ColourId_1': '-' };

  // Get unique items list first
  let items = [];
  for (let i = 0; i < Object.keys(formData).length; i++) {
    let item = Object.keys(formData)[i].split("_")[0];
    if (item != "datePickerInput"
      && item != "sellerId"
      && item != "temp"
      && items.indexOf(item) == -1) {
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
    while (formData[items[i] + "_QtyId_" + index]) {
      if (formData[items[i] + "_QtyId_" + index] != "") {
        // add row
        let row = [
          formData["datePickerInput"],
          formData["sellerId"],
          items[i].replaceAll("-", " "),
          formData[items[i] + "SizeId" + index],
          formData[items[i] + "ColourId" + index],
          formData[items[i] + "CashCardId" + index],
          formData[items[i] + "QtyId" + index],
          formData[items[i] + "StaffId" + index],
          formData[items[i] + "NotesId" + index],
          userAndTimestamp
        ]
      }

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
