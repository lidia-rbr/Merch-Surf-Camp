

function importOldJanData() {
  const oldSs = SpreadsheetApp.openById("1MNNPyj-aq9fv0R6N3KqyAmvZ7xSufbyW5JrWU-HSK3w");
  const oldJan = oldSs.getSheetByName("Jan");

  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  const currentJan = currentSs.getSheetByName("Jan");

  // Old : Date	XS	S	M	L	XL	XXL	Colour	Payment type	Staff?

  // Target: Sale date	Seller	Item	Size	Colour	Cash/card	Qty	Staff	Notes	Added by

  // Tee shirt A34:J75
  const oldTShirtData = oldJan.getRange("A34:J75").getDisplayValues();
  const oldTShirtDataHeaders = oldTShirtData[0];

  const newFormattedData = [];

  for (let i = 1; i < oldTShirtData.length; i++) {
    for (j = 1; j < oldTShirtData[0].length - 3; j++) {
      if (oldTShirtData[i][j] != "") {
        let staff = (oldTShirtData[i][oldTShirtDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 0, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "T-Shirt",
          oldTShirtDataHeaders[j],
          oldTShirtData[i][oldTShirtDataHeaders.indexOf("Colour")],
          oldTShirtData[i][oldTShirtDataHeaders.indexOf("Payment type")],
          oldTShirtData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }

  // Hoodies L34:U54
  const oldHoodiesData = oldJan.getRange("L34:U54").getDisplayValues();
  const oldHoodiesDataHeaders = oldHoodiesData[0];

  for (let i = 1; i < oldHoodiesData.length; i++) {
    for (j = 1; j < oldHoodiesData[0].length - 3; j++) {
      if (oldHoodiesData[i][j] != "") {
        let staff = (oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 0, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "Hoodie",
          oldHoodiesDataHeaders[j],
          oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Colour")],
          oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Payment type")],
          oldHoodiesData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }

  // Beanie W34:AF38
  const oldBeaniesData = oldJan.getRange("W34:AF38").getDisplayValues();
  const oldBeaniesDataHeaders = oldBeaniesData[0];

  for (let i = 1; i < oldBeaniesData.length; i++) {
    for (j = 1; j < oldBeaniesData[0].length - 3; j++) {
      if (oldBeaniesData[i][j] != "") {
        let staff = (oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 0, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "Beanie",
          oldBeaniesDataHeaders[j],
          oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Colour")],
          oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Payment type")],
          oldBeaniesData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }
  currentJan.getRange(4, 13, newFormattedData.length, newFormattedData[0].length).setValues(newFormattedData)

}



function importOldFebData() {
  const oldSs = SpreadsheetApp.openById("1MNNPyj-aq9fv0R6N3KqyAmvZ7xSufbyW5JrWU-HSK3w");
  const oldJan = oldSs.getSheetByName("Feb");

  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  const currentJan = currentSs.getSheetByName("Feb");

  // Old : Date	XS	S	M	L	XL	XXL	Colour	Payment type	Staff?

  // Target: Sale date	Seller	Item	Size	Colour	Cash/card	Qty	Staff	Notes	Added by

  // Tee shirt A34:J75
  const oldTShirtData = oldJan.getRange("A34:J72").getDisplayValues();
  const oldTShirtDataHeaders = oldTShirtData[0];

  const newFormattedData = [];

  for (let i = 1; i < oldTShirtData.length; i++) {
    for (j = 1; j < oldTShirtData[0].length - 3; j++) {
      if (oldTShirtData[i][j] != "") {
        let staff = (oldTShirtData[i][oldTShirtDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 1, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "T-Shirt",
          oldTShirtDataHeaders[j],
          oldTShirtData[i][oldTShirtDataHeaders.indexOf("Colour")],
          oldTShirtData[i][oldTShirtDataHeaders.indexOf("Payment type")],
          oldTShirtData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }

  // Hoodies L34:U54
  const oldHoodiesData = oldJan.getRange("L34:U52").getDisplayValues();
  const oldHoodiesDataHeaders = oldHoodiesData[0];

  for (let i = 1; i < oldHoodiesData.length; i++) {
    for (j = 1; j < oldHoodiesData[0].length - 3; j++) {
      if (oldHoodiesData[i][j] != "") {
        let staff = (oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 1, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "Hoodie",
          oldHoodiesDataHeaders[j],
          oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Colour")],
          oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Payment type")],
          oldHoodiesData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }

  // Beanie W34:AF38
  const oldBeaniesData = oldJan.getRange("W34:AF43").getDisplayValues();
  const oldBeaniesDataHeaders = oldBeaniesData[0];

  for (let i = 1; i < oldBeaniesData.length; i++) {
    for (j = 1; j < oldBeaniesData[0].length - 3; j++) {
      if (oldBeaniesData[i][j] != "") {
        let staff = (oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 1, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "Beanie",
          oldBeaniesDataHeaders[j],
          oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Colour")],
          oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Payment type")],
          oldBeaniesData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }
  currentJan.getRange(4, 13, newFormattedData.length, newFormattedData[0].length).setValues(newFormattedData)

}



function importOldMarchData() {
  const oldSs = SpreadsheetApp.openById("1MNNPyj-aq9fv0R6N3KqyAmvZ7xSufbyW5JrWU-HSK3w");
  const oldJan = oldSs.getSheetByName("Mar");

  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  const currentJan = currentSs.getSheetByName("March");

  // Old : Date	XS	S	M	L	XL	XXL	Colour	Payment type	Staff?

  // Target: Sale date	Seller	Item	Size	Colour	Cash/card	Qty	Staff	Notes	Added by

  // Tee shirt A34:J75
  const oldTShirtData = oldJan.getRange("A34:J54").getDisplayValues();
  const oldTShirtDataHeaders = oldTShirtData[0];

  const newFormattedData = [];

  for (let i = 1; i < oldTShirtData.length; i++) {
    for (j = 1; j < oldTShirtData[0].length - 3; j++) {
      if (oldTShirtData[i][j] != "") {
        let staff = (oldTShirtData[i][oldTShirtDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 2, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "T-Shirt",
          oldTShirtDataHeaders[j],
          oldTShirtData[i][oldTShirtDataHeaders.indexOf("Colour")],
          oldTShirtData[i][oldTShirtDataHeaders.indexOf("Payment type")],
          oldTShirtData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }

  // Hoodies L34:U54
  const oldHoodiesData = oldJan.getRange("L34:U47").getDisplayValues();
  const oldHoodiesDataHeaders = oldHoodiesData[0];

  for (let i = 1; i < oldHoodiesData.length; i++) {
    for (j = 1; j < oldHoodiesData[0].length - 3; j++) {
      if (oldHoodiesData[i][j] != "") {
        let staff = (oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 2, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "Hoodie",
          oldHoodiesDataHeaders[j],
          oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Colour")],
          oldHoodiesData[i][oldHoodiesDataHeaders.indexOf("Payment type")],
          oldHoodiesData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }

  // Beanie W34:AF38
  const oldBeaniesData = oldJan.getRange("W34:AF43").getDisplayValues();
  const oldBeaniesDataHeaders = oldBeaniesData[0];

  for (let i = 1; i < oldBeaniesData.length; i++) {
    for (j = 1; j < oldBeaniesData[0].length - 3; j++) {
      if (oldBeaniesData[i][j] != "") {
        let staff = (oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Payment type")] == "Y") ? true : false;
        let dateTab = oldTShirtData[i][0].split("-");// 14-03-24
        let fixedDate = new Date(2024, 2, parseInt(dateTab[0]) + 1);
        let row = [
          fixedDate,
          "??",
          "Beanie",
          oldBeaniesDataHeaders[j],
          oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Colour")],
          oldBeaniesData[i][oldBeaniesDataHeaders.indexOf("Payment type")],
          oldBeaniesData[i][j],
          staff,
          "",
          "Lidia (script)"
        ];
        newFormattedData.push(row);
      }
    }
  }
  currentJan.getRange(4, 13, newFormattedData.length, newFormattedData[0].length).setValues(newFormattedData)

}


function testDAte() {
  const oldSs = SpreadsheetApp.openById("1MNNPyj-aq9fv0R6N3KqyAmvZ7xSufbyW5JrWU-HSK3w");
  const oldJan = oldSs.getSheetByName("Mar");
  const oldJanData = oldJan.getDataRange().getDisplayValues();
  console.log(oldJanData[46]) // 14-03-24


}
