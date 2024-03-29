const MONTHS_ARRAY = [
  "Jan",
  "Feb",
  "March",
  "April",
  "May",
  "June",
  "July",
  "Aug",
  "Sept",
  "Oct",
  "Nov",
  "Dec",
];

const INVENTORY_CELL = "B7";
const NUMBER_OF_COL_TODELETE = 10;
const ITEM_COL = 1;
const COLOURS_COL = 2;

function formatFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Template");
  for (let i=0;i<MONTHS_ARRAY.length; i++) {
    let sheet = ss.getSheetByName(MONTHS_ARRAY[i]);
    ss.deleteSheet(sheet);
    templateSheet.copyTo(ss).setName(MONTHS_ARRAY[i])
  }
}