function myFunction() {
  // Assuming the form response is stored in a variable named formData
var formData = {
  datePickerInput: '2024-03-24',
  'T_Shirt-ColourId_2': ' White',
  'Beanie-QtyId': '',
  'T_Shirt-QtyId_1': '1',
  sellerId: 'Diogo',
  'T_Shirt-CashCardId_2': 'Cash',
  'Sunhat-SizeId': '-',
  'Beanie-ColourId': 'Grey',
  'T_Shirt-StaffId': false,
  'Hoodie-QtyId': '4',
  'T_Shirt-CashCardId_1': 'Cash',
  'T_Shirt-SizeId': 'XS',
  'T_Shirt-StaffId_1': false,
  'Hoodie-CashCardId': 'Cash',
  'temp-NotesId': '',
  'T_Shirt-NotesId_2': '',
  'T_Shirt-ColourId_1': ' Grey',
  'Hoodie-StaffId': false,
  'T_Shirt-NotesId_1': '',
  'T_Shirt-StaffId_2': false,
  'T_Shirt-SizeId_1': 'XS',
  'T_Shirt-SizeId_2': 'XS',
  'Hoodie-SizeId': 'XS',
  'T_Shirt-CashCardId': 'Cash',
  'Beanie-SizeId': 'XS',
  'Sunhat-QtyId': '',
  'Sunhat-StaffId': false,
  'temp-QtyId': '',
  'Hoodie-NotesId': '',
  'Hoodie-ColourId': 'Black',
  'temp-CashCardId': '',
  'temp-SizeId': '',
  'Sunhat-CashCardId': 'Cash',
  'Sunhat-ColourId': '-',
  'T_Shirt-QtyId': '2',
  'temp-ColourId': '',
  'T_Shirt-NotesId': '',
  'Beanie-NotesId': '',
  'T_Shirt-QtyId_2': '2',
  'Sunhat-NotesId': '',
  'T_Shirt-ColourId': 'Black',
  'Beanie-StaffId': false,
  'temp-StaffId': false,
  'Beanie-CashCardId': 'Cash'
};

// Initialize an empty array to store the formatted data
var dataArray = [];

// Function to extract product name from the key
function extractProductName(key) {
  return key.split("-")[0].replaceAll("_"," "); // Extracts the product name before the first hyphen
}

// Iterate over the keys in formData
for (var key in formData) {
  if (formData.hasOwnProperty(key)) {
    // Extract product name, size, colour, and quantity from the key
    var productName = extractProductName(key);
    console.log(productName)
  }
}

// Display the resulting array
console.log(dataArray);
}
