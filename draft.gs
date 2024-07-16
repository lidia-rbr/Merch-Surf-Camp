function myFunction() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const richText = sheet.getRange("F11").getRichTextValue();
    const value = sheet.getRange("F11").getValue();
    const r1c1 = sheet.getRange("F11").getFormulaR1C1();
    const formula = sheet.getRange("F11").getFormula()
    const dataSource = sheet.getRange("F11").getDataSourceFormula()
    const richTextD = sheet.getRange("D8").getRichTextValue();
    const valueD = sheet.getRange("D8").getValue();
    const r1c1D = sheet.getRange("D8").getFormulaR1C1();
    const formulaD = sheet.getRange("D8").getFormula()
    const dataSourceD = sheet.getRange("D8").getDataSourceFormula()
    console.log("richText", richText)
    console.log("value",value)
    console.log("r1c1",r1c1);
    console.log("formula", formula);
    console.log("dataSource",dataSource);

        console.log("richTextD", richTextD)
    console.log("valueD",valueD)
    console.log("r1c1D",r1c1D);
    console.log("formulaD", formulaD);
    console.log("dataSourceD",dataSourceD);


  
}
