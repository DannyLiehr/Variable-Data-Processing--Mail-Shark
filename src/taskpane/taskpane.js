/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

/*

Bookmark: code plan
,---.         |         ,---.|              
|    ,---.,---|,---.    |---'|    ,---.,---.
|    |   ||   ||---'    |    |    ,---||   |
`---'`---'`---'`---'    `    `---'`---^`   '

1. Create a count column. This is required for all sheets
2. If they wanted alternating rows, do that.
3. Split the data by how many up they need.
4. Detect specific columns and format them accordingly.

*/


document.getElementById("generate").addEventListener("click", async () => {
  try {
    await Excel.run(async (context) => {
      // Get the active workbook.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      // Load properties from the range
      range.load("address");
      // Update the fill color. This is for my own sanity.
      range.format.fill.color = "powderblue";
      await context.sync();
      // console.log(`The range address was ${range.address}.`); // The range address was:Sheet1!A1:O7407

      // Create a proper table from the sheet.
      let table= sheet.tables.add(range.address, true);
      table.name= "Data";
      await context.sync();

      // Grab all data in our sheet, and stick it in our ✨new✨ table.
      const index= new Array();
      let rangeVal= table.getDataBodyRange().load("values");
      await context.sync();
      let numZeros= String(rangeVal).length; // How many 0s
      console.log(numZeros)
      // Add the count + zeros (doesn't seem to keep the 0s tho)
      rangeVal.values.forEach((_row, i)=>{
        let n= String("0".repeat(numZeros))+String(i+1);  
        index.push([n]);
      });
      index.unshift(["index"]);
      table.columns.add(0, index);
      index.numberFormat = "0".repeat(numZeros);
      await context.sync();

      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
      sheet.activate();


      // SPLIT AT NUMBER

      

    });
  } catch (error) {
    console.error(error);
  }
})
