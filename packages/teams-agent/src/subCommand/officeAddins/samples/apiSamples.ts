export const apiSampleData = {
  "samples": [
    {
      "namespace": "Excel.Workbook",
      "name": "Worksheets",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-preview#excel-excel-workbook-worksheets-member",
      "sample": "await Excel.run(async (context) => {\n    let sheets = context.workbook.worksheets;\n    sheets.load('name');\n    await context.sync();\n    console.log(sheets.items);\n});",
      "scenario": "Get all worksheets's name in the workbook"
    },
    {
      "namespace": "Excel.Workbook.Worksheets",
      "name": "add",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview#excel-excel-worksheetcollection-add-member(1)",
      "sample": "await Excel.run(async (context) => {\n    let sheets = context.workbook.worksheets;\n    let sheet = sheets.add('Sheet2');\n    sheet.activate();\n    await context.sync();\n});",
      "scenario": "Add a new worksheet named 'Sheet2' to the workbook"
    },
    {
      "namespace": "Excel.Workbook.Worksheets",
      "name": "getItem",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview#excel-excel-worksheetcollection-getitem-member(1)",
      "sample": "await Excel.run(async (context) => {\n    let sheets = context.workbook.worksheets;\n    let sheet = sheets.getItem('Sheet1');\n    sheet.activate();\n    await context.sync();\n});",
      "scenario": "Get the worksheet named 'Sheet1' and activate it"
    },
    {
      "namespace": "Excel.Workbook.Worksheets",
      "name": "getItemOrNullObject",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview#excel-excel-worksheetcollection-getitemornullobject-member(1)",
      "sample": "await Excel.run(async (context) => {\n    let sheets = context.workbook.worksheets;\n    let sheet = sheets.getItemOrNullObject('Sheet1');\n    sheet.load('name');\n    await context.sync();\n    console.log(sheet.name);\n});",
      "scenario": "Get the worksheet named 'Sheet1' and print its name if the Sheet1 exists"
    },
    {
      "namespace": "Excel.Worksheet",
      "name": "getRange",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview#excel-excel-worksheet-getrange-member(1)",
      "sample": "await Excel.run(async (context) => {\n    let sheet = context.workbook.worksheets.getItem('Sheet1');\n    let range = sheet.getRange('A1:B2');\n    let twoDimissionArray = range.values;\n    await context.sync();\n});",
      "scenario": "Set the values of the range A1:B2 and assign to a two dimission array (The result of twoDimissionArray is [[1, 2], [3, 4]])"
    },
    {
      "namespace": "Excel.Range",
      "name": "values",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview#excel-excel-range-values-member",
      "sample": "await Excel.run(async (context) => {\n    let sheet = context.workbook.worksheets.getItem('Sheet1');\n    let range = sheet.getRange('A1:B2');\n    range.values = [[1, 2], [3, 4]];\n    await context.sync();\n});",
      "scenario": "Set the values of the range A1:B2 to [[1, 2], [3, 4]] (A1=1, B1=2, A2=3, B2=4)"
    },
    {
      "namespace": "Excel.Worksheet",
      "name": "activate",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview#excel-excel-worksheet-activate-member(1)",
      "sample": "await Excel.run(async (context) => {\n    let sheet = context.workbook.worksheets.getItem('Sheet1');\n    sheet.activate();\n    await context.sync();\n});",
      "scenario": "Activate the worksheet named 'Sheet1' to make it as current working worksheet"
    },
    {
      "namespace": "Excel.Worksheet",
      "name": "name",
      "docLink": "https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview#excel-excel-worksheet-name-member",
      "sample": "await Excel.run(async (context) => {\n    let sheet = context.workbook.worksheets.getItem('Sheet1');\n    sheet.load('name');\n    await context.sync();\n    console.log(sheet.name);\n});",
      "scenario": "Get the name of the worksheet named 'Sheet1'"
    }
  ]
}
