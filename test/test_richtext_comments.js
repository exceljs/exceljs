var Excel = require('exceljs');
 
var wb = new Excel.Workbook();
 
wb.xlsx.readFile("./data/richtext_comments.xlsx").then(function () {    
    let sheet = wb.getWorksheet('cn');
    sheet.eachRow(function (row, id) {
        for(let i = 1; i <= row.cellCount; i++) {
            let cell = row.getCell(i)
            if(cell.type == 8){ // Cell.Types.RichText
                cell.value = cell.value;
                //assigning value will create a new RichTextValue object with model.type set to Cell.Types.String
            }
        }
    });
    wb.xlsx.writeFile("./data/result.xlsx");
})