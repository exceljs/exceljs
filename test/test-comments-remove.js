const Excel = require('../lib/exceljs.nodejs');

const wb = new Excel.Workbook();
wb.xlsx
  .readFile(require.resolve('./data/comments.xlsx'))
  .then(() => {
    wb.worksheets.forEach(sheet => {
      sheet.getCell('A1').removeNote();
      sheet.getCell('B1').removeNote();
    });

    return wb.xlsx.writeFile(`${__dirname}/data/test.xlsx`);
  })
  .catch(console.error);
