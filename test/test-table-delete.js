/* eslint-disable eqeqeq */
const path = require('path');

const Excel = require('../lib/exceljs.nodejs.js');

const filename = 'DataTableTest.xlsx';
const filePath = path.resolve(__dirname, './data', filename);
const filePathSave = path.resolve(__dirname, './data', 'DataTableTestSave.xlsx');

(async () => {
  //read from file
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Tab1');
  worksheet.removeTable('Table1', true);

  await workbook.xlsx.writeFile(filePathSave);
})();
