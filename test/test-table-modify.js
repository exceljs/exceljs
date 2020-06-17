/* eslint-disable eqeqeq */
const path = require('path');
const util = require('util');

const Excel = require('../lib/exceljs.nodejs.js');

const filename = 'DataTableTest.xlsx';
const filePath = path.resolve(__dirname, './data', filename);
const filePathSave = path.resolve(__dirname, './data', 'DataTableTestSave.xlsx');

(async () => {
  //read from file
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Tab1');
  const dataTable = worksheet.getTable('Table1');
  dataTable.addRow(['Row4', 80]);
  dataTable.commit();
  //const tableData = worksheet.tables.Table1;
  console.log(util.inspect(dataTable, false, 5));
  await workbook.xlsx.writeFile(filePathSave);
})();
