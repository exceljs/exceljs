/* eslint-disable eqeqeq */
const path = require('path');
const util = require('util');

const Excel = require('../lib/exceljs.nodejs.js');

const filename = 'DataTableTestExcelJs.xlsx';
const filePath = path.resolve(__dirname, './data', filename);
//const filePath3 = path.resolve(__dirname, './data', 'DataTableTest.xlsx');

// const colCache = require('../lib/utils/col-cache');
// console.log(colCache.l2n('AA'), colCache.n2l(27));

(async () => {
  //read from file
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Tab1');
  const tableData = worksheet.tables.Table1;
  console.log(util.inspect(tableData.table, false, 5));
})();
