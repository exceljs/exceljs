/* eslint-disable eqeqeq */
const path = require('path');
const util = require('util');

const Excel = require('../lib/exceljs.nodejs.js');

const filename = 'DataTableTestExcelJs.xlsx';
const filePath = path.resolve(__dirname, './data', filename);
const filePath3 = path.resolve(__dirname, './data', 'DataTableTest.xlsx');

(async () => {
  // crate table and write to file
  const workbook1 = new Excel.Workbook();
  const worksheet1 = workbook1.addWorksheet('Tab1');
  worksheet1.addTable({
    name: 'Table1',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleLight1',
      showRowStripes: true,
    },
    columns: [
      {name: 'FirstCol', totalsRowLabel: 'Totals:', filterButton: true},
      {name: 'SecondCol', totalsRowFunction: 'sum', filterButton: false},
    ],
    rows: [
      ['row1', 70.1],
      ['row2', 70.2],
      ['row3', 70.3],
    ],
  });
  const tableDataWrite = worksheet1.tables.Table1;
  await workbook1.xlsx.writeFile(filePath);

  //read from file
  const workbook2 = new Excel.Workbook();
  await workbook2.xlsx.readFile(filePath);
  const worksheet2 = workbook2.getWorksheet('Tab1');
  const tableDataRead = worksheet2.tables.Table1;

  const workbook3 = new Excel.Workbook();
  await workbook3.xlsx.readFile(filePath3);
  const worksheet3 = workbook3.getWorksheet('Tab1');
  const tableDataExcel = worksheet3.tables.Table1;

  const keys = ['name', 'ref', 'headerRow', 'totalsRow', 'style', 'columns', 'rows', 'autoFilterRef', 'tableRef'];

  let hasErrors = false;
  keys.forEach(key => {
    if (Array.isArray(tableDataWrite.table[key])) {
      if (!Array.isArray(tableDataRead.table[key]) || tableDataWrite.table[key].length !== tableDataRead.table[key].length) {
        hasErrors = true;
        console.log('Error TableData read:', `"${key}"`);
      }
      if (!Array.isArray(tableDataExcel.table[key]) || tableDataWrite.table[key].length !== tableDataExcel.table[key].length) {
        hasErrors = true;
        console.log('Error TableData read Excel:', `"${key}"`);
      }
    } else if (typeof tableDataWrite.table[key] == 'object' && tableDataWrite.table[key] != null) {
      if (!deepEqual(tableDataWrite.table[key], tableDataRead.table[key])) {
        hasErrors = true;
        console.log('Error TableData read:', `"${key}"`);
      }
      if (!deepEqual(tableDataWrite.table[key], tableDataExcel.table[key])) {
        hasErrors = true;
        console.log('Error TableData read Excel:', `"${key}"`);
      }
    } else {
      if (tableDataWrite.table[key] !== tableDataRead.table[key]) {
        hasErrors = true;
        console.log('Error Table Data read:', key);
      }
      if (tableDataWrite.table[key] !== tableDataExcel.table[key]) {
        hasErrors = true;
        console.log('Error table Data read Excel:', key);
      }
    }
  });

  if (hasErrors) {
    console.log('\nTableData written:');
    console.log(util.inspect(tableDataWrite.table, false, 10));
    console.log('\nTableData read:');
    console.log(util.inspect(tableDataRead.table, false, 10));
    console.log('\nTableData read from Excel:');
    console.log(util.inspect(tableDataExcel.table, false, 10));
  } else {
    console.log('No errors');
  }
})();

function deepEqual(x, y) {
  return x && y && typeof x === 'object' && typeof y === 'object'
    ? Object.keys(x).length === Object.keys(y).length &&
        Object.keys(x).reduce(function(isEqual, key) {
          return isEqual && deepEqual(x[key], y[key]);
        }, true)
    : x === y;
}
