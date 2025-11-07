// Test pivot tables with null/undefined values
// This should not throw "undefined not in sharedItems" error

/* eslint-disable */

function main(filepath) {
  const Excel = require('../lib/exceljs.nodejs.js');

  const workbook = new Excel.Workbook();

  // Create data with null/undefined values
  const dataSheet = workbook.addWorksheet('Data');
  dataSheet.addRows([
    ['Region', 'Product', 'Territory', 'Quarter', 'Amount'],
    ['North', 'Widget A', 'NE', 'Q1', 1000],
    ['South', 'Widget B', null, 'Q1', 1500], // null territory
    ['North', 'Widget A', undefined, 'Q2', 1200], // undefined territory
    ['South', 'Widget B', 'SE', 'Q2', 1800],
    ['East', 'Widget C', null, 'Q1', 2000], // null territory
    ['West', 'Widget C', 'NW', 'Q2', 2200],
  ]);

  // Create pivot table that includes the field with null/undefined values
  const pivotSheet = workbook.addWorksheet('Pivot with Nulls');
  pivotSheet.addPivotTable({
    sourceSheet: dataSheet,
    rows: ['Region', 'Territory'], // Territory column has nulls
    columns: ['Quarter'],
    values: ['Amount'],
    metric: 'sum',
  });

  save(workbook, filepath);
}

function save(workbook, filepath) {
  const HrStopwatch = require('./utils/hr-stopwatch');
  const stopwatch = new HrStopwatch();
  stopwatch.start();

  workbook.xlsx
    .writeFile(filepath)
    .then(() => {
      const microseconds = stopwatch.microseconds;
      console.log('Done. Successfully generated pivot table with null values.');
      console.log('Time taken:', microseconds);
    })
    .catch(err => {
      console.error('ERROR:', err.message);
      process.exit(1);
    });
}

const [, , filepath] = process.argv;
main(filepath || 'test-pivot-nulls.xlsx');
