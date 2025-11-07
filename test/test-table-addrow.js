// Test table getTable().addRow() functionality
// This tests the fix from rmartin93/exceljs-fork commit 6b77cea
// Reproduces the workflow: load an Excel file with a table, then add rows to it

/* eslint-disable */

function main(filepath) {
  const Excel = require('../lib/exceljs.nodejs.js');

  const workbook = new Excel.Workbook();

  // Create a worksheet with a table
  const worksheet = workbook.addWorksheet('Data');

  // Add initial table with data
  worksheet.addTable({
    name: 'MyTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      theme: 'TableStyleMedium2',
      showRowStripes: true,
    },
    columns: [
      {name: 'Name', filterButton: true},
      {name: 'Age', filterButton: true},
      {name: 'Email', filterButton: true},
    ],
    rows: [
      ['Alice', 30, 'alice@example.com'],
      ['Bob', 25, 'bob@example.com'],
      ['Charlie', 35, 'charlie@example.com'],
    ],
  });

  console.log('Initial table created with 3 rows');

  // Test 1: Get the table and add a row (this used to crash)
  const table = worksheet.getTable('MyTable');

  console.log('Table retrieved:', table.name);
  console.log('Initial table ref:', table.ref);
  console.log('Initial table rows:', table.table.rows.length);
  console.log('Initial autoFilterRef:', table.autoFilterRef);

  // Add a new row using addRow (this is what was broken)
  console.log('\nAdding new row with addRow()...');
  table.addRow(['Diana', 28, 'diana@example.com']);

  console.log('After addRow:');
  console.log('Table ref:', table.ref);
  console.log('Table rows:', table.table.rows.length);
  console.log('AutoFilterRef:', table.autoFilterRef);

  // Add another row
  console.log('\nAdding another row...');
  table.addRow(['Eve', 32, 'eve@example.com']);

  console.log('After second addRow:');
  console.log('Table ref:', table.ref);
  console.log('Table rows:', table.table.rows.length);
  console.log('AutoFilterRef:', table.autoFilterRef);

  // Test 2: Create a second sheet to test loading a table
  const worksheet2 = workbook.addWorksheet('LoadTest');
  worksheet2.addTable({
    name: 'LoadTestTable',
    ref: 'B2',
    headerRow: true,
    totalsRow: false,
    columns: [
      {name: 'Product', filterButton: true},
      {name: 'Price', filterButton: true, style: {numFmt: '$#,##0.00'}},
      {name: 'Quantity', filterButton: true},
    ],
    rows: [
      ['Widget A', 19.99, 10],
      ['Widget B', 29.99, 5],
    ],
  });

  console.log('\n=== Test 2: Load and modify table ===');

  // Simulate loading: save and reload the workbook
  save(workbook, filepath, () => {
    console.log('\nWorkbook saved. Now testing load and modify...');

    const workbook2 = new Excel.Workbook();
    workbook2.xlsx
      .readFile(filepath)
      .then(() => {
        console.log('Workbook loaded successfully');

        const loadedSheet = workbook2.getWorksheet('LoadTest');
        const loadedTable = loadedSheet.getTable('LoadTestTable');

        console.log('Loaded table:', loadedTable.name);
        console.log('Loaded table ref:', loadedTable.ref);
        console.log('Loaded table rows:', loadedTable.table.rows.length);
        console.log('Loaded autoFilterRef:', loadedTable.autoFilterRef);

        // This is the critical test: adding rows to a loaded table
        console.log('\nAdding row to loaded table...');
        loadedTable.addRow(['Widget C', 39.99, 15]);

        console.log('After addRow to loaded table:');
        console.log('Table ref:', loadedTable.ref);
        console.log('Table rows:', loadedTable.table.rows.length);
        console.log('AutoFilterRef:', loadedTable.autoFilterRef);

        // Save the modified workbook
        const outputPath = filepath.replace('.xlsx', '-modified.xlsx');
        workbook2.xlsx.writeFile(outputPath).then(() => {
          console.log('\nModified workbook saved to:', outputPath);
          console.log('SUCCESS: All table addRow tests passed!');
        });
      })
      .catch(err => {
        console.error('ERROR loading workbook:', err);
        process.exit(1);
      });
  });
}

function save(workbook, filepath, callback) {
  const HrStopwatch = require('./utils/hr-stopwatch');
  const stopwatch = new HrStopwatch();
  stopwatch.start();

  workbook.xlsx
    .writeFile(filepath)
    .then(() => {
      const microseconds = stopwatch.microseconds;
      console.log('Initial save done.');
      console.log('Time taken:', microseconds);
      if (callback) callback();
    })
    .catch(err => {
      console.error('ERROR saving:', err);
      process.exit(1);
    });
}

const [, , filepath] = process.argv;
main(filepath || 'test-table-addrow.xlsx');
