const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('PR - Update objects ref', () => {
    it('Reading test-update-ref.xlsx - rows', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-update-ref.xlsx')
        .then(async () => {
          // Table tests
          let ws = wb.getWorksheet('Sheet1');
          expect(ws.getTable('Table1').table.tableRef).to.equal('B12:D17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B21:D24');

          // Insert rows tests
          // insert a single row at beginning of file
          ws.insertRow(3, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B13:D18');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B22:D25');

          // insert multiple rows at beginning of file
          ws.insertRows(5, [[], [], []]);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D21');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B25:D28');

          // insert a single row inside the first table
          ws.insertRow(17, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D22');
          expect(ws.getRow(17).getCell('B').value).to.equal(null);
          expect(ws.getTable('Table13').table.tableRef).to.equal('B26:D29');

          // insert multiple rows inside the first table
          ws.insertRows(17, [
            [null, 'test1'],
            [null, 'test2'],
          ]);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D24');
          expect(ws.getRow(17).getCell('B').value).to.equal('test1');
          expect(ws.getRow(18).getCell('B').value).to.equal('test2');
          expect(ws.getRow(19).getCell('B').value).to.equal(null);
          expect(ws.getTable('Table13').table.tableRef).to.equal('B28:D31');

          // insert a single row between the two tables
          ws.insertRow(26, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B29:D32');

          // insert multiple rows between the two tables
          ws.insertRows(27, [[], []]);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B31:D34');

          // Delete rows tests
          // Delete multiple rows between the two tables
          ws.spliceRows(27, 2);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B29:D32');

          // Delete a single row between the two tables
          ws.spliceRows(26, 1);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B28:D31');

          // Delete multiple rows that overlaps with a table
          ws.spliceRows(27, 3);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B27:D29');

          // Delete multiple rows inside the first table
          ws.spliceRows(18, 2);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D22');
          expect(ws.getRow(17).getCell('B').value).to.equal('test1');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B25:D27');

          // Delete a single row inside the first table
          ws.spliceRows(17, 1);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D21');
          expect(ws.getRow(17).getCell('B').value).to.equal(
            'Sukhoi Su-27 Flanker'
          );
          expect(ws.getTable('Table13').table.tableRef).to.equal('B24:D26');

          // delete multiple rows in a table and after
          ws.spliceRows(19, 5);
          expect(ws.getTable('Table1').table.tableRef).to.equal('B16:D18');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B19:D21');

          // Delete both tables
          ws.spliceRows(2, 50);
          expect(ws.getTables().length).to.equal(0);

          // Image tests
          ws = wb.getWorksheet('Sheet2');

          // Insert rows tests
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(11); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(19); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(16); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(27); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(38); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(51); // Image 3

          // Insert a single row at beginning of file
          ws.insertRow(5, []);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(12); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(20); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(17); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(28); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(39); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(52); // Image 3

          // Insert a single row at anchor point
          ws.insertRow(13, []);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(13); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(21); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(18); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(29); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(40); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(53); // Image 3

          // Insert multiple rows on first image
          ws.insertRows(15, [[], [], [], [], []]);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(13); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(21); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(23); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(34); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(45); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(58); // Image 3

          // Delete rows tests
          // Delete a single row at beginning of file
          ws.spliceRows(3, 1);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(12); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(20); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(22); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(33); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(44); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(57); // Image 3

          // Delete multiple rows at beginning of file
          ws.spliceRows(4, 5);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(7); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(15); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(17); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(28); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(39); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(52); // Image 3

          // Delete multiple rows at beginning of file to first anchor
          ws.spliceRows(4, 4);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(11); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(13); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(24); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(35); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(48); // Image 3

          // Delete multiple rows after first anchor
          ws.spliceRows(7, 10);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(11); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(6); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(17); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(25); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(38); // Image 3

          // Delete rows containing all images
          ws.spliceRows(1, 50);
          expect(Math.floor(ws.getImages()[0].range.tl.row)).to.equal(0); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.row)).to.equal(8); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.row)).to.equal(0); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.row)).to.equal(11); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.row)).to.equal(0); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.row)).to.equal(13); // Image 3
        });
    });

    it('Reading test-update-ref.xlsx - columns', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-update-ref.xlsx')
        .then(async () => {
          // Table tests
          let ws = wb.getWorksheet('Sheet1');
          // Table 1
          expect(ws.getTable('Table1').table.tableRef).to.equal('B12:D17');

          // insert a single column at beginning of file
          ws.spliceColumns(1, 0, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('C12:E17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('C21:E24');

          // insert multiple columns at beginning of file
          ws.spliceColumns(1, 0, [], [], []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:H17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:H24');

          // insert a single column at top left anchor
          ws.spliceColumns(6, 0, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:I17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:I24');

          // insert multiple columns at top left anchor
          ws.spliceColumns(6, 0, [], [], []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:L17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:L24');

          // insert a single column in table
          let values = Array(11).fill(null); // Define column values
          values.push('New Header');
          values.push('new value');
          ws.spliceColumns(11, 0, values);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:M17');
          expect(ws.getCell(12, 11).value).to.equal('New Header');
          expect(ws.getCell(13, 11).value).to.equal('new value');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:M24');

          // insert multiple columns in table
          values = Array(20).fill(null); // Define column values
          values.push('Header 1');
          values.push('value value');
          ws.spliceColumns(8, 0, [], values, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:P17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:P24');
          expect(ws.getCell(21, 9).value).to.equal('Header 1');
          expect(ws.getCell(22, 9).value).to.equal('value value');

          // insert a single column at bottom right anchor
          ws.spliceColumns(16, 0, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:Q17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:Q24');

          // insert a single column at bottom right anchor
          ws.spliceColumns(17, 0, [], [], []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:T17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:T24');

          // insert a single column at bottom right anchor
          ws.spliceColumns(21, 0, []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:T17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:T24');

          // insert a single column at bottom right anchor
          ws.spliceColumns(21, 0, [], [], []);
          expect(ws.getTable('Table1').table.tableRef).to.equal('F12:T17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('F21:T24');

          // Delete rows tests
          // delete a single column at beginning of file
          ws.spliceColumns(1, 1);
          expect(ws.getTable('Table1').table.tableRef).to.equal('E12:S17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('E21:S24');

          // delete multiple columns at beginning of file
          ws.spliceColumns(2, 2);
          expect(ws.getTable('Table1').table.tableRef).to.equal('C12:Q17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('C21:Q24');

          // delete a single column at the table anchor
          ws.spliceColumns(3, 1);
          expect(ws.getTable('Table1').table.tableRef).to.equal('C12:P17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('C21:P24');

          // delete multiple columns at the table anchor
          ws.spliceColumns(3, 3);
          expect(ws.getTable('Table1').table.tableRef).to.equal('C12:M17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('C21:M24');

          // delete multiple columns that overlaps with the table
          ws.spliceColumns(1, 5);
          expect(ws.getTable('Table1').table.tableRef).to.equal('A12:H17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('A21:H24');

          // delete a single column in a table
          expect(ws.getCell(13, 2).value).to.equal('new value');
          ws.spliceColumns(2, 1);
          expect(ws.getTable('Table1').table.tableRef).to.equal('A12:G17');
          expect(ws.getCell(13, 2).value).to.equal('16 000 kg');
          expect(ws.getTable('Table13').table.tableRef).to.equal('A21:G24');

          // delete multiple columns in a table and after
          ws.spliceColumns(6, 4);
          expect(ws.getTable('Table1').table.tableRef).to.equal('A12:E17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('A21:E24');

          // delete multiple columns in a table
          ws.spliceColumns(2, 2);
          expect(ws.getTable('Table1').table.tableRef).to.equal('A12:C17');
          expect(ws.getTable('Table13').table.tableRef).to.equal('A21:C24');

          // Delete both tables
          ws.spliceColumns(1, 20);
          expect(ws.getTables().length).to.equal(0);

          // Image tests
          ws = wb.getWorksheet('Sheet2');

          // Insert columns tests
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(2); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(5); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(10); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(13); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(5); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(9); // Image 3

          // Insert a single column at beginning of file
          ws.spliceColumns(1, 0, []);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(6); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(11); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(14); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(6); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(10); // Image 3

          // Insert multiple columns at beginning of file
          ws.spliceColumns(2, 0, [], [], []);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(6); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(9); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(14); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(17); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(9); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(13); // Image 3

          // Insert a single column at anchor point
          ws.spliceColumns(7, 0, []);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(7); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(10); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(15); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(18); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(10); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(14); // Image 3

          // Insert multiple columns on first image
          ws.spliceColumns(9, 0, [], [], [], []);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(7); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(10); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(19); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(22); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(14); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(18); // Image 3

          // Delete columns tests
          // Delete a single column at beginning of file
          ws.spliceColumns(3, 1);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(6); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(9); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(18); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(21); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(13); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(17); // Image 3

          // Delete multiple columns at beginning of file
          ws.spliceColumns(2, 3);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(6); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(15); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(18); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(10); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(14); // Image 3

          // Delete multiple columns at beginning of file to first anchor
          ws.spliceColumns(1, 5);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(0); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(10); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(13); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(5); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(9); // Image 3

          // Delete multiple columns after first anchor
          ws.spliceColumns(2, 5);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(0); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(5); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(8); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(0); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(4); // Image 3

          // Delete columns containing all images
          ws.spliceColumns(1, 20);
          expect(Math.floor(ws.getImages()[0].range.tl.col)).to.equal(0); // Image 1
          expect(Math.floor(ws.getImages()[0].range.br.col)).to.equal(3); // Image 1
          expect(Math.floor(ws.getImages()[1].range.tl.col)).to.equal(0); // Image 2
          expect(Math.floor(ws.getImages()[1].range.br.col)).to.equal(3); // Image 2
          expect(Math.floor(ws.getImages()[2].range.tl.col)).to.equal(0); // Image 3
          expect(Math.floor(ws.getImages()[2].range.br.col)).to.equal(4); // Image 3
        });
    });

    it('Reading test-update-ref.xlsx - cells', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-update-ref.xlsx')
        .then(async () => {
          // Table tests
          const ws = wb.getWorksheet('Sheet1');
          // Table 1
          expect(ws.getTable('Table1').table.tableRef).to.equal('B12:D17');

          // Change tl cell value (header)
          ws.getRow(12).getCell('B').value = 'Test';
          expect(ws.getTable('Table1').table.rows[0][0]).to.equal(
            'Sukhoi Su-27 Flanker'
          );

          // Change top left first cell value (row)
          ws.getRow(13).getCell('B').value = 'Fiesta';
          expect(ws.getTable('Table1').table.rows[0][0]).to.equal('Fiesta');

          // Change cell value in middle of the first table
          ws.getRow(15).getCell('C').value = 'Random info';
          expect(ws.getTable('Table1').table.rows[2][1]).to.equal(
            'Random info'
          );

          // Change br cell value
          ws.getRow(17).getCell('D').value = 'Another test';
          expect(ws.getTable('Table1').table.rows[4][2]).to.equal(
            'Another test'
          );
        });
    });
  });
});
