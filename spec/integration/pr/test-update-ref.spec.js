const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('PR - Update objects ref', () => {
    it('Reading test-update-ref.xlsx', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-update-ref.xlsx')
        .then(async () => {
          // Table tests
          let ws = wb.getWorksheet('Sheet1');
          // Table 1
          const table1 = ws.getTable('Table1');
          expect(table1.table.tableRef).to.equal('B12:D17');

          // Insert rows tests
          // insert a single row at beginning of file
          ws.insertRow(3, []);
          expect(table1.table.tableRef).to.equal('B13:D18');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B22:D25');

          // insert multiple rows at beginning of file
          ws.insertRows(5, [[], [], []]);
          expect(table1.table.tableRef).to.equal('B16:D21');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B25:D28');

          // insert a single row inside the first table
          ws.insertRow(17, []);
          expect(table1.table.tableRef).to.equal('B16:D22');
          expect(ws.getRow(17).getCell('B').value).to.equal(null);
          expect(ws.getTable('Table13').table.tableRef).to.equal('B26:D29');

          // insert multiple rows inside the first table
          ws.insertRows(17, [
            [null, 'test1'],
            [null, 'test2'],
          ]);
          expect(table1.table.tableRef).to.equal('B16:D24');
          expect(ws.getRow(17).getCell('B').value).to.equal('test1');
          expect(ws.getRow(18).getCell('B').value).to.equal('test2');
          expect(ws.getRow(19).getCell('B').value).to.equal(null);
          expect(ws.getTable('Table13').table.tableRef).to.equal('B28:D31');

          // insert a single row between the two tables
          ws.insertRow(26, []);
          expect(table1.table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B29:D32');

          // insert multiple rows between the two tables
          ws.insertRows(27, [[], []]);
          expect(table1.table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B31:D34');

          // Delete rows tests
          // Delete multiple rows between the two tables
          ws.spliceRows(27, 2);
          expect(table1.table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B29:D32');

          // Delete a single row between the two tables
          ws.spliceRows(26, 1);
          expect(table1.table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B28:D31');

          // Delete multiple rows that overlaps with a table
          ws.spliceRows(27, 3);
          expect(table1.table.tableRef).to.equal('B16:D24');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B27:D29');

          // Delete multiple rows inside the first table
          ws.spliceRows(18, 2);
          expect(table1.table.tableRef).to.equal('B16:D22');
          expect(ws.getRow(17).getCell('B').value).to.equal('test1');
          expect(ws.getTable('Table13').table.tableRef).to.equal('B25:D27');

          // Delete a single row inside the first table
          ws.spliceRows(17, 1);
          expect(table1.table.tableRef).to.equal('B16:D21');
          expect(ws.getRow(17).getCell('B').value).to.equal(
            'Sukhoi Su-27 Flanker'
          );
          expect(ws.getTable('Table13').table.tableRef).to.equal('B24:D26');

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
  });
});
