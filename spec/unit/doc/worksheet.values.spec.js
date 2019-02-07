var expect = require('chai').expect;

var _ = require('../../../lib/utils/under-dash');
var Excel = require('../../../excel');
var testUtils = require('../../utils/index');

describe('Worksheet', function() {
  describe('Values', function() {
    it('stores values properly', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      var now = new Date();

      // plain number
      ws.getCell('A1').value = 7;

      // simple string
      ws.getCell('B1').value = 'Hello, World!';

      // floating point
      ws.getCell('C1').value = 3.14;

      // 5 will be overwritten by the current date-time
      ws.getCell('D1').value = 5;
      ws.getCell('D1').value = now;

      // constructed string - will share recorded with B1
      ws.getCell('E1').value = ['Hello', 'World'].join(', ') + '!';

      // hyperlink
      ws.getCell('F1').value = {text: 'www.google.com', hyperlink: 'http://www.google.com'};

      // number formula
      ws.getCell('A2').value = {formula: 'A1', result: 7};

      // string formula
      ws.getCell('B2').value = {formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!'};

      // date formula
      ws.getCell('C2').value = {formula: 'D1', result: now};

      expect(ws.getCell('A1').value).to.equal(7);
      expect(ws.getCell('B1').value).to.equal('Hello, World!');
      expect(ws.getCell('C1').value).to.equal(3.14);
      expect(ws.getCell('D1').value).to.equal(now);
      expect(ws.getCell('E1').value).to.equal('Hello, World!');
      expect(ws.getCell('F1').value.text).to.equal('www.google.com');
      expect(ws.getCell('F1').value.hyperlink).to.equal('http://www.google.com');

      expect(ws.getCell('A2').value.formula).to.equal('A1');
      expect(ws.getCell('A2').value.result).to.equal(7);

      expect(ws.getCell('B2').value.formula).to.equal('CONCATENATE("Hello", ", ", "World!")');
      expect(ws.getCell('B2').value.result).to.equal('Hello, World!');

      expect(ws.getCell('C2').value.formula).to.equal('D1');
      expect(ws.getCell('C2').value.result).to.equal(now);
    });

    it('stores shared string values properly', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.getCell('A1').value = 'Hello, World!';

      ws.getCell('A2').value = 'Hello';
      ws.getCell('B2').value = 'World';
      ws.getCell('C2').value = {formula: 'CONCATENATE(A2, ", ", B2, "!")', result: 'Hello, World!'};

      ws.getCell('A3').value = ['Hello', 'World'].join(', ') + '!';

      // A1 and A3 should reference the same string object
      expect(ws.getCell('A1').value).to.equal(ws.getCell('A3').value);

      // A1 and C2 should not reference the same object
      expect(ws.getCell('A1').value).to.equal(ws.getCell('C2').value.result);
    });

    it('assigns cell types properly', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // plain number
      ws.getCell('A1').value = 7;

      // simple string
      ws.getCell('B1').value = 'Hello, World!';

      // floating point
      ws.getCell('C1').value = 3.14;

      // date-time
      ws.getCell('D1').value = new Date();

      // hyperlink
      ws.getCell('E1').value = {text: 'www.google.com', hyperlink: 'http://www.google.com'};

      // number formula
      ws.getCell('A2').value = {formula: 'A1', result: 7};

      // string formula
      ws.getCell('B2').value = {formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!'};

      // date formula
      ws.getCell('C2').value = {formula: 'D1', result: new Date()};

      expect(ws.getCell('A1').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('B1').type).to.equal(Excel.ValueType.String);
      expect(ws.getCell('C1').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('D1').type).to.equal(Excel.ValueType.Date);
      expect(ws.getCell('E1').type).to.equal(Excel.ValueType.Hyperlink);

      expect(ws.getCell('A2').type).to.equal(Excel.ValueType.Formula);
      expect(ws.getCell('B2').type).to.equal(Excel.ValueType.Formula);
      expect(ws.getCell('C2').type).to.equal(Excel.ValueType.Formula);
    });

    it('adds columns', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'name', width: 32 },
        { key: 'dob', width: 10 }
      ];

      expect(ws.getColumn('id').number).to.equal(1);
      expect(ws.getColumn('id').width).to.equal(10);
      expect(ws.getColumn('A')).to.equal(ws.getColumn('id'));
      expect(ws.getColumn(1)).to.equal(ws.getColumn('id'));

      expect(ws.getColumn('name').number).to.equal(2);
      expect(ws.getColumn('name').width).to.equal(32);
      expect(ws.getColumn('B')).to.equal(ws.getColumn('name'));
      expect(ws.getColumn(2)).to.equal(ws.getColumn('name'));

      expect(ws.getColumn('dob').number).to.equal(3);
      expect(ws.getColumn('dob').width).to.equal(10);
      expect(ws.getColumn('C')).to.equal(ws.getColumn('dob'));
      expect(ws.getColumn(3)).to.equal(ws.getColumn('dob'));
    });

    it('adds column headers', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.columns = [
        { header: 'Id', width: 10 },
        { header: 'Name', width: 32 },
        { header: 'D.O.B.', width: 10 }
      ];

      expect(ws.getCell('A1').value).to.equal('Id');
      expect(ws.getCell('B1').value).to.equal('Name');
      expect(ws.getCell('C1').value).to.equal('D.O.B.');
    });

    it('adds column headers by number', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // by defn
      ws.getColumn(1).defn = { key: 'id', header: 'Id', width: 10 };

      // by property
      ws.getColumn(2).key = 'name';
      ws.getColumn(2).header = 'Name';
      ws.getColumn(2).width = 32;

      expect(ws.getCell('A1').value).to.equal('Id');
      expect(ws.getCell('B1').value).to.equal('Name');

      expect(ws.getColumn('A').key).to.equal('id');
      expect(ws.getColumn(1).key).to.equal('id');
      expect(ws.getColumn(1).header).to.equal('Id');
      expect(ws.getColumn(1).headers).to.deep.equal(['Id']);
      expect(ws.getColumn(1).width).to.equal(10);

      expect(ws.getColumn(2).key).to.equal('name');
      expect(ws.getColumn(2).header).to.equal('Name');
      expect(ws.getColumn(2).headers).to.deep.equal(['Name']);
      expect(ws.getColumn(2).width).to.equal(32);
    });

    it('adds column headers by letter', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // by defn
      ws.getColumn('A').defn = { key: 'id', header: 'Id', width: 10 };

      // by property
      ws.getColumn('B').key = 'name';
      ws.getColumn('B').header = 'Name';
      ws.getColumn('B').width = 32;

      expect(ws.getCell('A1').value).to.equal('Id');
      expect(ws.getCell('B1').value).to.equal('Name');

      expect(ws.getColumn('A').key).to.equal('id');
      expect(ws.getColumn(1).key).to.equal('id');
      expect(ws.getColumn('A').header).to.equal('Id');
      expect(ws.getColumn('A').headers).to.deep.equal(['Id']);
      expect(ws.getColumn('A').width).to.equal(10);

      expect(ws.getColumn('B').key).to.equal('name');
      expect(ws.getColumn('B').header).to.equal('Name');
      expect(ws.getColumn('B').headers).to.deep.equal(['Name']);
      expect(ws.getColumn('B').width).to.equal(32);
    });

    it('adds rows by object', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // add columns to define column keys
      ws.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 32 },
        { header: 'D.O.B.', key: 'dob', width: 10 }
      ];

      var dateValue1 = new Date(1970, 1, 1);
      var dateValue2 = new Date(1965, 1, 7);

      ws.addRow({id: 1, name: 'John Doe', dob: dateValue1});
      ws.addRow({id: 2, name: 'Jane Doe', dob: dateValue2});

      expect(ws.getCell('A2').value).to.equal(1);
      expect(ws.getCell('B2').value).to.equal('John Doe');
      expect(ws.getCell('C2').value).to.equal(dateValue1);

      expect(ws.getCell('A3').value).to.equal(2);
      expect(ws.getCell('B3').value).to.equal('Jane Doe');
      expect(ws.getCell('C3').value).to.equal(dateValue2);

      expect(ws.getRow(2).values).to.deep.equal([, 1, 'John Doe', dateValue1]);
      expect(ws.getRow(3).values).to.deep.equal([, 2, 'Jane Doe', dateValue2]);

      var values = [,
        [, 'Id', 'Name', 'D.O.B.'],
        [, 1, 'John Doe', dateValue1],
        [, 2, 'Jane Doe', dateValue2]
      ];
      ws.eachRow(function(row, rowNumber) {
        expect(row.values).to.deep.equal(values[rowNumber]);
        row.eachCell(function(cell, colNumber) {
          expect(cell.value).to.equal(values[rowNumber][colNumber]);
        });
      });
    });

    it('adds rows by contiguous array', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      var dateValue1 = new Date(1970, 1, 1);
      var dateValue2 = new Date(1965, 1, 7);

      ws.addRow([1, 'John Doe', dateValue1]);
      ws.addRow([2, 'Jane Doe', dateValue2]);

      expect(ws.getCell('A1').value).to.equal(1);
      expect(ws.getCell('B1').value).to.equal('John Doe');
      expect(ws.getCell('C1').value).to.equal(dateValue1);

      expect(ws.getCell('A2').value).to.equal(2);
      expect(ws.getCell('B2').value).to.equal('Jane Doe');
      expect(ws.getCell('C2').value).to.equal(dateValue2);

      expect(ws.getRow(1).values).to.deep.equal([, 1, 'John Doe', dateValue1]);
      expect(ws.getRow(2).values).to.deep.equal([, 2, 'Jane Doe', dateValue2]);
    });

    it('adds rows by sparse array', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      var dateValue1 = new Date(1970, 1, 1);
      var dateValue2 = new Date(1965, 1, 7);
      var rows = [,
        [, 1, 'John Doe', , dateValue1],
        [, 2, 'Jane Doe', , dateValue2]
      ];
      var row3 = [];
      row3[1] = 3;
      row3[3] = 'Sam';
      row3[5] = dateValue1;
      rows.push(row3);
      rows.forEach(function(row) {
        if (row) {
          ws.addRow(row);
        }
      });

      expect(ws.getCell('A1').value).to.equal(1);
      expect(ws.getCell('B1').value).to.equal('John Doe');
      expect(ws.getCell('D1').value).to.equal(dateValue1);

      expect(ws.getCell('A2').value).to.equal(2);
      expect(ws.getCell('B2').value).to.equal('Jane Doe');
      expect(ws.getCell('D2').value).to.equal(dateValue2);

      expect(ws.getCell('A3').value).to.equal(3);
      expect(ws.getCell('C3').value).to.equal('Sam');
      expect(ws.getCell('E3').value).to.equal(dateValue1);

      expect(ws.getRow(1).values).to.deep.equal(rows[1]);
      expect(ws.getRow(2).values).to.deep.equal(rows[2]);
      expect(ws.getRow(3).values).to.deep.equal(rows[3]);

      ws.eachRow(function(row, rowNumber) {
        expect(row.values).to.deep.equal(rows[rowNumber]);
        row.eachCell(function(cell, colNumber) {
          expect(cell.value).to.equal(rows[rowNumber][colNumber]);
        });
      });
    });

    describe('Splice', function() {
      var options = {
        checkBadAlignments: false,
        checkSheetProperties: false,
        checkViews: false
      };
      describe('Rows', function() {
        it('Remove only', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.rows.removeOnly'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.rows.removeOnly'], options);
        });
        it('Remove and insert fewer', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.rows.insertFewer'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.rows.insertFewer'], options);
        });
        it('Remove and insert same', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.rows.insertSame'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.rows.insertSame'], options);
        });
        it('Remove and insert more', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.rows.insertMore'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.rows.insertMore'], options);
        });
      });
      describe('Columns', function() {
        it('splices columns', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.columns.removeOnly'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.columns.removeOnly'], options);
        });
        it('Remove and insert fewer', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.columns.insertFewer'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.columns.insertFewer'], options);
        });
        it('Remove and insert same', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.columns.insertSame'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.columns.insertSame'], options);
        });
        it('Remove and insert more', function() {
          var wb = new Excel.Workbook();
          testUtils.createTestBook(wb, 'xlsx', ['splice.columns.insertMore'], options);
          testUtils.checkTestBook(wb, 'xlsx', ['splice.columns.insertMore'], options);
        });
        it('handles column keys', function() {
          var wb = new Excel.Workbook();
          var ws = wb.addWorksheet('splice-column-insert-fewer');
          ws.columns = [
            { key: 'id', width: 10 },
            { key: 'dob', width: 20 },
            { key: 'name', width: 30 },
            { key: 'age', width: 40 }
          ];

          var values = [
            {id: '123', name: 'Jack', dob: new Date(), age: 0},
            {id: '124', name: 'Jill', dob: new Date(), age: 0},
          ];
          values.forEach(function(value) {
            ws.addRow(value);
          });

          ws.spliceColumns(2, 1, ['B1', 'B2'], ['C1', 'C2']);

          values.forEach(function(rowValues, index) {
            var row = ws.getRow(index + 1);
            _.each(rowValues, function(value, key) {
              if (key !== 'dob') {
                expect(row.getCell(key).value).to.equal(value);
              }
            });
          });

          expect(ws.getColumn(1).width).to.equal(10);
          expect(ws.getColumn(2).width).to.be.undefined();
          expect(ws.getColumn(3).width).to.be.undefined();
          expect(ws.getColumn(4).width).to.equal(30);
          expect(ws.getColumn(5).width).to.equal(40);
        });

        it('Splices to end', function() {
          var wb = new Excel.Workbook();
          var ws = wb.addWorksheet('splice-to-end');
          ws.columns = [
            { header: 'Col-1', width: 10 },
            { header: 'Col-2', width: 10 },
            { header: 'Col-3', width: 10 },
            { header: 'Col-4', width: 10 },
            { header: 'Col-5', width: 10 },
            { header: 'Col-6', width: 10 },
          ];

          ws.addRow([1,2,3,4,5,6]);
          ws.addRow([1,2,3,4,5,6]);

          // splice last 3 columns
          ws.spliceColumns(4, 3);
          expect(ws.getCell(1,1).value).to.equal('Col-1');
          expect(ws.getCell(1,2).value).to.equal('Col-2');
          expect(ws.getCell(1,3).value).to.equal('Col-3');
          expect(ws.getCell(1,4).value).to.be.null();
          expect(ws.getCell(1,5).value).to.be.null();
          expect(ws.getCell(1,6).value).to.be.null();
          expect(ws.getCell(1,7).value).to.be.null();
          expect(ws.getCell(2,1).value).to.equal(1);
          expect(ws.getCell(2,2).value).to.equal(2);
          expect(ws.getCell(2,3).value).to.equal(3);
          expect(ws.getCell(2,4).value).to.be.null();
          expect(ws.getCell(2,5).value).to.be.null();
          expect(ws.getCell(2,6).value).to.be.null();
          expect(ws.getCell(2,7).value).to.be.null();
          expect(ws.getCell(3,1).value).to.equal(1);
          expect(ws.getCell(3,2).value).to.equal(2);
          expect(ws.getCell(3,3).value).to.equal(3);
          expect(ws.getCell(3,4).value).to.be.null();
          expect(ws.getCell(3,5).value).to.be.null();
          expect(ws.getCell(3,6).value).to.be.null();
          expect(ws.getCell(3,7).value).to.be.null();

          expect(ws.getColumn(1).header).to.equal('Col-1');
          expect(ws.getColumn(2).header).to.equal('Col-2');
          expect(ws.getColumn(3).header).to.equal('Col-3');
          expect(ws.getColumn(4).header).to.be.undefined();
          expect(ws.getColumn(5).header).to.be.undefined();
          expect(ws.getColumn(6).header).to.be.undefined();
        });
        it('Splices past end', function() {
          var wb = new Excel.Workbook();
          var ws = wb.addWorksheet('splice-to-end');
          ws.columns = [
            { header: 'Col-1', width: 10 },
            { header: 'Col-2', width: 10 },
            { header: 'Col-3', width: 10 },
            { header: 'Col-4', width: 10 },
            { header: 'Col-5', width: 10 },
            { header: 'Col-6', width: 10 },
          ];

          ws.addRow([1,2,3,4,5,6]);
          ws.addRow([1,2,3,4,5,6]);

          // splice last 3 columns
          ws.spliceColumns(4, 4);
          expect(ws.getCell(1,1).value).to.equal('Col-1');
          expect(ws.getCell(1,2).value).to.equal('Col-2');
          expect(ws.getCell(1,3).value).to.equal('Col-3');
          expect(ws.getCell(1,4).value).to.be.null();
          expect(ws.getCell(1,5).value).to.be.null();
          expect(ws.getCell(1,6).value).to.be.null();
          expect(ws.getCell(1,7).value).to.be.null();
          expect(ws.getCell(2,1).value).to.equal(1);
          expect(ws.getCell(2,2).value).to.equal(2);
          expect(ws.getCell(2,3).value).to.equal(3);
          expect(ws.getCell(2,4).value).to.be.null();
          expect(ws.getCell(2,5).value).to.be.null();
          expect(ws.getCell(2,6).value).to.be.null();
          expect(ws.getCell(2,7).value).to.be.null();
          expect(ws.getCell(3,1).value).to.equal(1);
          expect(ws.getCell(3,2).value).to.equal(2);
          expect(ws.getCell(3,3).value).to.equal(3);
          expect(ws.getCell(3,4).value).to.be.null();
          expect(ws.getCell(3,5).value).to.be.null();
          expect(ws.getCell(3,6).value).to.be.null();
          expect(ws.getCell(3,7).value).to.be.null();

          expect(ws.getColumn(1).header).to.equal('Col-1');
          expect(ws.getColumn(2).header).to.equal('Col-2');
          expect(ws.getColumn(3).header).to.equal('Col-3');
          expect(ws.getColumn(4).header).to.be.undefined();
          expect(ws.getColumn(5).header).to.be.undefined();
          expect(ws.getColumn(6).header).to.be.undefined();
        });
        it('Splices almost to end', function() {
          var wb = new Excel.Workbook();
          var ws = wb.addWorksheet('splice-to-end');
          ws.columns = [
            { header: 'Col-1', width: 10 },
            { header: 'Col-2', width: 10 },
            { header: 'Col-3', width: 10 },
            { header: 'Col-4', width: 10 },
            { header: 'Col-5', width: 10 },
            { header: 'Col-6', width: 10 },
          ];

          ws.addRow([1,2,3,4,5,6]);
          ws.addRow([1,2,3,4,5,6]);

          // splice last 3 columns
          ws.spliceColumns(4, 2);
          expect(ws.getCell(1,1).value).to.equal('Col-1');
          expect(ws.getCell(1,2).value).to.equal('Col-2');
          expect(ws.getCell(1,3).value).to.equal('Col-3');
          expect(ws.getCell(1,4).value).to.equal('Col-6');
          expect(ws.getCell(1,5).value).to.be.null();
          expect(ws.getCell(1,6).value).to.be.null();
          expect(ws.getCell(1,7).value).to.be.null();
          expect(ws.getCell(2,1).value).to.equal(1);
          expect(ws.getCell(2,2).value).to.equal(2);
          expect(ws.getCell(2,3).value).to.equal(3);
          expect(ws.getCell(2,4).value).to.equal(6);
          expect(ws.getCell(2,5).value).to.be.null();
          expect(ws.getCell(2,6).value).to.be.null();
          expect(ws.getCell(2,7).value).to.be.null();
          expect(ws.getCell(3,1).value).to.equal(1);
          expect(ws.getCell(3,2).value).to.equal(2);
          expect(ws.getCell(3,3).value).to.equal(3);
          expect(ws.getCell(3,4).value).to.equal(6);
          expect(ws.getCell(3,5).value).to.be.null();
          expect(ws.getCell(3,6).value).to.be.null();
          expect(ws.getCell(3,7).value).to.be.null();

          expect(ws.getColumn(1).header).to.equal('Col-1');
          expect(ws.getColumn(2).header).to.equal('Col-2');
          expect(ws.getColumn(3).header).to.equal('Col-3');
          expect(ws.getColumn(4).header).to.equal('Col-6');
          expect(ws.getColumn(5).header).to.be.undefined();
          expect(ws.getColumn(6).header).to.be.undefined();
        });
      });
    });

    it('iterates over rows', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.getCell('A1').value = 1;
      ws.getCell('B2').value = 2;
      ws.getCell('D4').value = 4;
      ws.getCell('F6').value = 6;
      ws.eachRow(function(row, rowNumber) {
        expect(rowNumber).not.to.equal(3);
        expect(rowNumber).not.to.equal(5);
      });

      var count = 1;
      ws.eachRow({includeEmpty: true}, function(row, rowNumber) {
        expect(rowNumber).to.equal(count++);
      });
    });

    it('iterates over collumn cells', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('A4').value = 4;
      ws.getCell('A6').value = 6;
      var colA = ws.getColumn('A');
      colA.eachCell(function(cell, rowNumber) {
        expect(rowNumber).not.to.equal(3);
        expect(rowNumber).not.to.equal(5);
        expect(cell.value).to.equal(rowNumber);
      });

      var count = 1;
      colA.eachCell({includeEmpty: true}, function(cell, rowNumber) {
        expect(rowNumber).to.equal(count++);
      });
      expect(count).to.equal(7);
    });

    it('returns sheet values', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.getCell('A1').value = 11;
      ws.getCell('C1').value = 'C1';
      ws.getCell('A2').value = 21;
      ws.getCell('B2').value = 'B2';
      ws.getCell('A4').value = 'end';

      expect(ws.getSheetValues()).to.deep.equal([,
        [, 11,, 'C1'],
        [, 21, 'B2'],
        , // eslint-disable-line comma-style
        [, 'end']
      ]);
    });

    it('calculates rowCount and actualRowCount', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.getCell('A1').value = 'A1';
      ws.getCell('C1').value = 'C1';
      ws.getCell('A3').value = 'A3';
      ws.getCell('D3').value = 'D3';
      ws.getCell('A4').value = null;
      ws.getCell('B5').value = 'B5';

      expect(ws.rowCount).to.equal(5);
      expect(ws.actualRowCount).to.equal(3);
    });

    it('calculates columnCount and actualColumnCount', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.getCell('A1').value = 'A1';
      ws.getCell('C1').value = 'C1';
      ws.getCell('A3').value = 'A3';
      ws.getCell('D3').value = 'D3';
      ws.getCell('E4').value = null;
      ws.getCell('F5').value = 'F5';

      expect(ws.columnCount).to.equal(6);
      expect(ws.actualColumnCount).to.equal(4);
    });
  });
});
