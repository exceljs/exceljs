'use strict';

var expect = require('chai').expect;
var verquire = require('./verquire');

var Enums = verquire('doc/enums');

module.exports = {
  rows: {
    removeOnly: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-row-remove-only');

        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.addRow(['5,1', '5,2', '5,3']);

        ws.spliceRows(2, 1);
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-row-remove-only');
        expect(ws).to.not.be.undefined();

        expect(ws.getCell('A1').value).to.equal('1,1');
        expect(ws.getCell('A1').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('B1').value).to.equal('1,2');
        expect(ws.getCell('B1').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('C1').value).to.equal('1,3');
        expect(ws.getCell('C1').type).to.equal(Enums.ValueType.String);

        expect(ws.getCell('A2').type).to.equal(Enums.ValueType.Null);
        expect(ws.getCell('B2').type).to.equal(Enums.ValueType.Null);
        expect(ws.getCell('C2').type).to.equal(Enums.ValueType.Null);

        expect(ws.getCell('A3').value).to.equal(4.1);
        expect(ws.getCell('A3').type).to.equal(Enums.ValueType.Number);
        expect(ws.getCell('B3').type).to.equal(Enums.ValueType.Null);
        expect(ws.getCell('C3').value).to.equal(4.3);
        expect(ws.getCell('C3').type).to.equal(Enums.ValueType.Number);

        expect(ws.getCell('A4').value).to.equal('5,1');
        expect(ws.getCell('A4').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('B4').value).to.equal('5,2');
        expect(ws.getCell('B4').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('C4').value).to.equal('5,3');
        expect(ws.getCell('C4').type).to.equal(Enums.ValueType.String);

        ws.addRow(['5,1b', '5,2b', '5,3b']);
        expect(ws.getCell('A5').value).to.equal('5,1b');
        expect(ws.getCell('A5').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('B5').value).to.equal('5,2b');
        expect(ws.getCell('B5').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('C5').value).to.equal('5,3b');
        expect(ws.getCell('C5').type).to.equal(Enums.ValueType.String);
      }
    },
    insertFewer: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-row-insert-fewer');

        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.addRow(['5,1', '5,2', '5,3']);

        ws.spliceRows(2, 2, ['one', 'two', 'three']);
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-row-insert-fewer');
        expect(ws).to.not.be.undefined();

        expect(ws.getRow(1).values).to.deep.equal([, '1,1', '1,2', '1,3']);
        expect(ws.getRow(2).values).to.deep.equal([, 'one', 'two', 'three']);
        expect(ws.getRow(3).values).to.deep.equal([, 4.1,, 4.3]);
        expect(ws.getRow(4).values).to.deep.equal([, '5,1', '5,2', '5,3']);
      }
    },
    insertSame: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-row-insert-same');

        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.addRow(['5,1', '5,2', '5,3']);

        ws.spliceRows(2, 2, ['one', 'two', 'three'], ['une', 'deux', 'trois']);
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-row-insert-same');
        expect(ws).to.not.be.undefined();

        expect(ws.getRow(1).values).to.deep.equal([, '1,1', '1,2', '1,3']);
        expect(ws.getRow(2).values).to.deep.equal([, 'one', 'two', 'three']);
        expect(ws.getRow(3).values).to.deep.equal([, 'une', 'deux', 'trois']);
        expect(ws.getRow(4).values).to.deep.equal([, 4.1,, 4.3]);
        expect(ws.getRow(5).values).to.deep.equal([, '5,1', '5,2', '5,3']);
      }
    },
    insertMore: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-row-insert-more');

        ws.addRow(['1,1', '1,2', '1,3']);
        ws.addRow(['2,1', '2,2', '2,3']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.addRow(['5,1', '5,2', '5,3']);

        ws.spliceRows(2, 2, ['one', 'two', 'three'], ['une', 'deux', 'trois'], ['uno', 'due', 'tre']);
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-row-insert-more');
        expect(ws).to.not.be.undefined();

        expect(ws.getRow(1).values).to.deep.equal([, '1,1', '1,2', '1,3']);
        expect(ws.getRow(2).values).to.deep.equal([, 'one', 'two', 'three']);
        expect(ws.getRow(3).values).to.deep.equal([, 'une', 'deux', 'trois']);
        expect(ws.getRow(4).values).to.deep.equal([, 'uno', 'due', 'tre']);
        expect(ws.getRow(5).values).to.deep.equal([, 4.1,, 4.3]);
        expect(ws.getRow(6).values).to.deep.equal([, '5,1', '5,2', '5,3']);
      }
    }
  },
  columns: {
    removeOnly: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-column-remove-only');

        ws.columns = [
          { key: 'id', width: 10 },
          { key: 'name', width: 32 },
          { key: 'dob', width: 10 }
        ];

        ws.addRow({id: 'id1', name: 'name1', dob: 'dob1'});
        ws.addRow({id: 2, dob: 'dob2'});
        ws.addRow({name: 'name3', dob: 3});

        ws.spliceColumns(2, 1);
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-column-remove-only');
        expect(ws).to.not.be.undefined();

        expect(ws.getCell('A1').value).to.equal('id1');
        expect(ws.getCell('A1').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('B1').value).to.equal('dob1');
        expect(ws.getCell('B1').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('C1').type).to.equal(Enums.ValueType.Null);

        expect(ws.getCell('A2').value).to.equal(2);
        expect(ws.getCell('A2').type).to.equal(Enums.ValueType.Number);
        expect(ws.getCell('B2').value).to.equal('dob2');
        expect(ws.getCell('B2').type).to.equal(Enums.ValueType.String);
        expect(ws.getCell('C2').type).to.equal(Enums.ValueType.Null);

        expect(ws.getCell('A3').type).to.equal(Enums.ValueType.Null);
        expect(ws.getCell('B3').value).to.equal(3);
        expect(ws.getCell('B3').type).to.equal(Enums.ValueType.Number);
        expect(ws.getCell('C3').type).to.equal(Enums.ValueType.Null);
      }
    },
    insertFewer: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-column-insert-fewer');

        ws.addRow(['1,1', '1,2', '1,3', '1,4', '1,5']);
        ws.addRow(['2,1', '2,2', '2,3', '2,4', '2,5']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.getCell('E4').value = 4.5;
        ws.addRow(['5,1', '5,2', '5,3', '5,4', '5,5']);

        ws.spliceColumns(2, 2, ['one', 'two', 'three', 'four', 'five']);
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-column-insert-fewer');
        expect(ws).to.not.be.undefined();

        expect(ws.getRow(1).values).to.deep.equal([, '1,1', 'one', '1,4', '1,5']);
        expect(ws.getRow(2).values).to.deep.equal([, '2,1', 'two', '2,4', '2,5']);
        expect(ws.getRow(3).values).to.deep.equal([, , 'three', ]);
        expect(ws.getRow(4).values).to.deep.equal([, 4.1, 'four',, 4.5]);
        expect(ws.getRow(5).values).to.deep.equal([, '5,1', 'five', '5,4', '5,5']);
      }
    },
    insertSame: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-column-insert-same');

        ws.addRow(['1,1', '1,2', '1,3', '1,4', '1,5']);
        ws.addRow(['2,1', '2,2', '2,3', '2,4', '2,5']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.getCell('E4').value = 4.5;
        ws.addRow(['5,1', '5,2', '5,3', '5,4', '5,5']);

        ws.spliceColumns(2, 2,
          ['one', 'two', 'three', 'four', 'five'],
          ['une', 'deux', 'trois', 'quatre', 'cinq']
        );
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-column-insert-same');
        expect(ws).to.not.be.undefined();

        expect(ws.getRow(1).values).to.deep.equal([, '1,1', 'one', 'une', '1,4', '1,5']);
        expect(ws.getRow(2).values).to.deep.equal([, '2,1', 'two', 'deux', '2,4', '2,5']);
        expect(ws.getRow(3).values).to.deep.equal([,, 'three', 'trois', ]);
        expect(ws.getRow(4).values).to.deep.equal([, 4.1, 'four', 'quatre',, 4.5]);
        expect(ws.getRow(5).values).to.deep.equal([, '5,1', 'five', 'cinq', '5,4', '5,5']);
      }
    },
    insertMore: {
      addSheet: function(wb) {
        var ws = wb.addWorksheet('splice-column-insert-more');

        ws.addRow(['1,1', '1,2', '1,3', '1,4', '1,5']);
        ws.addRow(['2,1', '2,2', '2,3', '2,4', '2,5']);
        ws.getCell('A4').value = 4.1;
        ws.getCell('C4').value = 4.3;
        ws.getCell('E4').value = 4.5;
        ws.addRow(['5,1', '5,2', '5,3', '5,4', '5,5']);

        ws.spliceColumns(2, 2,
          ['one', 'two', 'three', 'four', 'five'],
          ['une', 'deux', 'trois', 'quatre', 'cinq'],
          ['uno', 'due', 'tre', 'quatro', 'cinque']
        );
      },

      checkSheet: function(wb) {
        var ws = wb.getWorksheet('splice-column-insert-more');
        expect(ws).to.not.be.undefined();

        expect(ws.getRow(1).values).to.deep.equal([, '1,1', 'one', 'une', 'uno', '1,4', '1,5']);
        expect(ws.getRow(2).values).to.deep.equal([, '2,1', 'two', 'deux', 'due', '2,4', '2,5']);
        expect(ws.getRow(3).values).to.deep.equal([,, 'three', 'trois', 'tre', ]);
        expect(ws.getRow(4).values).to.deep.equal([, 4.1, 'four', 'quatre', 'quatro',, 4.5]);
        expect(ws.getRow(5).values).to.deep.equal([, '5,1', 'five', 'cinq', 'cinque', '5,4', '5,5']);
      }
    }
  }
};
