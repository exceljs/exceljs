'use strict';

var expect = require('chai').expect;
var verquire = require('./verquire');
var tools = require('./tools');
var testValues = tools.fix(require('./data/sheet-values.json'));

var utils = verquire('utils/utils');
var Excel = verquire('excel');

function fillFormula(f) {
  return Object.assign({formula: undefined, result: undefined}, f);
}
var streamedValues = {
  B1: { sharedString: 0 },
  C1: utils.dateToExcel(testValues.date),
  D1: fillFormula(testValues.formulas[0]),
  E1: fillFormula(testValues.formulas[1]),
  F1: { sharedString: 1 },
  G1: { sharedString: 2 },
};
module.exports = {
  testValues: tools.fix(require('./data/sheet-values.json')),
  styles: tools.fix(require('./data/styles.json')),
  properties: tools.fix(require('./data/sheet-properties.json')),
  pageSetup: tools.fix(require('./data/page-setup.json')),

  checkBook: function(filename) {
    var wb = new Excel.stream.xlsx.WorkbookReader();

    // expectations
    var dateAccuracy = 0.00001;

    return new Promise(function(resolve, reject) {
      var rowCount = 0;

      wb.on('worksheet', function(ws) {
        // Sheet name stored in workbook. Not guaranteed here
        // expect(ws.name).to.equal('blort');
        ws.on('row', function(row) {
          rowCount++;
          try {
            switch (row.number) {
              case 1:
                expect(row.getCell('A').value).to.equal(7);
                expect(row.getCell('A').type).to.equal(Excel.ValueType.Number);
                expect(row.getCell('B').value).to.deep.equal(streamedValues.B1);
                expect(row.getCell('B').type).to.equal(Excel.ValueType.String);
                expect(Math.abs(row.getCell('C').value - streamedValues.C1)).to.be.below(dateAccuracy);
                expect(row.getCell('C').type).to.equal(Excel.ValueType.Number);

                expect(row.getCell('D').value).to.deep.equal(streamedValues.D1);
                expect(row.getCell('D').type).to.equal(Excel.ValueType.Formula);
                expect(row.getCell('E').value).to.deep.equal(streamedValues.E1);
                expect(row.getCell('E').type).to.equal(Excel.ValueType.Formula);
                expect(row.getCell('F').value).to.deep.equal(streamedValues.F1);
                expect(row.getCell('F').type).to.equal(Excel.ValueType.SharedString);
                expect(row.getCell('G').value).to.deep.equal(streamedValues.G1);
                break;

              case 2:
                // A2:B3
                expect(row.getCell('A').value).to.equal(5);
                expect(row.getCell('A').type).to.equal(Excel.ValueType.Number);

                expect(row.getCell('B').type).to.equal(Excel.ValueType.Null);

                // C2:D3
                expect(row.getCell('C').value).to.be.null();
                expect(row.getCell('C').type).to.equal(Excel.ValueType.Null);

                expect(row.getCell('D').value).to.be.null();
                expect(row.getCell('D').type).to.equal(Excel.ValueType.Null);

                break;

              case 3:
                expect(row.getCell('A').value).to.equal(null);
                expect(row.getCell('A').type).to.equal(Excel.ValueType.Null);

                expect(row.getCell('B').value).to.equal(null);
                expect(row.getCell('B').type).to.equal(Excel.ValueType.Null);

                expect(row.getCell('C').value).to.be.null();
                expect(row.getCell('C').type).to.equal(Excel.ValueType.Null);

                expect(row.getCell('D').value).to.be.null();
                expect(row.getCell('D').type).to.equal(Excel.ValueType.Null);
                break;

              case 4:
                expect(row.getCell('A').type).to.equal(Excel.ValueType.Number);
                expect(row.getCell('C').type).to.equal(Excel.ValueType.Number);
                break;

              case 5:
                // test fonts and formats
                expect(row.getCell('A').value).to.deep.equal(streamedValues.B1);
                expect(row.getCell('A').type).to.equal(Excel.ValueType.String);
                expect(row.getCell('B').value).to.deep.equal(streamedValues.B1);
                expect(row.getCell('B').type).to.equal(Excel.ValueType.String);
                expect(row.getCell('C').value).to.deep.equal(streamedValues.B1);
                expect(row.getCell('C').type).to.equal(Excel.ValueType.String);

                expect(Math.abs(row.getCell('D').value - 1.6)).to.be.below(0.00000001);
                expect(row.getCell('D').type).to.equal(Excel.ValueType.Number);

                expect(Math.abs(row.getCell('E').value - 1.6)).to.be.below(0.00000001);
                expect(row.getCell('E').type).to.equal(Excel.ValueType.Number);

                expect(Math.abs(row.getCell('F').value - streamedValues.C1)).to.be.below(dateAccuracy);
                expect(row.getCell('F').type).to.equal(Excel.ValueType.Number);
                break;

              case 6:
                expect(row.height).to.equal(42);
                break;

              case 7:
                break;

              case 8:
                expect(row.height).to.equal(40);
                break;

              default:
                break;
            }
          } catch (error) {
            reject(error);
          }
        });
      });
      wb.on('end', function() {
        try {
          expect(rowCount).to.equal(10);
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      wb.read(filename, {entries: 'emit', worksheets: 'emit'});
    });
  }
};
