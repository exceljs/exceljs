var expect = require('chai').expect;

var Excel = require('../../../excel');

describe('Worksheet', function() {
  describe('Shared Formulae', function() {
    it('Fills formula using 2D array values', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.fillFormula('A1:B2', 'ROW()+COLUMN()', [[2, 3], [3, 4]]);
      expect(ws.getCell('A1').value).to.deep.equal({ formula: 'ROW()+COLUMN()', result: 2 });
      expect(ws.getCell('B1').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('A2').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('B2').value).to.deep.equal({ sharedFormula: 'A1', result: 4 });
    });
    it('Fills formula down using 1D array values', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.fillFormula('A1:A4', 'ROW()+COLUMN()', [2, 3, 4, 5]);
      expect(ws.getCell('A1').value).to.deep.equal({ formula: 'ROW()+COLUMN()', result: 2 });
      expect(ws.getCell('A2').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('A3').value).to.deep.equal({ sharedFormula: 'A1', result: 4 });
      expect(ws.getCell('A4').value).to.deep.equal({ sharedFormula: 'A1', result: 5 });
    });
    it('Fills formula across using 1D array values', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.fillFormula('A1:D1', 'ROW()+COLUMN()', [2, 3, 4, 5]);
      expect(ws.getCell('A1').value).to.deep.equal({ formula: 'ROW()+COLUMN()', result: 2 });
      expect(ws.getCell('B1').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('C1').value).to.deep.equal({ sharedFormula: 'A1', result: 4 });
      expect(ws.getCell('D1').value).to.deep.equal({ sharedFormula: 'A1', result: 5 });
    });
    it('Fills formula down and across using 1D array values', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.fillFormula('A1:B2', 'ROW()+COLUMN()', [2, 3, 3, 4]);
      expect(ws.getCell('A1').value).to.deep.equal({ formula: 'ROW()+COLUMN()', result: 2 });
      expect(ws.getCell('B1').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('A2').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('B2').value).to.deep.equal({ sharedFormula: 'A1', result: 4 });
    });
    it('Fills formula using function', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.fillFormula('A1:B2', 'ROW()+COLUMN()', (r, c) => r + c);
      expect(ws.getCell('A1').value).to.deep.equal({ formula: 'ROW()+COLUMN()', result: 2 });
      expect(ws.getCell('B1').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('A2').value).to.deep.equal({ sharedFormula: 'A1', result: 3 });
      expect(ws.getCell('B2').value).to.deep.equal({ sharedFormula: 'A1', result: 4 });
    });
  });
});
