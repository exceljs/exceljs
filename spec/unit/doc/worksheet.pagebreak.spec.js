var expect = require('chai').expect;

var Excel = require('../../../excel');

describe('Worksheet', function() {
  describe('Page Breaks', function() {
    it('adds multiple row breaks', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // initial values
      ws.getCell('A1').value = 'A1';
      ws.getCell('B1').value = 'B1';
      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';
      
      var row = ws.getRow(1);
      row.addPageBreak();
      row = ws.getRow(2);
      row.addPageBreak();

      expect(ws.rowBreaks.length).to.equal(2);
    });
  });
});
