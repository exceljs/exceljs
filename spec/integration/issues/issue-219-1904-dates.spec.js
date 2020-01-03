'use strict';

var chai = require('chai');

var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

var expect = chai.expect;

// this file to contain integration tests created from github issues
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', function() {
  describe('issue 219 - 1904 dates not supported', function() {
    it('Reading 1904.xlsx', function() {
      var wb = new Excel.Workbook();
      return wb.xlsx.readFile('./spec/integration/data/1904.xlsx')
        .then(function() {
          expect(wb.properties.date1904).to.equal(true);

          var ws = wb.getWorksheet('Sheet1');
          expect(ws.getCell('B4').value.toISOString()).to.equal('1904-01-01T00:00:00.000Z');
        });
    });
    it('Writing and Reading', function() {
      var wb = new Excel.Workbook();
      wb.properties.date1904 = true;
      var ws = wb.addWorksheet('Sheet1');
      ws.getCell('B4').value = new Date('1904-01-01T00:00:00.000Z');
      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function(wb2) {
          expect(wb2.properties.date1904).to.equal(true);

          var ws2 = wb2.getWorksheet('Sheet1');
          expect(ws2.getCell('B4').value.toISOString()).to.equal('1904-01-01T00:00:00.000Z');
        });
    });
  });
});
