'use strict';

var chai = require('chai');

var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

var expect = chai.expect;

// this file to contain integration tests created from github issues
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', function() {
  it('issue 234 - Broken XLSX because of "vertical tab" ascii character in a cell', function() {
    var wb = new Excel.Workbook();
    var ws = wb.addWorksheet('Sheet1');

    // Start of Heading
    ws.getCell('A1').value = 'Hello, \x01World!';

    // Vertical Tab
    ws.getCell('A2').value = 'Hello, \x0bWorld!';

    return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet('Sheet1');
        expect(ws2.getCell('A1').value).to.equal('Hello, World!');
        expect(ws2.getCell('A2').value).to.equal('Hello, World!');
      });
  });
});
