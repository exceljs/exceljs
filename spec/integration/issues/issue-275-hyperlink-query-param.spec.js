'use strict';

var chai = require('chai');

var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

var expect = chai.expect;

// this file to contain integration tests created from github issues
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', function() {
  it('issue 275 - hyperlink with query arguments corrupts workbook', function() {
    var options = {
      filename: TEST_XLSX_FILE_NAME,
      useStyles: true
    };
    var wb = new Excel.stream.xlsx.WorkbookWriter(options);
    var ws = wb.addWorksheet('Sheet1');

    var hyperlink = {
      text: 'Somewhere with query params',
      hyperlink: 'www.somewhere.com?a=1&b=2&c=<>&d="\'"'
    };

    // Start of Heading
    ws.getCell('A1').value = hyperlink;
    ws.commit();

    return wb.commit()
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet('Sheet1');
        expect(ws2.getCell('A1').value).to.deep.equal(hyperlink);
      });
  });
});
