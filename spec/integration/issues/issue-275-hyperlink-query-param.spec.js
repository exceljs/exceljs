'use strict';

/**
 * Hyperlink in stream is broken after change the interface of hyperlink! Need a fix!
 */

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
      display: 'Somewhere with query params',
      target: 'www.somewhere.com?a=1&b=2&c=<>&d="\'"',
      mode: 'external'
    };

    // Start of Heading
    ws.getCell('A1').value = hyperlink.display;
    ws.getCell('A1').hyperlink = hyperlink;
    // TO-DO: this needs to commit hyperlink
    ws.commit();

    return wb.commit()
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet('Sheet1');
        expect(ws2.getCell('A1').value).to.deep.equal(hyperlink.display);
        // uncomment to test
        //expect(ws2.getCell('A1').hyperlink).to.deep.equal(hyperlink);
      });
  });
});
