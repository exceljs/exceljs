'use strict';

var verquire = require('../../utils/verquire');
var Excel = verquire('excel');
var TEST_XLSX_FILE_NAME = './spec/integration/data/test-pr-567.xlsx';

describe('pr related issues', function() {
  describe('pr 5676 whole column defined names', function() {
    it('Should be able to read this file', function() {
      var wb = new Excel.Workbook();
      return wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
    });
  });
});
