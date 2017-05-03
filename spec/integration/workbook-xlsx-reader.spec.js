'use strict';

var verquire = require('../utils/verquire');
var testutils = require('../utils/index');

var Excel = verquire('excel');

var TEST_FILE_NAME = './spec/out/wb.test.xlsx';

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe('WorkbookReader', function() {
  describe('Serialise', function() {
    it('xlsx file', function() {
      this.timeout(10000);
      var wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');

      return wb.xlsx.writeFile(TEST_FILE_NAME)
        .then(function() {
          return testutils.checkTestBookReader(TEST_FILE_NAME);
        });
    });
  });
});
