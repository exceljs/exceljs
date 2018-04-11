'use strict';

var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

describe('github issues', function() {
  describe('issue 539 - <contentType /> element', function() {
    it('Reading 1904.xlsx', function() {
      var wb = new Excel.Workbook();
      return wb.xlsx.readFile('./spec/integration/data/1519293514-KRISHNAPATNAM_LINE_UP.xlsx');
    });
  });
});
