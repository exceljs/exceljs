'use strict';

const verquire = require('../../utils/verquire');

const Excel = verquire('excel');

describe('github issues', () => {
  describe('issue 539 - <contentType /> element', () => {
    it('Reading 1904.xlsx', () => {
      const wb = new Excel.Workbook();
      return wb.xlsx.readFile(
        './spec/integration/data/1519293514-KRISHNAPATNAM_LINE_UP.xlsx'
      );
    });
  });
});
