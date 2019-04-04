'use strict';

const verquire = require('../../utils/verquire');

const Excel = verquire('excel');
const TEST_XLSX_FILE_NAME = './spec/integration/data/test-pr-567.xlsx';

describe('pr related issues', () => {
  describe('pr 5676 whole column defined names', () => {
    it('Should be able to read this file', () => {
      const wb = new Excel.Workbook();
      return wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
    });
  });
});
