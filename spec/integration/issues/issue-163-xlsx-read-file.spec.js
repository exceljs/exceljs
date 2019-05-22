'use strict';

const chai = require('chai');
const verquire = require('../../utils/verquire');

const Excel = verquire('excel');

const { expect } = chai;

describe('github issues', () => {
  it('issue 163 - Error while using xslx readFile method', () => {
    const wb = new Excel.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-163.xlsx')
      .then(() => {
        // arriving here is success
        expect(true).to.equal(true);
      });
  });
});
