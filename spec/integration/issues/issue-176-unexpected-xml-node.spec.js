'use strict';

const chai = require('chai');

const verquire = require('../../utils/verquire');

const Excel = verquire('excel');

const { expect } = chai;

describe('github issues', () => {
  it('issue 176 - Unexpected xml node in parseOpen', () => {
    const wb = new Excel.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-176.xlsx')
      .then(() => {
        // arriving here is success
        expect(true).to.equal(true);
      });
  });
});
