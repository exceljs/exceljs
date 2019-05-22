'use strict';

const chai = require('chai');
const verquire = require('../../utils/verquire');

const Excel = verquire('excel');

const { expect } = chai;

describe('github issues', () => {
  it('issue 257 - worksheet order is not respected', () => {
    const wb = new Excel.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-257.xlsx')
      .then(() => {
        expect(wb.worksheets.map(ws => ws.name)).to.deep.equal([
          'First',
          'Second',
        ]);
      });
  });
});
