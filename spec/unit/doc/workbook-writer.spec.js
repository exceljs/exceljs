'use strict';

const Stream = require('stream');

const expect = require('chai').expect;

const Excel = require('../../../excel');

describe('Workbook Writer', () => {
  it('returns undefined for non-existant sheet', () => {
    const stream = new Stream.Writable({ write: function noop() {} });
    const wb = new Excel.stream.xlsx.WorkbookWriter({
      stream,
    });
    wb.addWorksheet('first');
    expect(wb.getWorksheet('w00t')).to.equal(undefined);
  });
});
