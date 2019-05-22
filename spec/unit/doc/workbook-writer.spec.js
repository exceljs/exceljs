'use strict';

const Stream = require('stream');

const { expect } = require('chai');

const Excel = require('../../../lib/exceljs.nodejs');

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
