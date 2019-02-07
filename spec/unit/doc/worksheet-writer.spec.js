'use strict';

var expect = require('chai').expect;

var WorksheetWriter = require('../../../lib/stream/xlsx/worksheet-writer');
var StreamBuf = require('../../../lib/utils/stream-buf');

describe('Workbook Writer', function() {
  it('generates valid xml even when there is no data', function() {
    // issue: https://github.com/guyonroche/exceljs/issues/99
    // PR: https://github.com/guyonroche/exceljs/pull/255
    return new Promise(function(resolve, reject) {
      var mockWorkbook = {
        _openStream: function() {
          return this.stream;
        },
        stream: new StreamBuf()
      };
      mockWorkbook.stream.on('finish', function() {
        try {
          var xml = mockWorkbook.stream.read().toString();
          expect(xml).xml.to.be.valid();
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      var writer = new WorksheetWriter({
        id: 1,
        workbook: mockWorkbook
      });

      writer.commit();
    });
  });
});
