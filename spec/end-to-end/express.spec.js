const express = require('express');
const request = require('request');
const testutils = require('../utils/index');

const Excel = verquire('exceljs');

// =============================================================================
// Tests

describe('Express', () => {
  it('downloads a workbook', () => {
    const app = express();
    app.get('/workbook', (req, res) => {
      const wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
      wb.xlsx.write(res).then(() => {
        res.end();
      });
    });
    const server = app.listen(3003);

    return new Promise((resolve, reject) => {
      const r = request('http://127.0.0.1:3003/workbook');
      r.on('response', res => {
        const wb2 = new Excel.Workbook();
        const stream = wb2.xlsx.createInputStream();
        stream.on('done', () => {
          try {
            testutils.checkTestBook(wb2, 'xlsx');
            server.close();
            resolve();
          } catch (ex) {
            reject(ex);
          }
        });
        res.pipe(stream);
      });
    });
  });
});
