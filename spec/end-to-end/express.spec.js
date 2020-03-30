const express = require('express');
const got = require('got');
const testutils = require('../utils/index');

const Excel = verquire('exceljs');

// =============================================================================
// Tests

describe('Express', () => {
  it('downloads a workbook', async () => {
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

    const res = got.stream('http://127.0.0.1:3003/workbook');
    const wb2 = new Excel.Workbook();
    await wb2.xlsx.read(res);
    testutils.checkTestBook(wb2, 'xlsx');
    server.close();
  });
});
