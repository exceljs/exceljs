'use strict';

var verquire = require('../utils/verquire');
var testutils = require('../utils/index');

var express = require('express');
var request = require('request');

var Excel = verquire('excel');

// =============================================================================
// Tests

describe('Express', function() {
  it('downloads a workbook', function() {
    var app = express();
    app.get('/workbook', function(req, res) {
      var wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
      wb.xlsx.write(res)
        .then(function() {
          res.end();
        });
    });
    var server = app.listen(3003);

    return new Promise(function(resolve, reject) {
      var r = request('http://127.0.0.1:3003/workbook');
      r.on('response', function(res) {
        var wb2 = new Excel.Workbook();
        var stream = wb2.xlsx.createInputStream();
        stream.on('done', function() {
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
