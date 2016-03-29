'use strict';

var expect = require('chai').expect;
var Bluebird = require('bluebird');
var fs = require('fs');
var fsa = Bluebird.promisifyAll(fs);
var _ = require('underscore');
var Excel = require('../../excel');
var testutils = require('../testutils');

var express = require('express');
var request = require('request');

// =============================================================================
// Tests

describe('Express', function() {

  it('downloads a workbook', function() {
    var app = express();
    app.get('/workbook', function(req, res) {
      var wb = testutils.createTestBook(true, Excel.Workbook);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=' + 'Report.xlsx');
      wb.xlsx.write(res)
        .then(function() {
          res.end();
        });
    });
    var server = app.listen(3003);

    return new Bluebird(function(resolve, reject) {
      var r = request('http://localhost:3003/workbook');
      r.on('response', function(res) {
        var wb2 = new Excel.Workbook();
        var stream = wb2.xlsx.createInputStream();
        stream.on('done', function() {
          try {
            testutils.checkTestBook(wb2, 'xlsx', true);
          }
          catch(ex) {
            return reject(ex);
          }
          server.close();
          resolve();
        });
        res.pipe(stream);
      });
    });
  });

});
