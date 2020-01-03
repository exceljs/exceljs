'use strict';

var fs = require('fs');
var util = require('util');
var _ = require('underscore');
var Promise = require('bluebird');

var utils = require('./utils/utils');
var HrStopwatch = require('./utils/hr-stopwatch');
var ColumnSum = require('./utils/column-sum');

var Excel = require('../excel');

var Workbook = Excel.Workbook;
var WorkbookReader = Excel.stream.xlsx.WorkbookReader;

if (process.argv[2] === 'help') {
  console.log('Usage:');
  console.log('    node testBigBookIn filename reader plan');
  console.log('Where:');
  console.log('    reader is one of [stream, document]');
  console.log('    plan is one of [zero, one, two, hyperlinks]');
  process.exit(0);
}

var filename = process.argv[2];
var reader = process.argv.length > 3 ? process.argv[3] : 'stream';
var plan = process.argv.length > 4 ? process.argv[4] : 'one';

var useStream = (reader === 'stream');

function getTimeout() {
  var memory = process.memoryUsage();
  var heapSize = memory.heapTotal;
  if (heapSize < 300000000) {
    return 0;
  }
  if (heapSize < 600000000) {
    return heapSize / 5000000;
  }
  if (heapSize < 1000000000) {
    return heapSize / 2000000;
  }
  return heapSize / 1000000;
}

var options = {
  reader: (useStream ? 'stream' : 'document'),
  filename: filename,
  plan: plan,
  gc: {
    getTimeout: getTimeout
  }
};
console.log(JSON.stringify(options, null, '  '));

var stopwatch = new HrStopwatch();
stopwatch.start();

function logProgress(count) {
  var memory = process.memoryUsage();
  var txtCount = utils.fmt.number(count);
  var txtHeap = utils.fmt.number(memory.heapTotal);
  process.stdout.write('Count: ' + txtCount + ', Heap Size: ' + txtHeap + '\u001b[0G');
}

var colCount = new ColumnSum([3,6,7,8]);
var hyperlinkCount = 0;
function checkRow(row) {
  if (row.number > 1) {
    colCount.add(row);

    if (colCount.count % 1000 === 0) {
      logProgress(colCount.count);
    }
  }
  row.destroy();
}
function report() {
  console.log('Count: ' + colCount.count);
  console.log('Sums: ' + colCount);
  console.log('Hyperlinks: ' + hyperlinkCount);

  stopwatch.stop();
  console.log('Time: ' + stopwatch);
}

if (useStream) {
  var wb = new WorkbookReader(options);
  wb.on('end', function() { console.log('reached end of stream');});
  wb.on('finished', report);
  wb.on('worksheet', function(worksheet) {
    worksheet.on('row', checkRow);
  });
  wb.on('hyperlinks', function(hyperlinks) {
    hyperlinks.on('hyperlink', function(hyperlink) {
      hyperlinkCount++;
    });
  });
  wb.on('entry', function(entry) {
    console.log(JSON.stringify(entry));
  });
  switch (options.plan) {
    case 'zero':
      wb.read(filename, {
        entries: 'emit'
      });
      break;
    case 'one':
      wb.read(filename, {
        entries: 'emit',
        worksheets: 'emit'
      });
      break;
    case 'two':
      break;
    case 'hyperlinks':
      wb.read(filename, {
        hyperlinks: 'emit'
      });
      break;
    default:
      break;
  }
} else {
  var wb = new Workbook();
  wb.xlsx.readFile(filename)
    .then(function() {
      var ws = wb.getWorksheet('blort');
      ws.eachRow(checkRow);
    })
    .then(report);
}
