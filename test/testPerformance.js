var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var HrStopwatch = require('./utils/hr-stopwatch');

var Excel = require('../excel');

var Workbook = Excel.Workbook;
var WorkbookWriter = Excel.stream.xlsx.WorkbookWriter;

if (process.argv[2] === 'help') {
  console.log('Usage:');
  console.log('    node testPerformance resultFilename testFilename');
  process.exit(0);
}

var resultFilename = process.argv[2];
var testFilename = process.argv[3];
var sleepTime = 10000;

var resultBook = new WorkbookWriter({filename: resultFilename});
var resultSheet = resultBook.addWorksheet('results');
resultSheet.columns = [
  { header: 'Count', key:'count' },
  { header: 'DocSS', key:'dss' },
  { header: 'DocSO', key:'dso' },
  { header: 'DocPS', key:'dps' },
  { header: 'DocPO', key:'dpo' },
  { header: 'StmSS', key:'sss' },
  { header: 'StmSO', key:'sso' },
  { header: 'StmPS', key:'sps' },
  { header: 'StmPO', key:'spo' }
];

// =========================================================================
// Data generation functions
function randomName(length) {
  length = length || 5;
  var text = [];
  var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';

  for( var i = 0; i < length; i++)
    text.push(possible.charAt(Math.floor(Math.random() * possible.length)));

  return text.join('');
}
function randomNum(d) {
  return Math.round(Math.random() * d);
}

// =========================================================================
// Styles
var fonts = {
  arialBlackUI14: { name: 'Arial Black', family: 2, size: 14, underline: true, italic: true },
  comicSansUdB16: { name: 'Comic Sans MS', family: 4, size: 8, underline: 'double', bold: true }
};


// =========================================================================
// test parameters
var counts = [125, 250, 500, 1000, 2000, 4000, 8000, 16000, 32000];
// var counts = [125, 250, 500];
var workbooks = ['doc', 'stream'];
var styles = ['styled', 'plain'];
var strings = ['shared', 'own'];

var passes = 3;
function reduceResults(times) {
  // 3 results, sort numerically and return the median
  times.sort(function(a,b) { return a < b;});
  return times[1];
}

function execute(options) {
  console.log('Test Run ' + options.workbook + ', ' + options.style + ', ' + options.str + ', ' + options.count);

  var wbOptions = {
    filename: testFilename,
    useStyles: options.style === 'styled',
    useSharedStrings: options.str === 'shared'
  };
  var wb = (options.workbook === 'doc') ? new Workbook(wbOptions) : new WorkbookWriter(wbOptions);
  var ws = wb.addWorksheet('data');
  ws.columns = [
    { header: 'Col 1', key:'key', width: 25 },
    { header: 'Col 2', key:'name', width: 32 },
    { header: 'Col 3', key:'age', width: 21 },
    { header: 'Col 4', key:'addr1', width: 18 },
    { header: 'Col 5', key:'addr2', width: 8 },
    { header: 'Col 6', key:'num1', width: 8 },
    { header: 'Col 7', key:'num2', width: 8 },
    { header: 'Col 8', key:'num3', width: 32, style: { font: fonts.comicSansUdB16 } }
  ];
  for (var i = 0; i < options.count; i++) {
    ws.addRow({
      key: i,
      name: randomName(5),
      age: randomNum(100),
      addr1: randomName(16),
      addr2: randomName(10),
      num1: randomNum(10000),
      num2: randomNum(100000),
      num3: randomNum(1000000)
    }).commit();
  }
  if (options.workbook === 'doc') {
    console.log('Writing doc')
    return wb.xlsx.writeFile(testFilename, wbOptions);
  } else {
    console.log('Committing Writer')
    return wb.commit();
  }
}

function runTest(options) {
  var stopwatch = new HrStopwatch();
  stopwatch.start();
  return Promise.resolve(options)
    .then(execute)
    .then(function() {
      stopwatch.stop();
      console.log('Time: ' + stopwatch)
      return stopwatch.span;
    });
}

function runTests(options) {
  return function() {
    var results = [];
    var promise = Promise.resolve();
    for (var pass = 0; pass < passes; pass++) {
      // run each test with a 10 second pause between (to let GC do its stuff)
      promise = promise.then(function() {
        return runTest(options).then(function(result) {
          results.push(result);
        })
          .delay(sleepTime)
          .then(function() {
            try {
              fs.unlinkSync(testFilename);
            }
            catch(ex) {
              console.error('Error deleting file:' + ex.message);
            }
          })
          .delay(1000);
      });
    }
    return promise.then(function() {
      var testResult = reduceResults(results);
      var key = options.workbook[0] + options.style[0] + options.str[0]
      resultSheet.lastRow.getCell(key).value = testResult;
    });
  };
}

var mainPromise = Promise.resolve();
// var mainPromise = execute(125, 'stream', 'plain', 'own');
_.each(counts, function(count) {
  mainPromise = mainPromise.then(function() {
    resultSheet.addRow().getCell('count').value = count;
  });
  _.each(workbooks, function(workbook) {
    _.each(styles, function(style) {
      _.each(strings, function(str) {
        mainPromise = mainPromise.then(runTests({count: count, workbook: workbook, style: style, str: str}));
      });
    });
  });
  mainPromise = mainPromise.then(function() {
    resultSheet.lastRow.commit();
  });
});

mainPromise = mainPromise.then(function() {
  return resultBook.commit();
})
  .then(function() {
    console.log('All Done');
  })
  .catch(function(error) {
    console.log(error.message);
  });

console.log('May the tests begin...');
