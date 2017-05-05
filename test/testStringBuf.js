var Promise = require('bluebird');

var utils = require('./utils/utils');
var HrStopwatch = require('./utils/hr-stopwatch');

var StringBuf = require('../lib/utils/string-buf.js');

var SIZE = 1048576;

function testWrite(results) {
  var a = [];
  function test(size) {
    return function() {
      console.log('Write: ' + size);
      var text = utils.randomName(size);
      var sb = new StringBuf({size:SIZE + 10});
      var sw = new HrStopwatch();
      sw.start();
      while (sb.length < SIZE) {
        sb.addText(text);
      }
      sw.stop();
      a.push('' + size + ':' + (Math.round(sw.span*1000)));
    }
  }
  return Promise.resolve()
    .then(test(1))
    .delay(1000)
    .then(test(2))
    .delay(1000)
    .then(test(4))
    .delay(1000)
    .then(test(8))
    .delay(1000)
    .then(test(16))
    .delay(1000)
    .then(test(32))
    .delay(1000)
    .then(test(64))
    .delay(1000)
    .then(function() {
      results.write = a.join(', ');
      return results;
    });
}

function testGrow(results) {
  var a = [];
  function test(size) {
    return function() {
      console.log('Grow: ' + size);
      var text = utils.randomName(size);
      var sb = new StringBuf({size:8});
      var sw = new HrStopwatch();
      sw.start();
      while (sb.length < SIZE) {
        sb.addText(text);
      }
      sw.stop();
      a.push('' + size + ':' + (Math.round(sw.span*1000)));
    }
  }
  return Promise.resolve()
    .then(test(1))
    .delay(1000)
    .then(test(2))
    .delay(1000)
    .then(test(4))
    .delay(1000)
    .then(test(8))
    .delay(1000)
    .then(test(16))
    .delay(1000)
    .then(test(32))
    .delay(1000)
    .then(test(64))
    .delay(1000)
    .then(function() {
      results.grow = a.join(', ');
      return results;
    });
}

var results = {};
Promise.resolve(results)
  .then(testWrite)
  .then(testGrow)
  .then(function(r) {
    console.log(JSON.stringify(r, null, '  '));
  });