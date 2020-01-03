var Promise = require('bluebird');

var utils = require('./utils/utils');
var HrStopwatch = require('./utils/hr-stopwatch');

var StreamBuf = require('../lib/utils/stream-buf.js');

var sb = new StreamBuf({bufSize:64});
sb.write('Hello, World!');
console.log('Buffer after write: ' + sb.buffers[0].buffer);
var chunk = sb.read();
console.log('Chunk: ' + chunk);
console.log('to UTF8: ' + chunk.toString('UTF8'));