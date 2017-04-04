'use strict';

var fs = require('fs');

var DrawingXform = require('../../../../../lib/xlsx/xform/drawing/drawing-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Drawing 1',
    create: function() { return new DrawingXform({tag: 'xdr:from'}); },
    initialModel: require('./data/drawing.1.0.json'),
    preparedModel: require('./data/drawing.1.1.json'),
    xml: fs.readFileSync(__dirname + '/data/drawing.1.2.xml').toString(),
    parsedModel: require('./data/drawing.1.3.json'),
    tests: ['prepare', 'render', 'renderIn', 'parse']
  }
];

describe('DrawingXform', function() {
  testXformHelper(expectations);
});