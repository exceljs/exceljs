'use strict';

var fs = require('fs');

var DrawingXform = require('../../../../../lib/xlsx/xform/drawing/drawing-xform');
var testXformHelper = require('./../test-xform-helper');

const options = {
  rels: {
    rId1: { Target: '../media/image1.jpg' },
    rId2: { Target: '../media/image2.jpg' },
  },
  mediaIndex: { image1: 0, image2: 1 },
  media: [{}, {}],
};

var expectations = [
  {
    title: 'Drawing 1',
    create: function() { return new DrawingXform({tag: 'xdr:from'}); },
    initialModel: require('./data/drawing.1.0.json'),
    preparedModel: require('./data/drawing.1.1.json'),
    xml: fs.readFileSync(__dirname + '/data/drawing.1.2.xml').toString(),
    parsedModel: require('./data/drawing.1.3.json'),
    reconciledModel: require('./data/drawing.1.4.json'),
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options,
  }
];

describe('DrawingXform', function() {
  testXformHelper(expectations);
});
