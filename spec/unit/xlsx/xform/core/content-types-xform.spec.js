'use strict';

var fs = require('fs');

var ContentTypesXform = require('../../../../../lib/xlsx/xform/core/content-types-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Three Sheets',
    create: function() { return new ContentTypesXform(); },
    preparedModel: { worksheets: [{id: 1}, {id: 2}, {id: 3}, ], media: [], drawings: [] },
    xml: fs.readFileSync(__dirname + '/data/content-types.01.xml').toString().replace(/\r\n/g, '\n'),
    tests: ['render']
  },
  {
    title: 'Images',
    create: function() { return new ContentTypesXform(); },
    preparedModel: { worksheets: [{id: 1}, {id: 2}], media: [{type: 'image', extension: 'png'}, {type: 'image', extension: 'jpg'}], drawings: [] },
    xml: fs.readFileSync(__dirname + '/data/content-types.02.xml').toString().replace(/\r\n/g, '\n'),
    tests: ['render']
  }
];

describe('ContentTypesXform', function() {
  testXformHelper(expectations);
});
