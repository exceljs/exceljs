'use strict';

var fs = require('fs');

var RelationshipsXform = require('../../../../../lib/xlsx/xform/core/relationships-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'worksheet.rels',
    create: function() { return new RelationshipsXform(); },
    preparedModel: require('./data/worksheet.rels.1.json'),
    xml: fs.readFileSync(__dirname + '/data/worksheet.rels.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('RelationshipsXform', function() {
  testXformHelper(expectations);
});
