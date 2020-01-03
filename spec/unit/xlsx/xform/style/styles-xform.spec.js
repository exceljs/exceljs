'use strict';

var fs = require('fs');

var StylesXform = require('../../../../../lib/xlsx/xform/style/styles-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Styles',
    create: function() { return new StylesXform(); },
    preparedModel: require('./data/styles.1.json'),
    xml: fs.readFileSync(__dirname + '/data/styles.1.xml').toString(),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('StylesXform', function() {
  testXformHelper(expectations);
});
