'use strict';

var BooleanXform = require('../../../../../lib/xlsx/xform/simple/boolean-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'true',
    create: function() { return new BooleanXform({tag: 'boolean', attr: 'val'}); },
    preparedModel: true,
    get parsedModel() { return this.preparedModel; },
    xml: '<boolean/>',
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'false',
    create: function() { return new BooleanXform({tag: 'boolean', attr: 'val'}); },
    preparedModel: false,
    xml: '',
    tests: ['render', 'renderIn']
  },
  {
    title: 'undefined',
    create: function() { return new BooleanXform({tag: 'boolean', attr: 'val'}); },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn']
  }
];

describe('BooleanXform', function() {
  testXformHelper(expectations);
});
