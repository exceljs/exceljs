'use strict';

var DateXform = require('../../../../../lib/xlsx/xform/simple/date-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'date',
    create: function() { return new DateXform({tag: 'date', attr: 'val'}); },
    preparedModel: new Date('2016-07-13T00:00:00Z'),
    xml: '<date val="2016-07-13T00:00:00.000Z"/>',
    parsedModel: new Date('2016-07-13T00:00:00Z'),
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'iso-date',
    create: function() {
      return new DateXform({
        tag: 'date',
        attr: 'val',
        format: function(dt) { return dt.toISOString().split('T')[0]; },
        parse: function(value) { return new Date(value.replace('13', '14')); }
      });
    },
    preparedModel: new Date('2016-07-13T00:00:00Z'),
    xml: '<date val="2016-07-13"/>',
    parsedModel: new Date('2016-07-14T00:00:00Z'),
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'undefined',
    create: function() { return new DateXform({tag: 'date', attr: 'val'}); },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn']
  },
  {
    title: 'invalid date',
    create: function() { return new DateXform({tag: 'date', attr: undefined}); },
    preparedModel: new Date(undefined),
    xml: '<date />',
    tests: ['render']
  }


];

describe('DateXform', function() {
  testXformHelper(expectations);
});
