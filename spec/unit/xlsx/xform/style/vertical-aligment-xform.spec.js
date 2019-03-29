'use strict';

var StringXform = require('../../../../../lib/xlsx/xform/simple/string-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'superscript',
    create: function() { return new StringXform({ tag: 'vertAlign', attr: 'val' }); },
    preparedModel: 'superscript',
    xml: '<vertAlign val="superscript"/>',
    parsedModel: 'superscript',
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'subscript',
    create: function() { return new StringXform({ tag: 'vertAlign', attr: 'val' }); },
    preparedModel: 'subscript ',
    xml: '<vertAlign val="subscript"/>',
    parsedModel: 'subscript',
    tests: ['render', 'renderIn', 'parse']
  }
];
