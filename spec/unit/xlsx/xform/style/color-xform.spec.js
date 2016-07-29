'use strict';

var ColorXform = require('../../../../../lib/xlsx/xform/style/color-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'RGB',
    create:  function() { return new ColorXform()},
    preparedModel: {argb:'FF00FF00'},
    xml: '<color rgb="FF00FF00"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Theme',
    create:  function() { return new ColorXform()},
    preparedModel: {theme:1},
    xml: '<color theme="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Theme with Tint',
    create:  function() { return new ColorXform()},
    preparedModel: {theme:1, tint: 0.5},
    xml: '<color theme="1" tint="0.5"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Indexed',
    create:  function() { return new ColorXform()},
    preparedModel: {indexed: 1},
    xml: '<color indexed="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Undefined',
    create:  function() { return new ColorXform()},
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn']
  }
];

describe('ColorXform', function () {
  testXformHelper(expectations);
});
