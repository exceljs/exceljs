'use strict';

var DefinedNameXform = require('../../../../../lib/xlsx/xform/book/defined-name-xform');

var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Defined Names',
    create:  function() { return new DefinedNameXform()},
    preparedModel: {name: 'foo', ranges:['bar!$A1$C1']},
    get parsedModel() { return this.preparedModel; },
    xml: '<definedName name="foo">bar!$A1$C1</definedName>',
    tests: ['write', 'parse']
  }
];

describe('DefinedNameXform', function () {
  testXformHelper(expectations);
});
