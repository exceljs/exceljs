'use strict';

var DefinedNameXform = require('../../../../../lib/xlsx/xform/book/defined-name-xform');

var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Defined Names',
    create: function() { return new DefinedNameXform(); },
    preparedModel: {name: 'foo', ranges: ['bar!$A$1:$C$1'], formula: ''},
    xml: '<definedName name="foo">bar!$A$1:$C$1</definedName>',
    parsedModel: {name: 'foo', ranges: ['bar!$A$1:$C$1'], formula: ''},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Print Area',
    create: function() { return new DefinedNameXform(); },
    preparedModel: {name: '_xlnm.Print_Area', localSheetId: 0, ranges: ['bar!$A$1:$C$10'], formula: ''},
    xml: '<definedName name="_xlnm.Print_Area" localSheetId="0">bar!$A$1:$C$10</definedName>',
    parsedModel: {name: '_xlnm.Print_Area', localSheetId: 0, ranges: ['bar!$A$1:$C$10'], formula: ''},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('DefinedNameXform', function() {
  testXformHelper(expectations);
});
