'use strict';

var CellXform = require('../../../../../lib/xlsx/xform/sheet/cell-xform');
var testXformHelper = require('./../test-xform-helper');

var SharedStringsXform = require('../../../../../lib/xlsx/xform/strings/shared-strings-xform');
var Enums = require('../../../../../lib/doc/enums');

var fakeStyles = {
  addStyleModel: function() {return 0;},
  getStyleModel: function() {return null;}
};

var expectations = [
  {
    title: 'Number',
    create:  function() { return new CellXform()},
    preparedModel: {address: 'A1', type: Enums.ValueType.Number, value: 5},
    parsedModel: {address: 'A1', type: Enums.ValueType.Number, value: 5},
    xml: '<c r="A1"><v>5</v></c>',
    tests: ['write', 'parse']
  },
  {
    title: 'Inline String',
    create:  function() { return new CellXform()},
    preparedModel: {address: 'A1', type: Enums.ValueType.String, value: 'Foo'},
    xml: '<c r="A1" t="str"><v>Foo</v></c>',
    tests: ['write']
  },
  {
    title: 'Shared String',
    create:  function() { return new CellXform()},
    initialModel: {address: 'A1', type: Enums.ValueType.String, value: 'Foo'},
    preparedModel: {address: 'A1', type: Enums.ValueType.String, value: 'Foo', ssId: 0},
    xml: '<c r="A1" t="s"><v>0</v></c>',
    tests: ['prepare', 'write'],
    options: { sharedStrings: new SharedStringsXform(), styles: fakeStyles }
  },
  {
    title: 'Date',
    create:  function() { return new CellXform()},
    preparedModel: {address: 'A1', type: Enums.ValueType.Date, value: new Date('2016-06-09T00:00:00.000Z')},
    xml: '<c r="A1"><v>42530</v></c>',
    tests: ['write']
  },
  {
    title: 'Hyperlink',
    create:  function() { return new CellXform()},
    initialModel: {address: 'A1', type: Enums.ValueType.Hyperlink, hyperlink: 'http://www.foo.com', text: 'www.foo.com'},
    preparedModel: {address: 'A1', type: Enums.ValueType.Hyperlink, hyperlink: 'http://www.foo.com', text: 'www.foo.com', ssId: 0},
    xml: '<c r="A1" t="s"><v>0</v></c>',
    parsedModel: {address: 'A1',  type: Enums.ValueType.String, value: 0},
    reconciledModel: {address: 'A1', type: Enums.ValueType.Hyperlink, value: undefined, text: 'www.foo.com', hyperlink: 'http://www.foo.com'},
    tests: ['prepare', 'write', 'parse', 'reconcile'],
    options: { sharedStrings: new SharedStringsXform(), hyperlinks: [], hyperlinkMap: {A1: 'http://www.foo.com'}, styles: fakeStyles }
  },
  {
    title: 'String Formula',
    create:  function() { return new CellXform()},
    preparedModel: {address: 'A1', type: Enums.ValueType.Formula, formula: 'A2', result: 'Foo'},
    xml: '<c r="A1" t="str"><f>A2</f><v>Foo</v></c>',
    tests: ['write']
  },
  {
    title: 'Number Formula',
    create:  function() { return new CellXform()},
    preparedModel: {address: 'A1', type: Enums.ValueType.Formula, formula: 'A2', result: 7},
    xml: '<c r="A1"><f>A2</f><v>7</v></c>',
    tests: ['write']
  }
];

describe('CellXform', function () {
  testXformHelper(expectations);
});
