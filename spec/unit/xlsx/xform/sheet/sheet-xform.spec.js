'use strict';

var fs = require('fs');

var SheetXform = require('../../../../../lib/xlsx/xform/sheet/sheet-xform');
var testXformHelper = require('../test-xform-helper');

var SharedStringsXform = require('../../../../../lib/xlsx/xform/strings/shared-strings-xform');
var StylesXform = require('../../../../../lib/xlsx/xform/style/styles-xform');

function fixDate(model) {
  model.rows[3].cells[1].value = new Date(model.rows[3].cells[1].value);
  return model;
}
var models = [
  [
    fixDate(require('./data/sheet.1.0.json')),
    fixDate(require('./data/sheet.1.1.json'))
  ],
  [
    require('./data/sheet.2.0.json'),
    require('./data/sheet.2.1.json')
  ]
];

var expectations = [
  {
    title: 'Sheet 1',
    create:  function() { return new SheetXform(); },
    initialModel: models[0][0],
    preparedModel: models[0][1],
    xml: fs.readFileSync(__dirname + '/data/sheet.1.xml').toString(),
    tests: ['prepare', 'write'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  },
  {
    title: 'Sheet 2 - Data Validations',
    create:  function() { return new SheetXform(); },
    initialModel: models[1][0],
    preparedModel: models[1][1],
    xml: fs.readFileSync(__dirname + '/data/sheet.2.xml').toString(),
    tests: ['prepare', 'write'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  }
];

describe.only('SheetXform', function () {
  testXformHelper(expectations);
});
