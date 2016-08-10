'use strict';

var fs = require('fs');
var Enums = require('../../../../../lib/doc/enums');

var SheetXform = require('../../../../../lib/xlsx/xform/sheet/worksheet-xform');
var testXformHelper = require('../test-xform-helper');

var SharedStringsXform = require('../../../../../lib/xlsx/xform/strings/shared-strings-xform');
var StylesXform = require('../../../../../lib/xlsx/xform/style/styles-xform');

var fakeStyles = {
  addStyleModel: function(style, cellType) {
    if (cellType === Enums.ValueType.Date) {
      return 1;
    } else if (style && style.font) {
      return 2;
    } else {
      return 0;
    }
  },
  getStyleModel: function(id) {
    switch(id) {
      case 1:
        return {"numFmt": "mm-dd-yy" };
      case 2:
        return {"font": {"underline": true, "size": 11, "color": {"theme":10}, "name": "Calibri", "family": 2, "scheme": "minor"} };
      default:
        return null;
    }
  }
};

var fakeHyperlinkMap = {
  getHyperlink: function(address) {
    switch(address) {
      case 'B6':
        return 'https://www.npmjs.com/package/exceljs';
      default:
        return;
    }
  }
};

function fixDate(model) {
  model.rows[3].cells[1].value = new Date(model.rows[3].cells[1].value);
  return model;
}
var data = [
  [
    fixDate(require('./data/sheet.1.0.json')),
    fixDate(require('./data/sheet.1.1.json')),
    fs.readFileSync(__dirname + '/data/sheet.1.2.xml').toString(),
    require('./data/sheet.1.3.json'),
    fixDate(require('./data/sheet.1.4.json'))
  ],
  [
    require('./data/sheet.2.0.json'),
    require('./data/sheet.2.1.json'),
    fs.readFileSync(__dirname + '/data/sheet.2.2.xml').toString()
  ],
  [
    require('./data/sheet.3.1.json'),
    fs.readFileSync(__dirname + '/data/sheet.3.2.xml').toString()
  ]
];

var expectations = [
  {
    title: 'Sheet 1',
    create:  function() { return new SheetXform(); },
    initialModel: data[0][0],
    preparedModel: data[0][1],
    xml: data[0][2],
    parsedModel: data[0][3],
    reconciledModel: data[0][4],
    tests: ['prepare', 'render', 'parse'],
    options: { sharedStrings: new SharedStringsXform(), hyperlinks: [], hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles }
  },
  {
    title: 'Sheet 2 - Data Validations',
    create:  function() { return new SheetXform(); },
    initialModel: data[1][0],
    preparedModel: data[1][1],
    xml: data[1][2],
    tests: ['prepare', 'render'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  },
  {
    title: 'Sheet 3 - Empty Sheet',
    create:  function() { return new SheetXform(); },
    preparedModel: data[2][0],
    xml: data[2][1],
    tests: ['render'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  }
];

describe('SheetXform', function () {
  testXformHelper(expectations);
});
