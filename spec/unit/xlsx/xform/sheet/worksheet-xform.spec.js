'use strict';

var fs = require('fs');

var chai    = require('chai');
var expect  = chai.expect;

var Enums = require('../../../../../lib/doc/enums');
var XmlStream = require('../../../../../lib/utils/xml-stream');

var testXformHelper = require('../test-xform-helper');
var WorksheetXform = require('../../../../../lib/xlsx/xform/sheet/worksheet-xform');

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
  B6: 'https://www.npmjs.com/package/exceljs'
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
    create:  function() { return new WorksheetXform(); },
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
    create:  function() { return new WorksheetXform(); },
    initialModel: data[1][0],
    preparedModel: data[1][1],
    xml: data[1][2],
    tests: ['prepare', 'render'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  },
  {
    title: 'Sheet 3 - Empty Sheet',
    create:  function() { return new WorksheetXform(); },
    preparedModel: data[2][0],
    xml: data[2][1],
    tests: ['render'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  }
];

describe('WorksheetXform', function () {
  testXformHelper(expectations);

  it('hyperlinks must be after dataValidations', function() {
    var xform = new WorksheetXform();
    var model = require('./data/sheet.4.0.json');
    var xmlStream = new XmlStream();
    var options = { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []};
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    var xml = xmlStream.xml;
    var iHyperlinks = xml.indexOf('hyperlinks');
    var iDataValidations = xml.indexOf('dataValidations');
    expect(iHyperlinks).not.to.equal(-1);
    expect(iDataValidations).not.to.equal(-1);
    expect(iHyperlinks).to.be.greaterThan(iDataValidations);
  });
});
