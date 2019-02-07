'use strict';

var fs = require('fs');

var chai = require('chai');

var Enums = require('../../../../../lib/doc/enums');
var XmlStream = require('../../../../../lib/utils/xml-stream');

var testXformHelper = require('../test-xform-helper');
var WorksheetXform = require('../../../../../lib/xlsx/xform/sheet/worksheet-xform');

var SharedStringsXform = require('../../../../../lib/xlsx/xform/strings/shared-strings-xform');
var StylesXform = require('../../../../../lib/xlsx/xform/style/styles-xform');

var expect = chai.expect;

var fakeStyles = {
  addStyleModel: function(style, cellType) {
    if (cellType === Enums.ValueType.Date) {
      return 1;
    } else if (style && style.font) {
      return 2;
    }
    return 0;
  },
  getStyleModel: function(id) {
    switch (id) {
      case 1:
        return {numFmt: 'mm-dd-yy' };
      case 2:
        return {font: {underline: true, size: 11, color: { theme: 10}, name: 'Calibri', family: 2, scheme: 'minor'}};
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

var expectations = [
  {
    title: 'Sheet 1',
    create: () => new WorksheetXform(),
    initialModel: fixDate(require('./data/sheet.1.0.json')),
    preparedModel: fixDate(require('./data/sheet.1.1.json')),
    xml: fs.readFileSync(__dirname + '/data/sheet.1.2.xml').toString(),
    parsedModel: require('./data/sheet.1.3.json'),
    reconciledModel: fixDate(require('./data/sheet.1.4.json')),
    tests: ['prepare', 'render', 'parse'],
    options: { sharedStrings: new SharedStringsXform(), hyperlinks: [], hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles, formulae: {}, siFormulae: 0 }
  },
  {
    title: 'Sheet 2 - Data Validations',
    create: () => new WorksheetXform(),
    initialModel: require('./data/sheet.2.0.json'),
    preparedModel: require('./data/sheet.2.1.json'),
    xml: fs.readFileSync(__dirname + '/data/sheet.2.2.xml').toString(),
    tests: ['prepare', 'render'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: [], formulae: {}, siFormulae: 0 }
  },
  {
    title: 'Sheet 3 - Empty Sheet',
    create: () => new WorksheetXform(),
    preparedModel: require('./data/sheet.3.1.json'),
    xml: fs.readFileSync(__dirname + '/data/sheet.3.2.xml').toString(),
    tests: ['render'],
    options: { styles: new StylesXform(true), sharedStrings: new SharedStringsXform(), hyperlinks: []}
  },
  {
    title: 'Sheet 5 - Shared Formulas',
    create: () => new WorksheetXform(),
    initialModel: require('./data/sheet.5.0.json'),
    preparedModel: require('./data/sheet.5.1.json'),
    xml: fs.readFileSync(__dirname + '/data/sheet.5.2.xml').toString(),
    parsedModel: require('./data/sheet.5.3.json'),
    reconciledModel: require('./data/sheet.5.4.json'),
    tests: ['prepare-render', 'parse'],
    options: { sharedStrings: new SharedStringsXform(), hyperlinks: [], hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles, formulae: {}, siFormulae: 0 }
  },
  {
    title: 'Sheet 6 - AutoFilter',
    create: () => new WorksheetXform(),
    preparedModel: require('./data/sheet.6.1.json'),
    xml: fs.readFileSync(__dirname + '/data/sheet.6.2.xml').toString(),
    parsedModel: require('./data/sheet.6.3.json'),
    tests: ['render', 'parse'],
    options: { sharedStrings: new SharedStringsXform(), hyperlinks: [], hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles, formulae: {}, siFormulae: 0 }
  },
  {
    title: 'Sheet 7 - Row Breaks',
    create: () => new WorksheetXform(),
    initialModel: require('./data/sheet.7.0.json'),
    preparedModel: require('./data/sheet.7.1.json'),
    xml: fs.readFileSync(__dirname + '/data/sheet.7.2.xml').toString(),
    tests: ['prepare', 'render'],
    options: { sharedStrings: new SharedStringsXform(), hyperlinks: [], hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles, formulae: {}, siFormulae: 0 }
  },
];

describe('WorksheetXform', function() {
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
