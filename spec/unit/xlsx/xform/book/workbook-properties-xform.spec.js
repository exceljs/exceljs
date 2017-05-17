'use strict';

var WorkbookPropertiesXform = require('../../../../../lib/xlsx/xform/book/workbook-properties-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'default',
    create: function() { return new WorkbookPropertiesXform(); },
    preparedModel: {},
    xml: '<workbookPr defaultThemeVersion="164011" filterPrivacy="1"></workbookPr>',
    parsedModel: {},
    tests: ['render', 'renderIn']
  },
  {
    title: 'date1904',
    create: function() { return new WorkbookPropertiesXform(); },
    preparedModel: { date1904: true},
    xml: '<workbookPr date1904="1" defaultThemeVersion="164011" filterPrivacy="1"></workbookPr>',
    parsedModel: { date1904: true},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('WorkbookPropertiesXform', function() {
  testXformHelper(expectations);
});
