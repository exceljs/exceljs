'use strict';


var FontXform = require('../../../../../lib/xlsx/xform/style/font-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'green bold',
    create: function() { return new FontXform(); },
    preparedModel: {bold: true, size: 14, color: {argb: 'FF00FF00'}, name: 'Calibri', family: 2, scheme: 'minor'},
    xml: '<font><b/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><name val="Calibri"/></font>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'rPr tag',
    create: function() { return new FontXform({tagName: 'rPr', fontNameTag: 'rFont'}); },
    preparedModel: {italic: true, size: 14, color: {argb: 'FF00FF00'}, name: 'Calibri', family: 2, scheme: 'minor'},
    xml: '<rPr><i/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><rFont val="Calibri"/></rPr>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('FontXform', function() {
  testXformHelper(expectations);
});
