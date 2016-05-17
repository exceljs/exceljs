'use strict';


var FontXform = require('../../../../lib/xlsx/xform/font-xform');
var testXformHelper = require('./test-xform-helper');

var expectations = [
  {
    title: 'green bold',
    create: function() { return new FontXform(); },
    model: {bold: true, size: 14, color: {argb:'FF00FF00'}, name: 'Calibri', family: 2, scheme: 'minor'},
    xml: '<font><b/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><name val="Calibri"/></font>'
  },
  {
    title: 'rPr tag',
    create: function() { return new FontXform(undefined, {tagName: 'rPr',fontNameTag:  'rFont'}); },
    model: {italic: true, size: 14, color: {argb:'FF00FF00'}, name: 'Calibri', family: 2, scheme: 'minor'},
    xml: '<rPr><i/><color rgb="FF00FF00"/><family val="2"/><scheme val="minor"/><sz val="14"/><rFont val="Calibri"/></rPr>'
  }
];

testXformHelper('FontXform', expectations);
