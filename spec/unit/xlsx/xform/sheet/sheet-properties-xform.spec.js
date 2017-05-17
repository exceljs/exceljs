'use strict';

var SheetPropertiesXform = require('../../../../../lib/xlsx/xform/sheet/sheet-properties-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'empty',
    create: function() { return new SheetPropertiesXform(); },
    preparedModel: {},
    xml: '',
    parsedModel: {},
    tests: ['render', 'renderIn']
  },
  {
    title: 'tabColor',
    create: function() { return new SheetPropertiesXform(); },
    preparedModel: {tabColor: { argb: 'FFFF0000'}},
    xml: '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>',
    parsedModel: {tabColor: { argb: 'FFFF0000'}},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'pageSetup',
    create: function() { return new SheetPropertiesXform(); },
    preparedModel: {pageSetup: {fitToPage: true}},
    xml: '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
    parsedModel: {pageSetup: {fitToPage: true}},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'tabColor + pageSetup',
    create: function() { return new SheetPropertiesXform(); },
    preparedModel: {tabColor: { argb: 'FFFF0000'}, pageSetup: {fitToPage: true}},
    xml: '<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    parsedModel: {tabColor: { argb: 'FFFF0000'}, pageSetup: {fitToPage: true}},
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('SheetPropertiesXform', function() {
  testXformHelper(expectations);
});
