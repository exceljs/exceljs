'use strict';

var StyleXform = require('../../../../../lib/xlsx/xform/style/style-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'Default',
    create: function() { return new StyleXform(); },
    preparedModel: {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0},
    xml: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Default with xfId',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0},
    xml: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Aligned',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: { horizontal: 'center', vertical: 'middle' }},
    xml: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Font',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 0, fontId: 5, fillId: 0, borderId: 0, xfId: 0},
    xml: '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Border',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 0, fontId: 0, fillId: 0, borderId: 7, xfId: 0},
    xml: '<xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'NumFmt',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 1, fontId: 0, fillId: 0, borderId: 0, xfId: 0},
    xml: '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Fill',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 0, fontId: 0, fillId: 2, borderId: 0, xfId: 0},
    xml: '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'Protected',
    create: function() { return new StyleXform({xfId: true}); },
    preparedModel: {numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: { horizontal: 'center', vertical: 'middle' }},
    xml: '<xf borderId="1" fillId="10" fontId="3" numFmtId="0" xfId="0" applyProtection="1" applyAlignment="1" applyFill="1" applyBorder="1" applyFont="1"><protection locked="0"/><alignment horizontal="center" vertical="center"/></xf>',
    parsedModel: {borderId: 1, fillId: 10, fontId: 3, numFmtId: 0, xfId: 0, alignment: { horizontal: 'center', vertical: 'middle' }},
    tests: ['parse', 'parseIn']
  }
];

describe('StyleXform', function() {
  testXformHelper(expectations);
});
