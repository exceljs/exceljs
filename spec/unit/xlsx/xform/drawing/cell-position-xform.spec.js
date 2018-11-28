'use strict';

var CellPositionXform = require('../../../../../lib/xlsx/xform/drawing/cell-position-xform');
var Anchor = require('../../../../../lib/doc/anchor');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'integers',
    create: function() { return new CellPositionXform({tag: 'xdr:from'}); },
    preparedModel: new Anchor({ row: 5, col: 7 }),
    xml: '<xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    parsedModel:new Anchor({ row: 5, col: 7 }),
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'halves',
    create: function() { return new CellPositionXform({tag: 'xdr:to'}); },
    preparedModel: new Anchor({ row: 5.5, col: 7.5 }),
    xml: '<xdr:to><xdr:col>7</xdr:col><xdr:colOff>320000</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>90000</xdr:rowOff></xdr:to>',
    parsedModel: new Anchor({ row: 5.5, col: 7.5 }),
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('CellPositionXform', function() {
  testXformHelper(expectations);
});
