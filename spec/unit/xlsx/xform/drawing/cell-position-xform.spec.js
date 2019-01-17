'use strict';

const CellPositionXform = require('../../../../../lib/xlsx/xform/drawing/cell-position-xform');
const testXformHelper = require('./../test-xform-helper');

const expectations = [
  {
    title: 'integers',
    create() {
      return new CellPositionXform({ tag: 'xdr:from' });
    },
    preparedModel: { row: 5, col: 7 },
    xml:
      '<xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    parsedModel: { row: 5, col: 7 },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'halves',
    create() {
      return new CellPositionXform({ tag: 'xdr:to' });
    },
    preparedModel: { row: 5.5, col: 7.5 },
    xml:
      '<xdr:to><xdr:col>7</xdr:col><xdr:colOff>320000</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>90000</xdr:rowOff></xdr:to>',
    parsedModel: { row: 5.5, col: 7.5 },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('CellPositionXform', () => {
  testXformHelper(expectations);
});
