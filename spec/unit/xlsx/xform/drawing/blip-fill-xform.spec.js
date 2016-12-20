'use strict';

var BlipFillXform = require('../../../../../lib/xlsx/xform/drawing/blip-fill-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'empty',
    create: function() { return new BlipFillXform(); },
    preparedModel: { rId: 'rId1' },
    xml: '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" /></xdr:blipFill>',
    parsedModel: { rId: 'rId1' },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'full',
    create: function() { return new BlipFillXform(); },
    preparedModel: { rId: 'rId1' },
    xml: '<xdr:blipFill>' +
           '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print">' +
             '<a:extLst>' +
               '<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' +
                 '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>' +
               '</a:ext>' +
             '</a:extLst>' +
           '</a:blip>' +
         '</xdr:blipFill>',
    parsedModel: { rId: 'rId1' },
    tests: ['parse']
  }
];

describe('BlipFillXform', function() {
  testXformHelper(expectations);
});