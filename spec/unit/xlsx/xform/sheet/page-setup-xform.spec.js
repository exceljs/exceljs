'use strict';

var PageSetupXform = require('../../../../../lib/xlsx/xform/sheet/page-setup-xform');
var testXformHelper = require('./../test-xform-helper');

// -  "blackAndWhite": false
// -  "cellComments": "None"
// -  "draft": false
// -  "errors": "displayed"
// -  "fitToHeight": 1
// -  "fitToWidth": 1
// -  "horizontalDpi": "4294967295"
// +  "horizontalDpi": 4294967295
// "orientation": "portrait"
// -  "pageOrder": "downThenOver"
// "paperSize": 9
// -  "scale": 100
// -  "verticalDpi": "4294967295"
// +  "verticalDpi": 4294967295


var expectations = [
  {
    title: 'normal',
    create: () => new PageSetupXform(),
    preparedModel: {paperSize: 9, orientation: 'portrait', horizontalDpi: 4294967295, verticalDpi: 4294967295},
    xml: '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967295" verticalDpi="4294967295"/>',
    parsedModel: {
      paperSize: 9,
      orientation: 'portrait',
      horizontalDpi: 4294967295,
      verticalDpi: 4294967295,
      firstPageNumber: 1,
      useFirstPageNumber: false,
      usePrinterDefaults: false,
      copies: 1,
      blackAndWhite: false,
      cellComments: 'None',
      draft: false,
      errors: 'displayed',
      fitToHeight: 1,
      fitToWidth: 1,
      pageOrder: 'downThenOver',
      scale: 100
    },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'options',
    create: () => new PageSetupXform(),
    preparedModel: {paperSize: 119, pageOrder: 'overThenDown', orientation: 'portrait', blackAndWhite: true, draft: true, cellComments: 'atEnd', errors: 'dash'},
    xml: '<pageSetup paperSize="119" pageOrder="overThenDown" orientation="portrait" blackAndWhite="1" draft="1" cellComments="atEnd" errors="dash"/>',
    parsedModel: {paperSize: 119, pageOrder: 'overThenDown', orientation: 'portrait', blackAndWhite: true, draft: true, cellComments: 'atEnd', errors: 'dash',
      firstPageNumber: 1, useFirstPageNumber: false, usePrinterDefaults: false, copies: 1,
      horizontalDpi: 4294967295, verticalDpi: 4294967295, fitToHeight: 1, fitToWidth: 1, scale: 100},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'defaults',
    create: () => new PageSetupXform(),
    preparedModel: {pageOrder: 'downThenOver', orientation: 'portrait', cellComments: 'None', errors: 'displayed'},
    xml: '<pageSetup orientation="portrait"/>',
    parsedModel: {pageOrder: 'downThenOver', orientation: 'portrait', cellComments: 'None', errors: 'displayed',
      firstPageNumber: 1, useFirstPageNumber: false, usePrinterDefaults: false, copies: 1,
      horizontalDpi: 4294967295, verticalDpi: 4294967295, blackAndWhite: false, draft: false, fitToHeight: 1, fitToWidth: 1, scale: 100},
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'scale and fit',
    create: () => new PageSetupXform(),
    preparedModel: {paperSize: 119, scale: 95, fitToWidth: 2, fitToHeight: 3, orientation: 'landscape'},
    xml: '<pageSetup paperSize="119" scale="95" fitToWidth="2" fitToHeight="3" orientation="landscape"/>',
    parsedModel: {paperSize: 119, scale: 95, fitToWidth: 2, fitToHeight: 3, orientation: 'landscape',
      firstPageNumber: 1, useFirstPageNumber: false, usePrinterDefaults: false, copies: 1,
      horizontalDpi: 4294967295, verticalDpi: 4294967295, blackAndWhite: false, cellComments: 'None', draft: false, errors: 'displayed', pageOrder: 'downThenOver' },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('PageSetupXform', function() {
  testXformHelper(expectations);
});
