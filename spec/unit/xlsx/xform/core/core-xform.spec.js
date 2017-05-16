'use strict';

var fs = require('fs');

var CoreXform = require('../../../../../lib/xlsx/xform/core/core-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'core.xml',
    create: () => new CoreXform(),
    preparedModel: { creator: 'Guyon Roche', lastModifiedBy: 'Guyon Roche', created: new Date('2016-04-20T16:26:46Z'), modified: new Date('2016-05-12T06:52:49Z')},
    xml: fs.readFileSync(__dirname + '/data/core.01.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'core.xml',
    create: () => new CoreXform(),
    preparedModel: {
      creator: 'Guyon Roche',
      title: 'My Little Xlsx',
      subject: 'An Xlsx about Xlsxs',
      description: 'A lot of stuff',
      keywords: 'xlsx,test',
      category: 'unit test material',
      lastModifiedBy: 'Guyon Roche',
      created: new Date('2016-04-20T16:26:46Z'),
      modified: new Date('2016-05-12T06:52:49Z')
    },
    xml: fs.readFileSync(__dirname + '/data/core.02.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'core.xml - with cp:lastPrinted',
    create: () => new CoreXform(),
    preparedModel: { creator: 'Guyon Roche', lastModifiedBy: 'Guyon Roche', lastPrinted: new Date('2016-08-16T19:56:07Z'), created: new Date('2016-04-20T16:26:46Z'), modified: new Date('2016-05-12T06:52:49Z')},
    xml: fs.readFileSync(__dirname + '/data/core.03.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  },
  {
    title: 'core.xml - with cp:contentStatus',
    create: () => new CoreXform(),
    preparedModel: { creator: 'Guyon Roche', lastModifiedBy: 'Guyon Roche', contentStatus: 'Final', created: new Date('2016-04-20T16:26:46Z'), modified: new Date('2016-05-12T06:52:49Z')},
    xml: fs.readFileSync(__dirname + '/data/core.04.xml').toString().replace(/\r\n/g, '\n'),
    get parsedModel() { return this.preparedModel; },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('CoreXform', function() {
  testXformHelper(expectations);
});
