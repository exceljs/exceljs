'use strict';

var PhoneticTextXform = require('../../../../../lib/xlsx/xform/strings/phonetic-text-xform');
var testXformHelper = require('./../test-xform-helper');

var expectations = [
  {
    title: 'text',
    create:  function() { return new PhoneticTextXform(); },
    preparedModel: { text: 'Hello, World!', sb: 0, eb: 1 },
    xml: '<rPh sb="0" eb="1"><t>Hello, World!</t></rPh>',
    parsedModel: { text: 'Hello, World!', sb: 0, eb: 1 },
    tests: ['render', 'renderIn', 'parse']
  }
];

describe('PhoneticTextXform', function () {
  testXformHelper(expectations);
});
