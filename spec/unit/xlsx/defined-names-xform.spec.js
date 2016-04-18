'use strict';

var expect = require('chai').expect;

var _ = require('underscore');
var DefinedNamesXform = require('../../../lib/xlsx/defined-names-xform');

describe('DefinedNamesXform', function() {
  it('translate to xml', function() {
    var model = [
      {name: 'foo', ranges:['bar!$A1$C1']}
    ];
    var xml =
      '<definedNames>' +
        '<definedName name="foo">bar!$A1$C1</definedName>' +
      '</definedNames>';

    var dnx = new DefinedNamesXform(model);
      expect(dnx.xml).to.equal(xml);
  });
});