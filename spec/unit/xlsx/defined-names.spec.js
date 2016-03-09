'use strict';

var expect = require('chai').expect;

var _ = require('underscore');
var DefinedNames = require('../../../lib/xlsx/defined-names');

describe('DefinedNames', function() {

  it('Generates xml', function() {
    var dn = new DefinedNames();

    dn.add('blort!A1','foo');
    expect(dn.xml).to.equal('<definedNames><definedName name="foo">blort!$A$1</definedName>');
  });

});