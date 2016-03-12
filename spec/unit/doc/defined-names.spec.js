'use strict';

var expect = require('chai').expect;

var _ = require('underscore');
var DefinedNames = require('../../../lib/doc/defined-names');

describe('DefinedNames', function() {

  it('adds names for cells', function() {
    var dn = new DefinedNames();

    dn.add('blort!A1','foo');
    expect(dn.getNames('blort!A1')).to.deep.equal(['foo']);
  });

});