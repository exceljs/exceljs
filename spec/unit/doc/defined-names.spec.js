'use strict';

var expect = require('chai').expect;

var DefinedNames = require('../../../lib/doc/defined-names');

describe('DefinedNames', function() {
  it('adds names for cells', function() {
    var dn = new DefinedNames();

    dn.add('blort!A1', 'foo');
    expect(dn.getNames('blort!A1')).to.deep.equal(['foo']);
    expect(dn.getNames('blort!$A$1')).to.deep.equal(['foo']);

    dn.add('blort!$B$4', 'bar');
    expect(dn.getNames('blort!B4')).to.deep.equal(['bar']);
    expect(dn.getNames('blort!$B$4')).to.deep.equal(['bar']);

    dn.add("'blo rt'!$B$4", 'bar');
    expect(dn.getNames("'blo rt'!$B$4")).to.deep.equal(['bar']);
    dn.add("'blo ,!rt'!$B$4", 'bar');
    expect(dn.getNames("'blo ,!rt'!$B$4")).to.deep.equal(['bar']);
  });

  it('removes names for cells', function() {
    var dn = new DefinedNames();

    dn.add('blort!A1', 'foo');
    dn.add('blort!A1', 'bar');
    dn.remove('blort!A1', 'foo');

    expect(dn.getNames('blort!A1')).to.deep.equal(['bar']);
  });

  // get ranges example
  it('gets the right ranges for a name', function() {
    var dn = new DefinedNames();

    dn.add('blort!A1', 'vertical');
    dn.add('blort!A2', 'vertical');
    dn.add('blort!A3', 'vertical');

    dn.add('blort!C1', 'horizontal');
    dn.add('blort!D1', 'horizontal');
    dn.add('blort!E1', 'horizontal');

    dn.add('blort!C3', 'square');
    dn.add('blort!D3', 'square');
    dn.add('blort!C4', 'square');
    dn.add('blort!D4', 'square');

    dn.add('other!A1', 'single');

    expect(dn.getRanges('vertical')).to.deep.equal({name: 'vertical', ranges: ['blort!$A$1:$A$3']});
    expect(dn.getRanges('horizontal')).to.deep.equal({name: 'horizontal', ranges: ['blort!$C$1:$E$1']});
    expect(dn.getRanges('square')).to.deep.equal({name: 'square', ranges: ['blort!$C$3:$D$4']});
    expect(dn.getRanges('single')).to.deep.equal({name: 'single', ranges: ['other!$A$1']});
  });

  it('creates matrix from model', function() {
    var dn = new DefinedNames();

    dn.model = [];
    dn.add('blort!A1', 'bar');
    dn.remove('blort!A1', 'foo');

    expect(dn.getNames('blort!A1')).to.deep.equal(['bar']);
  });

  it('skips values with invalid range', function() {
    var dn = new DefinedNames();
    dn.model = [
      {name: 'eq', ranges: ['"="']},
      {name: 'ref', ranges: ['#REF!']},
      {name: 'single', ranges: ['Sheet3!$A$1']},
      {name: 'range', ranges: ['Sheet3!$A$2:$F$2228']}
    ];

    expect(dn.model).to.deep.equal([
      {name: 'single', ranges: ['Sheet3!$A$1']},
      {name: 'range', ranges: ['Sheet3!$A$2:$F$2228']}
    ]);
  });
});
