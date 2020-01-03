'use strict';

var expect = require('chai').expect;

var CellMatrix = require('../../../lib/utils/cell-matrix');

describe('CellMatrix', function() {
  it('getCell always returns a cell', function() {
    var cm = new CellMatrix();
    expect(cm.getCell('A1')).to.be.ok();
    expect(cm.getCell('$B$2')).to.be.ok();
    expect(cm.getCell('Sheet!$C$3')).to.be.ok();
  });
  it('findCell only returns known cells', function() {
    var cm = new CellMatrix();
    expect(cm.findCell('A1')).to.be.undefined();
    expect(cm.findCell('$B$2')).to.be.undefined();
    expect(cm.findCell('Sheet!$C$3')).to.be.undefined();
  });
});
