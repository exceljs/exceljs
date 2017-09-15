'use strict';

var fs = require('fs');
var chai = require('chai');
var expect = chai.expect;

var TwoCellAnchorXform = require('../../../../../lib/xlsx/xform/drawing/two-cell-anchor-xform');


describe('TwoCellAnchorXform', function() {
  describe('reconcile', function() {
    it('should not throw on null picture', function() {
      var twoCell = new TwoCellAnchorXform();
      expect(twoCell.reconcile({picture: null}, {})).to.not.throw;
    });
    it('should not throw on null tl', function() {
      var twoCell = new TwoCellAnchorXform();
      expect(twoCell.reconcile({tl: null}, {})).to.not.throw;
    });
  });
});
