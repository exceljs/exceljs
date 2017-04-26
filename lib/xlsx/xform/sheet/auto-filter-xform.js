/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var BaseXform = require('../base-xform');

var AutoFilterXform = module.exports = function() {
};

utils.inherits(AutoFilterXform, BaseXform, {

  get tag() { return 'autoFilter'; },

  render: function(xmlStream, model) {
    function getAddress (autoFilterObject) {
      if (autoFilterObject) {
        if (typeof autoFilterObject === 'string') {
            return colCache.getAddress(autoFilterObject);
        } else if (autoFilterObject.row && autoFilterObject.column) {
            return colCache.getAddress(autoFilterObject.row, autoFilterObject.column);
        }
      }
    }
    if (model) {
      var firstAddress = getAddress(model.from);
      var secondAddress = getAddress(model.to);
      if (firstAddress && secondAddress) {
        xmlStream.leafNode('autoFilter', {ref: firstAddress.address + ':' + secondAddress.address});
      }
    }
  },

  parseOpen: function(node) {
    if (node.name === 'autoFilter') {
      var cells = node.attributes.ref.split(':');
      if (cells.length === 2) {
        var firstAddress = colCache.getAddress(cells[0]);
        var secondAddress = colCache.getAddress(cells[1]);
        this.model = {
          from: {
            row: firstAddress.row,
            column: firstAddress.col
          },
          to: {
            row: secondAddress.row,
            column: secondAddress.col
          }
        }
      }
    }
  }

});