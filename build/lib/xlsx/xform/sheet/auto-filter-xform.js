/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var BaseXform = require('../base-xform');

var AutoFilterXform = module.exports = function () {};

utils.inherits(AutoFilterXform, BaseXform, {

  get tag() {
    return 'autoFilter';
  },

  render: function render(xmlStream, model) {
    if (model) {
      if (typeof model === 'string') {
        // assume range
        xmlStream.leafNode('autoFilter', { ref: model });
      } else {
        var getAddress = function getAddress(addr) {
          if (typeof addr === 'string') {
            return addr;
          }
          return colCache.getAddress(addr.row, addr.column).address;
        };

        var firstAddress = getAddress(model.from);
        var secondAddress = getAddress(model.to);
        if (firstAddress && secondAddress) {
          xmlStream.leafNode('autoFilter', { ref: firstAddress + ':' + secondAddress });
        }
      }
    }
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'autoFilter') {
      this.model = node.attributes.ref;
    }
  }
});
//# sourceMappingURL=auto-filter-xform.js.map
