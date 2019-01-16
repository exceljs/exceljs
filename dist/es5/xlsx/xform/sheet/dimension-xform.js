/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var DimensionXform = module.exports = function () {};

utils.inherits(DimensionXform, BaseXform, {

  get tag() {
    return 'dimension';
  },

  render: function render(xmlStream, model) {
    if (model) {
      xmlStream.leafNode('dimension', { ref: model });
    }
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'dimension') {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=dimension-xform.js.map
