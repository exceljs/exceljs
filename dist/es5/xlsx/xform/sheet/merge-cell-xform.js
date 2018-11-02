/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var MergeCellXform = module.exports = function () {};

utils.inherits(MergeCellXform, BaseXform, {

  get tag() {
    return 'mergeCell';
  },

  render: function render(xmlStream, model) {
    xmlStream.leafNode('mergeCell', { ref: model });
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'mergeCell') {
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
//# sourceMappingURL=merge-cell-xform.js.map
