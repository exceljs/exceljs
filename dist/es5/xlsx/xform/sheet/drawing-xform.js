/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var DrawingXform = module.exports = function () {};

utils.inherits(DrawingXform, BaseXform, {

  get tag() {
    return 'drawing';
  },

  render: function render(xmlStream, model) {
    if (model) {
      xmlStream.leafNode(this.tag, { 'r:id': model.rId });
    }
  },

  parseOpen: function parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:id']
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=drawing-xform.js.map
