/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var PageMarginsXform = module.exports = function () {};

utils.inherits(PageMarginsXform, BaseXform, {

  get tag() {
    return 'pageMargins';
  },

  render: function render(xmlStream, model) {
    if (model) {
      var attributes = {
        left: model.left,
        right: model.right,
        top: model.top,
        bottom: model.bottom,
        header: model.header,
        footer: model.footer
      };
      if (_.some(attributes, function (value) {
        return value !== undefined;
      })) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  },

  parseOpen: function parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          left: parseFloat(node.attributes.left || 0.7),
          right: parseFloat(node.attributes.right || 0.7),
          top: parseFloat(node.attributes.top || 0.75),
          bottom: parseFloat(node.attributes.bottom || 0.75),
          header: parseFloat(node.attributes.header || 0.3),
          footer: parseFloat(node.attributes.footer || 0.3)
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
//# sourceMappingURL=page-margins-xform.js.map
