/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

function booleanToXml(model) {
  return model ? '1' : undefined;
}

var PrintOptionsXform = module.exports = function () {};

utils.inherits(PrintOptionsXform, BaseXform, {

  get tag() {
    return 'printOptions';
  },

  render: function render(xmlStream, model) {
    if (model) {
      var attributes = {
        headings: booleanToXml(model.showRowColHeaders),
        gridLines: booleanToXml(model.showGridLines),
        horizontalCentered: booleanToXml(model.horizontalCentered),
        verticalCentered: booleanToXml(model.verticalCentered)
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
          showRowColHeaders: node.attributes.headings === '1',
          showGridLines: node.attributes.gridLines === '1',
          horizontalCentered: node.attributes.horizontalCentered === '1',
          verticalCentered: node.attributes.verticalCentered === '1'
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
//# sourceMappingURL=print-options-xform.js.map
