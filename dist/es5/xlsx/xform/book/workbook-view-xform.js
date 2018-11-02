/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var WorkbookViewXform = module.exports = function () {};

utils.inherits(WorkbookViewXform, BaseXform, {
  render: function render(xmlStream, model) {
    var attributes = {
      xWindow: model.x || 0,
      yWindow: model.y || 0,
      windowWidth: model.width || 12000,
      windowHeight: model.height || 24000,
      firstSheet: model.firstSheet,
      activeTab: model.activeTab
    };
    if (model.visibility && model.visibility !== 'visible') {
      attributes.visibility = model.visibility;
    }
    xmlStream.leafNode('workbookView', attributes);
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'workbookView') {
      var model = this.model = {};
      var addS = function addS(name, value, dflt) {
        var s = value !== undefined ? model[name] = value : dflt;
        if (s !== undefined) {
          model[name] = s;
        }
      };
      var addN = function addN(name, value, dflt) {
        var n = value !== undefined ? model[name] = parseInt(value, 10) : dflt;
        if (n !== undefined) {
          model[name] = n;
        }
      };
      addN('x', node.attributes.xWindow, 0);
      addN('y', node.attributes.yWindow, 0);
      addN('width', node.attributes.windowWidth, 25000);
      addN('height', node.attributes.windowHeight, 10000);
      addS('visibility', node.attributes.visibility, 'visible');
      addN('activeTab', node.attributes.activeTab, undefined);
      addN('firstSheet', node.attributes.firstSheet, undefined);
      return true;
    }
    return false;
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=workbook-view-xform.js.map
