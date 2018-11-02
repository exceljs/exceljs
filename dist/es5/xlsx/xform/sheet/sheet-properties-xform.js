/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var ColorXform = require('../style/color-xform');
var PageSetupPropertiesXform = require('./page-setup-properties-xform');
var OutlinePropertiesXform = require('./outline-properties-xform');

var SheetPropertiesXform = module.exports = function () {
  this.map = {
    tabColor: new ColorXform('tabColor'),
    pageSetUpPr: new PageSetupPropertiesXform(),
    outlinePr: new OutlinePropertiesXform()
  };
};

utils.inherits(SheetPropertiesXform, BaseXform, {

  get tag() {
    return 'sheetPr';
  },

  render: function render(xmlStream, model) {
    if (model) {
      xmlStream.addRollback();
      xmlStream.openNode('sheetPr');

      var inner = false;
      inner = this.map.tabColor.render(xmlStream, model.tabColor) || inner;
      inner = this.map.pageSetUpPr.render(xmlStream, model.pageSetup) || inner;
      inner = this.map.outlinePr.render(xmlStream, model.outlineProperties) || inner;

      if (inner) {
        xmlStream.closeNode();
        xmlStream.commit();
      } else {
        xmlStream.rollback();
      }
    }
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      this.reset();
      return true;
    }
    if (this.map[node.name]) {
      this.parser = this.map[node.name];
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
      return true;
    }
    return false;
  },
  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (this.map.tabColor.model || this.map.pageSetUpPr.model || this.map.outlinePr.model) {
      this.model = {};
      if (this.map.tabColor.model) {
        this.model.tabColor = this.map.tabColor.model;
      }
      if (this.map.pageSetUpPr.model) {
        this.model.pageSetup = this.map.pageSetUpPr.model;
      }
      if (this.map.outlinePr.model) {
        this.model.outlineProperties = this.map.outlinePr.model;
      }
    } else {
      this.model = null;
    }
    return false;
  }
});
//# sourceMappingURL=sheet-properties-xform.js.map
