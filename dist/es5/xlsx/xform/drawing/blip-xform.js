/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var BlipXform = module.exports = function () {};

utils.inherits(BlipXform, BaseXform, {

  get tag() {
    return 'a:blip';
  },

  render: function render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:embed': model.rId,
      cstate: 'print'
    });
    // TODO: handle children (e.g. a:extLst=>a:ext=>a14:useLocalDpi
  },

  parseOpen: function parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:embed']
        };
        return true;
      default:
        return true;
    }
  },

  parseText: function parseText() {},

  parseClose: function parseClose(name) {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes
        return true;
    }
  }
});
//# sourceMappingURL=blip-xform.js.map
