/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var ExtXform = module.exports = function(options) {
    this.tag = options.tag;
    this.map = {
  };
};

/** https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML */
const EMU_PER_PIXEL_AT_96_DPI = 9525

utils.inherits(ExtXform, BaseXform, {

  render: function(xmlStream, model) {
	  xmlStream.openNode(this.tag);

    var width = Math.floor(model.width * EMU_PER_PIXEL_AT_96_DPI);
    var height = Math.floor(model.height * EMU_PER_PIXEL_AT_96_DPI);
    
    xmlStream.addAttribute('cx', width);
    xmlStream.addAttribute('cy', height);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
	  if (node.name == this.tag) {
      this.model = {
        width: parseInt(node.attributes.cx || '0', 10) / EMU_PER_PIXEL_AT_96_DPI,
        height: parseInt(node.attributes.cy || '0', 10) / EMU_PER_PIXEL_AT_96_DPI,
      };
      return true;
    }
    return false;
  },

  parseText: function(text) {
  },

  parseClose: function(name) {
    return false;
  }
});
