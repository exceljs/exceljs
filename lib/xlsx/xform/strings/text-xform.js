/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

//   <t xml:space="preserve"> is </t>

var TextXform = module.exports = function() {
};

utils.inherits(TextXform, BaseXform, {

  get tag() { return 't'; },

  render: function(xmlStream, model) {
    xmlStream.openNode('t');
    if ((model[0] === ' ') || (model[model.length - 1] === ' ')) {
      xmlStream.addAttribute('xml:space', 'preserve');
    }
    xmlStream.writeText(model);
    xmlStream.closeNode();
  },
  
  get model() {
    return this._text.join('');
  },

  parseOpen: function(node) {
    switch (node.name) {
      case 't':
        this._text = [];
        return true;
      default:
        return false;
    }
  },
  parseText: function(text) {
    this._text.push(text);
  },
  parseClose: function() {
    return false;
  }
});
