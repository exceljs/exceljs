/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var HyperlinkXform = module.exports = function() {
};

utils.inherits(HyperlinkXform, BaseXform, {

  get tag() {
    return 'hyperlink';
  },

  render: function(xmlStream, model) {
    let content;
    if (model.rId) {
      content = {
        ref: model.address,
        'r:id': model.rId,
      };
    } else {
      content = {
        ref: model.address,
        location: model.location,
        display: model.display,
      };
    }
    xmlStream.leafNode('hyperlink', content);
  },

  parseOpen: function(node) {
    if (node.name === 'hyperlink') {
      // external link or file
      if (node.attributes['r:id']) {
        this.model = {
          address: node.attributes.ref,
          rId: node.attributes['r:id'],
        };
      } else {
        // internal link
        this.model = {
          address: node.attributes.ref,
          location: node.attributes.location,
          display: node.attributes.display,
        };
      }
      return true;
    }
    return false;
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
