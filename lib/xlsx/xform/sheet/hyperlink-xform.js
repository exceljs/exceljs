'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var HyperlinkXform = module.exports = function() {
};

utils.inherits(HyperlinkXform, BaseXform, {

  get tag() { return 'hyperlink'; },

  render: function(xmlStream, model) {
    xmlStream.leafNode('hyperlink', Object.assign({
      ref: model.address,
      'r:id': model.rId,
    }, model.tooltip ? {tooltip: model.tooltip} : {}));
  },

  parseOpen: function(node) {
    if (node.name === 'hyperlink') {
      this.model = Object.assign({
        address: node.attributes.ref,
        rId: node.attributes['r:id'],
      }, node.attributes.tooltip ? {tooltip: node.attributes.tooltip} : {});
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
