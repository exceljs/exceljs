'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var MergeCellXform = module.exports = function() {
};

utils.inherits(MergeCellXform, BaseXform, {

  get tag() { return 'mergeCell'; },

  render: function(xmlStream, model) {
    xmlStream.leafNode('mergeCell', {ref: model});
  },

  parseOpen: function(node) {
    if (node.name === 'mergeCell') {
      this.model = node.attributes.ref;
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
