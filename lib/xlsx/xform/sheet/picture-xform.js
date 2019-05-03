'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var PictureXform = module.exports = function() {
};

utils.inherits(PictureXform, BaseXform, {

  get tag() { return 'picture'; },

  render: function(xmlStream, model) {
    if (model) {
      xmlStream.leafNode(this.tag, {'r:id': model.rId});
    }
  },

  parseOpen: function(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:id']
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
