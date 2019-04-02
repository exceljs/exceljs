'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var RelationshipXform = module.exports = function() {
};

utils.inherits(RelationshipXform, BaseXform, {
  render: function(xmlStream, model) {
    xmlStream.leafNode('Relationship', model);
  },

  parseOpen: function(node) {
    switch (node.name) {
      case 'Relationship':
        this.model = node.attributes;
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
