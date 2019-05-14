'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const RelationshipXform = (module.exports = function() {});

utils.inherits(RelationshipXform, BaseXform, {
  render(xmlStream, model) {
    xmlStream.leafNode('Relationship', model);
  },

  parseOpen(node) {
    switch (node.name) {
      case 'Relationship':
        this.model = node.attributes;
        return true;
      default:
        return false;
    }
  },
  parseText() {},
  parseClose() {
    return false;
  },
});
