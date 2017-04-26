/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var RelationshipXform = module.exports = function() {
};

utils.inherits(RelationshipXform, BaseXform, {
  render: function(xmlStream, model) {
    xmlStream.leafNode('Relationship', {
      Id: model.rId,
      Type: model.type,
      Target: model.target,
      TargetMode: model.targetMode
    });
  },
  
  parseOpen: function(node) {
    switch(node.name) {
      case 'Relationship':
        this.model = {
          rId: node.attributes.Id,
          type: node.attributes.Type,
          target: node.attributes.Target
        };
        if (node.attributes.TargetMode) {
          this.model.targetMode = node.attributes.TargetMode;
        }
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
