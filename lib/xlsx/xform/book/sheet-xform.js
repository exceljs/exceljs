'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var WorksheetXform = module.exports = function() {
};

utils.inherits(WorksheetXform, BaseXform, {

  render: function(xmlStream, model) {
    xmlStream.leafNode('sheet', {
      sheetId: model.id,
      name: model.name,
      state: model.state,
      'r:id': model.rId
    });
  },

  parseOpen: function(node) {
    if (node.name === 'sheet') {
      this.model = {
        name: utils.xmlDecode(node.attributes.name),
        id: parseInt(node.attributes.sheetId, 10),
        state: node.attributes.state,
        rId: node.attributes['r:id']
      };
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
