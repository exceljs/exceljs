/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

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
      'r:id': model.rId
    });
  },
  
  parseOpen: function(node) {
    if (node.name === 'sheet') {
      this.model = {
        name: utils.xmlDecode(node.attributes.name),
        id: parseInt(node.attributes.sheetId, 10),
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
