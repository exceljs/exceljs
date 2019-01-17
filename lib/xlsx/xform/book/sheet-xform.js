/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const WorksheetXform = (module.exports = function() {});

utils.inherits(WorksheetXform, BaseXform, {
  render(xmlStream, model) {
    xmlStream.leafNode('sheet', {
      sheetId: model.id,
      name: model.name,
      state: model.state,
      'r:id': model.rId,
    });
  },

  parseOpen(node) {
    if (node.name === 'sheet') {
      this.model = {
        name: utils.xmlDecode(node.attributes.name),
        id: parseInt(node.attributes.sheetId, 10),
        state: node.attributes.state,
        rId: node.attributes['r:id'],
      };
      return true;
    }
    return false;
  },
  parseText() {},
  parseClose() {
    return false;
  },
});
