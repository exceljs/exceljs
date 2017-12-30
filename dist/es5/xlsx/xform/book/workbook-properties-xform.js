/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var WorksheetPropertiesXform = module.exports = function () {};

utils.inherits(WorksheetPropertiesXform, BaseXform, {

  render: function render(xmlStream, model) {
    xmlStream.leafNode('workbookPr', {
      date1904: model.date1904 ? 1 : undefined,
      defaultThemeVersion: 164011,
      filterPrivacy: 1
    });
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'workbookPr') {
      this.model = {
        date1904: node.attributes.date1904 === '1'
      };
      return true;
    }
    return false;
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=workbook-properties-xform.js.map
