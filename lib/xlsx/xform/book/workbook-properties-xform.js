'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const WorksheetPropertiesXform = (module.exports = function() {});

utils.inherits(WorksheetPropertiesXform, BaseXform, {
  render(xmlStream, model) {
    xmlStream.leafNode('workbookPr', {
      date1904: model.date1904 ? 1 : undefined,
      defaultThemeVersion: 164011,
      filterPrivacy: 1,
    });
  },

  parseOpen(node) {
    if (node.name === 'workbookPr') {
      this.model = {
        date1904: node.attributes.date1904 === '1',
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
