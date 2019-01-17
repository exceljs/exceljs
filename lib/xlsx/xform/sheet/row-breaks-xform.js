/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const PageBreaksXform = require('./page-breaks-xform');

const utils = require('../../../utils/utils');
const ListXform = require('../list-xform');

const RowBreaksXform = (module.exports = function() {
  const options = { tag: 'rowBreaks', count: true, childXform: new PageBreaksXform() };
  ListXform.call(this, options);
});

utils.inherits(RowBreaksXform, ListXform, {
  // get tag() { return 'rowBreaks'; },

  render(xmlStream, model) {
    if (model && model.length) {
      xmlStream.openNode(this.tag, this.$);
      if (this.count) {
        xmlStream.addAttribute(this.$count, model.length);
        xmlStream.addAttribute('manualBreakCount', model.length);
      }

      const childXform = this.childXform;
      model.forEach(childModel => {
        childXform.render(xmlStream, childModel);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  },
});
