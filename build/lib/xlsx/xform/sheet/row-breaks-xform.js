/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var PageBreaksXform = require('./page-breaks-xform');

var utils = require('../../../utils/utils');
var ListXform = require('../list-xform');

var RowBreaksXform = module.exports = function () {
  var options = { tag: 'rowBreaks', count: true, childXform: new PageBreaksXform() };
  ListXform.call(this, options);
};

utils.inherits(RowBreaksXform, ListXform, {

  // get tag() { return 'rowBreaks'; },

  render: function render(xmlStream, model) {
    if (model && model.length) {
      xmlStream.openNode(this.tag, this.$);
      if (this.count) {
        xmlStream.addAttribute(this.$count, model.length);
        xmlStream.addAttribute('manualBreakCount', model.length);
      }

      var childXform = this.childXform;
      model.forEach(function (childModel) {
        childXform.render(xmlStream, childModel);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  }
});
//# sourceMappingURL=row-breaks-xform.js.map
