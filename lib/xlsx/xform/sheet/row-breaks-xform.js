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

      const { childXform } = this;
      model.forEach(childModel => {
        childXform.render(xmlStream, childModel);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  },
});
