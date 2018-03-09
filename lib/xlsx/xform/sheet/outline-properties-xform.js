/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var OutlinePropertiesXform = module.exports = function() {};

var isDefined = (attr) => typeof attr !== 'undefined';

utils.inherits(OutlinePropertiesXform, BaseXform, {

  get tag() { return 'outlinePr'; },

  render: function(xmlStream, model) {
    if (model && (isDefined(model.summaryBelow) || isDefined(model.summaryRight))) {
      xmlStream.leafNode(this.tag, {
        summaryBelow: isDefined(model.summaryBelow) ? Number(model.summaryBelow) : undefined,
        summaryRight: isDefined(model.summaryRight) ? Number(model.summaryRight) : undefined,
      });
      return true;
    }
    return false;
  },

  parseOpen: function(node) {
    if (node.name === this.tag) {
      this.model = {
        summaryBelow: isDefined(node.attributes.summaryBelow) ? Boolean(Number(node.attributes.summaryBelow)) : undefined,
        summaryRight: isDefined(node.attributes.summaryRight) ? Boolean(Number(node.attributes.summaryRight)) : undefined,
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
