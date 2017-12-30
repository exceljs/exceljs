/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var Enums = require('../../../doc/enums');

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var validation = {
  horizontalValues: ['left', 'center', 'right', 'fill', 'centerContinuous', 'distributed', 'justify'].reduce(function (p, v) {
    p[v] = true;return p;
  }, {}),
  horizontal: function horizontal(value) {
    return this.horizontalValues[value] ? value : undefined;
  },

  verticalValues: ['top', 'middle', 'bottom', 'distributed', 'justify'].reduce(function (p, v) {
    p[v] = true;return p;
  }, {}),
  vertical: function vertical(value) {
    if (value === 'middle') return 'center';
    return this.verticalValues[value] ? value : undefined;
  },
  wrapText: function wrapText(value) {
    return value ? true : undefined;
  },
  shrinkToFit: function shrinkToFit(value) {
    return value ? true : undefined;
  },
  textRotation: function textRotation(value) {
    switch (value) {
      case 'vertical':
        return value;
      default:
        value = utils.validInt(value);
        return value >= -90 && value <= 90 ? value : undefined;
    }
  },
  indent: function indent(value) {
    value = utils.validInt(value);
    return Math.max(0, value);
  },
  readingOrder: function readingOrder(value) {
    switch (value) {
      case 'ltr':
        return Enums.ReadingOrder.LeftToRight;
      case 'rtl':
        return Enums.ReadingOrder.RightToLeft;
      default:
        return undefined;
    }
  }
};

var textRotationXform = {
  toXml: function toXml(textRotation) {
    textRotation = validation.textRotation(textRotation);
    if (textRotation) {
      if (textRotation === 'vertical') {
        return 255;
      }

      var tr = Math.round(textRotation);
      if (tr >= 0 && tr <= 90) {
        return tr;
      }

      if (tr < 0 && tr >= -90) {
        return 90 - tr;
      }
    }
    return undefined;
  },
  toModel: function toModel(textRotation) {
    var tr = utils.validInt(textRotation);
    if (tr !== undefined) {
      if (tr === 255) {
        return 'vertical';
      }
      if (tr >= 0 && tr <= 90) {
        return tr;
      }
      if (tr > 90 && tr <= 180) {
        return 90 - tr;
      }
    }
    return undefined;
  }
};

// Alignment encapsulates translation from style.alignment model to/from xlsx
var AlignmentXform = module.exports = function () {};

utils.inherits(AlignmentXform, BaseXform, {

  get tag() {
    return 'alignment';
  },

  render: function render(xmlStream, model) {
    xmlStream.addRollback();
    xmlStream.openNode('alignment');

    var isValid = false;
    function add(name, value) {
      if (value) {
        xmlStream.addAttribute(name, value);
        isValid = true;
      }
    }
    add('horizontal', validation.horizontal(model.horizontal));
    add('vertical', validation.vertical(model.vertical));
    add('wrapText', validation.wrapText(model.wrapText) ? '1' : false);
    add('shrinkToFit', validation.shrinkToFit(model.shrinkToFit) ? '1' : false);
    add('indent', validation.indent(model.indent));
    add('textRotation', textRotationXform.toXml(model.textRotation));
    add('readingOrder', validation.readingOrder(model.readingOrder));

    xmlStream.closeNode();

    if (isValid) {
      xmlStream.commit();
    } else {
      xmlStream.rollback();
    }
  },

  parseOpen: function parseOpen(node) {
    var model = {};

    var valid = false;
    function add(truthy, name, value) {
      if (truthy) {
        model[name] = value;
        valid = true;
      }
    }
    add(node.attributes.horizontal, 'horizontal', node.attributes.horizontal);
    add(node.attributes.vertical, 'vertical', node.attributes.vertical === 'center' ? 'middle' : node.attributes.vertical);
    add(node.attributes.wrapText, 'wrapText', !!node.attributes.wrapText);
    add(node.attributes.shrinkToFit, 'shrinkToFit', !!node.attributes.shrinkToFit);
    add(node.attributes.indent, 'indent', parseInt(node.attributes.indent, 10));
    add(node.attributes.textRotation, 'textRotation', textRotationXform.toModel(node.attributes.textRotation));
    add(node.attributes.readingOrder, 'readingOrder', node.attributes.readingOrder === '2' ? 'rtl' : 'ltr');

    this.model = valid ? model : null;
  },

  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=alignment-xform.js.map
