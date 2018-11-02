/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var Enums = require('../../../doc/enums');
var Range = require('../../../doc/range');

var RichTextXform = require('../strings/rich-text-xform');

function getValueType(v) {
  if (v === null || v === undefined) {
    return Enums.ValueType.Null;
  } else if (v instanceof String || typeof v === 'string') {
    return Enums.ValueType.String;
  } else if (typeof v === 'number') {
    return Enums.ValueType.Number;
  } else if (typeof v === 'boolean') {
    return Enums.ValueType.Boolean;
  } else if (v instanceof Date) {
    return Enums.ValueType.Date;
  } else if (v.text && v.hyperlink) {
    return Enums.ValueType.Hyperlink;
  } else if (v.formula) {
    return Enums.ValueType.Formula;
  } else if (v.error) {
    return Enums.ValueType.Error;
  }
  throw new Error('I could not understand type of value');
}

function getEffectiveCellType(cell) {
  switch (cell.type) {
    case Enums.ValueType.Formula:
      return getValueType(cell.result);
    default:
      return cell.type;
  }
}

var CellXform = module.exports = function () {
  this.richTextXForm = new RichTextXform();
};

utils.inherits(CellXform, BaseXform, {

  get tag() {
    return 'c';
  },

  prepare: function prepare(model, options) {
    var styleId = options.styles.addStyleModel(model.style || {}, getEffectiveCellType(model));
    if (styleId) {
      model.styleId = styleId;
    }

    switch (model.type) {
      case Enums.ValueType.String:
        if (options.sharedStrings) {
          model.ssId = options.sharedStrings.add(model.value);
        }
        break;
      case Enums.ValueType.Date:
        if (options.date1904) {
          model.date1904 = true;
        }
        break;
      case Enums.ValueType.Hyperlink:
        if (options.sharedStrings) {
          model.ssId = options.sharedStrings.add(model.text);
        }
        options.hyperlinks.push({
          address: model.address,
          target: model.hyperlink
        });
        break;
      case Enums.ValueType.Merge:
        options.merges.add(model);
        break;
      case Enums.ValueType.Formula:
        if (options.date1904) {
          // in case valueType is date
          model.date1904 = true;
        }
        if (model.formula) {
          options.formulae[model.address] = model;
        } else if (model.sharedFormula) {
          var master = options.formulae[model.sharedFormula];
          if (!master) {
            throw new Error('Shared Formula master must exist above and or left of clone');
          }
          if (master.si !== undefined) {
            model.si = master.si;
            master.ref.expandToAddress(model.address);
          } else {
            model.si = master.si = options.siFormulae++;
            master.ref = new Range(master.address, model.address);
          }
        }
        break;
      default:
        break;
    }
  },

  renderFormula: function renderFormula(xmlStream, model) {
    var attrs = null;
    if (model.ref) {
      attrs = {
        t: 'shared',
        ref: model.ref.range,
        si: model.si
      };
    } else if (model.si !== undefined) {
      attrs = {
        t: 'shared',
        si: model.si
      };
    }
    switch (getValueType(model.result)) {
      case Enums.ValueType.Null:
        // ?
        xmlStream.leafNode('f', attrs, model.formula);
        break;
      case Enums.ValueType.String:
        // oddly, formula results don't ever use shared strings
        xmlStream.addAttribute('t', 'str');
        xmlStream.leafNode('f', attrs, model.formula);
        xmlStream.leafNode('v', null, model.result);
        break;
      case Enums.ValueType.Number:
        xmlStream.leafNode('f', attrs, model.formula);
        xmlStream.leafNode('v', null, model.result);
        break;
      case Enums.ValueType.Boolean:
        xmlStream.addAttribute('t', 'b');
        xmlStream.leafNode('f', attrs, model.formula);
        xmlStream.leafNode('v', null, model.result ? 1 : 0);
        break;
      case Enums.ValueType.Error:
        xmlStream.addAttribute('t', 'e');
        xmlStream.leafNode('f', attrs, model.formula);
        xmlStream.leafNode('v', null, model.result.error);
        break;
      case Enums.ValueType.Date:
        xmlStream.leafNode('f', attrs, model.formula);
        xmlStream.leafNode('v', null, utils.dateToExcel(model.result, model.date1904));
        break;
      // case Enums.ValueType.Hyperlink: // ??
      // case Enums.ValueType.Formula:
      default:
        throw new Error('I could not understand type of value');
    }
  },

  render: function render(xmlStream, model) {
    if (model.type === Enums.ValueType.Null && !model.styleId) {
      // if null and no style, exit
      return;
    }

    xmlStream.openNode('c');
    xmlStream.addAttribute('r', model.address);

    if (model.styleId) {
      xmlStream.addAttribute('s', model.styleId);
    }

    switch (model.type) {
      case Enums.ValueType.Null:
        break;

      case Enums.ValueType.Number:
        xmlStream.leafNode('v', null, model.value);
        break;

      case Enums.ValueType.Boolean:
        xmlStream.addAttribute('t', 'b');
        xmlStream.leafNode('v', null, model.value ? '1' : '0');
        break;

      case Enums.ValueType.Error:
        xmlStream.addAttribute('t', 'e');
        xmlStream.leafNode('v', null, model.value.error);
        break;

      case Enums.ValueType.String:
        if (model.ssId !== undefined) {
          xmlStream.addAttribute('t', 's');
          xmlStream.leafNode('v', null, model.ssId);
        } else {
          if (model.value && model.value.richText) {
            xmlStream.addAttribute('t', 'inlineStr');
            xmlStream.openNode('is');
            var self = this;
            model.value.richText.forEach(function (text) {
              self.richTextXForm.render(xmlStream, text);
            });
            xmlStream.closeNode('is');
          } else {
            xmlStream.addAttribute('t', 'str');
            xmlStream.leafNode('v', null, model.value);
          }
        }
        break;

      case Enums.ValueType.Date:
        xmlStream.leafNode('v', null, utils.dateToExcel(model.value, model.date1904));
        break;

      case Enums.ValueType.Hyperlink:
        if (model.ssId !== undefined) {
          xmlStream.addAttribute('t', 's');
          xmlStream.leafNode('v', null, model.ssId);
        } else {
          xmlStream.addAttribute('t', 'str');
          xmlStream.leafNode('v', null, model.text);
        }
        break;

      case Enums.ValueType.Formula:
        this.renderFormula(xmlStream, model);
        break;

      case Enums.ValueType.Merge:
        // nothing to add
        break;

      default:
        break;
    }

    xmlStream.closeNode(); // </c>
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'c':
        // var address = colCache.decodeAddress(node.attributes.r);
        var model = this.model = {
          address: node.attributes.r
        };
        this.t = node.attributes.t;
        if (node.attributes.s) {
          model.styleId = parseInt(node.attributes.s, 10);
        }
        return true;

      case 'f':
        this.currentNode = 'f';
        this.model.si = node.attributes.si;
        if (node.attributes.t === 'shared') {
          this.model.sharedFormula = true;
        }
        this.model.ref = node.attributes.ref;
        return true;

      case 'v':
        this.currentNode = 'v';
        return true;

      case 't':
        this.currentNode = 't';
        return true;

      case 'r':
        this.parser = this.richTextXForm;
        this.parser.parseOpen(node);
        return true;

      default:
        return false;
    }
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
      return;
    }
    switch (this.currentNode) {
      case 'f':
        this.model.formula = this.model.formula ? this.model.formula + text : text;
        break;
      case 'v':
      case 't':
        if (this.model.value && this.model.value.richText) {
          this.model.value.richText.text = this.model.value.richText.text ? this.model.value.richText.text + text : text;
        } else {
          this.model.value = this.model.value ? this.model.value + text : text;
        }
        break;
      default:
        break;
    }
  },
  parseClose: function parseClose(name) {
    switch (name) {
      case 'c':
        var model = this.model;

        // first guess on cell type
        if (model.formula || model.sharedFormula) {
          model.type = Enums.ValueType.Formula;
          if (model.value) {
            if (this.t === 'str') {
              model.result = utils.xmlDecode(model.value);
            } else if (this.t === 'b') {
              model.result = parseInt(model.value, 10) !== 0;
            } else if (this.t === 'e') {
              model.result = { error: model.value };
            } else {
              model.result = parseFloat(model.value);
            }
            model.value = undefined;
          }
        } else if (model.value !== undefined) {
          switch (this.t) {
            case 's':
              model.type = Enums.ValueType.String;
              model.value = parseInt(model.value, 10);
              break;
            case 'str':
              model.type = Enums.ValueType.String;
              model.value = utils.xmlDecode(model.value);
              break;
            case 'inlineStr':
              model.type = Enums.ValueType.String;
              break;
            case 'b':
              model.type = Enums.ValueType.Boolean;
              model.value = parseInt(model.value, 10) !== 0;
              break;
            case 'e':
              model.type = Enums.ValueType.Error;
              model.value = { error: model.value };
              break;
            default:
              model.type = Enums.ValueType.Number;
              model.value = parseFloat(model.value);
              break;
          }
        } else if (model.styleId) {
          model.type = Enums.ValueType.Null;
        } else {
          model.type = Enums.ValueType.Merge;
        }
        return false;
      case 'f':
      case 'v':
      case 'is':
        this.currentNode = undefined;
        return true;
      case 't':
        if (this.parser) {
          this.parser.parseClose(name);
          return true;
        } else {
          this.currentNode = undefined;
          return true;
        }
      case 'r':
        this.model.value = this.model.value || {};
        this.model.value.richText = this.model.value.richText || [];
        this.model.value.richText.push(this.parser.model);
        this.parser = undefined;
        this.currentNode = undefined;
        return true;
      default:
        if (this.parser) {
          this.parser.parseClose(name);
          return true;
        }
        return false;
    }
  },

  reconcile: function reconcile(model, options) {
    var style = model.styleId && options.styles.getStyleModel(model.styleId);
    if (style) {
      model.style = style;
    }
    if (model.styleId !== undefined) {
      model.styleId = undefined;
    }

    switch (model.type) {
      case Enums.ValueType.String:
        if (typeof model.value === 'number') {
          model.value = options.sharedStrings.getString(model.value);
        }
        if (model.value.richText) {
          model.type = Enums.ValueType.RichText;
        }
        break;
      case Enums.ValueType.Number:
        if (style && utils.isDateFmt(style.numFmt)) {
          model.type = Enums.ValueType.Date;
          model.value = utils.excelToDate(model.value, options.date1904);
        }
        break;
      case Enums.ValueType.Formula:
        if (model.result !== undefined && style && utils.isDateFmt(style.numFmt)) {
          model.result = utils.excelToDate(model.result, options.date1904);
        }
        if (model.sharedFormula) {
          if (model.formula) {
            options.formulae[model.si] = model;
            delete model.sharedFormula;
          } else {
            model.sharedFormula = options.formulae[model.si].address;
          }
          delete model.si;
        }
        break;
      default:
        break;
    }

    // look for hyperlink
    var hyperlink = options.hyperlinkMap[model.address];
    if (hyperlink) {
      if (model.type === Enums.ValueType.Formula) {
        model.text = model.result;
        model.result = undefined;
      } else {
        model.text = model.value;
        model.value = undefined;
      }
      model.type = Enums.ValueType.Hyperlink;
      model.hyperlink = hyperlink;
    }
  }
});
//# sourceMappingURL=cell-xform.js.map
