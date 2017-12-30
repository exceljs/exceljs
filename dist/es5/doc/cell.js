/**
 * Copyright (c) 2014-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var colCache = require('../utils/col-cache');
var Enums = require('./enums');

var _require = require('../utils/shared-formula'),
    slideFormula = _require.slideFormula;

// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

var Cell = module.exports = function (row, column, address) {
  if (!row || !column) {
    throw new Error('A Cell needs a Row');
  }

  this._row = row;
  this._column = column;

  colCache.validateAddress(address);
  this._address = address;

  // TODO: lazy evaluation of this._value
  this._value = Value.create(Cell.Types.Null, this);

  this.style = this._mergeStyle(row.style, column.style, {});

  this._mergeCount = 0;
};

Cell.Types = Enums.ValueType;

Cell.prototype = {

  get worksheet() {
    return this._row.worksheet;
  },
  get workbook() {
    return this._row.worksheet.workbook;
  },

  // help GC by removing cyclic (and other) references
  destroy: function destroy() {
    delete this.style;
    delete this._value;
    delete this._row;
    delete this._column;
    delete this._address;
  },

  // =========================================================================
  // Styles stuff
  get numFmt() {
    return this.style.numFmt;
  },
  set numFmt(value) {
    this.style.numFmt = value;
  },
  get font() {
    return this.style.font;
  },
  set font(value) {
    this.style.font = value;
  },
  get alignment() {
    return this.style.alignment;
  },
  set alignment(value) {
    this.style.alignment = value;
  },
  get border() {
    return this.style.border;
  },
  set border(value) {
    this.style.border = value;
  },
  get fill() {
    return this.style.fill;
  },
  set fill(value) {
    this.style.fill = value;
  },

  _mergeStyle: function _mergeStyle(rowStyle, colStyle, style) {
    var numFmt = rowStyle && rowStyle.numFmt || colStyle && colStyle.numFmt;
    if (numFmt) style.numFmt = numFmt;

    var font = rowStyle && rowStyle.font || colStyle && colStyle.font;
    if (font) style.font = font;

    var alignment = rowStyle && rowStyle.alignment || colStyle && colStyle.alignment;
    if (alignment) style.alignment = alignment;

    var border = rowStyle && rowStyle.border || colStyle && colStyle.border;
    if (border) style.border = border;

    var fill = rowStyle && rowStyle.fill || colStyle && colStyle.fill;
    if (fill) style.fill = fill;

    return style;
  },

  // =========================================================================
  // return the address for this cell
  get address() {
    return this._address;
  },

  get row() {
    return this._row.number;
  },
  get col() {
    return this._column.number;
  },

  get $col$row() {
    return '$' + this._column.letter + '$' + this.row;
  },

  // =========================================================================
  // Value stuff

  get type() {
    return this._value.type;
  },

  get effectiveType() {
    return this._value.effectiveType;
  },

  toCsvString: function toCsvString() {
    return this._value.toCsvString();
  },

  // =========================================================================
  // Merge stuff

  addMergeRef: function addMergeRef() {
    this._mergeCount++;
  },
  releaseMergeRef: function releaseMergeRef() {
    this._mergeCount--;
  },
  get isMerged() {
    return this._mergeCount > 0 || this.type === Cell.Types.Merge;
  },
  merge: function merge(master) {
    this._value.release();
    this._value = Value.create(Cell.Types.Merge, this, master);
    this.style = master.style;
  },
  unmerge: function unmerge() {
    if (this.type === Cell.Types.Merge) {
      this._value.release();
      this._value = Value.create(Cell.Types.Null, this);
      this.style = this._mergeStyle(this._row.style, this._column.style, {});
    }
  },
  isMergedTo: function isMergedTo(master) {
    if (this._value.type !== Cell.Types.Merge) return false;
    return this._value.isMergedTo(master);
  },
  get master() {
    if (this.type === Cell.Types.Merge) {
      return this._value.master;
    }
    return this; // an unmerged cell is its own master
  },

  get isHyperlink() {
    return this._value.type === Cell.Types.Hyperlink;
  },
  get hyperlink() {
    return this._value.hyperlink;
  },

  // return the value
  get value() {
    return this._value.value;
  },
  // set the value - can be number, string or raw
  set value(v) {
    // special case - merge cells set their master's value
    if (this.type === Cell.Types.Merge) {
      this._value.master.value = v;
      return;
    }

    this._value.release();

    // assign value
    this._value = Value.create(Value.getType(v), this, v);
  },
  get text() {
    return this._value.toString();
  },
  toString: function toString() {
    return this.text;
  },

  _upgradeToHyperlink: function _upgradeToHyperlink(hyperlink) {
    // if this cell is a string, turn it into a Hyperlink
    if (this.type === Cell.Types.String) {
      this._value = Value.create(Cell.Types.Hyperlink, this, {
        text: this._value.value,
        hyperlink: hyperlink
      });
    }
  },

  // =========================================================================
  // Formula stuff
  get formula() {
    return this._value.formula;
  },
  get result() {
    return this._value.result;
  },
  get formulaType() {
    return this._value.formulaType;
  },

  // =========================================================================
  // Name stuff
  get fullAddress() {
    var worksheet = this._row.worksheet;
    return {
      sheetName: worksheet.name,
      address: this.address,
      row: this.row,
      col: this.col
    };
  },
  get name() {
    return this.names[0];
  },
  set name(value) {
    this.names = [value];
  },
  get names() {
    return this.workbook.definedNames.getNamesEx(this.fullAddress);
  },
  set names(value) {
    var self = this;
    var definedNames = this.workbook.definedNames;
    this.workbook.definedNames.removeAllNames(self.fullAddress);
    value.forEach(function (name) {
      definedNames.addEx(self.fullAddress, name);
    });
  },
  addName: function addName(name) {
    this.workbook.definedNames.addEx(this.fullAddress, name);
  },
  removeName: function removeName(name) {
    this.workbook.definedNames.removeEx(this.fullAddress, name);
  },
  removeAllNames: function removeAllNames() {
    this.workbook.definedNames.removeAllNames(this.fullAddress);
  },

  // =========================================================================
  // Data Validation stuff
  get _dataValidations() {
    return this.worksheet.dataValidations;
  },

  get dataValidation() {
    return this._dataValidations.find(this.address);
  },

  set dataValidation(value) {
    this._dataValidations.add(this.address, value);
  },

  // =========================================================================
  // Model stuff

  get model() {
    var model = this._value.model;
    model.style = this.style;
    return model;
  },
  set model(value) {
    this._value.release();
    this._value = Value.create(value.type, this);
    this._value.model = value;
    if (value.style) {
      this.style = value.style;
    } else {
      this.style = {};
    }
  }
};

// =============================================================================
// Internal Value Types

var NullValue = function NullValue(cell) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Null
  };
};
NullValue.prototype = {
  get value() {
    return null;
  },
  set value(value) {
    // nothing to do
  },
  get type() {
    return Cell.Types.Null;
  },
  get effectiveType() {
    return Cell.Types.Null;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return '';
  },
  release: function release() {},
  toString: function toString() {
    return '';
  }
};

var NumberValue = function NumberValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Number,
    value: value
  };
};
NumberValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  get type() {
    return Cell.Types.Number;
  },
  get effectiveType() {
    return Cell.Types.Number;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return '' + this.model.value;
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value.toString();
  }
};

var StringValue = function StringValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.String,
    value: value
  };
};
StringValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  get type() {
    return Cell.Types.String;
  },
  get effectiveType() {
    return Cell.Types.String;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return '"' + this.model.value.replace(/"/g, '""') + '"';
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value;
  }
};

var RichTextValue = function RichTextValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.String,
    value: value
  };
};
RichTextValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  toString: function toString() {
    return this.model.value.richText.map(function (t) {
      return t.text;
    }).join('');
  },
  get type() {
    return Cell.Types.RichText;
  },
  get effectiveType() {
    return Cell.Types.RichText;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return '"' + this.text.replace(/"/g, '""') + '"';
  },
  release: function release() {}
};

var DateValue = function DateValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Date,
    value: value
  };
};
DateValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  get type() {
    return Cell.Types.Date;
  },
  get effectiveType() {
    return Cell.Types.Date;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return this.model.value.toISOString();
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value.toString();
  }
};

var HyperlinkValue = function HyperlinkValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Hyperlink,
    text: value ? value.text : undefined,
    hyperlink: value ? value.hyperlink : undefined
  };
};
HyperlinkValue.prototype = {
  get value() {
    return {
      text: this.model.text,
      hyperlink: this.model.hyperlink
    };
  },
  set value(value) {
    this.model.text = value.text;
    this.model.hyperlink = value.hyperlink;
  },

  get text() {
    return this.model.text;
  },
  set text(value) {
    this.model.text = value;
  },
  get hyperlink() {
    return this.model.hyperlink;
  },
  set hyperlink(value) {
    this.model.hyperlink = value;
  },
  get type() {
    return Cell.Types.Hyperlink;
  },
  get effectiveType() {
    return Cell.Types.Hyperlink;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return this.model.hyperlink;
  },
  release: function release() {},
  toString: function toString() {
    return this.model.text;
  }
};

var MergeValue = function MergeValue(cell, master) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Merge,
    master: master ? master.address : undefined
  };
  this._master = master;
  if (master) {
    master.addMergeRef();
  }
};
MergeValue.prototype = {
  get value() {
    return this._master.value;
  },
  set value(value) {
    if (value instanceof Cell) {
      if (this._master) {
        this._master.releaseMergeRef();
      }
      value.addMergeRef();
      this._master = value;
    } else {
      this._master.value = value;
    }
  },

  isMergedTo: function isMergedTo(master) {
    return master === this._master;
  },
  get master() {
    return this._master;
  },
  get type() {
    return Cell.Types.Merge;
  },
  get effectiveType() {
    return this._master.effectiveType;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return '';
  },
  release: function release() {
    this._master.releaseMergeRef();
  },
  toString: function toString() {
    return this.value.toString();
  }
};

var FormulaValue = function FormulaValue(cell, value) {
  this.cell = cell;
  this.model = {
    address: cell.address,
    type: Cell.Types.Formula,
    formula: value ? value.formula : undefined,
    sharedFormula: value ? value.sharedFormula : undefined,
    result: value ? value.result : undefined
  };
};
FormulaValue.prototype = {
  get value() {
    return this.model.formula ? {
      formula: this.model.formula,
      result: this.model.result
    } : {
      sharedFormula: this.model.sharedFormula,
      result: this.model.result
    };
  },
  set value(value) {
    this.model.formula = value.formula;
    this.model.sharedFormula = value.sharedFormula;
    this.model.result = value.result;
  },
  validate: function validate(value) {
    switch (Value.getType(value)) {
      case Cell.Types.Null:
      case Cell.Types.String:
      case Cell.Types.Number:
      case Cell.Types.Date:
        break;
      case Cell.Types.Hyperlink:
      case Cell.Types.Formula:
      default:
        throw new Error('Cannot process that type of result value');
    }
  },
  get dependencies() {
    // find all the ranges and cells mentioned in the formula
    var ranges = this.formula.match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g);
    var cells = this.formula.replace(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g, '').match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g);
    return {
      ranges: ranges,
      cells: cells
    };
  },
  get formula() {
    return this.model.formula || this._getTranslatedFormula();
  },
  set formula(value) {
    this.model.formula = value;
  },
  get formulaType() {
    if (this.model.formula) {
      return Enums.FormulaType.Master;
    }
    if (this.model.sharedFormula) {
      return Enums.FormulaType.Shared;
    }
    return Enums.FormulaType.None;
  },
  get result() {
    return this.model.result;
  },
  set result(value) {
    this.model.result = value;
  },
  get type() {
    return Cell.Types.Formula;
  },
  get effectiveType() {
    var v = this.model.result;
    if (v === null || v === undefined) {
      return Enums.ValueType.Null;
    } else if (v instanceof String || typeof v === 'string') {
      return Enums.ValueType.String;
    } else if (typeof v === 'number') {
      return Enums.ValueType.Number;
    } else if (v instanceof Date) {
      return Enums.ValueType.Date;
    } else if (v.text && v.hyperlink) {
      return Enums.ValueType.Hyperlink;
    } else if (v.formula) {
      return Enums.ValueType.Formula;
    }

    return Enums.ValueType.Null;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  _getTranslatedFormula: function _getTranslatedFormula() {
    if (!this._translatedFormula && this.model.sharedFormula) {
      var worksheet = this.cell.worksheet;
      var master = worksheet.findCell(this.model.sharedFormula);
      this._translatedFormula = master && slideFormula(master.formula, master.address, this.model.address);
    }
    return this._translatedFormula;
  },

  toCsvString: function toCsvString() {
    return '' + (this.model.result || '');
  },
  release: function release() {},
  toString: function toString() {
    return this.model.result ? this.model.result.toString() : '';
  }
};

var SharedStringValue = function SharedStringValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.SharedString,
    value: value
  };
};
SharedStringValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  get type() {
    return Cell.Types.SharedString;
  },
  get effectiveType() {
    return Cell.Types.SharedString;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return '' + this.model.value;
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value.toString();
  }
};

var BooleanValue = function BooleanValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Boolean,
    value: value
  };
};
BooleanValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  get type() {
    return Cell.Types.Boolean;
  },
  get effectiveType() {
    return Cell.Types.Boolean;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return this.model.value ? 1 : 0;
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value.toString();
  }
};

var ErrorValue = function ErrorValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Error,
    value: value
  };
};
ErrorValue.prototype = {
  get value() {
    return this.model.value;
  },
  set value(value) {
    this.model.value = value;
  },
  get type() {
    return Cell.Types.Error;
  },
  get effectiveType() {
    return Cell.Types.Error;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return this.toString();
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value.error.toString();
  }
};

var JSONValue = function JSONValue(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.String,
    value: JSON.stringify(value),
    rawValue: value
  };
};
JSONValue.prototype = {
  get value() {
    return this.model.rawValue;
  },
  set value(value) {
    this.model.rawValue = value;
    this.model.value = JSON.stringify(value);
  },
  get type() {
    return Cell.Types.String;
  },
  get effectiveType() {
    return Cell.Types.String;
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    this.model.address = value;
  },
  toCsvString: function toCsvString() {
    return this.model.value;
  },
  release: function release() {},
  toString: function toString() {
    return this.model.value;
  }
};

// Value is a place to hold common static Value type functions
var Value = {
  getType: function getType(value) {
    if (value === null || value === undefined) {
      return Cell.Types.Null;
    } else if (value instanceof String || typeof value === 'string') {
      return Cell.Types.String;
    } else if (typeof value === 'number') {
      return Cell.Types.Number;
    } else if (typeof value === 'boolean') {
      return Cell.Types.Boolean;
    } else if (value instanceof Date) {
      return Cell.Types.Date;
    } else if (value.text && value.hyperlink) {
      return Cell.Types.Hyperlink;
    } else if (value.formula || value.sharedFormula) {
      return Cell.Types.Formula;
    } else if (value.richText) {
      return Cell.Types.RichText;
    } else if (value.sharedString) {
      return Cell.Types.SharedString;
    } else if (value.error) {
      return Cell.Types.Error;
    }
    return Cell.Types.JSON;
  },

  // map valueType to constructor
  types: [{ t: Cell.Types.Null, f: NullValue }, { t: Cell.Types.Number, f: NumberValue }, { t: Cell.Types.String, f: StringValue }, { t: Cell.Types.Date, f: DateValue }, { t: Cell.Types.Hyperlink, f: HyperlinkValue }, { t: Cell.Types.Formula, f: FormulaValue }, { t: Cell.Types.Merge, f: MergeValue }, { t: Cell.Types.JSON, f: JSONValue }, { t: Cell.Types.SharedString, f: SharedStringValue }, { t: Cell.Types.RichText, f: RichTextValue }, { t: Cell.Types.Boolean, f: BooleanValue }, { t: Cell.Types.Error, f: ErrorValue }].reduce(function (p, t) {
    p[t.t] = t.f;return p;
  }, []),

  create: function create(type, cell, value) {
    var T = this.types[type];
    if (!T) {
      throw new Error('Could not create Value of type ' + type);
    }
    return new T(cell, value);
  }
};
//# sourceMappingURL=cell.js.map
