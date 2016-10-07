/**
 * Copyright (c) 2014 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var colCache = require('./../utils/col-cache');
var utils = require('./../utils/utils');
var Enums = require('./enums');

// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

var Cell = module.exports = function(row, column, address) {
  if (!row || !column) {
    throw new Error('A Cell needs a Row');
  }

  //if (!(row instanceof Row)) {
  //  throw new Error('Expected row to be Row, was ' + typeof row);
  //}
  //if (!(column instanceof Column)) {
  //  throw new Error('Expected col to be Column, was ' + typeof column);
  //}

  this._row = row;
  this._column = column;

  colCache.validateAddress(address);
  this._address = address;

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
  destroy: function() {
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
    return this.style.numFmt = value;
  },
  get font() {
    return this.style.font;
  },
  set font(value) {
    return this.style.font = value;
  },
  get alignment() {
    return this.style.alignment;
  },
  set alignment(value) {
    return this.style.alignment = value;
  },
  get border() {
    return this.style.border;
  },
  set border(value) {
    return this.style.border = value;
  },
  get fill() {
    return this.style.fill;
  },
  set fill(value) {
    return this.style.fill = value;
  },

  _mergeStyle: function(rowStyle, colStyle, style) {
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

  toCsvString: function() {
    return this._value.toCsvString();
  },

  // =========================================================================
  // Merge stuff

  addMergeRef: function() {
    this._mergeCount++;
  },
  releaseMergeRef: function() {
    this._mergeCount--;
  },
  get isMerged() {
    return (this._mergeCount > 0) || (this.type == Cell.Types.Merge);
  },
  merge: function(master) {
    this._value.release();
    this._value = Value.create(Cell.Types.Merge, this, master);
    this.style = master.style;
  },
  unmerge: function() {
    if (this.type == Cell.Types.Merge) {
      this._value.release();
      this._value = Value.create(Cell.Types.Null, this);
      this.style = this._mergeStyle(this._row.style, this._column.style, {});
    }
  },
  isMergedTo: function(master) {
    if (this._value.type != Cell.Types.Merge) return false;
    return this._value.isMergedTo(master);
  },
  get master() {
    if (this.type === Cell.Types.Merge) return this._value.master;
    else return this; // an unmerged cell is its own master
  },

  get isHyperlink() {
    return this._value.type == Cell.Types.Hyperlink;
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
      return this._value.master.value = v;
    }

    this._value.release();

    // assign value
    this._value = Value.create(Value.getType(v), this, v);
    return v;
  },
  get text() {
    return this._value.toString();
  } ,
  toString: function() {
    return this.text;
  },

  _upgradeToHyperlink: function(hyperlink) {
    // if this cell is a string, turn it into a Hyperlink
    if (this.type == Cell.Types.String) {
      this._value = Value.create(Cell.Types.Hyperlink, this, {
        text: this._value._value,
        hyperlink: hyperlink
      });
    }
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
    value.forEach(function(name) {
      definedNames.addEx(self.fullAddress, name);
    });
  },
  addName: function(name) {
    this.workbook.definedNames.addEx(this.fullAddress, name);
  },
  removeName: function(name) {
    this.workbook.definedNames.removeEx(this.fullAddress, name);
  },
  removeAllNames: function() {
    this.workbook.definedNames.removeAllNames(this.fullAddress);
  },

  // =========================================================================
  // Data Validation stuff
  get _dataValidations() {
    return this.worksheet.dataValidations;
  },

  get dataValidation() {
    return this._dataValidations.find(this.address)
  },

  set dataValidation(value) {
    return this._dataValidations.add(this.address, value);
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
    return value;
  }
};

// =============================================================================
// Internal Value Types

var NullValue = function(cell) {
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
    return value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return '';
  },
  release: function() {
  },
  toString: function() {
    return '';
  }
};

var NumberValue = function(cell, value) {
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
    return this.model.value = value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return '' + this.model.value;
  },
  release: function() {
  },
  toString: function() {
    return this.model.value.toString();
  }
};

var StringValue = function(cell, value) {
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
    return this.model.value = value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return '"' + this.model.value.replace(/"/g, '""') + '"';
  },
  release: function() {
  },
  toString: function() {
    return this.model.value;
  }
};

var RichTextValue = function(cell, value) {
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
    return this.model.value = value;
  },
  toString: function() {
    return this.model.value.richText.map(function(t) {
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return '"' + this.text.replace(/"/g, '""') + '"';
  },
  release: function() {
  }
};

var DateValue = function(cell, value) {
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
    return this.model.value = value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return this.model.value.toISOString();
  },
  release: function() {
  },
  toString: function() {
    return this.model.value.toString();
  }
};

var HyperlinkValue = function(cell, value) {
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
    return value;
  },

  get text() {
    return this.model.text;
  },
  set text(value) {
    return this.model.text = value;
  },
  get hyperlink() {
    return this.model.hyperlink;
  },
  set hyperlink(value) {
    return this.model.hyperlink = value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return this.model.hyperlink;
  },
  release: function() {
  },
  toString: function() {
    return this.model.text;
  }
};

var MergeValue = function(cell, master) {
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
      return this._master = value;
    } else {
      return this._master.value = value;
    }
  },

  isMergedTo: function(master) {
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return '';
  },
  release: function() {
    this._master.releaseMergeRef();
  },
  toString: function() {
    return this.value.toString();
  }
};

var FormulaValue = function(cell, value) {
  this.model = {
    address: cell.address,
    type: Cell.Types.Formula,
    formula: value ? value.formula : undefined,
    result: value ? value.result : undefined
  };
};
FormulaValue.prototype = {
  get value() {
    return {
      formula: this.model.formula,
      result: this.model.result
    };
  },
  set value(value) {
    this.model.formula = value.formula;
    this.model.result = value.result;
    return value;
  },
  validate: function(value) {
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
    var ranges = this.formula.match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g);
    var cells = this.formula.replace(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}:[A-Z]{1,3}\d{1,4}/g, '')
      .match(/([a-zA-Z0-9]+!)?[A-Z]{1,3}\d{1,4}/g);
    return {
      ranges: ranges,
      cells: cells
    };
  },
  get formula() {
    return this.model.formula;
  },
  set formula(value) {
    return this.model.formula = value;
  },
  get result() {
    return this.model.result;
  },
  set result(value) {
    return this.model.result = value;
  },
  get type() {
    return Cell.Types.Formula;
  },
  get effectiveType() {
    var v = this.model.result;
    if ((v === null) || (v === undefined)) {
      return Enums.ValueType.Null;
    } else if ((v instanceof String) || (typeof v == 'string')) {
      return Enums.ValueType.String;
    } else if (typeof v == 'number') {
      return Enums.ValueType.Number;
    } else if (v instanceof Date) {
      return Enums.ValueType.Date;
    } else if (v.text && v.hyperlink) {
      return Enums.ValueType.Hyperlink;
    } else if (v.formula) {
      return Enums.ValueType.Formula;
    } else {
      return Enums.ValueType.Null;
    }
  },
  get address() {
    return this.model.address;
  },
  set address(value) {
    return this.model.address = value;
  },
  toCsvString: function() {
    return '' + (this.model.result || '');
  },
  release: function() {
  },
  toString: function() {
    return this.model.result ? this.model.result.toString() : '';
  }
};

var SharedStringValue = function(cell, value) {
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
    return this.model.value = value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return '' + this.model.value;
  },
  release: function() {
  },
  toString: function() {
    return this.model.value.toString();
  }
};

var JSONValue = function(cell, value) {
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
    return value;
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
    return this.model.address = value;
  },
  toCsvString: function() {
    return this.model.value;
  },
  release: function() {
  },
  toString: function() {
    return this.model.value;
  }
};

// Value is a place to hold common static Value type functions
var Value = {
  getType:  function(value) {
    if ((value === null) || (value === undefined)) {
      return Cell.Types.Null;
    } else if ((value instanceof String) || (typeof value === 'string')) {
      return Cell.Types.String;
    } else if (typeof value === 'number') {
      return Cell.Types.Number;
    } else if (value instanceof Date) {
      return Cell.Types.Date;
    } else if (value.text && value.hyperlink) {
      return Cell.Types.Hyperlink;
    } else if (value.formula) {
      return Cell.Types.Formula;
    } else if (value.richText) {
      return Cell.Types.RichText;
    } else if (value.sharedString) {
      return Cell.Types.SharedString;
    } else {
      return Cell.Types.JSON;
      //console.log('Error: value=' + value + ', type=' + typeof value)
      // throw new Error('I could not understand type of value: ' +  JSON.stringify(value) + ' - typeof: ' + typeof value);
    }
  },

  // map valueType to constructor
  types: [
    {t:Cell.Types.Null, f:NullValue},
    {t:Cell.Types.Number, f:NumberValue},
    {t:Cell.Types.String, f:StringValue},
    {t:Cell.Types.Date, f:DateValue},
    {t:Cell.Types.Hyperlink, f:HyperlinkValue},
    {t:Cell.Types.Formula, f:FormulaValue},
    {t:Cell.Types.Merge, f:MergeValue},
    {t:Cell.Types.JSON, f:JSONValue},
    {t:Cell.Types.SharedString, f:SharedStringValue},
    {t:Cell.Types.RichText, f:RichTextValue}
  ].reduce(function(p,t){p[t.t]=t.f; return p;}, []),

  create: function(type, cell, value) {
    var t = this.types[type];
    if (!t) {
      throw new Error('Could not create Value of type ' + type);
    }
    return new t(cell, value);
  }
};
