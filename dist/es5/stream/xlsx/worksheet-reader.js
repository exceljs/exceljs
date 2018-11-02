/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var events = require('events');
var Sax = require('sax');

var _ = require('../../utils/under-dash');
var utils = require('../../utils/utils');
var colCache = require('../../utils/col-cache');
var Dimensions = require('../../doc/range');

var Row = require('../../doc/row');
var Column = require('../../doc/column');

var WorksheetReader = module.exports = function (workbook, id) {
  this.workbook = workbook;
  this.id = id;

  // and a name
  this.name = 'Sheet' + this.id;

  // column definitions
  this._columns = null;
  this._keys = {};

  // keep a record of dimensions
  this._dimensions = new Dimensions();
};

utils.inherits(WorksheetReader, events.EventEmitter, {

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy: function destroy() {
    throw new Error('Invalid Operation: destroy');
  },

  // return the current dimensions of the writer
  get dimensions() {
    return this._dimensions;
  },

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns() {
    return this._columns;
  },

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn: function getColumn(c) {
    if (typeof c === 'string') {
      // if it matches a key'd column, return that
      var col = this._keys[c];
      if (col) {
        return col;
      }

      // otherise, assume letter
      c = colCache.l2n(c);
    }
    if (!this._columns) {
      this._columns = [];
    }
    if (c > this._columns.length) {
      var n = this._columns.length + 1;
      while (n <= c) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[c - 1];
  },

  getColumnKey: function getColumnKey(key) {
    return this._keys[key];
  },
  setColumnKey: function setColumnKey(key, value) {
    this._keys[key] = value;
  },
  deleteColumnKey: function deleteColumnKey(key) {
    delete this._keys[key];
  },
  eachColumnKey: function eachColumnKey(f) {
    _.each(this._keys, f);
  },


  // =========================================================================
  // Read

  _emitRow: function _emitRow(row) {
    this.emit('row', row);
  },

  read: function read(entry, options) {
    var _this = this;

    var emitSheet = false;
    var emitHyperlinks = false;
    var hyperlinks = null;
    switch (options.worksheets) {
      case 'emit':
        emitSheet = true;
        break;
      case 'prep':
        break;
      default:
        break;
    }
    switch (options.hyperlinks) {
      case 'emit':
        emitHyperlinks = true;
        break;
      case 'cache':
        this.hyperlinks = hyperlinks = {};
        break;
      default:
        break;
    }
    if (!emitSheet && !emitHyperlinks && !hyperlinks) {
      entry.autodrain();
      this.emit('finished');
      return;
    }

    // references
    var sharedStrings = this.workbook.sharedStrings;
    var styles = this.workbook.styles;
    var properties = this.workbook.properties;

    // xml position
    var inCols = false;
    var inRows = false;
    var inHyperlinks = false;

    // parse state
    var cols = null;
    var row = null;
    var c = null;
    var current = null;

    var parser = Sax.createStream(true, {});
    parser.on('opentag', function (node) {
      if (emitSheet) {
        switch (node.name) {
          case 'cols':
            inCols = true;
            cols = [];
            break;
          case 'sheetData':
            inRows = true;
            break;

          case 'col':
            if (inCols) {
              cols.push({
                min: parseInt(node.attributes.min, 10),
                max: parseInt(node.attributes.max, 10),
                width: parseFloat(node.attributes.width),
                styleId: parseInt(node.attributes.style || '0', 10)
              });
            }
            break;

          case 'row':
            if (inRows) {
              var r = parseInt(node.attributes.r, 10);
              row = new Row(_this, r);
              if (node.attributes.ht) {
                row.height = parseFloat(node.attributes.ht);
              }
              if (node.attributes.s) {
                var styleId = parseInt(node.attributes.s, 10);
                var style = styles.getStyleModel(styleId);
                if (style) {
                  row.style = style;
                }
              }
            }
            break;
          case 'c':
            if (row) {
              c = {
                ref: node.attributes.r,
                s: parseInt(node.attributes.s, 10),
                t: node.attributes.t
              };
            }
            break;
          case 'f':
            if (c) {
              current = c.f = { text: '' };
            }
            break;
          case 'v':
            if (c) {
              current = c.v = { text: '' };
            }
            break;
          case 'mergeCell':
            break;
          default:
            break;
        }
      }

      // =================================================================
      //
      if (emitHyperlinks || hyperlinks) {
        switch (node.name) {
          case 'hyperlinks':
            inHyperlinks = true;
            break;
          case 'hyperlink':
            if (inHyperlinks) {
              var hyperlink = {
                ref: node.attributes.ref,
                rId: node.attributes['r:id']
              };
              if (emitHyperlinks) {
                _this.emit('hyperlink', hyperlink);
              } else {
                hyperlinks[hyperlink.ref] = hyperlink;
              }
            }
            break;
          default:
            break;
        }
      }
    });

    // only text data is for sheet values
    parser.on('text', function (text) {
      if (emitSheet) {
        if (current) {
          current.text += text;
        }
      }
    });

    parser.on('closetag', function (name) {
      if (emitSheet) {
        switch (name) {
          case 'cols':
            inCols = false;
            _this._columns = Column.fromModel(cols);
            break;
          case 'sheetData':
            inRows = false;
            break;

          case 'row':
            _this._dimensions.expandRow(row);
            _this._emitRow(row);
            row = null;
            break;

          case 'c':
            if (row && c) {
              var address = colCache.decodeAddress(c.ref);
              var cell = row.getCell(address.col);
              if (c.s) {
                var style = styles.getStyleModel(c.s);
                if (style) {
                  cell.style = style;
                }
              }

              if (c.f) {
                var value = {
                  formula: c.f.text
                };
                if (c.v) {
                  if (c.t === 'str') {
                    value.result = utils.xmlDecode(c.v.text);
                  } else {
                    value.result = parseFloat(c.v.text);
                  }
                }
                cell.value = value;
              } else if (c.v) {
                switch (c.t) {
                  case 's':
                    var index = parseInt(c.v.text, 10);
                    if (sharedStrings) {
                      cell.value = sharedStrings[index];
                    } else {
                      cell.value = {
                        sharedString: index
                      };
                    }
                    break;
                  case 'str':
                    cell.value = utils.xmlDecode(c.v.text);
                    break;
                  case 'e':
                    cell.value = { error: c.v.text };
                    break;
                  case 'b':
                    cell.value = parseInt(c.v.text, 10) !== 0;
                    break;
                  default:
                    if (utils.isDateFmt(cell.numFmt)) {
                      cell.value = utils.excelToDate(parseFloat(c.v.text), properties.model.date1904);
                    } else {
                      cell.value = parseFloat(c.v.text);
                    }
                    break;
                }
              }
              if (hyperlinks) {
                var hyperlink = hyperlinks[c.ref];
                if (hyperlink) {
                  cell.text = cell.value;
                  cell.value = undefined;
                  cell.hyperlink = hyperlink;
                }
              }
              c = null;
            }
            break;
          default:
            break;
        }
      }
      if (emitHyperlinks || hyperlinks) {
        switch (name) {
          case 'hyperlinks':
            inHyperlinks = false;
            break;
          default:
            break;
        }
      }
    });
    parser.on('error', function (error) {
      _this.emit('error', error);
    });
    parser.on('end', function () {
      _this.emit('finished');
    });

    // create a down-stream flow-control to regulate the stream
    var flowControl = this.workbook.flowControl.createChild();
    flowControl.pipe(parser, { sync: true });
    entry.pipe(flowControl);
  }
});
//# sourceMappingURL=worksheet-reader.js.map
