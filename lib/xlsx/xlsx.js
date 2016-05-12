/**
 * Copyright (c) 2014 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software'), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var fs = require('fs');
var Bluebird = require('bluebird');
var _ = require('underscore');
var Archiver = require('archiver');
var unzip = require('unzip2');
var Sax = require('sax');

var utils = require('../utils/utils');
var colCache = require('../utils/col-cache');
var Enums = require('../doc/enums');
var Range = require('../doc/range');

var StyleManager = require('./stylemanager');

var DataValidationsXform = require('./xform/data-validations-xform');
var DefinedNamesXform = require('./xform/defined-names-xform');
var SharedStringsXform = require('./xform/shared-strings-xform');

var XLSX = module.exports = function (workbook) {
  this.workbook = workbook;
};

XLSX.RelType = require('./rel-type');


XLSX.prototype = {

// ===============================================================================
// Workbook
  // =========================================================================
  // Read

  readFile: function (filename) {
    var self = this;
    var stream;
    return utils.fs.exists(filename)
      .then(function (exists) {
        if (!exists) {
          throw new Error('File not found: ' + filename);
        }
        stream = fs.createReadStream(filename);
        return self.read(stream);
      })
      .then(function (workbook) {
        stream.close();
        return workbook;
      });
  },
  parseRels: function (stream) {
    return new Bluebird(function(resolve, reject) {
      
      var relationships = {};

      var parser = Sax.createStream(true, {});
      parser.on('opentag', function (node) {
        if (node.name === 'Relationship') {
          var rId = node.attributes.Id;
          switch (node.attributes.Type) {
            case XLSX.RelType.OfficeDocument:
              relationships[rId] = {
                type: Enums.RelationshipType.OfficeDocument,
                rId: rId,
                target: node.attributes.Target
              };
              break;
            case XLSX.RelType.Worksheet:
              relationships[rId] = {
                type: Enums.RelationshipType.Worksheet,
                rId: rId,
                target: node.attributes.Target,
                index: parseInt(node.attributes.Target.replace('worksheets/sheet', '').replace('.xml', ''))
              };
              break;
            case XLSX.RelType.CalcChain:
              relationships[rId] = {
                type: Enums.RelationshipType.CalcChain,
                rId: rId,
                target: node.attributes.Target
              };
              break;
            case XLSX.RelType.SharedStrings:
              relationships[rId] = {
                type: Enums.RelationshipType.SharedStrings,
                rId: rId,
                target: node.attributes.Target
              };
              break;
            case XLSX.RelType.Styles:
              relationships[rId] = {
                type: Enums.RelationshipType.Styles,
                rId: rId,
                target: node.attributes.Target
              };
              break;
            case XLSX.RelType.Theme:
              relationships[rId] = {
                type: Enums.RelationshipType.Theme,
                rId: rId,
                target: node.attributes.Target,
                index: parseInt(node.attributes.Target.replace('theme/theme', '').replace('.xml', ''))
              };
              break;
            case XLSX.RelType.Hyperlink:
              relationships[rId] = {
                type: Enums.RelationshipType.Styles,
                rId: rId,
                target: node.attributes.Target,
                targetMode: node.attributes.TargetMode
              };
              break;
            default:
              // ignore this?
              break;
          }
        }
      });
      
      parser.on('end', function () {
        resolve(relationships);
      });
      parser.on('error', function (error) {
        reject(error);
      });

      stream.pipe(parser); 
    });
  },
  parseWorkbook: function (stream) {
    return new Bluebird(function(resolve, reject) {
      var parser = Sax.createStream(true, {});
      var sheetDefs = {};
      var definedNames = new DefinedNamesXform();

      parser.on('opentag', function (node) {
        switch (node.name) {
          case 'sheet':
            // add sheet to sheetDefs indexed by rId
            // <sheet name="<%=worksheet.name%>" sheetId="<%=worksheet.id%>" r:id="<%=worksheet.rId%>"/>
            var name = node.attributes.name;
            var id = parseInt(node.attributes.sheetId);
            var rId = node.attributes['r:id'];
            sheetDefs[rId] = {
              name: name,
              id: id,
              rId: rId
            };
            break;

          case 'definedNames':
            definedNames.isActive = true;
            break;

          case 'definedName':
            definedNames.parseOpen(node);
            break;
        }
      });
      parser.on('text', function (text) {
        if (definedNames.isActive) {
          definedNames.parseText(text);
        }
      });
      parser.on('closetag', function (name) {
        switch (name) {
          case 'definedNames':
            definedNames.isActive = false;
            break;

          case 'definedName':
            definedNames.parseClose(name);
            break;
        }
      });
      parser.on('end', function () {
        resolve({sheetDefs: sheetDefs, definedNames: definedNames.model});
      });
      parser.on('error', function (error) {
        reject(error);
      });
      stream.pipe(parser);
    });
  },
  parseSharedStrings: function (stream) {
    var sharedStringsXform = new SharedStringsXform();
    return sharedStringsXform.parseStream(stream);
  },
  parseWorksheet: function (stream) {
    // <cols><col min="1" max="1" width="25" style="?" />
    // <sheetData><row><c r="A1" s="1" t="s"><f>A2</f><v>value</v></c>
    // <mergeCells><mergeCell ref="A1:B2" />
    // <hyperlinks><hyperlink ref="A1" r:id="rId1" />

    return new Bluebird(function(resolve, reject) {

      var parser = Sax.createStream(true, {});

      var dataValidationsXform = new DataValidationsXform();

      var model = {
        cols: [],
        rows: [],
        hyperlinks: [],
        merges: []
      };

      function parseBool(attribute) {
        return !!attribute;
      }

      var current;
      var c; // cell stuff
      var inDataValidations = false;
      parser.on('opentag', function (node) {
        switch (node.name) {
          case 'col':
            model.cols.push({
              min: parseInt(node.attributes.min),
              max: parseInt(node.attributes.max),
              width: parseFloat(node.attributes.width),
              styleId: parseInt(node.attributes.style || '0'),
              hidden: parseBool(node.attributes.hidden)
            });
            break;
          case 'row':
            var r = parseInt(node.attributes.r);
            var row = {
              number: r,
              cells: [],
              min: 0,
              max: 0,
              hidden: parseBool(node.attributes.hidden)
            };
            if (node.attributes.ht) {
              row.height = parseFloat(node.attributes.ht);
            }
            if (node.attributes.s) {
              row.styleId = parseInt(node.attributes.s);
            }
            model.rows[r] = row;
            break;
          case 'c':
            c = {
              ref: node.attributes.r,
              s: parseInt(node.attributes.s),
              t: node.attributes.t
            };
            break;
          case 'f':
            current = c.f = {text: ''};
            break;
          case 'v':
            current = c.v = {text: ''};
            break;
          case 'mergeCell':
            model.merges.push(node.attributes.ref);
            break;
          case 'hyperlink':
            model.hyperlinks.push({
              ref: node.attributes.ref,
              rId: node.attributes['r:id']
            });
            break;
          case 'dataValidations':
            inDataValidations = true;
            break;
          case 'dataValidation':
          case 'formula1':
          case 'formula2':
            dataValidationsXform.parseOpen(node);
            break;
        }
      });
      parser.on('text', function (text) {
        if (inDataValidations) {
          dataValidationsXform.parseText(text);
        } else if (current) {
          current.text += text;
        } 
      });

      parser.on('closetag', function (name) {
        switch (name) {
          case 'c':
            var address = colCache.decodeAddress(c.ref);
            var cell = {
              address: c.ref,
              s: c.s
            };
            // Note: merged cells (if they've been merged already will not have v or f nodes)
            if (c.f) {
              cell.type = Enums.ValueType.Formula;
              cell.formula = c.f.text;
              if (c.v) {
                if (c.t === 'str') {
                  cell.result = utils.xmlDecode(c.v.text);
                } else {
                  cell.result = parseFloat(c.v.text);
                }
              }
            } else if (c.v) {
              switch (c.t) {
                case 's':
                  cell.type = Enums.ValueType.String;
                  cell.value = parseInt(c.v.text);
                  cell.shared = true;
                  break;
                case 'str':
                  cell.type = Enums.ValueType.String;
                  cell.value = utils.xmlDecode(c.v.text);
                  break;
                default:
                  cell.type = Enums.ValueType.Number;
                  cell.value = parseFloat(c.v.text);
                  break;
              }
            } else {
              cell.type = Enums.ValueType.Merge;
            }

            var row = model.rows[address.row];
            row.cells[address.col] = cell;
            row.min = Math.min(row.min, address.col) || address.col;
            row.max = Math.max(row.max, address.col);
            break;

          case 'dataValidations':
            inDataValidations = false;
            break;
          case 'dataValidation':
          case 'formula1':
          case 'formula2':
            dataValidationsXform.parseClose(name);
            break;
        }
      });

      parser.on('end', function () {
        // if cols empty then can remove it from model
        if (!model.cols.length) {
          delete model.cols;
        }
        model.dataValidations = dataValidationsXform.model;

        resolve(model);
      });
      parser.on('error', function (error) {
        reject(error);
      });
      stream.pipe(parser);
    });
  },
  prepareModel: function (model) {
    // reconcile sheet ids, rIds and names
    _.each(model.sheetDefs, function (sheetDef) {
      var rel = model.workbookRels[sheetDef.rId];
      if (!rel) {
        return;
      }
      var worksheet = model.worksheetHash['xl/' + rel.target];
      worksheet.name = sheetDef.name;
      worksheet.id = sheetDef.id;
    });

    // reconcile worksheets
    _.each(model.worksheets, function (worksheet) {
      // reconcile column styles
      _.each(worksheet.cols, function (col) {
        if (col.styleId) {
          col.style = model.styles.getStyleModel(col.styleId);
        }
      });

      // reconcile merge list with merge cells
      _.each(worksheet.merges, function (merge) {
        var dimensions = colCache.decode(merge);
        for (var i = dimensions.top; i <= dimensions.bottom; i++) {
          var row = worksheet.rows[i];
          for (var j = dimensions.left; j <= dimensions.right; j++) {
            var cell = row.cells[j];
            if (!cell) {
              // nulls are not included in document - so if master cell has no value - add a null one here
              row.cells[j] = {
                type: Enums.ValueType.Null,
                address: colCache.encodeAddress(i, j)
              };
            } else if (cell.type === Enums.ValueType.Merge) {
              cell.master = dimensions.tl;
            }
          }
        }
      });

      // if worksheet rels, merge them in
      var relationships = model.worksheetRels[worksheet.sheetNo];
      if (relationships) {
        var hyperlinks = {};
        _.each(relationships, function (relationship) {
          hyperlinks[relationship.rId] = relationship.target;
        });

        _.each(worksheet.hyperlinks, function (hyperlink) {
          var address = colCache.decodeAddress(hyperlink.ref);
          var row = worksheet.rows[address.row];
          var cell = row.cells[address.col];
          cell.type = Enums.ValueType.Hyperlink;
          cell.text = cell.value;
          cell.hyperlink = hyperlinks[hyperlink.rId];
        });
        // no need for them any more
        delete worksheet.hyperlinks;
      }

      // compact the rows and calculate dimensions
      var dimensions = new Range();
      worksheet.rows = worksheet.rows.filter(function (row) {
        return row;
      });
      _.each(worksheet.rows, function (row) {
        dimensions.expand(row.number, row.min, row.number, row.max);
        row.cells = row.cells.filter(function (cell) {
          return cell;
        });
      });
      worksheet.dimensions = dimensions.model;
      delete worksheet.sheetNo;

      // reconcile cell values, styles
      _.each(worksheet.rows, function (row) {
        row.style = row.styleId ? model.styles.getStyleModel(row.styleId) : {};

        _.each(row.cells, function (cell) {
          // get style model from style manager
          var style = cell.s ? model.styles.getStyleModel(cell.s) : {};
          //console.log(cell.address + ' ' + cell.s + ' ' + JSON.stringify(style));
          delete cell.s;
          if (style) {
            cell.style = style;
          }
          switch (cell.type) {
            case Enums.ValueType.String:
              if (cell.shared) {
                cell.value = model.sharedStrings.values[cell.value];
              }
              if (cell.value.richText) {
                cell.type = Enums.ValueType.RichText;
              }
              break;
            case Enums.ValueType.Hyperlink:
              cell.text = cell.shared ? model.sharedStrings.values[cell.value] : cell.value;
              delete cell.value;
              break;
            case Enums.ValueType.Number:
              if (style && utils.isDateFmt(style.numFmt)) {
                cell.type = Enums.ValueType.Date;
                cell.value = utils.excelToDate(cell.value);
              }
              break;
            case Enums.ValueType.Formula:
              if ((cell.result !== undefined) && style && utils.isDateFmt(style.numFmt)) {
                cell.result = utils.excelToDate(cell.result);
              }
              break;
          }
        });
      });
    });

    // delete unnecessary parts
    delete model.worksheetHash;
    delete model.worksheetRels;
    delete model.globalRels;
    delete model.sharedStrings;
    delete model.workbookRels;
    delete model.sheetDefs;
    delete model.styles;
  },
  createInputStream: function () {
    var self = this;
    var model = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: []
    };

    // we have to be prepared to read the zip entries in whatever order they arrive
    var promises = [];
    var stream = unzip.Parse();
    stream.on('entry', function (entry) {
      var promise = null;
      var match, sheetNo;
      switch (entry.path) {
        case '_rels/.rels':
          promise = self.parseRels(entry)
            .then(function (relationships) {
              model.globalRels = relationships;
            });
          break;
        case 'xl/workbook.xml':
          promise = self.parseWorkbook(entry)
            .then(function (workbook) {
              model.sheetDefs = workbook.sheetDefs;
              model.definedNames = workbook.definedNames;
            });
          break;
        case 'xl/_rels/workbook.xml.rels':
          promise = self.parseRels(entry)
            .then(function (relationships) {
              model.workbookRels = relationships;
            });
          break;
        case 'xl/sharedStrings.xml':
          promise = self.parseSharedStrings(entry)
            .then(function (sharedStrings) {
              model.sharedStrings = sharedStrings;
            });
          break;
        case 'xl/styles.xml':
          model.styles = new StyleManager();
          promise = model.styles.parse(entry);
          break;
        default:
          if (entry.path.match(/xl\/worksheets\/sheet\d+\.xml/)) {
            match = entry.path.match(/xl\/worksheets\/sheet(\d+)\.xml/);
            sheetNo = match[1];
            promise = self.parseWorksheet(entry)
              .then(function (worksheet) {
                worksheet.sheetNo = sheetNo;
                model.worksheetHash[entry.path] = worksheet;
                model.worksheets.push(worksheet);
              });
          } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+\.xml.rels/)) {
            match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)\.xml.rels/);
            sheetNo = match[1];
            promise = self.parseRels(entry)
              .then(function (relationships) {
                model.worksheetRels[sheetNo] = relationships;
              });
          } else {
            entry.autodrain();
          }
          break;
      }

      if (promise) {
        promises.push(promise);
        promise = null;
      }
    });
    stream.on('close', function () {
      Bluebird.all(promises)
        .then(function () {
          self.prepareModel(model);

          // apply model
          self.workbook.model = model;
        })
        .then(function () {
          stream.emit('done');
        })
        .catch(function (error) {
          stream.emit('error', error);
        });
    });
    return stream;
  },

  read: function (stream) {
    var self = this;
    var zipStream = this.createInputStream();
    return new Bluebird(function(resolve, reject) {
      zipStream.on('done', function () {
        resolve(self.workbook);
      }).on('error', function (error) {
        reject(error);
      });
      stream.pipe(zipStream);
    });
  },

  // =========================================================================
  // Write

  addContentTypes: function (zip, model) {
    return utils.fetchTemplate(require.resolve('./content-types.xml'))
      .then(function (template) {
        return template(model);
      })
      .then(function (data) {
        zip.append(data, {name: '[Content_Types].xml'});
        return zip;
      });
  },

  addApp: function (zip, model) {
    return utils.fetchTemplate(require.resolve('./app.xml'))
      .then(function (template) {
        return template(model);
      })
      .then(function (data) {
        zip.append(data, {name: 'docProps/app.xml'});
        return zip;
      });
  },

  addCore: function (zip, model) {
    return utils.fetchTemplate(require.resolve('./core.xml'))
      .then(function (template) {
        return template(model);
      })
      .then(function (data) {
        zip.append(data, {name: 'docProps/core.xml'});
        return zip;
      });
  },

  addThemes: function (zip) {
    return utils.readModuleFile(require.resolve('./theme1.xml'))
      .then(function (data) {
        zip.append(data, {name: 'xl/theme/theme1.xml'});
        return zip;
      });
  },

  //<Relationship Id="rId1"
  //        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  //        Target="xl/workbook.xml"/>
  addOfficeRels: function (zip, model) {
    model._rels = {
      relationships: [
        {rId: 'rId1', type: XLSX.RelType.OfficeDocument, target: 'xl/workbook.xml'}
      ]
    };
    return utils.fetchTemplate(require.resolve('./.rels'))
      .then(function (template) {
        return template(model._rels);
      })
      .then(function (data) {
        zip.append(data, {name: '/_rels/.rels'});
      });
  },

  addWorkbookRels: function (zip, model) {
    var count = 1;
    model.workbookRels = {
      relationships: [
        {rId: 'rId' + (count++), type: XLSX.RelType.Styles, target: 'styles.xml'},
        {rId: 'rId' + (count++), type: XLSX.RelType.Theme, target: 'theme/theme1.xml'}
      ]
    };
    if (model.sharedStrings.count) {
      model.workbookRels.relationships.push(
        {rId: 'rId' + (count++), type: XLSX.RelType.SharedStrings, target: 'sharedStrings.xml'}
      );
    }
    _.each(model.worksheets, function (worksheet) {
      worksheet.rId = 'rId' + (count++);
      model.workbookRels.relationships.push(
        {rId: worksheet.rId, type: XLSX.RelType.Worksheet, target: 'worksheets/sheet' + worksheet.id + '.xml'}
      );
    });
    return utils.fetchTemplate(require.resolve('./.rels'))
      .then(function (template) {
        return template(model.workbookRels);
      })
      .then(function (data) {
        zip.append(data, {name: '/xl/_rels/workbook.xml.rels'});
      });
  },
  addSharedStrings: function (zip, model) {
    if (!model.sharedStrings || !model.sharedStrings.count) {
      return Bluebird.resolve();
    } else {
      return new Bluebird(function(resolve, reject) {
        zip.append(model.sharedStrings.xml, {name: '/xl/sharedStrings.xml'});
        resolve();
      });
    }
  },
  addWorkbook: function (zip, model) {
    return utils.fetchTemplate(require.resolve('./workbook.xml'))
      .then(function (template) {
        return template(model);
      })
      .then(function (data) {
        zip.append(data, {name: '/xl/workbook.xml'});
        return zip;
      });
  },
  getValueType: function (v) {
    if ((v === null) || (v === undefined)) {
      return Enums.ValueType.Null;
    } else if ((v instanceof String) || (typeof v === 'string')) {
      return Enums.ValueType.String;
    } else if (typeof v === 'number') {
      return Enums.ValueType.Number;
    } else if (v instanceof Date) {
      return Enums.ValueType.Date;
    } else if (v.text && v.hyperlink) {
      return Enums.ValueType.Hyperlink;
    } else if (v.formula) {
      return Enums.ValueType.Formula;
    } else {
      throw new Error('I could not understand type of value');
    }
  },
  getEffectiveCellType: function (cell) {
    switch (cell.type) {
      case Enums.ValueType.Formula:
        return this.getValueType(cell.result);
      default:
        return cell.type;
    }
  },
  addWorksheets: function (zip, model) {
    var self = this;
    var promises = [];

    function buildC(cell, t, close) {
      var c = ['<c r="'];
      c.push(cell.address);
      c.push('"');
      if (t) {
        c.push(' t="' + t + '"');
      }

      var s = model.styles.addStyleModel(cell.style, self.getEffectiveCellType(cell));
      if (s) {
        c.push(' s="' + s + '"');
      }

      c.push(close ? '/>' : '>');
      return c.join('');
    }

    _.each(model.worksheets, function (worksheet) {
      worksheet.dataValidations = new DataValidationsXform(worksheet.dataValidations);

      // TODO: process worksheet model and prepare for template
      if (!worksheet.cols) {
        worksheet.cols = false;
      } else {
        _.each(worksheet.cols, function (column) {
          column.styleId = model.styles.addStyleModel(column.style);
        });
      }

      worksheet.hyperlinks = worksheet.relationships = [];
      var hyperlinkCount = 1;
      _.each(worksheet.rows, function (row) {

        row.styleId = model.styles.addStyleModel(row.style);

        _.each(row.cells, function (cell) {
          switch (cell.type) {
            case Enums.ValueType.Number:
              cell.xml = buildC(cell) + '<v>' + cell.value + '</v></c>';
              break;
            case Enums.ValueType.String:
              if (model.useSharedStrings) {
                cell.xml = buildC(cell, 's') +
                  '<v>' + model.sharedStrings.add(cell.value) + '</v>' +
                  '</c>';
              } else {
                cell.xml = buildC(cell, 'str') +
                  '<v>' + utils.xmlEncode(cell.value) + '</v>' +
                  '</c>';
              }
              break;
            case Enums.ValueType.Date:
              cell.xml = buildC(cell) +
                '<v>' + utils.dateToExcel(cell.value) + '</v>' +
                '</c>';
              break;
            case Enums.ValueType.Hyperlink:
              if (model.useSharedStrings) {
                cell.xml = buildC(cell, 's') +
                  '<v>' + model.sharedStrings.add(cell.text) + '</v>' +
                  '</c>';
              } else {
                cell.xml = buildC(cell, 'str') +
                  '<v>' + utils.xmlEncode(cell.value) + '</v>' +
                  '</c>';
              }
              worksheet.hyperlinks.push({
                address: cell.address,
                rId: 'rId' + (hyperlinkCount++),
                type: XLSX.RelType.Hyperlink,
                target: utils.xmlEncode(cell.hyperlink),
                targetMode: 'External'
              });
              break;
            case Enums.ValueType.Formula:
              switch (self.getValueType(cell.result)) {
                case Enums.ValueType.Null: // ?
                  cell.xml = buildC(cell) +
                    '<f>' + utils.xmlEncode(cell.formula) + '</f>' +
                    '</c>';
                  break;
                case Enums.ValueType.String:
                  // oddly, formula results don't ever use shared strings
                  cell.xml = buildC(cell, 'str') +
                    '<f>' + utils.xmlEncode(cell.formula) + '</f>' +
                    '<v>' + utils.xmlEncode(cell.result) + '</v>' +
                    '</c>';
                  break;
                case Enums.ValueType.Number:
                  cell.xml = buildC(cell) +
                    '<f>' + utils.xmlEncode(cell.formula) + '</f>' +
                    '<v>' + cell.result + '</v>' +
                    '</c>';
                  break;
                case Enums.ValueType.Date:
                  cell.xml = buildC(cell) +
                    '<f>' + utils.xmlEncode(cell.formula) + '</f>' +
                    '<v>' + utils.dateToExcel(cell.result) + '</v>' +
                    '</c>';
                  break;
                // case Enums.ValueType.Hyperlink: // ??
                // case Enums.ValueType.Formula:
                default:
                  throw new Error('I could not understand type of value');
              }
              break;
            case Enums.ValueType.Merge:
              cell.xml = buildC(cell, undefined, true);
              break;
          }
        });
      });

      if (!worksheet.merges.length) {
        worksheet.merges = false;
      }
      if (!worksheet.hyperlinks.length) {
        worksheet.hyperlinks = false;
      }

      promises.push(utils.fetchTemplate(require.resolve('./sheet.xml'))
        .then(function (template) {
          return template(worksheet);
        })
        .then(function (data) {
          zip.append(data, {name: '/xl/worksheets/sheet' + worksheet.id + '.xml'});
          return zip;
        }));

      if (worksheet.hyperlinks) {
        promises.push(utils.fetchTemplate(require.resolve('./.rels'))
          .then(function (template) {
            return template(worksheet);
          })
          .then(function (data) {
            zip.append(data, {name: '/xl/worksheets/_rels/sheet' + worksheet.id + '.xml.rels'});
          }));
      }
    });

    return Bluebird.all(promises);
  },
  _finalize: function (zip) {
    var self = this;

    return new Bluebird(function(resolve, reject) {

      zip.on('end', function () {
        resolve(self);
      });
      zip.on('error', function (error) {
        reject(error);
      });

      zip.finalize();

    });

  },
  write: function (stream, options) {
    options = options || {};
    var self = this;
    var model = self.workbook.model;
    var zip = Archiver('zip');
    zip.pipe(stream);

    // ensure following properties have sane values
    model.creator = model.creator || 'ExcelJS';
    model.lastModifiedBy = model.lastModifiedBy || 'ExcelJS';
    model.created = model.created || new Date();
    model.modified = model.modified || new Date();

    model.useSharedStrings = options.useSharedStrings !== undefined ?
      options.useSharedStrings :
      true;
    model.useStyles = options.useStyles !== undefined ?
      options.useStyles :
      true;

    // Manage the shared strings
    model.sharedStrings = new SharedStringsXform();
    model.definedNames = new DefinedNamesXform(model.definedNames);

    // add a style manager to handle cell formats, fonts, etc.
    model.styles = model.useStyles ? new StyleManager() : new StyleManager.Mock();

    var promises = [
      self.addContentTypes(zip, model),
      self.addApp(zip, model),
      self.addCore(zip, model),
      self.addThemes(zip),
      self.addOfficeRels(zip, model)
    ];
    return Bluebird.all(promises)
      .then(function () {
        return self.addWorksheets(zip, model);
      })
      .then(function () {
        // Some things can only be done after all the worksheets have been processed
        var afters = [
          self.addSharedStrings(zip, model),
          model.styles.addToZip(zip),
          self.addWorkbookRels(zip, model)
        ];
        return Bluebird.all(afters);
      })
      .then(function () {
        return self.addWorkbook(zip, model);
      })
      .then(function () {
        return self._finalize(zip);
      });
  },
  writeFile: function (filename, options) {
    var self = this;
    var stream = fs.createWriteStream(filename);

    return new Bluebird(function(resolve, reject) {
      stream.on('finish', function () {
        resolve();
      });
      stream.on('error', function (error) {
        reject(error);
      });

      self.write(stream, options)
        .then(function () {
          stream.end();
        })
        .catch(function (error) {
          reject(error);
        });
    });
  }
};
