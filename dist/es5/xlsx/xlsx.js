/**
 * Copyright (c) 2014 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var fs = require('fs');
var ZipStream = require('../utils/zip-stream');
var StreamBuf = require('../utils/stream-buf');
var PromishLib = require('../utils/promish');

var utils = require('../utils/utils');
var XmlStream = require('../utils/xml-stream');

var StylesXform = require('./xform/style/styles-xform');

var CoreXform = require('./xform/core/core-xform');
var SharedStringsXform = require('./xform/strings/shared-strings-xform');
var RelationshipsXform = require('./xform/core/relationships-xform');
var ContentTypesXform = require('./xform/core/content-types-xform');
var AppXform = require('./xform/core/app-xform');
var WorkbookXform = require('./xform/book/workbook-xform');
var WorksheetXform = require('./xform/sheet/worksheet-xform');
var DrawingXform = require('./xform/drawing/drawing-xform');

var theme1Xml = require('./xml/theme1.js');

var XLSX = module.exports = function (workbook) {
  this.workbook = workbook;
};

function fsReadFileAsync(filename, options) {
  return new PromishLib.Promish(function (resolve, reject) {
    fs.readFile(filename, options, function (error, data) {
      if (error) {
        reject(error);
      } else {
        resolve(data);
      }
    });
  });
}

XLSX.RelType = require('./rel-type');

XLSX.prototype = {
  // ===============================================================================
  // Workbook
  // =========================================================================
  // Read

  readFile: function readFile(filename, options) {
    var self = this;
    var stream;
    return utils.fs.exists(filename).then(function (exists) {
      if (!exists) {
        throw new Error('File not found: ' + filename);
      }
      stream = fs.createReadStream(filename);
      return self.read(stream, options).catch(function (error) {
        stream.close();
        throw error;
      });
    }).then(function (workbook) {
      stream.close();
      return workbook;
    });
  },
  parseRels: function parseRels(stream) {
    var xform = new RelationshipsXform();
    return xform.parseStream(stream);
  },
  parseWorkbook: function parseWorkbook(stream) {
    var xform = new WorkbookXform();
    return xform.parseStream(stream);
  },
  parseSharedStrings: function parseSharedStrings(stream) {
    var xform = new SharedStringsXform();
    return xform.parseStream(stream);
  },
  reconcile: function reconcile(model, options) {
    var workbookXform = new WorkbookXform();
    var worksheetXform = new WorksheetXform(options);
    var drawingXform = new DrawingXform();

    workbookXform.reconcile(model);

    // reconcile drawings with their rels
    var drawingOptions = {
      media: model.media,
      mediaIndex: model.mediaIndex
    };
    Object.keys(model.drawings).forEach(function (name) {
      var drawing = model.drawings[name];
      var drawingRel = model.drawingRels[name];
      if (drawingRel) {
        drawingOptions.rels = drawingRel.reduce(function (o, rel) {
          o[rel.Id] = rel;return o;
        }, {});
        drawingXform.reconcile(drawing, drawingOptions);
      }
    });

    var sheetOptions = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      mediaIndex: model.mediaIndex,
      date1904: model.properties && model.properties.date1904,
      drawings: model.drawings
    };
    model.worksheets.forEach(function (worksheet) {
      worksheet.relationships = model.worksheetRels[worksheet.sheetNo];
      worksheetXform.reconcile(worksheet, sheetOptions);
    });

    // delete unnecessary parts
    delete model.worksheetHash;
    delete model.worksheetRels;
    delete model.globalRels;
    delete model.sharedStrings;
    delete model.workbookRels;
    delete model.sheetDefs;
    delete model.styles;
    delete model.mediaIndex;
    delete model.drawings;
    delete model.drawingRels;
  },
  processWorksheetEntry: function processWorksheetEntry(entry, model, options) {
    var match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
    if (match) {
      var sheetNo = match[1];
      var xform = new WorksheetXform(options);
      return xform.parseStream(entry).then(function (worksheet) {
        worksheet.sheetNo = sheetNo;
        model.worksheetHash[entry.path] = worksheet;
        model.worksheets.push(worksheet);
      });
    }
    return undefined;
  },
  processWorksheetRelsEntry: function processWorksheetRelsEntry(entry, model) {
    var match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
    if (match) {
      var sheetNo = match[1];
      var xform = new RelationshipsXform();
      return xform.parseStream(entry).then(function (relationships) {
        model.worksheetRels[sheetNo] = relationships;
      });
    }
    return undefined;
  },
  processMediaEntry: function processMediaEntry(entry, model) {
    var match = entry.path.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
    if (match) {
      var filename = match[1];
      var lastDot = filename.lastIndexOf('.');
      if (lastDot === -1) {
        // if we can't determine extension, ignore it
        return undefined;
      }
      var extension = filename.substr(lastDot + 1);
      var name = filename.substr(0, lastDot);
      return new PromishLib.Promish(function (resolve, reject) {
        var streamBuf = new StreamBuf();
        streamBuf.on('finish', function () {
          model.mediaIndex[filename] = model.media.length;
          model.mediaIndex[name] = model.media.length;
          var medium = {
            type: 'image',
            name: name,
            extension: extension,
            buffer: streamBuf.toBuffer()
          };
          model.media.push(medium);
          resolve();
        });
        entry.on('error', function (error) {
          reject(error);
        });
        entry.pipe(streamBuf);
      });
    }
    return undefined;
  },
  processDrawingEntry: function processDrawingEntry(entry, model) {
    var match = entry.path.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
    if (match) {
      var name = match[1];
      var xform = new DrawingXform();
      return xform.parseStream(entry).then(function (drawing) {
        model.drawings[name] = drawing;
      });
    }
    return undefined;
  },
  processDrawingRelsEntry: function processDrawingRelsEntry(entry, model) {
    var match = entry.path.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
    if (match) {
      var name = match[1];
      var xform = new RelationshipsXform();
      return xform.parseStream(entry).then(function (relationships) {
        model.drawingRels[name] = relationships;
      });
    }
    return undefined;
  },
  processThemeEntry: function processThemeEntry(entry, model) {
    var match = entry.path.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
    if (match) {
      return new PromishLib.Promish(function (resolve, reject) {
        var name = match[1];
        // TODO: stream entry into buffer and store the xml in the model.themes[]
        var stream = new StreamBuf();
        entry.on('error', reject);
        stream.on('error', reject);
        stream.on('finish', function () {
          model.themes[name] = stream.read().toString();
          resolve();
        });
        entry.pipe(stream);
      });
    }
    return undefined;
  },
  processIgnoreEntry: function processIgnoreEntry(entry) {
    entry.autodrain();
  },
  createInputStream: function createInputStream(options) {
    var self = this;
    var model = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: [],
      themes: {},
      media: [],
      mediaIndex: {},
      drawings: {},
      drawingRels: {}
    };

    // we have to be prepared to read the zip entries in whatever order they arrive
    var promises = [];
    var stream = new ZipStream.ZipReader({
      getEntryType: function getEntryType(path) {
        return path.match(/xl\/media\//) ? 'nodebuffer' : 'string';
      }
    });
    stream.on('entry', function (entry) {
      var promise = null;

      var entryPath = entry.path;
      if (entryPath[0] === '/') {
        entryPath = entryPath.substr(1);
      }
      switch (entryPath) {
        case '_rels/.rels':
          promise = self.parseRels(entry).then(function (relationships) {
            model.globalRels = relationships;
          });
          break;

        case 'xl/workbook.xml':
          promise = self.parseWorkbook(entry).then(function (workbook) {
            model.sheets = workbook.sheets;
            model.definedNames = workbook.definedNames;
            model.views = workbook.views;
            model.properties = workbook.properties;
          });
          break;

        case 'xl/_rels/workbook.xml.rels':
          promise = self.parseRels(entry).then(function (relationships) {
            model.workbookRels = relationships;
          });
          break;

        case 'xl/sharedStrings.xml':
          model.sharedStrings = new SharedStringsXform();
          promise = model.sharedStrings.parseStream(entry);
          break;

        case 'xl/styles.xml':
          model.styles = new StylesXform();
          promise = model.styles.parseStream(entry);
          break;

        case 'docProps/app.xml':
          var appXform = new AppXform();
          promise = appXform.parseStream(entry).then(function (appProperties) {
            Object.assign(model, {
              company: appProperties.company,
              manager: appProperties.manager
            });
          });
          break;

        case 'docProps/core.xml':
          var coreXform = new CoreXform();
          promise = coreXform.parseStream(entry).then(function (coreProperties) {
            Object.assign(model, coreProperties);
          });
          break;

        default:
          promise = self.processWorksheetEntry(entry, model, options) || self.processWorksheetRelsEntry(entry, model) || self.processThemeEntry(entry, model) || self.processMediaEntry(entry, model) || self.processDrawingEntry(entry, model) || self.processDrawingRelsEntry(entry, model) || self.processIgnoreEntry(entry);
          break;
      }

      if (promise) {
        promise = promise.catch(function (error) {
          stream.destroy(error);
          throw error;
        });

        promises.push(promise);
        promise = null;
      }
    });
    stream.on('finished', function () {
      PromishLib.Promish.all(promises).then(function () {
        self.reconcile(model, options);

        // apply model
        self.workbook.model = model;
      }).then(function () {
        stream.emit('done');
      }).catch(function (error) {
        stream.emit('error', error);
      });
    });
    return stream;
  },

  read: function read(stream, options) {
    options = options || {};
    var self = this;
    var zipStream = this.createInputStream(options);
    return new PromishLib.Promish(function (resolve, reject) {
      zipStream.on('done', function () {
        resolve(self.workbook);
      }).on('error', function (error) {
        reject(error);
      });
      stream.pipe(zipStream);
    });
  },

  load: function load(data, options) {
    var self = this;
    if (options === undefined) {
      options = {};
    }
    var zipStream = this.createInputStream();
    return new PromishLib.Promish(function (resolve, reject) {
      zipStream.on('done', function () {
        resolve(self.workbook);
      }).on('error', function (error) {
        reject(error);
      });

      if (options.base64) {
        var buffer = new Buffer(data.toString(), 'base64');
        zipStream.write(buffer);
      } else {
        zipStream.write(data);
      }
      zipStream.end();
    });
  },

  // =========================================================================
  // Write

  addMedia: function addMedia(zip, model) {
    return PromishLib.Promish.all(model.media.map(function (medium) {
      if (medium.type === 'image') {
        var filename = 'xl/media/' + medium.name + '.' + medium.extension;
        if (medium.filename) {
          return fsReadFileAsync(medium.filename).then(function (data) {
            zip.append(data, { name: filename });
          });
        }
        if (medium.buffer) {
          return new PromishLib.Promish(function (resolve) {
            zip.append(medium.buffer, { name: filename });
            resolve();
          });
        }
        if (medium.base64) {
          return new PromishLib.Promish(function (resolve) {
            var dataimg64 = medium.base64;
            var content = dataimg64.substring(dataimg64.indexOf(',') + 1);
            zip.append(content, { name: filename, base64: true });
            resolve();
          });
        }
      }
      return PromishLib.Promish.reject(new Error('Unsupported media'));
    }));
  },

  addDrawings: function addDrawings(zip, model) {
    var drawingXform = new DrawingXform();
    var relsXform = new RelationshipsXform();
    var promises = [];

    model.worksheets.forEach(function (worksheet) {
      var drawing = worksheet.drawing;
      if (drawing) {
        promises.push(new PromishLib.Promish(function (resolve) {
          drawingXform.prepare(drawing, {});
          var xml = drawingXform.toXml(drawing);
          zip.append(xml, { name: 'xl/drawings/' + drawing.name + '.xml' });

          xml = relsXform.toXml(drawing.rels);
          zip.append(xml, { name: 'xl/drawings/_rels/' + drawing.name + '.xml.rels' });

          resolve();
        }));
      }
    });

    return PromishLib.Promish.all(promises);
  },

  addContentTypes: function addContentTypes(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      var xform = new ContentTypesXform();
      var xml = xform.toXml(model);
      zip.append(xml, { name: '[Content_Types].xml' });
      resolve();
    });
  },

  addApp: function addApp(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      var xform = new AppXform();
      var xml = xform.toXml(model);
      zip.append(xml, { name: 'docProps/app.xml' });
      resolve();
    });
  },

  addCore: function addCore(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      var coreXform = new CoreXform();
      zip.append(coreXform.toXml(model), { name: 'docProps/core.xml' });
      resolve();
    });
  },

  addThemes: function addThemes(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      var themes = model.themes || { theme1: theme1Xml };
      Object.keys(themes).forEach(function (name) {
        var xml = themes[name];
        var path = 'xl/theme/' + name + '.xml';
        zip.append(xml, { name: path });
      });
      resolve();
    });
  },

  addOfficeRels: function addOfficeRels(zip) {
    return new PromishLib.Promish(function (resolve) {
      var xform = new RelationshipsXform();
      var xml = xform.toXml([{ Id: 'rId1', Type: XLSX.RelType.OfficeDocument, Target: 'xl/workbook.xml' }, { Id: 'rId2', Type: XLSX.RelType.CoreProperties, Target: 'docProps/core.xml' }, { Id: 'rId3', Type: XLSX.RelType.ExtenderProperties, Target: 'docProps/app.xml' }]);
      zip.append(xml, { name: '_rels/.rels' });
      resolve();
    });
  },

  addWorkbookRels: function addWorkbookRels(zip, model) {
    var count = 1;
    var relationships = [{ Id: 'rId' + count++, Type: XLSX.RelType.Styles, Target: 'styles.xml' }, { Id: 'rId' + count++, Type: XLSX.RelType.Theme, Target: 'theme/theme1.xml' }];
    if (model.sharedStrings.count) {
      relationships.push({ Id: 'rId' + count++, Type: XLSX.RelType.SharedStrings, Target: 'sharedStrings.xml' });
    }
    model.worksheets.forEach(function (worksheet) {
      worksheet.rId = 'rId' + count++;
      relationships.push({ Id: worksheet.rId, Type: XLSX.RelType.Worksheet, Target: 'worksheets/sheet' + worksheet.id + '.xml' });
    });
    return new PromishLib.Promish(function (resolve) {
      var xform = new RelationshipsXform();
      var xml = xform.toXml(relationships);
      zip.append(xml, { name: 'xl/_rels/workbook.xml.rels' });
      resolve();
    });
  },
  addSharedStrings: function addSharedStrings(zip, model) {
    if (!model.sharedStrings || !model.sharedStrings.count) {
      return PromishLib.Promish.resolve();
    }
    return new PromishLib.Promish(function (resolve) {
      zip.append(model.sharedStrings.xml, { name: 'xl/sharedStrings.xml' });
      resolve();
    });
  },
  addStyles: function addStyles(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      var xml = model.styles.xml;
      if (xml) {
        zip.append(xml, { name: 'xl/styles.xml' });
      }
      resolve();
    });
  },
  addWorkbook: function addWorkbook(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      var xform = new WorkbookXform();
      zip.append(xform.toXml(model), { name: 'xl/workbook.xml' });
      resolve();
    });
  },
  addWorksheets: function addWorksheets(zip, model) {
    return new PromishLib.Promish(function (resolve) {
      // preparation phase
      var worksheetXform = new WorksheetXform();
      var relationshipsXform = new RelationshipsXform();

      // write sheets
      model.worksheets.forEach(function (worksheet) {
        var xmlStream = new XmlStream();
        worksheetXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, { name: 'xl/worksheets/sheet' + worksheet.id + '.xml' });

        if (worksheet.rels && worksheet.rels.length) {
          xmlStream = new XmlStream();
          relationshipsXform.render(xmlStream, worksheet.rels);
          zip.append(xmlStream.xml, { name: 'xl/worksheets/_rels/sheet' + worksheet.id + '.xml.rels' });
        }
      });

      resolve();
    });
  },
  _finalize: function _finalize(zip) {
    var _this = this;

    return new PromishLib.Promish(function (resolve, reject) {
      zip.on('finish', function () {
        resolve(_this);
      });
      zip.on('error', reject);
      zip.finalize();
    });
  },
  prepareModel: function prepareModel(model, options) {
    // ensure following properties have sane values
    model.creator = model.creator || 'ExcelJS';
    model.lastModifiedBy = model.lastModifiedBy || 'ExcelJS';
    model.created = model.created || new Date();
    model.modified = model.modified || new Date();

    model.useSharedStrings = options.useSharedStrings !== undefined ? options.useSharedStrings : true;
    model.useStyles = options.useStyles !== undefined ? options.useStyles : true;

    // Manage the shared strings
    model.sharedStrings = new SharedStringsXform();

    // add a style manager to handle cell formats, fonts, etc.
    model.styles = model.useStyles ? new StylesXform(true) : new StylesXform.Mock();

    // prepare all of the things before the render
    var workbookXform = new WorkbookXform();
    var worksheetXform = new WorksheetXform();

    workbookXform.prepare(model);

    var worksheetOptions = {
      sharedStrings: model.sharedStrings,
      styles: model.styles,
      date1904: model.properties.date1904,
      drawingsCount: 0,
      media: model.media
    };
    worksheetOptions.drawings = model.drawings = [];
    model.worksheets.forEach(function (worksheet) {
      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // TODO: workbook drawing list
  },
  write: function write(stream, options) {
    var _this2 = this;

    options = options || {};
    var model = this.workbook.model;
    var zip = new ZipStream.ZipWriter();
    zip.pipe(stream);

    this.prepareModel(model, options);

    // render
    return PromishLib.Promish.resolve().then(function () {
      return _this2.addContentTypes(zip, model);
    }).then(function () {
      return _this2.addOfficeRels(zip, model);
    }).then(function () {
      return _this2.addWorkbookRels(zip, model);
    }).then(function () {
      return _this2.addWorksheets(zip, model);
    }).then(function () {
      return _this2.addSharedStrings(zip, model);
    }) // always after worksheets
    .then(function () {
      return _this2.addDrawings(zip, model);
    }).then(function () {
      var promises = [_this2.addThemes(zip, model), _this2.addStyles(zip, model)];
      return PromishLib.Promish.all(promises);
    }).then(function () {
      return _this2.addMedia(zip, model);
    }).then(function () {
      var afters = [_this2.addApp(zip, model), _this2.addCore(zip, model)];
      return PromishLib.Promish.all(afters);
    }).then(function () {
      return _this2.addWorkbook(zip, model);
    }).then(function () {
      return _this2._finalize(zip);
    });
  },
  writeFile: function writeFile(filename, options) {
    var self = this;
    var stream = fs.createWriteStream(filename);

    return new PromishLib.Promish(function (resolve, reject) {
      stream.on('finish', function () {
        resolve();
      });
      stream.on('error', function (error) {
        reject(error);
      });

      self.write(stream, options).then(function () {
        stream.end();
      }).catch(function (error) {
        reject(error);
      });
    });
  },
  writeBuffer: function writeBuffer(options) {
    var self = this;
    var stream = new StreamBuf();
    return self.write(stream, options).then(function () {
      return stream.read();
    });
  }
};
//# sourceMappingURL=xlsx.js.map
