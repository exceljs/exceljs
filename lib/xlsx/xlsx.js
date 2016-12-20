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
var _ = require('lodash');
var Archiver = require('archiver');
var unzip = require('unzip2');

var utils = require('../utils/utils');
var XmlStream = require('../utils/xml-stream');
var StreamBuf = require('../utils/stream-buf');

var StylesXform = require('./xform/style/styles-xform');

var CoreXform = require('./xform/core/core-xform');
var SharedStringsXform = require('./xform/strings/shared-strings-xform');
var RelationshipsXform = require('./xform/core/relationships-xform');
var ContentTypesXform = require('./xform/core/content-types-xform');
var AppXform = require('./xform/core/app-xform');
var WorkbookXform = require('./xform/book/workbook-xform');
var WorksheetXform = require('./xform/sheet/worksheet-xform');

var fsReadFileAsync = Bluebird.promisify(fs.readFile);

var XLSX = module.exports = function (workbook) {
  this.workbook = workbook;
};


var getTheme1Xml = utils.readModuleFile(require.resolve('./xml/theme1.xml'));

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
    var xform = new RelationshipsXform();
    return xform.parseStream(stream);
  },
  parseWorkbook: function (stream) {
    var xform = new WorkbookXform();
    return xform.parseStream(stream);
  },
  parseSharedStrings: function (stream) {
    var xform = new SharedStringsXform();
    return xform.parseStream(stream);
  },
  parseWorksheet: function (stream) {
    var xform = new WorksheetXform();
    return xform.parseStream(stream);
  },
  reconcile: function (model) {
    var workbookXform = new WorkbookXform();
    var worksheetXform = new WorksheetXform();
    
    workbookXform.reconcile(model);
    var sheetOptions = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      date1904: model.properties.date1904
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
  },
  createInputStream: function () {
    var self = this;
    var model = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: [],
      media: {}
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
              model.sheets = workbook.sheets;
              model.definedNames = workbook.definedNames;
              model.views = workbook.views;
              model.properties = workbook.properties;
            });
          break;
        case 'xl/_rels/workbook.xml.rels':
          promise = self.parseRels(entry)
            .then(function (relationships) {
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
          promise = appXform.parseStream(entry)
            .then(function(appProperties) {
              _.extend(model, {
                company: appProperties.company,
                manager: appProperties.manager
              });
            });
          break;
        case 'docProps/core.xml':
          var coreXform = new CoreXform();
          promise = coreXform.parseStream(entry)
            .then(function(coreProperties) {
              _.extend(model, coreProperties);
            });
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
          } else if (entry.path.match(/xl\/media\/.+/)) {
            var name = entry.path.substr(9);
            return new Bluebird(function(resolve, reject) {
              var streamBuf = new StreamBuf();
              entry.on('end', function() {
                model.media[name] = {
                  type: 'image',
                  image: {
                    type: name.split('.')[1],
                    buffer: streamBuf.toBuffer()
                  }
                };
                resolve();
              });
              entry.on('error', function(error) {
                reject(error);
              });
              entry.pipe(streamBuf);
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
          self.reconcile(model);

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

  addMedia: function(zip, model) {
    return Bluebird.all(model.media.map(function(medium) {
      if (medium.type === 'image') {
        var image = medium.image;
        var filename = 'xl/media/image' + medium.index + '.' + image.type;
        if (image.filename) {
          return fsReadFileAsync(image.filename)
            .then(function(data) {
              zip.append(data, {name: filename});
            });
        } else {
          // assume image buffer
          return new Bluebird(function(resolve) {
            zip.append(image.buffer, {name: filename});
            resolve();
          });
        }
      } else {
        return Bluebird.resolve();
      }
    }));
  },

  addContentTypes: function (zip, model) {
    return new Bluebird(function(resolve) {
      var xform = new ContentTypesXform();
      var xml = xform.toXml(model);
      zip.append(xml, {name: '[Content_Types].xml'});
      resolve();
    });
  },

  addApp: function (zip, model) {
    return new Bluebird(function(resolve) {
      var xform = new AppXform();
      var xml = xform.toXml(model);
      zip.append(xml, {name: 'docProps/app.xml'});
      resolve();
    });
  },

  addCore: function (zip, model) {
    return new Bluebird(function(resolve) {
      var coreXform = new CoreXform();
      zip.append(coreXform.toXml(model), {name: 'docProps/core.xml'});
      resolve();
    });
  },

  addThemes: function (zip) {
    return getTheme1Xml
      .then(function (data) {
        zip.append(data, {name: 'xl/theme/theme1.xml'});
        return zip;
      });
  },

  addOfficeRels: function (zip) {
    return new Bluebird(function(resolve) {
      var xform = new RelationshipsXform();
      var xml = xform.toXml([
          {Id: 'rId1', Type: XLSX.RelType.OfficeDocument, Target: 'xl/workbook.xml'}
        ]);
      zip.append(xml, {name: '/_rels/.rels'});
      resolve();
    });
  },

  addWorkbookRels: function (zip, model) {
    var count = 1;
    var relationships = [
        {Id: 'rId' + (count++), Type: XLSX.RelType.Styles, Target: 'styles.xml'},
        {Id: 'rId' + (count++), Type: XLSX.RelType.Theme, Target: 'theme/theme1.xml'}
    ];
    if (model.sharedStrings.count) {
      relationships.push(
        {Id: 'rId' + (count++), Type: XLSX.RelType.SharedStrings, Target: 'sharedStrings.xml'}
      );
    }
    model.worksheets.forEach(function (worksheet) {
      worksheet.rId = 'rId' + (count++);
      relationships.push(
        {Id: worksheet.rId, Type: XLSX.RelType.Worksheet, Target: 'worksheets/sheet' + worksheet.id + '.xml'}
      );
    });
    return new Bluebird(function(resolve) {
      var xform = new RelationshipsXform();
      var xml = xform.toXml(relationships);
      zip.append(xml, {name: '/xl/_rels/workbook.xml.rels'});
      resolve();
    });
  },
  addSharedStrings: function (zip, model) {
    if (!model.sharedStrings || !model.sharedStrings.count) {
      return Bluebird.resolve();
    } else {
      return new Bluebird(function(resolve) {
        zip.append(model.sharedStrings.xml, {name: '/xl/sharedStrings.xml'});
        resolve();
      });
    }
  },
  addStyles: function(zip, model) {
    return new Bluebird(function(resolve) {
      var xml = model.styles.xml;
      if (xml) {
        zip.append(xml, {name: '/xl/styles.xml'});
      }
      resolve();
    });
  },
  addWorkbook: function (zip, model, xform) {
    return new Bluebird(function(resolve) {
      zip.append(xform.toXml(model), {name: '/xl/workbook.xml'});
      resolve();
    });
  },
  addWorksheets: function (zip, model, worksheetXform) {
    return new Bluebird(function(resolve) {

      // preparation phase
      var relationshipsXform = new RelationshipsXform();

      // write sheets
      model.worksheets.forEach(function (worksheet) {
        var xmlStream = new XmlStream();
        worksheetXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, {name: '/xl/worksheets/sheet' + worksheet.id + '.xml'});

        if (worksheet.rels && worksheet.rels.length) {
          xmlStream = new XmlStream();
          relationshipsXform.render(xmlStream, worksheet.rels);
          zip.append(xmlStream.xml, {name: '/xl/worksheets/_rels/sheet' + worksheet.id + '.xml.rels'});
        }
      });

      resolve();
    });
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

    // add a style manager to handle cell formats, fonts, etc.
    model.styles = model.useStyles ? new StylesXform(true) : new StylesXform.Mock();

    // prepare all of the things before the render
    var workbookXform = new WorkbookXform();
    var worksheetXform = new WorksheetXform();
    var prepareOptions = {
      sharedStrings: model.sharedStrings,
      styles: model.styles,
      date1904: model.properties.date1904
    };
    workbookXform.prepare(model);
    model.worksheets.forEach(function (worksheet) {
      worksheetXform.prepare(worksheet, prepareOptions);
    });

    // render
    var promises = [
      self.addMedia(zip, model),
      self.addContentTypes(zip, model),
      self.addApp(zip, model),
      self.addCore(zip, model),
      self.addThemes(zip),
      self.addOfficeRels(zip, model)
    ];
    return Bluebird.all(promises)
      .then(function () {
        return self.addWorksheets(zip, model, worksheetXform);
      })
      .then(function () {
        // Some things can only be done after all the worksheets have been processed
        var afters = [
          self.addSharedStrings(zip, model),
          self.addStyles(zip, model),
          self.addWorkbookRels(zip, model)
        ];
        return Bluebird.all(afters);
      })
      .then(function () {
        return self.addWorkbook(zip, model, workbookXform);
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
