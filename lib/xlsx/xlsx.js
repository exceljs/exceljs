'use strict';

const fs = require('fs');
const ZipStream = require('../utils/zip-stream');
const StreamBuf = require('../utils/stream-buf');
const PromishLib = require('../utils/promish');

const utils = require('../utils/utils');
const XmlStream = require('../utils/xml-stream');

const StylesXform = require('./xform/style/styles-xform');

const CoreXform = require('./xform/core/core-xform');
const SharedStringsXform = require('./xform/strings/shared-strings-xform');
const RelationshipsXform = require('./xform/core/relationships-xform');
const ContentTypesXform = require('./xform/core/content-types-xform');
const AppXform = require('./xform/core/app-xform');
const WorkbookXform = require('./xform/book/workbook-xform');
const WorksheetXform = require('./xform/sheet/worksheet-xform');
const DrawingXform = require('./xform/drawing/drawing-xform');

const theme1Xml = require('./xml/theme1.js');

const XLSX = (module.exports = function(workbook) {
  this.workbook = workbook;
});

function fsReadFileAsync(filename, options) {
  return new PromishLib.Promish((resolve, reject) => {
    fs.readFile(filename, options, (error, data) => {
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

  readFile(filename, options) {
    const self = this;
    let stream;
    return utils.fs
      .exists(filename)
      .then(exists => {
        if (!exists) {
          throw new Error(`File not found: ${filename}`);
        }
        stream = fs.createReadStream(filename);
        return self.read(stream, options).catch(error => {
          stream.close();
          throw error;
        });
      })
      .then(workbook => {
        stream.close();
        return workbook;
      });
  },
  parseRels(stream) {
    const xform = new RelationshipsXform();
    return xform.parseStream(stream);
  },
  parseWorkbook(stream) {
    const xform = new WorkbookXform();
    return xform.parseStream(stream);
  },
  parseSharedStrings(stream) {
    const xform = new SharedStringsXform();
    return xform.parseStream(stream);
  },
  reconcile(model, options) {
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform(options);
    const drawingXform = new DrawingXform();

    workbookXform.reconcile(model);

    // reconcile drawings with their rels
    const drawingOptions = {
      media: model.media,
      mediaIndex: model.mediaIndex,
    };
    Object.keys(model.drawings).forEach(name => {
      const drawing = model.drawings[name];
      const drawingRel = model.drawingRels[name];
      if (drawingRel) {
        drawingOptions.rels = drawingRel.reduce((o, rel) => {
          o[rel.Id] = rel;
          return o;
        }, {});
        drawingXform.reconcile(drawing, drawingOptions);
      }
    });

    const sheetOptions = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      mediaIndex: model.mediaIndex,
      date1904: model.properties && model.properties.date1904,
      drawings: model.drawings,
    };
    model.worksheets.forEach(worksheet => {
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
  processWorksheetEntry(entry, model, options) {
    const match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
    if (match) {
      const sheetNo = match[1];
      const xform = new WorksheetXform(options);
      return xform.parseStream(entry).then(worksheet => {
        worksheet.sheetNo = sheetNo;
        model.worksheetHash[entry.path] = worksheet;
        model.worksheets.push(worksheet);
      });
    }
    return undefined;
  },
  processWorksheetRelsEntry(entry, model) {
    const match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
    if (match) {
      const sheetNo = match[1];
      const xform = new RelationshipsXform();
      return xform.parseStream(entry).then(relationships => {
        model.worksheetRels[sheetNo] = relationships;
      });
    }
    return undefined;
  },
  processMediaEntry(entry, model) {
    const match = entry.path.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
    if (match) {
      const filename = match[1];
      const lastDot = filename.lastIndexOf('.');
      if (lastDot === -1) {
        // if we can't determine extension, ignore it
        return undefined;
      }
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      return new PromishLib.Promish((resolve, reject) => {
        const streamBuf = new StreamBuf();
        streamBuf.on('finish', () => {
          model.mediaIndex[filename] = model.media.length;
          model.mediaIndex[name] = model.media.length;
          const medium = {
            type: 'image',
            name,
            extension,
            buffer: streamBuf.toBuffer(),
          };
          model.media.push(medium);
          resolve();
        });
        entry.on('error', error => {
          reject(error);
        });
        entry.pipe(streamBuf);
      });
    }
    return undefined;
  },
  processDrawingEntry(entry, model) {
    const match = entry.path.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
    if (match) {
      const name = match[1];
      const xform = new DrawingXform();
      return xform.parseStream(entry).then(drawing => {
        model.drawings[name] = drawing;
      });
    }
    return undefined;
  },
  processDrawingRelsEntry(entry, model) {
    const match = entry.path.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
    if (match) {
      const name = match[1];
      const xform = new RelationshipsXform();
      return xform.parseStream(entry).then(relationships => {
        model.drawingRels[name] = relationships;
      });
    }
    return undefined;
  },
  processThemeEntry(entry, model) {
    const match = entry.path.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
    if (match) {
      return new PromishLib.Promish((resolve, reject) => {
        const name = match[1];
        // TODO: stream entry into buffer and store the xml in the model.themes[]
        const stream = new StreamBuf();
        entry.on('error', reject);
        stream.on('error', reject);
        stream.on('finish', () => {
          model.themes[name] = stream.read().toString();
          resolve();
        });
        entry.pipe(stream);
      });
    }
    return undefined;
  },
  processIgnoreEntry(entry) {
    entry.autodrain();
  },
  createInputStream(options) {
    const self = this;
    const model = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: [],
      themes: {},
      media: [],
      mediaIndex: {},
      drawings: {},
      drawingRels: {},
    };

    // we have to be prepared to read the zip entries in whatever order they arrive
    const promises = [];
    const stream = new ZipStream.ZipReader({
      getEntryType: path => (path.match(/xl\/media\//) ? 'nodebuffer' : 'string'),
    });
    stream.on('entry', entry => {
      let promise = null;

      let entryPath = entry.path;
      if (entryPath[0] === '/') {
        entryPath = entryPath.substr(1);
      }
      switch (entryPath) {
        case '_rels/.rels':
          promise = self.parseRels(entry).then(relationships => {
            model.globalRels = relationships;
          });
          break;

        case 'xl/workbook.xml':
          promise = self.parseWorkbook(entry).then(workbook => {
            model.sheets = workbook.sheets;
            model.definedNames = workbook.definedNames;
            model.views = workbook.views;
            model.properties = workbook.properties;
          });
          break;

        case 'xl/_rels/workbook.xml.rels':
          promise = self.parseRels(entry).then(relationships => {
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

        case 'docProps/app.xml': {
          const appXform = new AppXform();
          promise = appXform.parseStream(entry).then(appProperties => {
            Object.assign(model, {
              company: appProperties.company,
              manager: appProperties.manager,
            });
          });
          break;
        }

        case 'docProps/core.xml': {
          const coreXform = new CoreXform();
          promise = coreXform.parseStream(entry).then(coreProperties => {
            Object.assign(model, coreProperties);
          });
          break;
        }

        default:
          promise =
            self.processWorksheetEntry(entry, model, options) ||
            self.processWorksheetRelsEntry(entry, model) ||
            self.processThemeEntry(entry, model) ||
            self.processMediaEntry(entry, model) ||
            self.processDrawingEntry(entry, model) ||
            self.processDrawingRelsEntry(entry, model) ||
            self.processIgnoreEntry(entry);
          break;
      }

      if (promise) {
        promise = promise.catch(error => {
          stream.destroy(error);
          throw error;
        });

        promises.push(promise);
        promise = null;
      }
    });
    stream.on('finished', () => {
      PromishLib.Promish.all(promises)
        .then(() => {
          self.reconcile(model, options);

          // apply model
          self.workbook.model = model;
        })
        .then(() => {
          stream.emit('done');
        })
        .catch(error => {
          stream.emit('error', error);
        });
    });
    return stream;
  },

  read(stream, options) {
    options = options || {};
    const self = this;
    const zipStream = this.createInputStream(options);
    return new PromishLib.Promish((resolve, reject) => {
      zipStream
        .on('done', () => {
          resolve(self.workbook);
        })
        .on('error', error => {
          reject(error);
        });
      stream.pipe(zipStream);
    });
  },

  load(data, options) {
    const self = this;
    if (options === undefined) {
      options = {};
    }
    const zipStream = this.createInputStream();
    return new PromishLib.Promish((resolve, reject) => {
      zipStream
        .on('done', () => {
          resolve(self.workbook);
        })
        .on('error', error => {
          reject(error);
        });

      if (options.base64) {
        const buffer = new Buffer(data.toString(), 'base64');
        zipStream.write(buffer);
      } else {
        zipStream.write(data);
      }
      zipStream.end();
    });
  },

  // =========================================================================
  // Write

  addMedia(zip, model) {
    return PromishLib.Promish.all(
      model.media.map(medium => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}.${medium.extension}`;
          if (medium.filename) {
            return fsReadFileAsync(medium.filename).then(data => {
              zip.append(data, { name: filename });
            });
          }
          if (medium.buffer) {
            return new PromishLib.Promish(resolve => {
              zip.append(medium.buffer, { name: filename });
              resolve();
            });
          }
          if (medium.base64) {
            return new PromishLib.Promish(resolve => {
              const dataimg64 = medium.base64;
              const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
              zip.append(content, { name: filename, base64: true });
              resolve();
            });
          }
        }
        return PromishLib.Promish.reject(new Error('Unsupported media'));
      })
    );
  },

  addDrawings(zip, model) {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();
    const promises = [];

    model.worksheets.forEach(worksheet => {
      const drawing = worksheet.drawing;
      if (drawing) {
        promises.push(
          new PromishLib.Promish(resolve => {
            drawingXform.prepare(drawing, {});
            let xml = drawingXform.toXml(drawing);
            zip.append(xml, { name: `xl/drawings/${drawing.name}.xml` });

            xml = relsXform.toXml(drawing.rels);
            zip.append(xml, { name: `xl/drawings/_rels/${drawing.name}.xml.rels` });

            resolve();
          })
        );
      }
    });

    return PromishLib.Promish.all(promises);
  },

  addContentTypes(zip, model) {
    return new PromishLib.Promish(resolve => {
      const xform = new ContentTypesXform();
      const xml = xform.toXml(model);
      zip.append(xml, { name: '[Content_Types].xml' });
      resolve();
    });
  },

  addApp(zip, model) {
    return new PromishLib.Promish(resolve => {
      const xform = new AppXform();
      const xml = xform.toXml(model);
      zip.append(xml, { name: 'docProps/app.xml' });
      resolve();
    });
  },

  addCore(zip, model) {
    return new PromishLib.Promish(resolve => {
      const coreXform = new CoreXform();
      zip.append(coreXform.toXml(model), { name: 'docProps/core.xml' });
      resolve();
    });
  },

  addThemes(zip, model) {
    return new PromishLib.Promish(resolve => {
      const themes = model.themes || { theme1: theme1Xml };
      Object.keys(themes).forEach(name => {
        const xml = themes[name];
        const path = `xl/theme/${name}.xml`;
        zip.append(xml, { name: path });
      });
      resolve();
    });
  },

  addOfficeRels(zip) {
    return new PromishLib.Promish(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml([
        { Id: 'rId1', Type: XLSX.RelType.OfficeDocument, Target: 'xl/workbook.xml' },
        { Id: 'rId2', Type: XLSX.RelType.CoreProperties, Target: 'docProps/core.xml' },
        { Id: 'rId3', Type: XLSX.RelType.ExtenderProperties, Target: 'docProps/app.xml' },
      ]);
      zip.append(xml, { name: '_rels/.rels' });
      resolve();
    });
  },

  addWorkbookRels(zip, model) {
    let count = 1;
    const relationships = [
      { Id: `rId${count++}`, Type: XLSX.RelType.Styles, Target: 'styles.xml' },
      { Id: `rId${count++}`, Type: XLSX.RelType.Theme, Target: 'theme/theme1.xml' },
    ];
    if (model.sharedStrings.count) {
      relationships.push({ Id: `rId${count++}`, Type: XLSX.RelType.SharedStrings, Target: 'sharedStrings.xml' });
    }
    model.worksheets.forEach(worksheet => {
      worksheet.rId = `rId${count++}`;
      relationships.push({ Id: worksheet.rId, Type: XLSX.RelType.Worksheet, Target: `worksheets/sheet${worksheet.id}.xml` });
    });
    return new PromishLib.Promish(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml(relationships);
      zip.append(xml, { name: 'xl/_rels/workbook.xml.rels' });
      resolve();
    });
  },
  addSharedStrings(zip, model) {
    if (!model.sharedStrings || !model.sharedStrings.count) {
      return PromishLib.Promish.resolve();
    }
    return new PromishLib.Promish(resolve => {
      zip.append(model.sharedStrings.xml, { name: 'xl/sharedStrings.xml' });
      resolve();
    });
  },
  addStyles(zip, model) {
    return new PromishLib.Promish(resolve => {
      const xml = model.styles.xml;
      if (xml) {
        zip.append(xml, { name: 'xl/styles.xml' });
      }
      resolve();
    });
  },
  addWorkbook(zip, model) {
    return new PromishLib.Promish(resolve => {
      const xform = new WorkbookXform();
      zip.append(xform.toXml(model), { name: 'xl/workbook.xml' });
      resolve();
    });
  },
  addWorksheets(zip, model) {
    return new PromishLib.Promish(resolve => {
      // preparation phase
      const worksheetXform = new WorksheetXform();
      const relationshipsXform = new RelationshipsXform();

      // write sheets
      model.worksheets.forEach(worksheet => {
        let xmlStream = new XmlStream();
        worksheetXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, { name: `xl/worksheets/sheet${worksheet.id}.xml` });

        if (worksheet.rels && worksheet.rels.length) {
          xmlStream = new XmlStream();
          relationshipsXform.render(xmlStream, worksheet.rels);
          zip.append(xmlStream.xml, { name: `xl/worksheets/_rels/sheet${worksheet.id}.xml.rels` });
        }
      });

      resolve();
    });
  },
  _finalize(zip) {
    return new PromishLib.Promish((resolve, reject) => {
      zip.on('finish', () => {
        resolve(this);
      });
      zip.on('error', reject);
      zip.finalize();
    });
  },
  prepareModel(model, options) {
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
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorksheetXform();

    workbookXform.prepare(model);

    const worksheetOptions = {
      sharedStrings: model.sharedStrings,
      styles: model.styles,
      date1904: model.properties.date1904,
      drawingsCount: 0,
      media: model.media,
    };
    worksheetOptions.drawings = model.drawings = [];
    model.worksheets.forEach(worksheet => {
      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // TODO: workbook drawing list
  },
  write(stream, options) {
    options = options || {};
    const model = this.workbook.model;
    const zip = new ZipStream.ZipWriter();
    zip.pipe(stream);

    this.prepareModel(model, options);

    // render
    return PromishLib.Promish.resolve()
      .then(() => this.addContentTypes(zip, model))
      .then(() => this.addOfficeRels(zip, model))
      .then(() => this.addWorkbookRels(zip, model))
      .then(() => this.addWorksheets(zip, model))
      .then(() => this.addSharedStrings(zip, model)) // always after worksheets
      .then(() => this.addDrawings(zip, model))
      .then(() => {
        const promises = [this.addThemes(zip, model), this.addStyles(zip, model)];
        return PromishLib.Promish.all(promises);
      })
      .then(() => this.addMedia(zip, model))
      .then(() => {
        const afters = [this.addApp(zip, model), this.addCore(zip, model)];
        return PromishLib.Promish.all(afters);
      })
      .then(() => this.addWorkbook(zip, model))
      .then(() => this._finalize(zip));
  },
  writeFile(filename, options) {
    const self = this;
    const stream = fs.createWriteStream(filename);

    return new PromishLib.Promish((resolve, reject) => {
      stream.on('finish', () => {
        resolve();
      });
      stream.on('error', error => {
        reject(error);
      });

      self
        .write(stream, options)
        .then(() => {
          stream.end();
        })
        .catch(error => {
          reject(error);
        });
    });
  },
  writeBuffer(options) {
    const self = this;
    const stream = new StreamBuf();
    return self.write(stream, options).then(() => stream.read());
  },
};
