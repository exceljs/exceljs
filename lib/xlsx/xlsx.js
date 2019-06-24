const fs = require('fs');
const ZipStream = require('../utils/zip-stream');
const StreamBuf = require('../utils/stream-buf');
const PromiseLib = require('../utils/promise');

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
const CommentsXform = require('./xform/comment/comments-xform');
const VmlNotesXform = require('./xform/comment/vml-notes-xform');

const theme1Xml = require('./xml/theme1.js');

function fsReadFileAsync(filename, options) {
  return new PromiseLib.Promise((resolve, reject) => {
    fs.readFile(filename, options, (error, data) => {
      if (error) {
        reject(error);
      } else {
        resolve(data);
      }
    });
  });
}

class XLSX {
  constructor(workbook) {
    this.workbook = workbook;
  }

  // ===============================================================================
  // Workbook
  // =========================================================================
  // Read

  readFile(filename, options) {
    let stream;
    return utils.fs
      .exists(filename)
      .then(exists => {
        if (!exists) {
          throw new Error(`File not found: ${filename}`);
        }
        stream = fs.createReadStream(filename);
        return this.read(stream, options).catch(error => {
          stream.close();
          throw error;
        });
      })
      .then(workbook => {
        stream.close();
        return workbook;
      });
  }

  parseRels(stream) {
    const xform = new RelationshipsXform();
    return xform.parseStream(stream);
  }

  parseWorkbook(stream) {
    const xform = new WorkbookXform();
    return xform.parseStream(stream);
  }

  parseSharedStrings(stream) {
    const xform = new SharedStringsXform();
    return xform.parseStream(stream);
  }

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
      comments: model.comments,
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
  }

  processWorksheetEntry(entry, model, options) {
    const match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
    if (match) {
      const sheetNo = match[1];
      const xform = new WorksheetXform(options);
      return xform.parseStream(entry)
        .then(worksheet => {
          worksheet.sheetNo = sheetNo;
          model.worksheetHash[entry.path] = worksheet;
          model.worksheets.push(worksheet);
        });
    }
    return undefined;
  }

  processCommentEntry(entry, model) {
    const match = entry.path.match(/xl\/(comments\d+)[.]xml/);
    if (match) {
      const name = match[1];
      const xform = new CommentsXform();
      return xform.parseStream(entry)
        .then(comments => {
          model.comments[`../${name}.xml`] = comments;
        });
    }
    return undefined;
  }

  processWorksheetRelsEntry(entry, model) {
    const match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
    if (match) {
      const sheetNo = match[1];
      const xform = new RelationshipsXform();
      return xform.parseStream(entry)
        .then(relationships => {
          model.worksheetRels[sheetNo] = relationships;
        });
    }
    return undefined;
  }

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
      return new PromiseLib.Promise((resolve, reject) => {
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
  }

  processDrawingEntry(entry, model) {
    const match = entry.path.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
    if (match) {
      const name = match[1];
      const xform = new DrawingXform();
      return xform.parseStream(entry)
        .then(drawing => {
          model.drawings[name] = drawing;
        });
    }
    return undefined;
  }

  processDrawingRelsEntry(entry, model) {
    const match = entry.path.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
    if (match) {
      const name = match[1];
      const xform = new RelationshipsXform();
      return xform.parseStream(entry)
        .then(relationships => {
          model.drawingRels[name] = relationships;
        });
    }
    return undefined;
  }

  processThemeEntry(entry, model) {
    const match = entry.path.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
    if (match) {
      return new PromiseLib.Promise((resolve, reject) => {
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
  }

  processIgnoreEntry(entry) {
    entry.autodrain();
  }

  createInputStream(options) {
    const model = {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: [],
      themes: {},
      media: [],
      mediaIndex: {},
      drawings: {},
      drawingRels: {},
      comments: {},
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
          promise = this.parseRels(entry)
            .then(relationships => {
              model.globalRels = relationships;
            });
          break;

        case 'xl/workbook.xml':
          promise = this.parseWorkbook(entry)
            .then(workbook => {
              model.sheets = workbook.sheets;
              model.definedNames = workbook.definedNames;
              model.views = workbook.views;
              model.properties = workbook.properties;
            });
          break;

        case 'xl/_rels/workbook.xml.rels':
          promise = this.parseRels(entry)
            .then(relationships => {
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
          promise = appXform.parseStream(entry)
            .then(appProperties => {
              Object.assign(model, {
                company: appProperties.company,
                manager: appProperties.manager,
              });
            });
          break;
        }

        case 'docProps/core.xml': {
          const coreXform = new CoreXform();
          promise = coreXform.parseStream(entry)
            .then(coreProperties => {
              Object.assign(model, coreProperties);
            });
          break;
        }

        default:
          promise =
            this.processWorksheetEntry(entry, model, options) ||
            this.processWorksheetRelsEntry(entry, model) ||
            this.processThemeEntry(entry, model) ||
            this.processMediaEntry(entry, model) ||
            this.processDrawingEntry(entry, model) ||
            this.processCommentEntry(entry, model) ||
            this.processDrawingRelsEntry(entry, model) ||
            this.processIgnoreEntry(entry);
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
      PromiseLib.Promise.all(promises)
        .then(() => {
          this.reconcile(model, options);

          // apply model
          this.workbook.model = model;
        })
        .then(() => {
          stream.emit('done');
        })
        .catch(error => {
          stream.emit('error', error);
        });
    });
    return stream;
  }

  read(stream, options) {
    options = options || {};
    const zipStream = this.createInputStream(options);
    return new PromiseLib.Promise((resolve, reject) => {
      zipStream
        .on('done', () => {
          resolve(this.workbook);
        })
        .on('error', error => {
          reject(error);
        });
      stream.pipe(zipStream);
    });
  }

  load(data, options) {
    if (options === undefined) {
      options = {};
    }
    const zipStream = this.createInputStream();
    return new PromiseLib.Promise((resolve, reject) => {
      zipStream
        .on('done', () => {
          resolve(this.workbook);
        })
        .on('error', error => {
          reject(error);
        });

      if (options.base64) {
        const buffer = Buffer.from(data.toString(), 'base64');
        zipStream.write(buffer);
      } else {
        zipStream.write(data);
      }
      zipStream.end();
    });
  }

  // =========================================================================
  // Write

  addMedia(zip, model) {
    return PromiseLib.Promise.all(
      model.media.map(medium => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}.${medium.extension}`;
          if (medium.filename) {
            return fsReadFileAsync(medium.filename)
              .then(data => {
                zip.append(data, {name: filename});
              });
          }
          if (medium.buffer) {
            return new PromiseLib.Promise(resolve => {
              zip.append(medium.buffer, {name: filename});
              resolve();
            });
          }
          if (medium.base64) {
            return new PromiseLib.Promise(resolve => {
              const dataimg64 = medium.base64;
              const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
              zip.append(content, {name: filename, base64: true});
              resolve();
            });
          }
        }
        return PromiseLib.Promise.reject(new Error('Unsupported media'));
      })
    );
  }

  addDrawings(zip, model) {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();
    const promises = [];

    model.worksheets.forEach(worksheet => {
      const {drawing} = worksheet;
      if (drawing) {
        promises.push(
          new PromiseLib.Promise(resolve => {
            drawingXform.prepare(drawing, {});
            let xml = drawingXform.toXml(drawing);
            zip.append(xml, {name: `xl/drawings/${drawing.name}.xml`});

            xml = relsXform.toXml(drawing.rels);
            zip.append(xml, {name: `xl/drawings/_rels/${drawing.name}.xml.rels`});

            resolve();
          })
        );
      }
    });

    return PromiseLib.Promise.all(promises);
  }

  addContentTypes(zip, model) {
    return new PromiseLib.Promise(resolve => {
      const xform = new ContentTypesXform();
      const xml = xform.toXml(model);
      zip.append(xml, {name: '[Content_Types].xml'});
      resolve();
    });
  }

  addApp(zip, model) {
    return new PromiseLib.Promise(resolve => {
      const xform = new AppXform();
      const xml = xform.toXml(model);
      zip.append(xml, {name: 'docProps/app.xml'});
      resolve();
    });
  }

  addCore(zip, model) {
    return new PromiseLib.Promise(resolve => {
      const coreXform = new CoreXform();
      zip.append(coreXform.toXml(model), {name: 'docProps/core.xml'});
      resolve();
    });
  }

  addThemes(zip, model) {
    return new PromiseLib.Promise(resolve => {
      const themes = model.themes || {theme1: theme1Xml};
      Object.keys(themes).forEach(name => {
        const xml = themes[name];
        const path = `xl/theme/${name}.xml`;
        zip.append(xml, {name: path});
      });
      resolve();
    });
  }

  addOfficeRels(zip) {
    return new PromiseLib.Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml([
        {Id: 'rId1', Type: XLSX.RelType.OfficeDocument, Target: 'xl/workbook.xml'},
        {Id: 'rId2', Type: XLSX.RelType.CoreProperties, Target: 'docProps/core.xml'},
        {Id: 'rId3', Type: XLSX.RelType.ExtenderProperties, Target: 'docProps/app.xml'},
      ]);
      zip.append(xml, {name: '_rels/.rels'});
      resolve();
    });
  }

  addWorkbookRels(zip, model) {
    let count = 1;
    const relationships = [
      {Id: `rId${count++}`, Type: XLSX.RelType.Styles, Target: 'styles.xml'},
      {Id: `rId${count++}`, Type: XLSX.RelType.Theme, Target: 'theme/theme1.xml'},
    ];
    if (model.sharedStrings.count) {
      relationships.push({Id: `rId${count++}`, Type: XLSX.RelType.SharedStrings, Target: 'sharedStrings.xml'});
    }
    model.worksheets.forEach(worksheet => {
      worksheet.rId = `rId${count++}`;
      relationships.push({Id: worksheet.rId, Type: XLSX.RelType.Worksheet, Target: `worksheets/sheet${worksheet.id}.xml`});
    });
    return new PromiseLib.Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml(relationships);
      zip.append(xml, {name: 'xl/_rels/workbook.xml.rels'});
      resolve();
    });
  }

  addSharedStrings(zip, model) {
    if (!model.sharedStrings || !model.sharedStrings.count) {
      return PromiseLib.Promise.resolve();
    }
    return new PromiseLib.Promise(resolve => {
      zip.append(model.sharedStrings.xml, {name: 'xl/sharedStrings.xml'});
      resolve();
    });
  }

  addStyles(zip, model) {
    return new PromiseLib.Promise(resolve => {
      const {xml} = model.styles;
      if (xml) {
        zip.append(xml, {name: 'xl/styles.xml'});
      }
      resolve();
    });
  }

  addWorkbook(zip, model) {
    return new PromiseLib.Promise(resolve => {
      const xform = new WorkbookXform();
      zip.append(xform.toXml(model), {name: 'xl/workbook.xml'});
      resolve();
    });
  }

  addWorksheets(zip, model) {
    return new PromiseLib.Promise(resolve => {
      // preparation phase
      const worksheetXform = new WorksheetXform();
      const relationshipsXform = new RelationshipsXform();
      const commentsXform = new CommentsXform();
      const vmlNotesXform = new VmlNotesXform();

      // write sheets
      model.worksheets.forEach(worksheet => {
        let xmlStream = new XmlStream();
        worksheetXform.render(xmlStream, worksheet);
        zip.append(xmlStream.xml, {name: `xl/worksheets/sheet${worksheet.id}.xml`});

        if (worksheet.rels && worksheet.rels.length) {
          xmlStream = new XmlStream();
          relationshipsXform.render(xmlStream, worksheet.rels);
          zip.append(xmlStream.xml, {name: `xl/worksheets/_rels/sheet${worksheet.id}.xml.rels`});
        }

        if (worksheet.comments.length > 0) {
          xmlStream = new XmlStream();
          commentsXform.render(xmlStream, worksheet);
          zip.append(xmlStream.xml, {name: `xl/comments${worksheet.id}.xml`});

          xmlStream = new XmlStream();
          vmlNotesXform.render(xmlStream, worksheet);
          zip.append(xmlStream.xml, {name: `xl/drawings/vmlDrawing${worksheet.id}.vml`});
        }
      });

      resolve();
    });
  }

  _finalize(zip) {
    return new PromiseLib.Promise((resolve, reject) => {
      zip.on('finish', () => {
        resolve(this);
      });
      zip.on('error', reject);
      zip.finalize();
    });
  }

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
    worksheetOptions.commentRefs = model.commentRefs = [];
    model.worksheets.forEach(worksheet => {
      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // TODO: workbook drawing list
  }

  write(stream, options) {
    options = options || {};
    const {model} = this.workbook;
    const zip = new ZipStream.ZipWriter(options.zip);
    zip.pipe(stream);

    this.prepareModel(model, options);

    // render
    return PromiseLib.Promise.resolve()
      .then(() => this.addContentTypes(zip, model))
      .then(() => this.addOfficeRels(zip, model))
      .then(() => this.addWorkbookRels(zip, model))
      .then(() => this.addWorksheets(zip, model))
      .then(() => this.addSharedStrings(zip, model)) // always after worksheets
      .then(() => this.addDrawings(zip, model))
      .then(() => {
        const promises = [this.addThemes(zip, model), this.addStyles(zip, model)];
        return PromiseLib.Promise.all(promises);
      })
      .then(() => this.addMedia(zip, model))
      .then(() => {
        const afters = [this.addApp(zip, model), this.addCore(zip, model)];
        return PromiseLib.Promise.all(afters);
      })
      .then(() => this.addWorkbook(zip, model))
      .then(() => this._finalize(zip));
  }

  writeFile(filename, options) {
    const stream = fs.createWriteStream(filename);

    return new PromiseLib.Promise((resolve, reject) => {
      stream.on('finish', () => {
        resolve();
      });
      stream.on('error', error => {
        reject(error);
      });

      this
        .write(stream, options)
        .then(() => {
          stream.end();
        })
        .catch(error => {
          reject(error);
        });
    });
  }

  writeBuffer(options) {
    const stream = new StreamBuf();
    return this.write(stream, options)
      .then(() => stream.read());
  }
}

XLSX.RelType = require('./rel-type');

module.exports = XLSX;
