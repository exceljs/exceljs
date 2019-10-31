const fs = require('fs');
const ZipStream = require('../utils/zip-stream');
const StreamBuf = require('../utils/stream-buf');

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
const TableXform = require('./xform/table/table-xform');
const CommentsXform = require('./xform/comment/comments-xform');
const VmlNotesXform = require('./xform/comment/vml-notes-xform');

const theme1Xml = require('./xml/theme1.js');

function fsReadFileAsync(filename, options) {
  return new Promise((resolve, reject) => {
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

  async readFile(filename, options) {
    if (!(await utils.fs.exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const stream = fs.createReadStream(filename);
    try {
      const workbook = await this.read(stream, options);
      stream.close();
      return workbook;
    } catch (error) {
      stream.close();
      throw error;
    }
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
    const tableXform = new TableXform();

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

    // reconcile tables with the default styles
    const tableOptions = {
      styles: model.styles,
    };
    Object.values(model.tables).forEach(table => {
      tableXform.reconcile(table, tableOptions);
    });

    const sheetOptions = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      mediaIndex: model.mediaIndex,
      date1904: model.properties && model.properties.date1904,
      drawings: model.drawings,
      comments: model.comments,
      tables: model.tables,
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

  async _processWorksheetEntry(entry, model, sheetNo, options) {
    const xform = new WorksheetXform(options);
    const worksheet = await xform.parseStream(entry);
    worksheet.sheetNo = sheetNo;
    model.worksheetHash[entry.path] = worksheet;
    model.worksheets.push(worksheet);
  }

  async _processCommentEntry(entry, model, name) {
    const xform = new CommentsXform();
    const comments = await xform.parseStream(entry);
    model.comments[`../${name}.xml`] = comments;
  }

  async _processTableEntry(entry, model, name) {
    const xform = new TableXform();
    const table = await xform.parseStream(entry);
    model.tables[`../tables/${name}.xml`] = table;
  }

  async _processWorksheetRelsEntry(entry, model, sheetNo) {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry);
    model.worksheetRels[sheetNo] = relationships;
  }

  async _processMediaEntry(entry, model, filename) {
    const lastDot = filename.lastIndexOf('.');
    // if we can't determine extension, ignore it
    if (lastDot >= 1) {
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      await new Promise((resolve, reject) => {
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
  }

  async _processDrawingEntry(entry, model, name) {
    const xform = new DrawingXform();
    const drawing = await xform.parseStream(entry);
    model.drawings[name] = drawing;
  }

  async _processDrawingRelsEntry(entry, model, name) {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry);
    model.drawingRels[name] = relationships;
  }

  async _processThemeEntry(entry, model, name) {
    await new Promise((resolve, reject) => {
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
      tables: {},
    };

    // we have to be prepared to read the zip entries in whatever order they arrive
    const promises = [];
    const stream = new ZipStream.ZipReader({
      getEntryType: path => (path.match(/xl\/media\//) ? 'nodebuffer' : 'string'),
    });
    stream.on('entry', entry => {
      promises.push(
        (async () => {
          try {
            let entryPath = entry.path;
            if (entryPath[0] === '/') {
              entryPath = entryPath.substr(1);
            }
            switch (entryPath) {
              case '_rels/.rels':
                model.globalRels = await this.parseRels(entry);
                break;

              case 'xl/workbook.xml': {
                const workbook = await this.parseWorkbook(entry);
                model.sheets = workbook.sheets;
                model.definedNames = workbook.definedNames;
                model.views = workbook.views;
                model.properties = workbook.properties;
                model.calcProperties = workbook.calcProperties;
                break;
              }

              case 'xl/_rels/workbook.xml.rels':
                model.workbookRels = await this.parseRels(entry);
                break;

              case 'xl/sharedStrings.xml':
                model.sharedStrings = new SharedStringsXform();
                await model.sharedStrings.parseStream(entry);
                break;

              case 'xl/styles.xml':
                model.styles = new StylesXform();
                await model.styles.parseStream(entry);
                break;

              case 'docProps/app.xml': {
                const appXform = new AppXform();
                const appProperties = await appXform.parseStream(entry);
                model.company = appProperties.company;
                model.manager = appProperties.manager;
                break;
              }

              case 'docProps/core.xml': {
                const coreXform = new CoreXform();
                const coreProperties = await coreXform.parseStream(entry);
                Object.assign(model, coreProperties);
                break;
              }

              default: {
                let match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
                if (match) {
                  await this._processWorksheetEntry(entry, model, match[1], options);
                  break;
                }
                match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
                if (match) {
                  await this._processWorksheetRelsEntry(entry, model, match[1]);
                  break;
                }
                match = entry.path.match(/xl\/theme\/([a-zA-Z0-9]+)[.]xml/);
                if (match) {
                  await this._processThemeEntry(entry, model, match[1]);
                  break;
                }
                match = entry.path.match(/xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/);
                if (match) {
                  await this._processMediaEntry(entry, model, match[1]);
                  break;
                }
                match = entry.path.match(/xl\/drawings\/([a-zA-Z0-9]+)[.]xml/);
                if (match) {
                  await this._processDrawingEntry(entry, model, match[1]);
                  break;
                }
                match = entry.path.match(/xl\/(comments\d+)[.]xml/);
                if (match) {
                  await this._processCommentEntry(entry, model, match[1]);
                  break;
                }
                match = entry.path.match(/xl\/tables\/(table\d+)[.]xml/);
                if (match) {
                  await this._processTableEntry(entry, model, match[1]);
                  break;
                }
                match = entry.path.match(/xl\/drawings\/_rels\/([a-zA-Z0-9]+)[.]xml[.]rels/);
                if (match) {
                  await this._processDrawingRelsEntry(entry, model, match[1]);
                  break;
                }

                entry.autodrain();
              }
            }
          } catch (error) {
            stream.destroy(error);
            throw error;
          }
        })()
      );
    });
    stream.on('finished', async () => {
      try {
        await Promise.all(promises);
        this.reconcile(model, options);

        // apply model
        this.workbook.model = model;
        stream.emit('done');
      } catch (error) {
        stream.emit('error', error);
      }
    });
    return stream;
  }

  read(stream, options) {
    return new Promise((resolve, reject) => {
      options = options || {};
      const zipStream = this.createInputStream(options);
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
    return new Promise((resolve, reject) => {
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

  async addMedia(zip, model) {
    await Promise.all(
      model.media.map(async medium => {
        if (medium.type === 'image') {
          const filename = `xl/media/${medium.name}.${medium.extension}`;
          if (medium.filename) {
            const data = await fsReadFileAsync(medium.filename);
            return zip.append(data, {name: filename});
          }
          if (medium.buffer) {
            return zip.append(medium.buffer, {name: filename});
          }
          if (medium.base64) {
            const dataimg64 = medium.base64;
            const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
            return zip.append(content, {name: filename, base64: true});
          }
        }
        throw new Error('Unsupported media');
      })
    );
  }

  addDrawings(zip, model) {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    model.worksheets.forEach(worksheet => {
      const {drawing} = worksheet;
      if (drawing) {
        drawingXform.prepare(drawing, {});
        let xml = drawingXform.toXml(drawing);
        zip.append(xml, {name: `xl/drawings/${drawing.name}.xml`});

        xml = relsXform.toXml(drawing.rels);
        zip.append(xml, {name: `xl/drawings/_rels/${drawing.name}.xml.rels`});
      }
    });
  }

  addTables(zip, model) {
    const tableXform = new TableXform();

    model.worksheets.forEach(worksheet => {
      const {tables} = worksheet;
      tables.forEach(table => {
        tableXform.prepare(table, {});
        const tableXml = tableXform.toXml(table);
        zip.append(tableXml, {name: `xl/tables/${table.target}`});
      });
    });
  }

  async addContentTypes(zip, model) {
    const xform = new ContentTypesXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: '[Content_Types].xml'});
  }

  async addApp(zip, model) {
    const xform = new AppXform();
    const xml = xform.toXml(model);
    zip.append(xml, {name: 'docProps/app.xml'});
  }

  async addCore(zip, model) {
    const coreXform = new CoreXform();
    zip.append(coreXform.toXml(model), {name: 'docProps/core.xml'});
  }

  async addThemes(zip, model) {
    const themes = model.themes || {theme1: theme1Xml};
    Object.keys(themes).forEach(name => {
      const xml = themes[name];
      const path = `xl/theme/${name}.xml`;
      zip.append(xml, {name: path});
    });
  }

  async addOfficeRels(zip) {
    const xform = new RelationshipsXform();
    const xml = xform.toXml([
      {Id: 'rId1', Type: XLSX.RelType.OfficeDocument, Target: 'xl/workbook.xml'},
      {Id: 'rId2', Type: XLSX.RelType.CoreProperties, Target: 'docProps/core.xml'},
      {Id: 'rId3', Type: XLSX.RelType.ExtenderProperties, Target: 'docProps/app.xml'},
    ]);
    zip.append(xml, {name: '_rels/.rels'});
  }

  async addWorkbookRels(zip, model) {
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
    const xform = new RelationshipsXform();
    const xml = xform.toXml(relationships);
    zip.append(xml, {name: 'xl/_rels/workbook.xml.rels'});
  }

  async addSharedStrings(zip, model) {
    if (model.sharedStrings && model.sharedStrings.count) {
      zip.append(model.sharedStrings.xml, {name: 'xl/sharedStrings.xml'});
    }
  }

  async addStyles(zip, model) {
    const {xml} = model.styles;
    if (xml) {
      zip.append(xml, {name: 'xl/styles.xml'});
    }
  }

  async addWorkbook(zip, model) {
    const xform = new WorkbookXform();
    zip.append(xform.toXml(model), {name: 'xl/workbook.xml'});
  }

  async addWorksheets(zip, model) {
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
  }

  _finalize(zip) {
    return new Promise((resolve, reject) => {
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
    let tableCount = 0;
    model.tables = [];
    model.worksheets.forEach(worksheet => {
      // assign unique filenames to tables
      worksheet.tables.forEach(table => {
        tableCount++;
        table.target = `table${tableCount}.xml`;
        table.id = tableCount;
        model.tables.push(table);
      });

      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // TODO: workbook drawing list
  }

  async write(stream, options) {
    options = options || {};
    const {model} = this.workbook;
    const zip = new ZipStream.ZipWriter(options.zip);
    zip.pipe(stream);

    this.prepareModel(model, options);

    // render
    await this.addContentTypes(zip, model);
    await this.addOfficeRels(zip, model);
    await this.addWorkbookRels(zip, model);
    await this.addWorksheets(zip, model);
    await this.addSharedStrings(zip, model); // always after worksheets
    await this.addDrawings(zip, model);
    await this.addTables(zip, model);
    await Promise.all([this.addThemes(zip, model), this.addStyles(zip, model)]);
    await this.addMedia(zip, model);
    await Promise.all([this.addApp(zip, model), this.addCore(zip, model)]);
    await this.addWorkbook(zip, model);
    return this._finalize(zip);
  }

  writeFile(filename, options) {
    const stream = fs.createWriteStream(filename);

    return new Promise((resolve, reject) => {
      stream.on('finish', () => {
        resolve();
      });
      stream.on('error', error => {
        reject(error);
      });

      this.write(stream, options)
        .then(() => {
          stream.end();
        });
    });
  }

  async writeBuffer(options) {
    const stream = new StreamBuf();
    await this.write(stream, options);
    return stream.read();
  }
}

XLSX.RelType = require('./rel-type');

module.exports = XLSX;
