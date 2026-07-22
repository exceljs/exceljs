const fs = require('fs');
const {EventEmitter} = require('events');
const {Readable} = require('readable-stream');
const nodeStream = require('stream');
const unzip = require('unzipper');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
const RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

class WorkbookReader extends EventEmitter {
  constructor(input, options = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  _getStream(input) {
    if (input instanceof nodeStream.Readable || input instanceof Readable) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error(`Could not recognise input: ${input}`);
  }

  async read(input, options) {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async _openZip(source) {
    if (typeof source === 'string') {
      return unzip.Open.file(source);
    }
    // unzip.Open reads the archive's central directory, which needs the whole
    // (compressed) archive addressable, so buffer stream input here. This does NOT
    // materialize the much larger decompressed workbook: worksheets below are still
    // streamed row-by-row via iterateStream.
    let buffer;
    if (Buffer.isBuffer(source)) {
      buffer = source;
    } else {
      const stream = this._getStream(source);
      const chunks = [];
      for await (const chunk of stream) {
        chunks.push(chunk);
      }
      buffer = Buffer.concat(chunks);
    }
    return unzip.Open.buffer(buffer);
  }

  async *parse(input, options) {
    if (options) this.options = options;
    const source = input || this.input;

    // Read the archive through its central directory instead of a single streaming
    // unzip pass. The streaming pass emits entries in stored order and can lose the
    // final entry under async iteration on Node >= 18; because xl/workbook.xml is
    // written last, this.model was frequently never set and _parseWorksheet threw on
    // `this.model.sheets`. Opening the central directory makes every part addressable,
    // so rels/workbook/sharedStrings/styles are always parsed before any worksheet.
    // Worksheets are still consumed lazily, so the full workbook is never held in memory.
    const directory = await this._openZip(source);
    const byPath = new Map();
    for (const file of directory.files) {
      if (file.type === 'File') {
        byPath.set(file.path, file);
      }
    }

    const rels = byPath.get('xl/_rels/workbook.xml.rels');
    if (rels) {
      await this._parseRels(rels.stream());
    }

    const workbook = byPath.get('xl/workbook.xml');
    if (workbook) {
      await this._parseWorkbook(workbook.stream());
    }

    const sharedStrings = byPath.get('xl/sharedStrings.xml');
    if (sharedStrings) {
      yield* this._parseSharedStrings(sharedStrings.stream());
    }

    const styles = byPath.get('xl/styles.xml');
    if (styles) {
      await this._parseStyles(styles.stream());
    }

    const worksheetEntries = [];
    const worksheetRels = new Map();
    for (const file of byPath.values()) {
      let match = file.path.match(/^xl\/worksheets\/sheet(\d+)[.]xml$/);
      if (match) {
        worksheetEntries.push({sheetNo: match[1], file});
        // eslint-disable-next-line no-continue
        continue;
      }
      match = file.path.match(/^xl\/worksheets\/_rels\/sheet(\d+)[.]xml[.]rels$/);
      if (match) {
        worksheetRels.set(match[1], file);
      }
    }
    worksheetEntries.sort((a, b) => Number(a.sheetNo) - Number(b.sheetNo));

    for (const {sheetNo, file} of worksheetEntries) {
      const relFile = worksheetRels.get(sheetNo);
      if (relFile) {
        yield* this._parseHyperlinks(iterateStream(relFile.stream()), sheetNo);
      }
      yield* this._parseWorksheet(iterateStream(file.stream()), sheetNo);
    }
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseRels(entry) {
    const xform = new RelationshipsXform();
    this.workbookRels = await xform.parseStream(iterateStream(entry));
  }

  async _parseWorkbook(entry) {
    this._emitEntry({type: 'workbook'});

    const workbook = new WorkbookXform();
    await workbook.parseStream(iterateStream(entry));

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  async *_parseSharedStrings(entry) {
    this._emitEntry({type: 'shared-strings'});
    switch (this.options.sharedStrings) {
      case 'cache':
        this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        return;
    }

    let text = null;
    let richText = [];
    let index = 0;
    let font = null;
    for await (const events of parseSax(iterateStream(entry))) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          const node = value;
          switch (node.name) {
            case 'b':
              font = font || {};
              font.bold = true;
              break;
            case 'charset':
              font = font || {};
              font.charset = parseInt(node.attributes.charset, 10);
              break;
            case 'color':
              font = font || {};
              font.color = {};
              if (node.attributes.rgb) {
                font.color.argb = node.attributes.argb;
              }
              if (node.attributes.val) {
                font.color.argb = node.attributes.val;
              }
              if (node.attributes.theme) {
                font.color.theme = node.attributes.theme;
              }
              break;
            case 'family':
              font = font || {};
              font.family = parseInt(node.attributes.val, 10);
              break;
            case 'i':
              font = font || {};
              font.italic = true;
              break;
            case 'outline':
              font = font || {};
              font.outline = true;
              break;
            case 'rFont':
              font = font || {};
              font.name = node.value;
              break;
            case 'si':
              font = null;
              richText = [];
              text = null;
              break;
            case 'sz':
              font = font || {};
              font.size = parseInt(node.attributes.val, 10);
              break;
            case 'strike':
              break;
            case 't':
              text = null;
              break;
            case 'u':
              font = font || {};
              font.underline = true;
              break;
            case 'vertAlign':
              font = font || {};
              font.vertAlign = node.attributes.val;
              break;
          }
        } else if (eventType === 'text') {
          text = text ? text + value : value;
        } else if (eventType === 'closetag') {
          const node = value;
          switch (node.name) {
            case 'r':
              richText.push({
                font,
                text,
              });

              font = null;
              text = null;
              break;
            case 'si':
              if (this.options.sharedStrings === 'cache') {
                this.sharedStrings.push(richText.length ? {richText} : text);
              } else if (this.options.sharedStrings === 'emit') {
                yield {index: index++, text: richText.length ? {richText} : text};
              }

              richText = [];
              font = null;
              text = null;
              break;
          }
        }
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(iterateStream(entry));
    }
  }

  *_parseWorksheet(iterator, sheetNo) {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });

    const matchingRel = (this.workbookRels || []).find(rel => rel.Target === `worksheets/sheet${sheetNo}.xml`);
    const matchingSheet = matchingRel && (this.model.sheets || []).find(sheet => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(iterator, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

module.exports = WorkbookReader;
