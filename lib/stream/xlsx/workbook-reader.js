const fs = require('fs');
const {EventEmitter} = require('events');
const {PassThrough, Readable} = require('readable-stream');
const nodeStream = require('stream');
const unzip = require('@zklogic/unzipper');
const tmp = require('tmp');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
const RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

tmp.setGracefulCleanup();

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

  async *parse(input, options) {
    if (options) this.options = options;
    const stream = (this.stream = this._getStream(input || this.input));
    const zip = unzip.Parse({forceStream: true});
    stream.pipe(zip);

    // Collect all entries first (fully consuming the ZIP), then process in-order
    // so sharedStrings/styles/rels are available before worksheet parsing.
    const allEntries = [];
    const entryPromises = [];
    await new Promise((resolve, reject) => {
      const handleEntry = entry => {
        if (!entry || !entry.path) {
          return;
        }
        const entryRecord = {path: entry.path, entry, buffer: null};
        allEntries.push(entryRecord);

        // Must consume each entry immediately so unzipper advances.
        const chunks = [];
        const entryPromise = new Promise((entryResolve, entryReject) => {
          entry.on('data', chunk => chunks.push(chunk));
          entry.on('end', () => {
            entryRecord.buffer = Buffer.concat(chunks);
            entryResolve();
          });
          entry.on('error', entryReject);
        });
        entryPromises.push(entryPromise);
      };

      zip.on('data', handleEntry);
      zip.on('error', reject);
      zip.on('close', resolve);
    });
    await Promise.all(entryPromises);

    // Phase 2: process entries in zip order now that all data is buffered
    const makeReadable = buf => {
      const r = new Readable();
      r.push(buf);
      r.push(null);
      return r;
    };
    const waitingWorkSheets = [];
    for (const {path: entryPath, buffer} of allEntries) {
      let match;
      let sheetNo;
      switch (entryPath) {
        case '_rels/.rels':
          break;
        case 'xl/_rels/workbook.xml.rels':
          // eslint-disable-next-line no-await-in-loop
          await this._parseRels(makeReadable(buffer));
          break;
        case 'xl/workbook.xml':
          // eslint-disable-next-line no-await-in-loop
          await this._parseWorkbook(makeReadable(buffer));
          break;
        case 'xl/sharedStrings.xml':
          // eslint-disable-next-line no-await-in-loop
          await this._drainSharedStrings(makeReadable(buffer));
          break;
        case 'xl/styles.xml':
          // eslint-disable-next-line no-await-in-loop
          await this._parseStyles(makeReadable(buffer));
          break;
        default:
          if (entryPath.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
            match = entryPath.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
            sheetNo = match[1];
            waitingWorkSheets.push({sheetNo, buffer});
          } else if (entryPath.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
            match = entryPath.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
            sheetNo = match[1];
            yield* this._parseHyperlinks(iterateStream(makeReadable(buffer)), sheetNo);
          }
          break;
      }
    }

    for (const {sheetNo, buffer} of waitingWorkSheets) {
      let fileStream = makeReadable(buffer);
      if (!fileStream[Symbol.asyncIterator]) {
        fileStream = fileStream.pipe(new PassThrough());
      }
      yield* this._parseWorksheet(iterateStream(fileStream), sheetNo);
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
    // Emit a final marker to ensure cache mode parsing completes
    if (this.options.sharedStrings === 'cache' && this.sharedStrings) {
      yield {eventType: 'shared-strings-cached', count: this.sharedStrings.length};
    }
  }

  // Fully drive _parseSharedStrings to completion as a plain async fn,
  // ensuring this.sharedStrings is populated before any awaiting worksheet
  // is processed — even when sharedStrings:'cache' yields nothing externally.
  async _drainSharedStrings(entry) {
    const gen = this._parseSharedStrings(entry);
    let result = await gen.next();
    while (!result.done) {
      // eslint-disable-next-line no-await-in-loop
      result = await gen.next();
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
    const matchingSheet =
      matchingRel && ((this.model && this.model.sheets) || []).find(sheet => sheet.rId === matchingRel.Id);
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
