const fs = require('fs');
const {EventEmitter} = require('events');
const {PassThrough, Readable} = require('readable-stream');
const nodeStream = require('stream');
const unzip = require('unzipper');
const tmp = require('tmp');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookPropertiesManager = require('../../xlsx/xform/book/workbook-properties-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

tmp.setGracefulCleanup();

class WorkbookReader extends EventEmitter {
  constructor(input, options) {
    super();

    this.input = input;

    this.options = options || {
      hyperlinks: 'cache',
      sharedStrings: 'cache',
      worksheets: 'emit',
    };

    this.styles = new StyleManager();
    this.styles.init();

    this.properties = new WorkbookPropertiesManager();
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
            for (const val of value) {
              this.emit(eventType, val);
            }
            break;
          case 'worksheet': {
            this.emit(eventType, value);
            await value.read();
            break;
          }
          case 'hyperlinks': {
            this.emit(eventType, value);
            break;
          }
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
    const zip = (this.zip = unzip.Parse({forceStream: true}));
    // TODO: Remove once node v8 is deprecated
    // Detect and upgrade old streams
    let pipedZipStream = stream.pipe(zip);
    if (!pipedZipStream[Symbol.asyncIterator]) {
      pipedZipStream = pipedZipStream.pipe(new PassThrough({writableObjectMode: true, readableObjectMode: true}));
    }

    // worksheets, deferred for parsing after shared strings reading
    const waitingWorkSheets = [];

    for await (const entry of pipedZipStream) {
      let match;
      let sheetNo;
      switch (entry.path) {
        case '_rels/.rels':
        case 'xl/_rels/workbook.xml.rels':
          entry.autodrain();
          break;
        case 'xl/workbook.xml':
          await this._parseWorkbook(entry);
          break;
        case 'xl/sharedStrings.xml':
          yield* this._parseSharedStrings(entry);
          break;
        case 'xl/styles.xml':
          await this._parseStyles(entry);
          break;
        default:
          if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
            match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
            sheetNo = match[1];
            if (this.sharedStrings) {
              yield* this._parseWorksheet(entry, sheetNo);
            } else {
              // create temp file for each worksheet
              await new Promise((resolve, reject) => {
                tmp.file((err, path, fd, tempFileCleanupCallback) => {
                  if (err) {
                    return reject(err);
                  }
                  waitingWorkSheets.push({sheetNo, path, tempFileCleanupCallback});

                  const tempStream = fs.createWriteStream(path);
                  entry.pipe(tempStream);
                  return tempStream.on('finish', () => {
                    return resolve();
                  });
                });
              });
            }
          } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
            match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
            sheetNo = match[1];
            yield* this._parseHyperlinks(entry, sheetNo);
          } else {
            entry.autodrain();
          }
          break;
      }
    }

    for (const {sheetNo, path, tempFileCleanupCallback} of waitingWorkSheets) {
      let entry = fs.createReadStream(path);
      // TODO: Remove once node v8 is deprecated
      // Detect and upgrade old streams
      if (!entry[Symbol.asyncIterator]) {
        entry = entry.pipe(new PassThrough());
      }
      yield* this._parseWorksheet(entry, sheetNo);
      tempFileCleanupCallback();
    }
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseWorkbook(entry) {
    this._emitEntry({type: 'workbook'});
    await this.properties.parseStream(entry);
  }

  async *_parseSharedStrings(entry) {
    this._emitEntry({type: 'shared-strings'});
    let sharedStrings = null;
    switch (this.options.sharedStrings) {
      case 'cache':
        sharedStrings = this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        entry.autodrain();
        return;
    }

    let inT = false;
    let text = null;
    let index = 0;
    for await (const events of parseSax(entry)) {
      const sharedStringEvents = [];
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          const node = value;
          if (node.name === 't') {
            text = null;
            inT = true;
          }
        } else if (eventType === 'text') {
          text = text ? text + value : value;
        } else if (eventType === 'closetag') {
          const node = value;
          if (inT && node.name === 't') {
            if (this.options.sharedStrings === 'cache') {
              sharedStrings.push(text);
            } else {
              sharedStringEvents.push({index: index++, text});
            }
            text = null;
          }
        }
      }
      if (this.options.sharedStrings === 'emit') {
        yield {eventType: 'shared-string', value: sharedStringEvents};
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({type: 'styles'});
    if (this.options.styles !== 'cache') {
      entry.autodrain();
      return;
    }
    this.styles = new StyleManager();
    await this.styles.parseStream(entry);
  }

  *_parseWorksheet(entry, sheetNo) {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({workbook: this, id: sheetNo, entry, options: this.options});
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(entry, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({workbook: this, id: sheetNo, entry, options: this.options});
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  entries: ['emit'],
  sharedStrings: ['cache', 'emit'],
  styles: ['cache'],
  hyperlinks: ['cache', 'emit'],
  worksheets: ['emit'],
};

module.exports = WorkbookReader;
