const fs = require('fs');
const {EventEmitter} = require('events');
const {PassThrough, Readable} = require('readable-stream');
const nodeStream = require('stream');
const unzip = require('unzipper');
const tmp = require('tmp');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookPropertiesManager = require('../../xlsx/xform/book/workbook-properties-xform');

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

    // worksheets, deferred for parsing after shared strings reading
    const waitingWorkSheets = [];

    for await (const entry of iterateStream(zip)) {
      let match;
      let sheetNo;
      switch (entry.path) {
        case '_rels/.rels':
        case 'xl/_rels/workbook.xml.rels':
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
              yield* this._parseWorksheet(iterateStream(entry), sheetNo);
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
            yield* this._parseHyperlinks(iterateStream(entry), sheetNo);
          }
          break;
      }
      entry.autodrain();
    }

    for (const {sheetNo, path, tempFileCleanupCallback} of waitingWorkSheets) {
      let fileStream = fs.createReadStream(path);
      // TODO: Remove once node v8 is deprecated
      // Detect and upgrade old fileStreams
      if (!fileStream[Symbol.asyncIterator]) {
        fileStream = fileStream.pipe(new PassThrough());
      }
      yield* this._parseWorksheet(fileStream, sheetNo);
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
    await this.properties.parseStream(iterateStream(entry));
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

    let inT = false;
    let text = null;
    let index = 0;
    for await (const events of parseSax(iterateStream(entry))) {
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
              this.sharedStrings.push(text);
            } else if (this.options.sharedStrings === 'emit') {
              yield {index: index++, text};
            }
            text = null;
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
    const worksheetReader = new WorksheetReader({workbook: this, id: sheetNo, iterator, options: this.options});
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(iterator, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({workbook: this, id: sheetNo, iterator, options: this.options});
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
