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
  constructor(options) {
    super();

    this.options = options = options || {};

    this.styles = new StyleManager();
    this.styles.init();

    this.properties = new WorkbookPropertiesManager();

    // worksheet readers, indexed by sheetNo
    this.worksheetReaders = {};

    // hyperlink readers, indexed by sheetNo
    this.hyperlinkReaders = {};

    // worksheets, deferred for parsing after shared strings reading
    this.waitingWorkSheets = [];

    // callbacks for temp files cleanup
    this.tempFileCleanupCallbacks = [];
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
      for await (const {eventType, value, entry} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            for (const val of value) {
              this.emit(eventType, val);
            }
            break;
          case 'worksheet': {
            this.emit(eventType, value);
            await value.read(entry, options);
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

  async *parse(input, options) {
    const stream = (this.stream = this._getStream(input));
    const zip = (this.zip = unzip.Parse({forceStream: true}));
    // TODO: Remove once node v8 is deprecated
    // Detect and upgrade old streams
    let pipedStream = stream.pipe(zip);
    if (!pipedStream[Symbol.asyncIterator]) {
      pipedStream = pipedStream.pipe(new PassThrough({writableObjectMode: true, readableObjectMode: true}));
    }
    for await (const entry of pipedStream) {
      let match;
      let sheetNo;
      switch (entry.path) {
        case '_rels/.rels':
        case 'xl/_rels/workbook.xml.rels':
          entry.autodrain();
          break;
        case 'xl/workbook.xml':
          await this._parseWorkbook(entry, options);
          break;
        case 'xl/sharedStrings.xml':
          yield* this._parseSharedStrings(entry, options);
          break;
        case 'xl/styles.xml':
          await this._parseStyles(entry, options);
          break;
        default:
          if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
            match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
            sheetNo = match[1];
            if (this.sharedStrings) {
              yield* this._parseWorksheet(entry, sheetNo, options);
            } else {
              // create temp file for each worksheet
              await new Promise((resolve, reject) => {
                tmp.file((err, path, fd, cleanupCallback) => {
                  if (err) {
                    return reject(err);
                  }

                  const tempStream = fs.createWriteStream(path);

                  this.waitingWorkSheets.push({sheetNo, path});
                  entry.pipe(tempStream);

                  this.tempFileCleanupCallbacks.push(cleanupCallback);

                  return tempStream.on('finish', () => {
                    return resolve();
                  });
                });
              });
            }
          } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
            match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
            sheetNo = match[1];
            yield* this._parseHyperlinks(entry, sheetNo, options);
          } else {
            entry.autodrain();
          }
          break;
      }
    }

    for (const {sheetNo, path} of this.waitingWorkSheets) {
      let entry = fs.createReadStream(path);
      // TODO: Remove once node v8 is deprecated
      // Detect and upgrade old streams
      if (!entry[Symbol.asyncIterator]) {
        entry = entry.pipe(new PassThrough());
      }
      yield* this._parseWorksheet(entry, sheetNo, options);
    }
    for (const tempFileCleanupCallback of this.tempFileCleanupCallbacks) {
      tempFileCleanupCallback();
    }
    this.tempFileCleanupCallbacks = [];
  }

  _emitEntry(options, payload) {
    if (options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseWorkbook(entry, options) {
    this._emitEntry(options, {type: 'workbook'});
    await this.properties.parseStream(entry);
  }

  async *_parseSharedStrings(entry, options) {
    this._emitEntry(options, {type: 'shared-strings'});
    let sharedStrings = null;
    switch (options.sharedStrings) {
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
            if (options.sharedStrings === 'cache') {
              sharedStrings.push(text);
            } else {
              sharedStringEvents.push({index: index++, text});
            }
            text = null;
          }
        }
      }
      if (options.sharedStrings === 'emit') {
        yield {eventType: 'shared-string', value: sharedStringEvents, entry};
      }
    }
  }

  async _parseStyles(entry, options) {
    this._emitEntry(options, {type: 'styles'});
    if (options.styles !== 'cache') {
      entry.autodrain();
      return;
    }
    this.styles = new StyleManager();
    await this.styles.parseStream(entry);
  }

  _getReader(Type, collection, sheetNo) {
    let reader = collection[sheetNo];
    if (!reader) {
      reader = new Type(this, sheetNo);
      collection[sheetNo] = reader;
    }
    return reader;
  }

  *_parseWorksheet(entry, sheetNo, options) {
    this._emitEntry(options, {type: 'worksheet', id: sheetNo});
    const worksheetReader = this._getReader(WorksheetReader, this.worksheetReaders, sheetNo);
    if (options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader, entry};
    }
  }

  *_parseHyperlinks(entry, sheetNo, options) {
    this._emitEntry(options, {type: 'hyerlinks', id: sheetNo});
    const hyperlinksReader = this._getReader(HyperlinkReader, this.hyperlinkReaders, sheetNo);
    if (options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader, entry};
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
