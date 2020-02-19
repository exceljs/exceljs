const fs = require('fs');
const {EventEmitter} = require('events');
const Stream = require('stream');
const unzip = require('unzipper');
const tmp = require('tmp');
const SAXStream = require('../../utils/sax-stream');

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

    // count the open readers
    this.readers = 0;

    // end of stream check
    this.atEnd = false;

    // worksheets, deferred for parsing after shared strings reading
    this.waitingWorkSheets = [];

    // callbacks for temp files cleanup
    this.tempFileCleanupCallbacks = [];
  }

  _getStream(input) {
    if (input instanceof Stream.Readable) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error('Could not recognise input');
  }

  async read(input, options) {
    try {
      const stream = (this.stream = this._getStream(input));
      const zip = (this.zip = unzip.Parse());
      const passThrough = new Stream.PassThrough({objectMode: true});
      zip.pipe(passThrough);
      stream.pipe(zip);
      for await (const entry of passThrough) {
        let match;
        let sheetNo;
        switch (entry.path) {
          case '_rels/.rels':
          case 'xl/_rels/workbook.xml.rels':
            entry.autodrain();
            break;
          case 'xl/workbook.xml':
            this._parseWorkbook(entry, options);
            break;
          case 'xl/sharedStrings.xml':
            this._parseSharedStrings(entry, options);
            break;
          case 'xl/styles.xml':
            this._parseStyles(entry, options);
            break;
          default:
            if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
              match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
              sheetNo = match[1];
              if (this.sharedStrings) {
                this._parseWorksheet(entry, sheetNo, options);
              } else {
                // create temp file for each worksheet
                await new Promise((resolve, reject) => {
                  tmp.file((err, path, fd, cleanupCallback) => {
                    if (err) {
                      return reject(err);
                    }

                    const tempStream = fs.createWriteStream(path);

                    this.waitingWorkSheets.push({sheetNo, options, path});
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
              this._parseHyperlinks(entry, sheetNo, options);
            } else {
              entry.autodrain();
            }
            break;
        }
      }

      if (this.waitingWorkSheets.length) {
        let currentBook = 0;

        const processBooks = () => {
          const worksheetInfo = this.waitingWorkSheets[currentBook];
          const entry = fs.createReadStream(worksheetInfo.path);

          const {sheetNo} = worksheetInfo;
          const worksheet = this._parseWorksheet(entry, sheetNo, worksheetInfo.options);

          worksheet.on('finished', () => {
            ++currentBook;
            if (currentBook === this.waitingWorkSheets.length) {
              // temp files cleaning up
              this.tempFileCleanupCallbacks.forEach(cb => {
                cb();
              });

              this.tempFileCleanupCallbacks = [];

              this.emit('end');
              this.atEnd = true;
              if (!this.readers) {
                this.emit('finished');
              }
            } else {
              setImmediate(processBooks);
            }
          });
        };
        setImmediate(processBooks);
      } else {
        this.emit('end');
        this.atEnd = true;
        if (!this.readers) {
          this.emit('finished');
        }
      }
    } catch (error) {
      this.emit('error', error);
    }
  }

  _emitEntry(options, payload) {
    if (options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  _parseWorkbook(entry, options) {
    this._emitEntry(options, {type: 'workbook'});
    this.properties.parseStream(entry);
  }

  async _parseSharedStrings(entry, options) {
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
    let t = null;
    let index = 0;
    try {
      for await (const {event, value} of entry.pipe(new SAXStream())) {
        if (event === 'opentag') {
          const node = value;
          if (node.name === 't') {
            t = null;
            inT = true;
          }
        } else if (event === 'text') {
          t = t ? t + value : value;
        } else if (event === 'closetag') {
          const node = value;
          if (inT && node.name === 't') {
            if (sharedStrings) {
              sharedStrings.push(t);
            } else {
              this.emit('shared-string', {index: index++, text: t});
            }
            t = null;
          }
        }
      }
    } catch (error) {
      this.emit('error', error);
    }
  }

  _parseStyles(entry, options) {
    this._emitEntry(options, {type: 'styles'});
    if (options.styles !== 'cache') {
      entry.autodrain();
      return;
    }
    this.styles = new StyleManager();
    this.styles.parseStream(entry);
  }

  _getReader(Type, collection, sheetNo) {
    let reader = collection[sheetNo];
    if (!reader) {
      reader = new Type(this, sheetNo);
      this.readers++;
      reader.on('finished', () => {
        if (!--this.readers) {
          if (this.atEnd) {
            this.emit('finished');
          }
        }
      });
      collection[sheetNo] = reader;
    }
    return reader;
  }

  _parseWorksheet(entry, sheetNo, options) {
    this._emitEntry(options, {type: 'worksheet', id: sheetNo});
    const worksheetReader = this._getReader(WorksheetReader, this.worksheetReaders, sheetNo);
    if (options.worksheets === 'emit') {
      this.emit('worksheet', worksheetReader);
    }
    worksheetReader.read(entry, options, this.hyperlinkReaders[sheetNo]);
    return worksheetReader;
  }

  _parseHyperlinks(entry, sheetNo, options) {
    this._emitEntry(options, {type: 'hyerlinks', id: sheetNo});
    const hyperlinksReader = this._getReader(HyperlinkReader, this.hyperlinkReaders, sheetNo);
    if (options.hyperlinks === 'emit') {
      this.emit('hyperlinks', hyperlinksReader);
    }
    hyperlinksReader.read(entry, options);
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
