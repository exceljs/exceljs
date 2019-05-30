'use strict';

const fs = require('fs');
const events = require('events');
const Stream = require('stream');
const unzip = require('unzipper');
const Sax = require('sax');
const tmp = require('tmp');

const utils = require('../../utils/utils');
const FlowControl = require('../../utils/flow-control');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookPropertiesManager = require('../../xlsx/xform/book/workbook-properties-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

tmp.setGracefulCleanup();

const WorkbookReader = (module.exports = function(options) {
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
});

utils.inherits(WorkbookReader, events.EventEmitter, {
  _getStream(input) {
    if (input instanceof Stream.Readable) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error('Could not recognise input');
  },

  get flowControl() {
    if (!this._flowControl) {
      this._flowControl = new FlowControl(this.options);
    }
    return this._flowControl;
  },

  options: {
    entries: ['emit'],
    sharedStrings: ['cache', 'emit'],
    styles: ['cache'],
    hyperlinks: ['cache', 'emit'],
    worksheets: ['emit'],
  },
  read(input, options) {
    const stream = (this.stream = this._getStream(input));
    const zip = (this.zip = unzip.Parse());

    zip.on('entry', entry => {
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
              tmp.file((err, path, fd, cleanupCallback) => {
                if (err) throw err;

                const tempStream = fs.createWriteStream(path);

                this.waitingWorkSheets.push({ sheetNo, options, path });
                entry.pipe(tempStream);

                this.tempFileCleanupCallbacks.push(cleanupCallback);
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
    });

    zip.on('close', () => {
      if (this.waitingWorkSheets.length) {
        let currentBook = 0;

        const processBooks = () => {
          const worksheetInfo = this.waitingWorkSheets[currentBook];
          const entry = fs.createReadStream(worksheetInfo.path);

          const {sheetNo} = worksheetInfo;
          const worksheet = this._parseWorksheet(
            entry,
            sheetNo,
            worksheetInfo.options
          );

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
    });

    zip.on('error', err => {
      this.emit('error', err);
    });

    // Pipe stream into top flow-control
    // this.flowControl.pipe(zip);
    stream.pipe(zip);
  },
  _emitEntry(options, payload) {
    if (options.entries === 'emit') {
      this.emit('entry', payload);
    }
  },
  _parseWorkbook(entry, options) {
    this._emitEntry(options, { type: 'workbook' });
    this.properties.parseStream(entry);
  },
  _parseSharedStrings(entry, options) {
    this._emitEntry(options, { type: 'shared-strings' });
    const self = this;
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

    const parser = Sax.createStream(true, {});
    let inT = false;
    let t = null;
    let index = 0;
    parser.on('opentag', node => {
      if (node.name === 't') {
        t = null;
        inT = true;
      }
    });
    parser.on('closetag', name => {
      if (inT && name === 't') {
        if (sharedStrings) {
          sharedStrings.push(t);
        } else {
          self.emit('shared-string', { index: index++, text: t });
        }
        t = null;
      }
    });
    parser.on('text', text => {
      t = t ? t + text : text;
    });
    parser.on('error', error => {
      self.emit('error', error);
    });
    entry.pipe(parser);
  },
  _parseStyles(entry, options) {
    this._emitEntry(options, { type: 'styles' });
    if (options.styles !== 'cache') {
      entry.autodrain();
      return;
    }
    this.styles = new StyleManager();
    this.styles.parseStream(entry);
  },
  _getReader(Type, collection, sheetNo) {
    const self = this;
    let reader = collection[sheetNo];
    if (!reader) {
      reader = new Type(this, sheetNo);
      self.readers++;
      reader.on('finished', () => {
        if (!--self.readers) {
          if (self.atEnd) {
            self.emit('finished');
          }
        }
      });
      collection[sheetNo] = reader;
    }
    return reader;
  },
  _parseWorksheet(entry, sheetNo, options) {
    this._emitEntry(options, { type: 'worksheet', id: sheetNo });
    const worksheetReader = this._getReader(WorksheetReader, this.worksheetReaders, sheetNo);
    if (options.worksheets === 'emit') {
      this.emit('worksheet', worksheetReader);
    }
    worksheetReader.read(entry, options, this.hyperlinkReaders[sheetNo]);
    return worksheetReader;
  },
  _parseHyperlinks(entry, sheetNo, options) {
    this._emitEntry(options, { type: 'hyerlinks', id: sheetNo });
    const hyperlinksReader = this._getReader(HyperlinkReader, this.hyperlinkReaders, sheetNo);
    if (options.hyperlinks === 'emit') {
      this.emit('hyperlinks', hyperlinksReader);
    }
    hyperlinksReader.read(entry, options);
  },
});
