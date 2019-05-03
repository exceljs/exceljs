'use strict';

var fs = require('fs');
var events = require('events');
var Stream = require('stream');
var unzip = require('unzipper');
var Sax = require('sax');
var tmp = require('tmp');

var utils = require('../../utils/utils');
var FlowControl = require('../../utils/flow-control');

var StyleManager = require('../../xlsx/xform/style/styles-xform');
var WorkbookPropertiesManager = require('../../xlsx/xform/book/workbook-properties-xform');

var WorksheetReader = require('./worksheet-reader');
var HyperlinkReader = require('./hyperlink-reader');

tmp.setGracefulCleanup();

var WorkbookReader = module.exports = function(options) {
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
};

utils.inherits(WorkbookReader, events.EventEmitter, {
  _getStream: function(input) {
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
    worksheets: ['emit']
  },
  read: function(input, options) {
    var stream = this.stream = this._getStream(input);
    var zip = this.zip = unzip.Parse();

    zip.on('entry', entry => {
      var match, sheetNo;
      // console.log(entry.path);
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

                var tempStream = fs.createWriteStream(path);

                this.waitingWorkSheets.push({sheetNo: sheetNo, options: options, path: path});
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
          var currentBook = 0;

          var processBooks = () => {
              var worksheetInfo = this.waitingWorkSheets[currentBook];
              var entry = fs.createReadStream(worksheetInfo.path);

              var sheetNo = worksheetInfo.sheetNo;
              var options = worksheetInfo.options;
              var worksheet = this._parseWorksheet(entry, sheetNo, options);

              worksheet.on('finished', (node) => {
                  ++currentBook;
                  if (currentBook === this.waitingWorkSheets.length) {
                    // temp files cleaning up
                    this.tempFileCleanupCallbacks.forEach(function (cb) {
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
              })
          }
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
  _emitEntry: function(options, payload) {
    if (options.entries === 'emit') {
      this.emit('entry', payload);
    }
  },
  _parseWorkbook: function(entry, options) {
    this._emitEntry(options, {type: 'workbook'});
    this.properties.parseStream(entry);
  },
  _parseSharedStrings: function(entry, options) {
    this._emitEntry(options, {type: 'shared-strings'});
    var self = this;
    var sharedStrings = null;
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

    var parser = Sax.createStream(true, {});
    var inT = false;
    var t = null;
    var index = 0;
    parser.on('opentag', function(node) {
      if (node.name === 't') {
        t = null;
        inT = true;
      }
    });
    parser.on('closetag', function(name) {
      if (inT && (name === 't')) {
        if (sharedStrings) {
          sharedStrings.push(t);
        } else {
          self.emit('shared-string', { index: index++, text: t});
        }
        t = null;
      }
    });
    parser.on('text', function(text) {
      t = t ? t + text : text;
    });
    parser.on('error', function(error) {
      self.emit('error', error);
    });
    entry.pipe(parser);
  },
  _parseStyles: function(entry, options) {
    this._emitEntry(options, {type: 'styles'});
    if (options.styles !== 'cache') {
      entry.autodrain();
      return;
    }
    this.styles = new StyleManager();
    this.styles.parseStream(entry);
  },
  _getReader: function(Type, collection, sheetNo) {
    var self = this;
    var reader = collection[sheetNo];
    if (!reader) {
      reader = new Type(this, sheetNo);
      self.readers++;
      reader.on('finished', function() {
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
  _parseWorksheet: function(entry, sheetNo, options) {
    this._emitEntry(options, {type: 'worksheet', id: sheetNo});
    var worksheetReader = this._getReader(WorksheetReader, this.worksheetReaders, sheetNo);
    if (options.worksheets === 'emit') {
      this.emit('worksheet', worksheetReader);
    }
    worksheetReader.read(entry, options, this.hyperlinkReaders[sheetNo]);
    return worksheetReader;
  },
  _parseHyperlinks: function(entry, sheetNo, options) {
    this._emitEntry(options, {type: 'hyerlinks', id: sheetNo});
    var hyperlinksReader = this._getReader(HyperlinkReader, this.hyperlinkReaders, sheetNo);
    if (options.hyperlinks === 'emit') {
      this.emit('hyperlinks', hyperlinksReader);
    }
    hyperlinksReader.read(entry, options);
  }
});
