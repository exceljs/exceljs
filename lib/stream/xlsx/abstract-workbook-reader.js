const {EventEmitter} = require('events');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
const RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

class AbstractWorkbookReader extends EventEmitter {

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
AbstractWorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

module.exports = AbstractWorkbookReader;
