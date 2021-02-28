const BlobZipStream = require('../../utils/blob-zip-stream');
const iterateStream = require('../../utils/iterate-stream');
const AbstractWorkbookReader = require('./abstract-workbook-reader');

class WorkbookReaderBlob extends AbstractWorkbookReader {
  constructor(input, options = {}) {
    super(input, options);
    if (typeof FileReader === 'undefined') {
      throw new TypeError('FileReader is required');
    }
  }

  _getZipEntryStream(zip, entity) {
    return new Promise((resolve, reject) => {
      zip.stream(entity, (err, stream) => {
        if (err) {
          reject(err);
        } else {
          resolve(stream);
        }
      });
    });
  }

  async _getZipStream(input) {
    // eslint-disable-next-line no-undef
    if (!(input instanceof Blob)) {
      throw new Error(`Could not recognise input: ${input}`);
    }

    // Note: jszip would read the complete blob into memory
    // since blob is seekable, we can read is by part
    const zip = new BlobZipStream({file: this.input});

    // wait for zip central directory parse complete
    await new Promise((resolve, reject) => {
      function ready() {
        zip.removeListener('error', error);
        zip.removeListener('ready', ready);
        resolve();
      }

      function error(err) {
        zip.removeListener('ready', ready);
        zip.removeListener('error', error);
        reject(err);
      }

      zip.once('ready', ready);
      zip.once('error', error);
    });

    return zip;
  }

  async* [Symbol.asyncIterator]() {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input, options) {
    if (options) this.options = options;
    const zip = await this._getZipStream(this.input);
    const entities = zip.entries();

    // parse the shared strings before parsing worksheets
    if (entities['xl/sharedStrings.xml']) {
      yield* this._parseSharedStrings(
              await this._getZipEntryStream(zip, entities['xl/sharedStrings.xml']));
    }

    // parse the workbook before parsing worksheets to get worksheet name
    // https://github.com/exceljs/exceljs/pull/1478
    if (entities['xl/workbook.xml']) {
      await this._parseWorkbook(
              await this._getZipEntryStream(zip, entities['xl/workbook.xml']));
    }

    // parse the workbook rels before parsing worksheets to get worksheet name
    if (entities['xl/_rels/workbook.xml.rels']) {
      await this._parseRels(
              await this._getZipEntryStream(zip, entities['xl/_rels/workbook.xml.rels']));
    }

    // parse anything else
    for (const key of Object.keys(entities)) {
      const entry = entities[key];
      if (!entry.isFile) continue;
      let match;
      let sheetNo;
      let stream;
      switch (key) {
        case 'xl/sharedStrings.xml':
        case 'xl/_rels/workbook.xml.rels':
        case 'xl/workbook.xml':
          // already parsed
          break;
        case '_rels/.rels':
          if (this.workbookRels) {
            break;
          }
          // eslint-disable-next-line no-await-in-loop
          stream = await this._getZipEntryStream(zip, entry);
          // eslint-disable-next-line no-await-in-loop
          await this._parseRels(stream);
          break;
        case 'xl/styles.xml':
          // eslint-disable-next-line no-await-in-loop
          stream = await this._getZipEntryStream(zip, entry);
          // eslint-disable-next-line no-await-in-loop
          await this._parseStyles(stream);
          break;
        default:
          if (key.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
            match = key.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
            sheetNo = match[1];
            // eslint-disable-next-line no-await-in-loop
            stream = await this._getZipEntryStream(zip, entry);
            yield* this._parseWorksheet(iterateStream(stream), sheetNo);
          } else if (key.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
            match = key.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
            sheetNo = match[1];
            // eslint-disable-next-line no-await-in-loop
            stream = await this._getZipEntryStream(zip, entry);
            yield* this._parseHyperlinks(iterateStream(stream), sheetNo);
          }
          break;
      }

    }
    zip.close();
  }

}

// for reference - these are the valid values for options
WorkbookReaderBlob.Options = AbstractWorkbookReader.Options;

module.exports = WorkbookReaderBlob;
