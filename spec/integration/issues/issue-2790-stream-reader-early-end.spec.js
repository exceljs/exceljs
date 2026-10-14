const fs = require('fs');
const {Readable} = require('stream');

const ExcelJS = verquire('exceljs');

// Makes every file write finish late, as on a slow or busy disk. The streaming
// reader writes worksheets to temporary files while it reads the zip, and
// unzipper emits 'end' as soon as its input is consumed, while the entries it
// parsed last are still buffered.
function slowFileWrites() {
  const {createWriteStream} = fs;
  fs.createWriteStream = function(...args) {
    const stream = createWriteStream.apply(fs, args);
    const on = stream.on.bind(stream);
    stream.on = (event, listener) =>
      event === 'finish'
        ? on(event, (...finishArgs) => setTimeout(() => listener(...finishArgs), 200))
        : on(event, listener);
    return stream;
  };
  return () => {
    fs.createWriteStream = createWriteStream;
  };
}

async function readWorksheetNames(buffer, sharedStrings) {
  const input = new Readable({read() {}});
  input.push(buffer);
  input.push(null);
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, {
    worksheets: 'emit',
    sharedStrings,
  });
  const names = [];
  for await (const worksheet of reader) {
    names.push(worksheet.name);
    // eslint-disable-next-line no-unused-vars
    for await (const row of worksheet) {
      // the rows are read to the end
    }
  }
  return names;
}

describe('github issues', () => {
  describe('issue 2790 - xlsx stream reader drops the entries buffered when unzipper ends', () => {
    let buffer;

    before(async () => {
      const workbook = new ExcelJS.Workbook();
      ['first', 'second', 'third'].forEach((name, i) => {
        workbook.addWorksheet(name).addRow([name, i]);
      });
      buffer = Buffer.from(await workbook.xlsx.writeBuffer());
    });

    ['cache', 'ignore'].forEach(sharedStrings => {
      it(`reads every worksheet when writing the temporary files is slow (sharedStrings: ${sharedStrings})`, async () => {
        const restore = slowFileWrites();
        try {
          expect(await readWorksheetNames(buffer, sharedStrings)).to.deep.equal([
            'first',
            'second',
            'third',
          ]);
        } finally {
          restore();
        }
      });
    });
  });
});
