const {PassThrough} = require('stream');
const tmp = require('tmp');

const ExcelJS = verquire('exceljs');

// A workbook written by ExcelJS stores its worksheets before sharedStrings.xml and
// workbook.xml, so the streaming reader defers each worksheet to a temp file. If
// creating that temp file is slower than unzipping the rest of the archive, the unzip
// stream emits 'end' while the remaining entries are still buffered on it.
const TEMP_FILE_DELAY_MS = 50;

describe('WorkbookReader', () => {
  describe('when a worksheet is deferred to a slowly created temp file', () => {
    let originalTmpFile;

    beforeEach(() => {
      originalTmpFile = tmp.file;
      tmp.file = (...args) => {
        const callback = args.pop();
        originalTmpFile(...args, (...result) =>
          setTimeout(() => callback(...result), TEMP_FILE_DELAY_MS)
        );
      };
    });

    afterEach(() => {
      tmp.file = originalTmpFile;
    });

    it('still reads the shared strings and the workbook that follow it', async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet With Strings');
      worksheet.getCell('A1').value = 'first';
      worksheet.getCell('B1').value = 42;
      worksheet.getCell('A2').value = 'second';
      const buffer = await workbook.xlsx.writeBuffer();

      const input = new PassThrough();
      input.end(Buffer.from(buffer));
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(input, {
        sharedStrings: 'cache',
        styles: 'ignore',
        hyperlinks: 'ignore',
      });

      const names = [];
      const rows = [];
      for await (const worksheetReader of workbookReader) {
        names.push(worksheetReader.name);
        for await (const row of worksheetReader) {
          rows.push(row.values.slice(1));
        }
      }

      expect(names).to.deep.equal(['Sheet With Strings']);
      expect(rows).to.deep.equal([['first', 42], ['second']]);
    });
  });
});
