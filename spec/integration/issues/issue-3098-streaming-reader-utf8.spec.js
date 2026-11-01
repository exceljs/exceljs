const {Readable} = require('stream');

const ExcelJS = verquire('exceljs');

class ChunkedReadable extends Readable {
  constructor(buffer) {
    super();
    this.buffer = buffer;
    this.offset = 0;
  }

  _read() {
    if (this.offset < this.buffer.length) {
      this.push(this.buffer.slice(this.offset, ++this.offset));
    } else {
      this.push(null);
    }
  }
}

describe('github issues', () => {
  it('issue 3098 - preserves UTF-8 text split across streaming chunks', async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet('Sheet 1').getCell('A1').value = '中文';
    const buffer = await workbook.xlsx.writeBuffer({useSharedStrings: false});
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(
      new ChunkedReadable(buffer),
      {
        worksheets: 'emit',
        sharedStrings: 'cache',
      }
    );
    const values = [];

    for await (const worksheet of reader) {
      for await (const row of worksheet) {
        values.push(row.values[1]);
      }
    }

    expect(values).to.deep.equal(['中文']);
  });
});
