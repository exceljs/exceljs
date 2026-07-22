const {PassThrough} = require('stream');

const ExcelJS = verquire('exceljs');

function bufferToStream(buffer) {
  const stream = new PassThrough();
  stream.end(Buffer.from(buffer));
  return stream;
}

// Regression: the streaming WorkbookReader read every part from a single streaming
// unzip pass. That pass emits entries in stored order and, on Node >= 18, frequently
// lost the final entry under async iteration. Because xl/workbook.xml is written last,
// this.model was often never set and _parseWorksheet threw
// "Cannot read properties of undefined (reading 'sheets')" (~90% of reads).
describe('WorkbookReader - streaming resolves the workbook model', () => {
  async function buildBuffer() {
    const wb = new ExcelJS.Workbook();
    const s1 = wb.addWorksheet('First');
    s1.addRow(['id', 'name']);
    s1.addRow([1, 'Alpha']);
    const s2 = wb.addWorksheet('Second');
    s2.addRow(['id', 'name']);
    s2.addRow([2, 'Beta']);
    s2.addRow([3, 'Gamma']);
    return wb.xlsx.writeBuffer();
  }

  async function readSheets(buffer) {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(
      bufferToStream(buffer),
      {
        worksheets: 'emit',
        sharedStrings: 'cache',
        entries: 'ignore',
      }
    );
    const sheets = [];
    for await (const worksheet of reader) {
      const rows = [];
      for await (const row of worksheet) {
        rows.push(row.values.slice(1).map(value => `${value}`));
      }
      sheets.push({name: worksheet.name, rows});
    }
    return sheets;
  }

  it('exposes sheet names and rows on every read', async function() {
    this.timeout(20000);
    const buffer = await buildBuffer();

    // The failure was non-deterministic (~90% per read), so repeat the read to keep
    // a regression from slipping through as an occasional pass.
    for (let i = 0; i < 15; i += 1) {
      // eslint-disable-next-line no-await-in-loop
      const sheets = await readSheets(buffer);
      expect(sheets.map(sheet => sheet.name)).to.deep.equal([
        'First',
        'Second',
      ]);
      expect(sheets[0].rows).to.deep.equal([
        ['id', 'name'],
        ['1', 'Alpha'],
      ]);
      expect(sheets[1].rows).to.deep.equal([
        ['id', 'name'],
        ['2', 'Beta'],
        ['3', 'Gamma'],
      ]);
    }
  });
});
