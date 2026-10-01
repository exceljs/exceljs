const {PassThrough} = require('stream');

const ExcelJS = verquire('exceljs');

describe('CSV', () => {
  ['constructor', 'toString', 'hasOwnProperty', '__proto__'].forEach(text => {
    it(`preserves ${text} as text when reading`, async () => {
      const workbook = new ExcelJS.Workbook();
      const stream = new PassThrough();
      const worksheetPromise = workbook.csv.read(stream);
      stream.end(`${text}\n`);

      const worksheet = await worksheetPromise;
      expect(worksheet.getCell('A1').value).to.equal(text);
    });
  });

  it('still recognizes supported boolean and error values', async () => {
    const workbook = new ExcelJS.Workbook();
    const stream = new PassThrough();
    const worksheetPromise = workbook.csv.read(stream);
    stream.end('true,false,#N/A,#REF!\n');

    const worksheet = await worksheetPromise;
    expect(worksheet.getRow(1).values.slice(1)).to.deep.equal([
      true,
      false,
      {error: '#N/A'},
      {error: '#REF!'},
    ]);
  });
});
