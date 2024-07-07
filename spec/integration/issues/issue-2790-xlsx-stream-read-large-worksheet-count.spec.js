const {createReadStream} = require('fs');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  let worsheetCount = 0;
  before(
    () =>
      new Promise((resolve, reject) => {
        const stream = createReadStream(
          './spec/integration/data/WORKBOOK_WITH_188_SHEETS.xlsx'
        );
        const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(stream, {
          worksheets: 'emit',
        });
        workbookReader.read();
        workbookReader.on('worksheet', () => (worsheetCount += 1));
        workbookReader.on('end', resolve);
        workbookReader.on('error', reject);
      })
  );

  it('will read all worksheets', async () => {
    const ACTUAL_WORKSHEET_COUNT = 188;
    expect(worsheetCount).to.equal(ACTUAL_WORKSHEET_COUNT);
  });
});
