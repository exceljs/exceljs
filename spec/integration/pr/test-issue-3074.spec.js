const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 3074 - streaming reader should keep an error result on a formula cell', async () => {
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: './test-issue-3074.xlsx',
    });
    const sheet = workbook.addWorksheet('data');
    sheet.getCell('A1').value = {formula: 'C1/D1', result: {error: '#REF!'}};
    sheet.commit();
    await workbook.commit();

    return new Promise((resolve, reject) => {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        './test-issue-3074.xlsx',
        {
          entries: 'emit',
          sharedStrings: 'cache',
          styles: 'cache',
          worksheets: 'emit',
        }
      );

      workbookReader.on('worksheet', worksheet =>
        worksheet.on('row', row => {
          expect(row.getCell(1).value).to.eql({
            formula: 'C1/D1',
            result: {error: '#REF!'},
          });
          resolve();
        })
      );
      workbookReader.on('error', reject);

      workbookReader.read();
    });
  });
});
