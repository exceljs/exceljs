const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('pull request 1262 - protect should not be undefined', async () => {
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: './test.xlsx',
    });

    const sheet = workbook.addWorksheet('data');
    const row = sheet.addRow(['readonly cell']);
    row.getCell(1).protection = {
      locked: true,
    };

    expect(sheet.protect).to.exist();
  });
});
