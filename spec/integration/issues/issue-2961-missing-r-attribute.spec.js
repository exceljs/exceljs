const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 2961 - Invalid row number in model when r attribute missing', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-2961.xlsx')
      .then(() => {
        const sheet = wb.getWorksheet(1);

        // Test that cells are parsed correctly even without r attributes
        expect(sheet.getCell('A1').text).to.equal('id');
        expect(sheet.getCell('B1').text).to.equal('orderCode');
        expect(sheet.getCell('C1').text).to.equal('orderNumber');

        expect(sheet.getCell('A2').text).to.equal('1');
        expect(sheet.getCell('B2').text).to.equal('IM20220302624042440003');
        expect(sheet.getCell('C2').text).to.equal('lp002');

        const sheet2 = wb.getWorksheet(2);
        expect(sheet2.getCell('A1').text).to.match(/^SELECT/);
      });
  });
});
