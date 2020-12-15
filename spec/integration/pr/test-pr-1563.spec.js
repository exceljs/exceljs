const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('col listed in different order in the xml', () => {
    it('should parse correctly', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-pr-9999.xlsx')
        .then(() => {
          expect(wb.worksheets[0].columns[0].width).to.equal(23.6640625);
          expect(wb.worksheets[0].columns[5].hidden).to.equal(false);
          expect(wb.worksheets[0].columns[6].hidden).to.equal(true);
        });
    });
  });
});
