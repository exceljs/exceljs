const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('pull request 1936 - fix load of tables with computed columns', () => {
    it('Reading test-pr-1936.xlsx', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-pr-1936.xlsx')
        .then(() => {
          const ws = wb.getWorksheet('Sheet1');
          const table = ws.getTable('Table1');
          expect(table.model.columns.length).to.equal(3);
        });
    });
  });
});
