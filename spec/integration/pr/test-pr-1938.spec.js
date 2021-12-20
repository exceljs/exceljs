const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('pull request 1938 - Load table rows and ref', () => {
    it('Reading test-issue-1936.xlsx', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-pr-1936.xlsx')
        .then(() => {
          const ws = wb.getWorksheet('Sheet1');
          const table = ws.getTable('Table1');
          expect(table.model.ref).to.be.equal('C4');
          expect(table.model.rows.length).equal(8);
          expect(table.model.rows[3][2]).equal('4');
        });
    });
  });
});
