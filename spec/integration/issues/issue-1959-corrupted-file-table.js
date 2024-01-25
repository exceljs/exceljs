const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('issue 1959 - Corrupted file generating by reading and modifing an existing XLSX file with tables', () => {
    it('Reading test-issue-1669.xlsx, changing reference and writing test-issue-1959.xlsx', () => {
      const wb = new ExcelJS.Workbook();
      return wb.xlsx
        .readFile('./spec/integration/data/test-issue-1959.xlsx')
        .then(async () => {
          const ws = wb.getWorksheet('Sheet1');
          // Table 1
          const table1 = ws.getTable('Table1');
          expect(table1.ref).to.equal('B12:D17');
          table1.ref = 'F21:H26';
          table1.commit();
          expect(table1.ref).to.equal('F21:H26');
          // Table 13
          const table13 = ws.getTable('Table13');
          expect(table13.ref).to.equal('B21:D24');

          await wb.xlsx.writeFile('./spec/out/test-issue-1959.xlsx');
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx
            .readFile('./spec/out/test-issue-1959.xlsx')
            .then(() => {
              const worksheet = wb.getWorksheet('Sheet1');
              expect(worksheet.getTable('Table1').ref).to.equal('F21:H26');
              expect(worksheet.getTable('Table13').ref).to.equal('B21:D24');
            });
        });
    });
  });
});
