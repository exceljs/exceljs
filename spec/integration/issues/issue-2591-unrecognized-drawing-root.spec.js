const ExcelJS = verquire('exceljs');

// Both fixtures below exercise the fix for
// https://github.com/exceljs/exceljs/issues/2591:
//
// - `test-issue-2591-chart-user-shapes.xlsx` reproduces the stack trace
//   from the report verbatim — a workbook with a chart whose
//   `<c:userShapes>` overlay drawing lands at `xl/drawings/*.xml`, parses
//   to a null model, and then crashes `XLSX.reconcile` on
//   `(drawing.anchors || []).forEach`.
//
// - `test-issue-2591-default-namespace.xlsx` is a separate real-world
//   workbook that hits the same downstream null-drawing guards via a
//   different XML trigger: its drawing root is `<wsDr>` without the
//   canonical `xdr:` prefix, which DrawingXform doesn't recognise either.
//   The original #2591 report doesn't describe this variant, but the same
//   fix covers it.
describe('github issues', () => {
  it('issue 2591 - tolerates a chart user shapes overlay drawing', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile(
        './spec/integration/data/test-issue-2591-chart-user-shapes.xlsx'
      )
      .then(() => {
        const [sheet] = wb.worksheets;
        expect(sheet.name).to.equal('Employees');
        expect(sheet.getCell('A1').value).to.equal('Employee ID');
        expect(sheet.getCell('A2').value).to.equal('123');
        expect(sheet.getCell('B2').value).to.equal('Alex');
      });
  });

  it('issue 2591 - tolerates a drawing that uses the default namespace instead of xdr:', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile(
        './spec/integration/data/test-issue-2591-default-namespace.xlsx'
      )
      .then(() => {
        const [sheet] = wb.worksheets;
        expect(sheet.name).to.equal('Employees');
        expect(sheet.getCell('A1').value).to.equal('Employee ID');
        expect(sheet.getCell('A2').value).to.equal('123');
        expect(sheet.getCell('B2').value).to.equal('Alex');
      });
  });
});
