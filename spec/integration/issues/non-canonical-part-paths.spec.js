const ExcelJS = verquire('exceljs');

// Several rel types in a worksheet point at parts the loader indexes by
// filename pattern (`xl/commentsN.xml`, `xl/drawings/vmlDrawingN.vml`,
// `xl/tables/tableN.xml`, …). When a producing tool stores those parts at a
// non-canonical path, the loader's regex misses them and the corresponding
// entry in `options.<bag>[rel.Target]` is undefined at reconcile time.
// WorkSheetXform now filters such dangling rels out before any consumer
// dereferences them, so the workbook reads without error.
//
// All three fixtures below were synthesized — the comments case was
// reconstructed from a stack trace of an in-the-wild crash; the table and
// vml-drawing cases are anticipatory hardening of the same bug shape. Each
// was built by writing a small workbook with exceljs and then post-processing
// the zip to rename one part to a non-canonical filename and patch the
// worksheet rels (and `[Content_Types].xml` for the table case) to match —
// yielding fully valid OOXML files that exceljs's filename-pattern loader
// can't index.
describe('github issues', () => {
  it('tolerates a worksheet rel that points at a non-canonical comments part', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-non-canonical-comments.xlsx')
      .then(() => {
        const [sheet] = wb.worksheets;
        expect(sheet.name).to.equal('Sheet1');
        expect(sheet.getCell('A1').value).to.equal('email');
        expect(sheet.getCell('B1').value).to.equal('firstName');
        expect(sheet.getCell('A2').value).to.equal('a@b.c');
        // The comments part lives at xl/sheet1_comments.xml; the loader's
        // regex misses it, so the note on A2 should not be readable.
        expect(sheet.getCell('A2').note).to.equal(undefined);
      });
  });

  it('tolerates a worksheet rel that points at a non-canonical table part', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-non-canonical-table.xlsx')
      .then(() => {
        const [sheet] = wb.worksheets;
        expect(sheet.name).to.equal('Sheet1');
        expect(sheet.getCell('A1').value).to.equal('Id');
        expect(sheet.getCell('B1').value).to.equal('Name');
        expect(sheet.getCell('A2').value).to.equal(1);
        expect(sheet.getCell('B3').value).to.equal('Bob');
        // The table part lives at xl/tables/customers.xml; the loader's
        // regex misses it, so no Table object should be reconstructed.
        expect(Object.keys(sheet.tables)).to.have.lengthOf(0);
      });
  });

  it('tolerates a worksheet rel that points at a non-canonical vmlDrawing part', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-non-canonical-vml-drawing.xlsx')
      .then(() => {
        const [sheet] = wb.worksheets;
        expect(sheet.name).to.equal('Sheet1');
        expect(sheet.getCell('A1').value).to.equal('hello');
        expect(sheet.getCell('A2').value).to.equal('world');
        // The comments part itself lives at the canonical path and loads
        // fine; only the vmlDrawing (the note's visual frame) is at a
        // non-canonical path, so the note text should still be readable.
        expect(sheet.getCell('A1').note).to.equal('a note');
      });
  });
});
