const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/test-issue-1751.xlsx';

describe('github issues', () => {
  it('issue 1751 - xm:f in cfRule', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/test-issue-1751.xlsx');

    checkFormattings(wb.worksheets[0]);

    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

    const wbNew = new ExcelJS.Workbook();
    await wbNew.xlsx.readFile(TEST_XLSX_FILE_NAME);

    checkFormattings(wbNew.worksheets[0]);
  }).timeout(6000);
});

function checkFormattings(ws) {
  const cfs = ws.conditionalFormattings;
  expect(cfs).to.have.lengthOf(2);
  const cfB = cfs[0];
  const cfC = cfs[1];

  expect(cfB.ref).to.equal('B1');
  expect(cfB.rules).to.have.lengthOf(1);
  expect(cfB.rules[0].type).to.equal('expression');
  expect(cfB.rules[0].formulae).to.deep.equal(['A1=1']);

  expect(cfC.ref).to.equal('C1');
  expect(cfC.rules).to.have.lengthOf(1);
  expect(cfC.rules[0].type).to.equal('expression');
  expect(cfC.rules[0].formulae).to.deep.equal(['A1<>0']);
}
