const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('pull request 1743 - fix CSV reading large number', () => {
    it('Reading test-pr-1743.csv', () => {
      const wb = new ExcelJS.Workbook();
      return wb.csv
        .readFile('./spec/integration/data/test-pr-1743.csv')
        .then(ws => {
          expect(ws.getCell('A1').value).to.equal('A');
          expect(ws.getCell('A2').value).to.equal('B');
          expect(ws.getCell('A3').value).to.equal('Some123String');
          expect(ws.getCell('A4').value).to.equal(-1);
          expect(ws.getCell('A5').value).to.equal(0);
          expect(ws.getCell('A6').value).to.equal(1);
          expect(ws.getCell('A7').value).to.equal(100);
          expect(ws.getCell('A8').value).to.equal(1);
          expect(ws.getCell('A9').value).to.equal(-1);
          expect(ws.getCell('A10').value).to.equal(1.1);
          expect(ws.getCell('A11').value).to.equal(1.1);
          expect(ws.getCell('A12').value).to.equal(1.10101);
          expect(ws.getCell('A13').value).to.equal(0);
          expect(ws.getCell('A14').value).to.equal(1000000);
          expect(ws.getCell('A15').value).to.equal('56343416020533614003');
          expect(ws.getCell('A16').value).to.equal('56343416020533614003.0');
          expect(ws.getCell('A17').value).to.equal('56343416020533614003.1');
        });
    });
  });
});
