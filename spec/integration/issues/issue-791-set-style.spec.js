const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 791 - modify a cell style on existing excel file may affect other cells', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx
      .readFile('./spec/integration/data/test-issue-791.xlsx')
      .then(() => {
        wb.worksheets.forEach(sheet => {
          // supposed to change A1 to E10 font color
          expect(sheet.getCell('B2').style).to.be.equal(
            sheet.getCell('D5').style
          );
          const b2 = sheet.getCell('B2');
          b2.font = {...b2.font, color: {argb: 'FFFFAA00'}};
          expect(sheet.getCell('B2').style).to.be.equal(
            sheet.getCell('D5').style
          );

          // supposed to change C12 font color only

          expect(sheet.getCell('C12').style).to.be.equal(
            sheet.getCell('A13').style
          );
          const c12 = sheet.getCell('C12');
          c12.setStyle({font: {...c12.font, color: {argb: 'FFFFAA00'}}});
          expect(sheet.getCell('C12').style).to.not.be.equal(
            sheet.getCell('A13').style
          );

          // supposed to change F1 to J10 background color

          expect(sheet.getCell('G2').style).to.be.equal(
            sheet.getCell('I9').style
          );
          sheet.getCell('G2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {argb: 'FF789ABC'},
          };
          expect(sheet.getCell('G2').style).to.be.equal(
            sheet.getCell('I9').style
          );

          // supposed to change F13 background color only

          expect(sheet.getCell('F13').style).to.be.equal(
            sheet.getCell('D12').style
          );
          sheet.getCell('F13').setStyle({
            fill: {
              type: 'pattern',
              pattern: 'solid',
              fgColor: {argb: 'FF789ABC'},
            },
          });
          expect(sheet.getCell('F13').style).to.not.be.equal(
            sheet.getCell('D12').style
          );
        });

        // return wb.xlsx.writeFile('./spec/integration/data/test-issue-791-result.xlsx');
      });
  });
});
