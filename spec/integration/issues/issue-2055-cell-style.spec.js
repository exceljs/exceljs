const ExcelJS = verquire('exceljs');

const fileName = './spec/integration/data/test-issue-2055.xlsx';

describe('github issues', () => {
  describe('issue 2055 - Cells from loaded files reuse the same style object', () => {
    describe('when cells share the same style', () => {
      it('expects to those cells to use the same style object', async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(fileName);

        const sheet = wb.getWorksheet(1);

        const A1 = sheet.getCell('A1');
        const B1 = sheet.getCell('B1');
        const C1 = sheet.getCell('C1');

        expect(A1.style).to.equal(B1.style);
        expect(A1.style).to.equal(C1.style);
      });

      describe('when changing any attribute of a cell\'s style', () => {
        let workbook;
        beforeEach(async () => {
          workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(fileName);
        });

        it('should clone the style and set the new attribute values', () => {
          const sheet = workbook.getWorksheet(1);
          const A1 = sheet.getCell('A1');
          const B1 = sheet.getCell('B1');
          const C1 = sheet.getCell('C1');
          const originalStyle = A1.style;

          B1.fill = {
            type: 'pattern',
            pattern: 'solid',
            bgColor: {
              argb: 'ff123456',
            },
            fgColor: {
              argb: 'ff123456',
            },
          };

          expect(A1.style).to.equal(originalStyle);
          expect(C1.style).to.equal(originalStyle);

          expect(B1.style).to.not.equal(originalStyle);
          expect(B1.style).to.deep.equal({
            ...originalStyle,
            fill: {
              type: 'pattern',
              pattern: 'solid',
              bgColor: {
                argb: 'ff123456',
              },
              fgColor: {
                argb: 'ff123456',
              },
            },
          });
        });
      });
    });
  });
});
