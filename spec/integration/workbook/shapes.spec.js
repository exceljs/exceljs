const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Shapes', () => {
    it('stores shape', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      let wb2;
      let ws2;

      ws.addShape({type: 'line'}, 'B2:D6');

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('sheet');
          expect(ws2).to.not.be.undefined();
          const shapes = ws2.getShapes();
          expect(shapes.length).to.equal(1);

          const shape = shapes[0];
          expect(shape.props.type).to.equal('line');
        });
    });

    it('stores shape with oneCell', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      let wb2;
      let ws2;

      ws.addShape(
        {type: 'rect'},
        {
          tl: {col: 0.1125, row: 0.4},
          br: {col: 2.101046875, row: 3.4},
          editAs: 'oneCell',
        }
      );

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('sheet');
          expect(ws2).to.not.be.undefined();
          const shapes = ws2.getShapes();
          expect(shapes.length).to.equal(1);

          const shape = shapes[0];
          expect(shape.range.editAs).to.equal('oneCell');
          expect(shape.props.type).to.equal('rect');
        });
    });
  });
});