const fs = require('fs');

const ExcelJS = verquire('exceljs');

const IMAGE_FILENAME = `${__dirname}/../data/image.png`;

function place(ws, ids) {
  ids.forEach((id, col) => ws.addImage(id, {tl: {col, row: 0}, ext: {width: 10, height: 10}}));
}

function placedIds(ws) {
  return ws
    .getImages()
    .sort((a, b) => a.range.tl.nativeCol - b.range.tl.nativeCol)
    .map(image => Number(image.imageId));
}

describe('github issues', () => {
  describe('issue 1804 - consecutive placements of the same image', () => {
    it('each placement keeps its own image', async () => {
      const wb = new ExcelJS.Workbook();
      const buffer = fs.readFileSync(IMAGE_FILENAME);
      const img1 = wb.addImage({buffer, extension: 'png'});
      const img2 = wb.addImage({buffer, extension: 'png'});

      const order = [img2, img1, img1, img2, img2, img1];
      place(wb.addWorksheet('Sheet1'), order);

      const loaded = new ExcelJS.Workbook();
      await loaded.xlsx.load(await wb.xlsx.writeBuffer());
      expect(placedIds(loaded.getWorksheet('Sheet1'))).to.deep.equal(order);
    });

    it('works when the same image ends one sheet and starts the next', async () => {
      const wb = new ExcelJS.Workbook();
      const buffer = fs.readFileSync(IMAGE_FILENAME);
      const img1 = wb.addImage({buffer, extension: 'png'});
      const img2 = wb.addImage({buffer, extension: 'png'});

      place(wb.addWorksheet('Sheet1'), [img1, img2]);
      place(wb.addWorksheet('Sheet2'), [img2, img1]);

      const loaded = new ExcelJS.Workbook();
      await loaded.xlsx.load(await wb.xlsx.writeBuffer());
      expect(placedIds(loaded.getWorksheet('Sheet1'))).to.deep.equal([img1, img2]);
      expect(placedIds(loaded.getWorksheet('Sheet2'))).to.deep.equal([img2, img1]);
    });
  });
});
