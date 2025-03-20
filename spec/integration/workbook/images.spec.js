const fs = require('fs');
const {promisify} = require('util');

const ExcelJS = verquire('exceljs');

const IMAGE_FILENAME = `${__dirname}/../data/image.png`;
const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';
const IMAGE_XLSX_FILE_NAME = `${__dirname}/../data/images.xlsx`;

const fsReadFileAsync = promisify(fs.readFile);

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Images', () => {
    it('stores background image', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      let wb2;
      let ws2;
      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.getCell('A1').value = 'Hello, World!';
      ws.addBackgroundImage(imageId);

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(imageData => {
          const backgroundId2 = ws2.getBackgroundImageId();
          const image = wb2.getImage(backgroundId2);

          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('stores embedded image and hyperlink', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      let wb2;
      let ws2;

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.getCell('A1').value = 'Hello, World!';
      ws.getCell('A2').value = {
        hyperlink: 'http://www.somewhere.com',
        text: 'www.somewhere.com',
      };
      ws.addImage(imageId, 'C3:E6');

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          expect(ws.getCell('A1').value).to.equal('Hello, World!');
          expect(ws.getCell('A2').value).to.deep.equal({
            hyperlink: 'http://www.somewhere.com',
            text: 'www.somewhere.com',
          });

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(imageData => {
          const images = ws2.getImages();
          expect(images.length).to.equal(1);

          const imageDesc = images[0];
          expect(imageDesc.range.tl.col).to.equal(2);
          expect(imageDesc.range.tl.row).to.equal(2);
          expect(imageDesc.range.br.col).to.equal(5);
          expect(imageDesc.range.br.row).to.equal(6);

          const image = wb2.getImage(imageDesc.imageId);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('stores embedded image with oneCell', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      let wb2;
      let ws2;

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.addImage(imageId, {
        tl: {col: 0.1125, row: 0.4},
        br: {col: 2.101046875, row: 3.4},
        editAs: 'oneCell',
      });

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(imageData => {
          const images = ws2.getImages();
          expect(images.length).to.equal(1);

          const imageDesc = images[0];
          expect(imageDesc.range.editAs).to.equal('oneCell');

          const image = wb2.getImage(imageDesc.imageId);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('stores embedded image with one-cell-anchor', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      let wb2;
      let ws2;

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.addImage(imageId, {
        tl: {col: 0.1125, row: 0.4},
        ext: {width: 100, height: 100},
        editAs: 'oneCell',
      });

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(imageData => {
          const images = ws2.getImages();
          expect(images.length).to.equal(1);

          const imageDesc = images[0];
          expect(imageDesc.range.editAs).to.equal('oneCell');
          expect(imageDesc.range.ext.width).to.equal(100);
          expect(imageDesc.range.ext.height).to.equal(100);

          const image = wb2.getImage(imageDesc.imageId);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('stores embedded image with hyperlinks', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      let wb2;
      let ws2;

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.addImage(imageId, {
        tl: {col: 0.1125, row: 0.4},
        ext: {width: 100, height: 100},
        editAs: 'absolute',
        hyperlinks: {
          hyperlink: 'http://www.somewhere.com',
          tooltip: 'www.somewhere.com',
        },
      });

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(imageData => {
          const images = ws2.getImages();
          expect(images.length).to.equal(1);

          const imageDesc = images[0];
          expect(imageDesc.range.editAs).to.equal('absolute');
          expect(imageDesc.range.ext.width).to.equal(100);
          expect(imageDesc.range.ext.height).to.equal(100);

          expect(imageDesc.range.hyperlinks).to.deep.equal({
            hyperlink: 'http://www.somewhere.com',
            tooltip: 'www.somewhere.com',
          });

          const image = wb2.getImage(imageDesc.imageId);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('image extensions should not be case sensitive', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      let wb2;
      let ws2;

      const imageId1 = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'PNG',
      });

      const imageId2 = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'JPEG',
      });

      ws.addImage(imageId1, {
        tl: {col: 0.1125, row: 0.4},
        ext: {width: 100, height: 100},
      });

      ws.addImage(imageId2, {
        tl: {col: 0.1125, row: 0.4},
        br: {col: 2.101046875, row: 3.4},
        editAs: 'oneCell',
      });

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(imageData => {
          const images = ws2.getImages();
          expect(images.length).to.equal(2);

          const imageDesc1 = images[0];
          expect(imageDesc1.range.ext.width).to.equal(100);
          expect(imageDesc1.range.ext.height).to.equal(100);
          const image1 = wb2.getImage(imageDesc1.imageId);

          const imageDesc2 = images[1];
          expect(imageDesc2.range.editAs).to.equal('oneCell');

          const image2 = wb2.getImage(imageDesc1.imageId);

          expect(Buffer.compare(imageData, image1.buffer)).to.equal(0);
          expect(Buffer.compare(imageData, image2.buffer)).to.equal(0);
        });
    });
  });

  describe('Image Manipulation', () => {
    describe('Functionality Tests', () => {
      it('should return sheetImageId when adding an image', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        const imageId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: 'jpeg',
        });

        const sheetImageId = ws.addImage(imageId, 'C3:E6');
        expect(sheetImageId).to.be.a('string');
        expect(
          ws.getImages().some(img => img.sheetImageId === sheetImageId)
        ).to.be.true();
      });

      it('should return sheetImageId for existing images in a file', () => {
        const wb = new ExcelJS.Workbook();

        return wb.xlsx.readFile(IMAGE_XLSX_FILE_NAME).then(() => {
          const ws = wb.getWorksheet(1); // 假设图片在第一个工作表中
          const images = ws.getImages();
          expect(images.length).to.be.greaterThan(0);

          images.forEach(image => {
            expect(image.sheetImageId).to.be.a('string');
          });
        });
      });
    });

    describe('Unit Tests', () => {
      it('should add and remove an embedded image', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        const imageId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: 'jpeg',
        });

        const sheetImageId = ws.addImage(imageId, 'C3:E6');
        expect(ws.getImages().length).to.equal(1);

        ws.removeImage(sheetImageId);
        expect(ws.getImages().length).to.equal(0);
      });

      it('should add and remove a background image', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        const imageId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: 'jpeg',
        });

        const sheetImageId = ws.addBackgroundImage(imageId);
        expect(ws.getBackgroundImageId()).to.equal(imageId);

        ws.removeImage(sheetImageId);
        expect(ws.getBackgroundImageId()).to.be.undefined();
      });
    });

    describe('Integration Tests', () => {
      it('should add and remove an embedded image without affecting the original', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        // Step 1: Add an initial image
        const wbImageId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: 'jpeg',
        });
        const firstSheetImageId = ws.addImage(wbImageId, 'C3:E6');
        expect(ws.getImages().length).to.equal(1);

        // Step 2: Retrieve the existing image and add a new one
        const secondSheetImageId = ws.addImage(wbImageId, 'F3:H6');
        expect(ws.getImages().length).to.equal(2);

        // Step 3: Remove the original sheet image
        ws.removeImage(firstSheetImageId);
        expect(ws.getImages().length).to.equal(1);

        // Step 4: Verify the original image is unaffected
        const remainingImage = ws.getImages()[0];
        expect(remainingImage.sheetImageId).to.equal(secondSheetImageId);
        expect(remainingImage.range.tl.col).to.equal(5);
        expect(remainingImage.range.tl.row).to.equal(2);
      });

      it('should add and remove a background image without affecting other images', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        // Add an embedded image
        const wbImageId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: 'jpeg',
        });
        const normalSheetImageId = ws.addImage(wbImageId, 'C3:E6');
        expect(ws.getImages().length).to.equal(1);

        // Add a background image
        const backgroundImageId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: 'jpeg',
        });
        const backgroundSheetImageId = ws.addBackgroundImage(backgroundImageId);
        expect(ws.getBackgroundImageId()).to.equal(backgroundImageId);

        // Remove the background image
        ws.removeImage(backgroundSheetImageId);
        expect(ws.getBackgroundImageId()).to.be.undefined();

        // Verify the normal image is unaffected
        expect(ws.getImages().length).to.equal(1);
        const remainingImage = ws.getImages()[0];
        expect(remainingImage.sheetImageId).to.equal(normalSheetImageId);
      });
    });
  });
});
