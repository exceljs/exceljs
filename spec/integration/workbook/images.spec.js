'use strict';

var fs = require('fs');
var Promish = require('promish');
var chai = require('chai');
var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

var expect = chai.expect;

const IMAGE_FILENAME = __dirname + '/../data/image.png';
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';
var fsReadFileAsync = Promish.promisify(fs.readFile);

// =============================================================================
// Tests

describe('Workbook', function() {
  describe('Images', function() {
    it('stores background image', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');
      var wb2, ws2;
      var imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.getCell('A1').value = 'Hello, World!';
      ws.addBackgroundImage(imageId);

      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function() {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(function(imageData) {
          const backgroundId2 = ws2.getBackgroundImageId();
          const image = wb2.getImage(backgroundId2);

          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('stores embedded image and hyperlink', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');
      var wb2, ws2;

      var imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.getCell('A1').value = 'Hello, World!';
      ws.getCell('A2').value = { hyperlink: 'http://www.somewhere.com', text: 'www.somewhere.com' };
      ws.addImage(imageId, 'C3:E6');

      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function() {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          expect(ws.getCell('A1').value).to.equal('Hello, World!');
          expect(ws.getCell('A2').value).to.deep.equal({
            hyperlink: 'http://www.somewhere.com',
            text: 'www.somewhere.com'
          });

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(function(imageData) {
          const images = ws2.getImages();
          expect(images.length).to.equal(1);

          const imageDesc = images[0];
          expect(imageDesc.range).to.equal('C3:E6');

          const image = wb2.getImage(imageDesc.imageId);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });

    it('stores embedded image with oneCell', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');
      var wb2, ws2;

      var imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        br: { col: 2.101046875, row: 3.4 },
        editAs: 'oneCell'
      });

      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function() {
          ws2 = wb2.getWorksheet('blort');
          expect(ws2).to.not.be.undefined();

          return fsReadFileAsync(IMAGE_FILENAME);
        })
        .then(function(imageData) {
          const images = ws2.getImages();
          expect(images.length).to.equal(1);

          const imageDesc = images[0];
          expect(imageDesc.range.editAs).to.equal('oneCell');

          const image = wb2.getImage(imageDesc.imageId);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        });
    });
  });
});
