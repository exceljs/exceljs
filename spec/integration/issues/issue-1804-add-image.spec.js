const ExcelJS = verquire('exceljs');

// this file to contain integration tests created from github issues
const TEST_XLSX_FILE_NAME = './spec/out/wb-issue-1804.test.xlsx';
const IMAGE_FILENAME1 = `${__dirname}/../data/image.png`;
const IMAGE_FILENAME2 = `${__dirname}/../data/image1.jpg`;

describe('github issues', () => {
  it('issue 1804 - add wrong image', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');

    const img1 = wb.addImage({
      filename: IMAGE_FILENAME1,
      extension: 'PNG',
    });

    const img2 = wb.addImage({
      filename: IMAGE_FILENAME2,
      extension: 'JPEG',
    });

    ws.addImage(img2, 'A1:A1');
    ws.addImage(img1, 'A3:A3');
    ws.addImage(img1, 'A5:A5');

    ws.getRow(1).getCell(2).value = 'image2';
    ws.getRow(3).getCell(2).value = 'image1';
    ws.getRow(5).getCell(2).value = 'image1';

    return wb.xlsx
      .writeFile(TEST_XLSX_FILE_NAME)
      .then(() => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then(wb2 => {
        const ws2 = wb2.getWorksheet('Sheet1');
        expect(ws2._media[1].imageId).to.equal(ws2._media[2].imageId);
      });
  }).timeout(6000);
});
