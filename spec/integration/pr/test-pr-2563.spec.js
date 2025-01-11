const ExcelJS = verquire('exceljs');

const COMMENTS_AND_HTIMAGE_XLSX_FILE_NAME = './spec/out/comments-headerImage.test.xlsx';
const COMMENTS_AND_HTIMAGE_AND_SHEETIMAGE_XLSX_FILE_NAME =
  './spec/out/comments-headerImage-sheetimage.test.xlsx';

describe('pull request 2563', () => {
  it('pull request 2563 - header and footer support image', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/comments-headerImage.xlsx');
    await wb.xlsx.writeFile(COMMENTS_AND_HTIMAGE_XLSX_FILE_NAME);
  });

  it('pull request 2563 - sheet image and hf image', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/comments-headerImage-sheetimage.xlsx');
    await wb.xlsx.writeFile(COMMENTS_AND_HTIMAGE_AND_SHEETIMAGE_XLSX_FILE_NAME);
  });
});
