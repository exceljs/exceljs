const ExcelJS = verquire('exceljs');

describe('github 2244', () => {
  it('pull request 2244', async () => {
    async function test() {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('sheet');
      const imageId1 = workbook.addImage({
        filename: 'path/to/image.jpg',
        extension: 'jpeg',
      });
      worksheet.addImage(imageId1, 'B2:D6');
      await workbook.xlsx.writeFile('test.xlsx');
    }
    let error;
    try {
      await test();
    } catch (err) {
      error = err;
    }
    expect(error).to.be.an('error');
    //   expect(test).to.throw(Error);
  });
});
