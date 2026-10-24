const ExcelJS = verquire('exceljs');

describe('pull request  2244', () => {
  async function test(options) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('sheet');
    const imageId1 = workbook.addImage({
      filename: 'path/to/image.jpg', // Non-existent file
      extension: 'jpeg',
    });
    worksheet.addImage(imageId1, 'B2:D6');
    await workbook.xlsx.writeFile('test.xlsx', options);
  }

  it('pull request 2244- Fix xlsx.writeFile() not catching error when error occurs', async () => {
    // openly fork: a missing image only fails the write when placeholders are off
    let error;
    try {
      await test({media: {placeholder: false}});
    } catch (err) {
      error = err;
    }
    expect(error).to.be.an('error');
  });

  it('openly fork: replaces a missing image with the placeholder by default', async () => {
    const errors = [];
    await test({media: {onError: err => errors.push(err.code)}});
    expect(errors).to.deep.equal(['ENOENT']);
  });
});
