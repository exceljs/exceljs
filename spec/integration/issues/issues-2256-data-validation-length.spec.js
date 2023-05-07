const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('issue 2256 - (dataValidation) The exported table inserts a drop-down menu for the cell. The cell cannot display drop-down options', () => {
    it('issue 2256 - (dataValidation) Input content is 266 characters, more than 255 characters, throw error', async () => {
      async function test() {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('sheet');
        const columns = [
          {
            header: 'Test1',
            key: 'test1',
            width: 25,
          },
        ];
        worksheet.columns = columns;
        const cell = worksheet.getCell('A2');
        const value =
          '"1234 Main St, Anytown 56789,5678 Elm Ave, Smallville 23456,9012 Oak Rd, Big City 34567,3456 Pine Dr, Tinytown 67890,7890 Maple Ln, Mediumtown 12345,2345 Hickory St, Suburbia 45678,6789 Cedar Cir, Urbanville 89012,0123 Birch Blvd, Countryside 23456,4567 Wal"';
        cell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [value],
        };
        await workbook.xlsx.writeFile(
          './spec/integration/data/test-issue-2256-1.xlsx'
        );
      }
      let error;
      try {
        await test();
      } catch (err) {
        error = err;
      }
      expect(error)
        .to.be.an('error')
        .with.property(
          'message',
          'The input cannot be larger than 255 characters. Please check the value of dataValidation.formulae'
        );
    });

    it('issue 2256 - (dataValidation) Input content is 255 characters, normal display', async () => {
      async function test() {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('sheet');
        const columns = [
          {
            header: 'Test1',
            key: 'test1',
            width: 25,
          },
        ];
        worksheet.columns = columns;
        const cell = worksheet.getCell('A2');
        const value =
          '"1234 Main St, Anytown 56789,5678 Elm Ave, Smallville 23456,9012 Oak Rd, Big City 34567,3456 Pine Dr, Tinytown 67890,7890 Maple Ln, Mediumtown 12345,2345 Hickory St, Suburbia 45678,6789 Cedar Cir, Urbanville 89012,0123 Birch Blvd, Countryside 23456,4567 Wa"';
        cell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: [value],
        };
        await workbook.xlsx.writeFile(
          './spec/integration/data/test-issue-2256-2.xlsx'
        );
      }
      await test();
    });
  });
});
