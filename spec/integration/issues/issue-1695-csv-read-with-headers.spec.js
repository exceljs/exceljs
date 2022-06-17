const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  const TEST_CSV_FILE = './spec/integration/data/test-issue-1695.csv';
  const TEST_CSV_WO_HEADERS_FILE =
    './spec/integration/data/test-issue-1695-without-headers.csv';
  const BIRTHDAY = new Date('2019-11-04T00:00:00').toString();

  it('issue 1695 - reading csv files with headers = true', () => {
    const wb = new ExcelJS.Workbook();
    const options = {
      parserOptions: {
        headers: true,
      },
    };
    return wb.csv.readFile(TEST_CSV_FILE, options).then(worksheet => {
      expect(worksheet.rowCount).to.equal(2);
      expect(worksheet.getColumn('Name').values[1]).to.equal('Name');
      expect(worksheet.getColumn('Name').values[2]).to.equal('Jim');
      expect(worksheet.getColumn('Full Name').values[2]).to.equal('Jim Green');
      expect(worksheet.getColumn('Birthday').values[2].toString()).to.equal(
        BIRTHDAY
      );
      expect(worksheet.getColumn('Age').values[2]).to.equal(3);
      expect(worksheet.getColumn('Address')).to.equal(worksheet.getColumn('E'));
    });
  });

  it('issue 1695 - reading csv files with transforming headers', () => {
    const wb = new ExcelJS.Workbook();
    const options = {
      parserOptions: {
        headers: headers => headers.map(h => h.toUpperCase()),
      },
    };
    return wb.csv.readFile(TEST_CSV_FILE, options).then(worksheet => {
      expect(worksheet.rowCount).to.equal(2);
      expect(worksheet.getColumn('NAME').values[1]).to.equal('NAME');
      expect(worksheet.getColumn('NAME').values[2]).to.equal('Jim');
      expect(worksheet.getColumn('FULL NAME').values[2]).to.equal('Jim Green');
      expect(worksheet.getColumn('BIRTHDAY').values[2].toString()).to.equal(
        BIRTHDAY
      );
      expect(worksheet.getColumn('AGE').values[2]).to.equal(3);
      expect(worksheet.getColumn('ADDRESS')).to.equal(worksheet.getColumn('E'));
    });
  });

  it('issue 1695 - reading csv files with rename headers', () => {
    const wb = new ExcelJS.Workbook();
    const options = {
      parserOptions: {
        headers: ['name', 'full_name', 'birthday', 'age', 'address'],
        renameHeaders: true,
      },
    };
    return wb.csv.readFile(TEST_CSV_FILE, options).then(worksheet => {
      expect(worksheet.rowCount).to.equal(2);
      expect(worksheet.getColumn('name').values[1]).to.equal('name');
      expect(worksheet.getColumn('name').values[2]).to.equal('Jim');
      expect(worksheet.getColumn('full_name').values[2]).to.equal('Jim Green');
      expect(worksheet.getColumn('birthday').values[2].toString()).to.equal(
        BIRTHDAY
      );
      expect(worksheet.getColumn('age').values[2]).to.equal(3);
      expect(worksheet.getColumn('address')).to.equal(worksheet.getColumn('E'));
    });
  });

  it('issue 1695 - reading csv files with custom headers', () => {
    const wb = new ExcelJS.Workbook();
    const options = {
      parserOptions: {
        headers: ['name', 'full_name', 'birthday', 'age', 'address'],
      },
    };
    return wb.csv
      .readFile(TEST_CSV_WO_HEADERS_FILE, options)
      .then(worksheet => {
        expect(worksheet.rowCount).to.equal(2);
        expect(worksheet.getColumn('name').values[1]).to.equal('name');
        expect(worksheet.getColumn('name').values[2]).to.equal('Jim');
        expect(worksheet.getColumn('full_name').values[2]).to.equal(
          'Jim Green'
        );
        expect(worksheet.getColumn('birthday').values[2].toString()).to.equal(
          BIRTHDAY
        );
        expect(worksheet.getColumn('age').values[2]).to.equal(3);
        expect(worksheet.getColumn('address')).to.equal(
          worksheet.getColumn('E')
        );
      });
  });

  it('issue 1695 - reading csv files skipping columns', () => {
    const wb = new ExcelJS.Workbook();
    const options = {
      parserOptions: {
        headers: ['name', 'full_name', 'birthday', undefined, 'address'],
        renameHeaders: true,
      },
    };
    return wb.csv.readFile(TEST_CSV_FILE, options).then(worksheet => {
      expect(worksheet.rowCount).to.equal(2);
      expect(worksheet.getColumn('name').values[1]).to.equal('name');
      expect(worksheet.getColumn('name').values[2]).to.equal('Jim');
      expect(worksheet.getColumn('full_name').values[2]).to.equal('Jim Green');
      expect(worksheet.getColumn('birthday').values[2].toString()).to.equal(
        BIRTHDAY
      );
      expect(worksheet.getColumn('address')).to.equal(worksheet.getColumn('E'));

      // empty column
      expect(worksheet.getColumn(4).values).to.empty();
      expect(worksheet.getColumn('D').values).to.empty();
      expect(() => worksheet.getColumn('age')).to.throw(
        Error,
        'Out of bounds. Invalid column letter: age'
      );
    });
  });

  // using custom headers (without "renameHeaders" option) on has header
  // row csv file, original header will be changed to 2nd row, custom
  // header will be 1st row.
  it('issue 1695 - reading has header row csv files with custom headers', () => {
    const wb = new ExcelJS.Workbook();
    const options = {
      parserOptions: {
        headers: ['name', 'full_name', 'birthday', 'age', 'address'],
      },
    };
    return wb.csv.readFile(TEST_CSV_FILE, options).then(worksheet => {
      expect(worksheet.rowCount).to.equal(3);
      expect(worksheet.getColumn('name').values[1]).to.equal('name');
      expect(worksheet.getColumn('name').values[2]).to.equal('Name');
      expect(worksheet.getColumn('name').values[3]).to.equal('Jim');
      expect(worksheet.getColumn('full_name').values[1]).to.equal('full_name');
      expect(worksheet.getColumn('full_name').values[2]).to.equal('Full Name');
      expect(worksheet.getColumn('full_name').values[3]).to.equal('Jim Green');
      expect(worksheet.getColumn('birthday').values[3].toString()).to.equal(
        BIRTHDAY
      );
      expect(worksheet.getColumn('age').values[3]).to.equal(3);
      expect(worksheet.getColumn('address')).to.equal(worksheet.getColumn('E'));
    });
  });
});
