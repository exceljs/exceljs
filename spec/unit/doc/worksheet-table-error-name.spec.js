const Excel = verquire('exceljs');

function addTable(ws, tableOptions) {
  return ws.addTable(tableOptions);
}

describe('Error handling of table creation', () => {
  let tableOptions;

  beforeEach(() => {
    // Resuse this object for multiple tests of table creation error handling
    tableOptions = {
      name: 'TestTable',
      ref: 'A1',
      headerRow: true,
      totalsRow: true,
      style: {
        theme: 'TableStyleDark3',
        showRowStripes: true,
      },
      columns: [
        {name: 'Date', totalsRowLabel: 'Totals', filterButton: true},
        {
          name: 'Id',
          totalsRowFunction: 'max',
          filterButton: true,
          totalsRowResult: 4,
        },
        {
          name: 'Word',
          filterButton: false,
          style: {font: {bold: true, name: 'Comic Sans MS'}},
        },
      ],
      rows: [
        [new Date('2020-01-01'), 1, 'one'],
        [new Date('2020-01-02'), 2, 'two'],
        [new Date('2020-01-03'), 3, 'three'],
      ],
    };
  });

  it('should create a table without error when all required options are provided', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    expect(() => addTable(ws, tableOptions)).to.not.throw();
  });

  it('should throw an error when adding a table with empty rows', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    tableOptions.name = '404 Errors'; // invalid name, shouldn't be allowed
    expect(() => addTable(ws, tableOptions)).to.throw(
      'Invalid Table name. See Microsoft documentation for requirements for table names.'
    );
  });
});
