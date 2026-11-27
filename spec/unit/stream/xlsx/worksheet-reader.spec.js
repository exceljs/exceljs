const WorksheetReader = verquire('stream/xlsx/worksheet-reader');

describe('WorksheetReader', () => {
  it('preserves the types of cached formula results', async () => {
    const xml = `<worksheet><sheetData><row r="1">
      <c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>
      <c r="B1" t="b"><f>1=1</f><v>1</v></c>
      <c r="C1" t="b"><f>1=2</f><v>0</v></c>
      <c r="D1"><f>1-1</f><v>0</v></c>
      <c r="E1" t="str"><f>"paid"</f><v>paid</v></c>
      <c r="F1"><f>2+2</f></c>
    </row></sheetData></worksheet>`;
    const reader = new WorksheetReader({
      workbook: {},
      id: 1,
      iterator: [xml],
      options: {worksheets: 'emit'},
    });
    const actual = [];
    for await (const row of reader) {
      row.eachCell(cell => actual.push([cell.formula, cell.result]));
    }
    expect(actual).to.deep.equal([
      ['1/0', {error: '#DIV/0!'}],
      ['1=1', true],
      ['1=2', false],
      ['1-1', 0],
      ['"paid"', 'paid'],
      ['2+2', undefined],
    ]);
  });
});
