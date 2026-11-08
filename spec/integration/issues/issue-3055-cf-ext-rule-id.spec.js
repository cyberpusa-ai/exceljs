const ExcelJS = verquire('exceljs');

const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;

describe('github issues', () => {
  it('issue 3055 - x14 conditional formatting rules get GUID ids without the uuid package', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {type: 'dataBar', priority: 1, cfvo: [{type: 'min'}, {type: 'max'}]},
        {
          type: 'iconSet',
          priority: 2,
          iconSet: '3Triangles',
          cfvo: [
            {type: 'percent', value: 0},
            {type: 'percent', value: 33},
            {type: 'percent', value: 67},
          ],
        },
      ],
    });

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);

    const [cf] = wb2.getWorksheet('Sheet1').conditionalFormattings;
    const ids = cf.rules.map(rule => rule.x14Id);
    expect(ids).to.have.length(2);
    ids.forEach(id => expect(id).to.match(GUID));
    expect(ids[0]).to.not.equal(ids[1]);
  });
});
