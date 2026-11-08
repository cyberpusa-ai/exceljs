const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 2233 - conditional formatting stopIfTrue survives write and read', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'expression',
          priority: 1,
          stopIfTrue: true,
          formulae: ['A1>5'],
          style: {font: {bold: true}},
        },
        {
          type: 'cellIs',
          priority: 2,
          operator: 'greaterThanOrEqual',
          formulae: ['3'],
          style: {font: {italic: true}},
        },
      ],
    });

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);

    const [cf] = wb2.getWorksheet('Sheet1').conditionalFormattings;
    expect(cf.rules[0].stopIfTrue).to.equal(true);
    expect(cf.rules[1].stopIfTrue).to.equal(undefined);
    expect(cf.rules[1].operator).to.equal('greaterThanOrEqual');

    const buffer2 = await wb2.xlsx.writeBuffer();
    const wb3 = new ExcelJS.Workbook();
    await wb3.xlsx.load(buffer2);
    expect(wb3.getWorksheet('Sheet1').conditionalFormattings[0].rules[0].stopIfTrue).to.equal(true);
  });
});
