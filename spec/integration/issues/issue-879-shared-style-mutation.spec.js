const ExcelJS = verquire('exceljs');

async function roundTrip(wb) {
  const buffer = await wb.xlsx.writeBuffer();
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer);
  return wb2;
}

describe('github issues', () => {
  describe('issue 879 - cell font styling is applied to entire sheet', () => {
    const props = {
      numFmt: '0.00%',
      font: {bold: true, color: {argb: 'FFFF0000'}},
      alignment: {horizontal: 'center'},
      border: {top: {style: 'thin'}},
      fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}},
      protection: {locked: false},
    };

    async function loadSharedStyleSheet() {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A2').value = 'a';
      ws.getCell('B2').value = 'b';
      ws.getCell('A3').value = 'c';
      // give the cells one shared, non-default xf
      ['A2', 'B2', 'A3'].forEach(address => {
        ws.getCell(address).font = {italic: true};
      });
      const wb2 = await roundTrip(wb);
      return wb2.getWorksheet('Sheet1');
    }

    Object.keys(props).forEach(name => {
      it(`setting ${name} on one loaded cell does not change other cells`, async () => {
        const ws = await loadSharedStyleSheet();
        const before = ws.getCell('B2')[name];
        ws.getCell('A2')[name] = props[name];
        expect(ws.getCell('A2')[name]).to.deep.equal(props[name]);
        expect(ws.getCell('B2')[name]).to.deep.equal(before);
        expect(ws.getCell('A3')[name]).to.deep.equal(before);
      });
    });

    it('styling one loaded cell survives a write without restyling others', async () => {
      const ws = await loadSharedStyleSheet();
      ws.getCell('A2').font = {bold: true};
      const wb3 = await roundTrip(ws.workbook);
      const ws3 = wb3.getWorksheet('Sheet1');
      expect(ws3.getCell('A2').font).to.deep.equal({bold: true});
      expect(ws3.getCell('B2').font).to.not.deep.equal({bold: true});
      expect(ws3.getCell('A3').font).to.not.deep.equal({bold: true});
    });

    it('setting a style on a loaded column does not change cells sharing its style', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      const font = {italic: true};
      ws.getColumn(1).font = font;
      ws.getColumn(2).font = font;
      ws.getCell('C1').font = font;
      ws.getCell('C1').value = 'x';
      const ws2 = (await roundTrip(wb)).getWorksheet('Sheet1');

      ws2.getColumn(1).font = {bold: true};
      expect(ws2.getColumn(1).font).to.deep.equal({bold: true});
      expect(ws2.getColumn(2).font).to.deep.equal(font);
      expect(ws2.getCell('C1').font).to.deep.equal(font);
    });

    it('setting a style on a column-styled cell does not change the column or its other cells', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ws.getColumn(1).font = {italic: true};
      ws.getRow(1).font = {underline: true};
      const a1 = ws.getCell('A1');
      const a2 = ws.getCell('A2');
      a1.font = {bold: true};
      expect(a1.font).to.deep.equal({bold: true});
      expect(a2.font).to.deep.equal({italic: true});
      expect(ws.getColumn(1).font).to.deep.equal({italic: true});
      expect(ws.getRow(1).font).to.deep.equal({underline: true});
    });

    it('merged cells keep sharing one style object when styled', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ws.mergeCells('A4:B5');
      ws.getCell('B5').font = {bold: true};
      ws.getCell('A4').fill = props.fill;
      ['A4', 'A5', 'B4', 'B5'].forEach(address => {
        expect(ws.getCell(address).style).to.equal(ws.getCell('A4').style);
      });
      expect(ws.getCell('A4').font).to.deep.equal({bold: true});
      expect(ws.getCell('B5').fill).to.deep.equal(props.fill);
      expect(ws.getCell('C4').font).to.be.undefined();
    });

    it('styling a loaded merged cell does not change unmerged cells with the same xf', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ['A1', 'B1', 'A2', 'B2', 'D1'].forEach(address => {
        ws.getCell(address).font = {italic: true};
      });
      ws.mergeCells('A1:B2');
      const ws2 = (await roundTrip(wb)).getWorksheet('Sheet1');

      ws2.getCell('B2').font = {bold: true};
      ['A1', 'B1', 'A2', 'B2'].forEach(address => {
        expect(ws2.getCell(address).font).to.deep.equal({bold: true});
        expect(ws2.getCell(address).style).to.equal(ws2.getCell('A1').style);
      });
      expect(ws2.getCell('D1').font).to.deep.equal({italic: true});
    });

    it('streaming writer keeps merged cells sharing one style object', () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({useStyles: true});
      const ws = wb.addWorksheet('Sheet1');
      ws.mergeCells('A1:B2');
      ws.getCell('A1').font = {bold: true};
      ['A1', 'B1', 'A2', 'B2'].forEach(address => {
        expect(ws.getCell(address).font).to.deep.equal({bold: true});
        expect(ws.getCell(address).style).to.equal(ws.getCell('A1').style);
      });
    });
  });
});
