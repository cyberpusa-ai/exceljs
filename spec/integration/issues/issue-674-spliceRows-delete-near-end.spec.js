const ExcelJS = verquire('exceljs');

function createSheet(nRows = 5) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('splice');
  for (let i = 1; i <= nRows; i++) {
    ws.addRow([`A${i}`, `B${i}`]);
    ws.getRow(i).getCell(1).font = {bold: i % 2 === 0};
  }
  return {wb, ws};
}

function columnA(ws) {
  const values = [];
  ws.eachRow(row => values.push(row.getCell(1).value));
  return values;
}

describe('github issues', () => {
  describe('issue 674 - spliceRows delete near the end of the sheet', () => {
    it('removes the requested rows when fewer rows remain below them', () => {
      const {ws} = createSheet();
      ws.spliceRows(4, 2);

      expect(ws.rowCount).to.equal(3);
      expect(columnA(ws)).to.deep.equal(['A1', 'A2', 'A3']);
    });

    it('does not leave stale rows behind the moved rows', () => {
      const {ws} = createSheet();
      ws.spliceRows(2, 3);

      expect(ws.rowCount).to.equal(2);
      expect(columnA(ws)).to.deep.equal(['A1', 'A5']);
      expect(ws.getCell('B2').value).to.equal('B5');
      expect(ws.getCell('A2').font).to.deep.equal({bold: false});
      expect(ws.findRow(3)).to.be.undefined();
    });

    it('handles repeated deletes at the end', () => {
      const {ws} = createSheet();
      ws.spliceRows(4, 1);
      expect(ws.rowCount).to.equal(4);
      expect(columnA(ws)).to.deep.equal(['A1', 'A2', 'A3', 'A5']);

      ws.spliceRows(4, 3);
      expect(ws.rowCount).to.equal(3);
      expect(columnA(ws)).to.deep.equal(['A1', 'A2', 'A3']);
    });

    it('deletes past the end of the sheet without growing it', () => {
      const {ws} = createSheet();
      ws.spliceRows(5, 10);

      expect(ws.rowCount).to.equal(4);
      expect(ws._rows.length).to.equal(5);
      expect(columnA(ws)).to.deep.equal(['A1', 'A2', 'A3', 'A4']);
    });

    it('removes rows that are replaced by fewer inserts', () => {
      const {ws} = createSheet();
      ws.spliceRows(3, 3, ['X3', 'Y3']);

      expect(ws.rowCount).to.equal(3);
      expect(columnA(ws)).to.deep.equal(['A1', 'A2', 'X3']);
    });

    it('keeps styles and merges above the splice', () => {
      const {ws} = createSheet();
      ws.mergeCells('A2:B2');
      ws.spliceRows(4, 2);

      expect(ws.getCell('A2').font).to.deep.equal({bold: true});
      expect(ws.getCell('A2').isMerged).to.be.true();
      expect(ws.getCell('B2').isMerged).to.be.true();
      expect(ws.getCell('B2').master.address).to.equal('A2');
    });

    it('writes only the remaining rows', async () => {
      const {wb, ws} = createSheet();
      ws.mergeCells('A2:B2');
      ws.spliceRows(3, 2);

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('splice');

      expect(ws2.rowCount).to.equal(3);
      expect(columnA(ws2)).to.deep.equal(['A1', 'A2', 'A5']);
      expect(ws2.getCell('A3').font.bold).to.not.be.ok();
      expect(ws2.getCell('B2').isMerged).to.be.true();
    });
  });
});
