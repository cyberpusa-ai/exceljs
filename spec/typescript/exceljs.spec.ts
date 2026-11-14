import 'regenerator-runtime/runtime';

import { expect } from 'chai';
import { readFileSync } from 'fs';
import { join } from 'path';
import { PassThrough } from 'stream';
import ExcelJS from '../../index';
// type-only import so the hand-written typings in index.d.ts are checked
import type { CellValue, Column, Row } from '../../index.d';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });
  it('can create and stream xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const stream = new PassThrough();
    const wb2 = new ExcelJS.Workbook();
    const [, read] = await Promise.all([
      wb.xlsx.write(stream).then(() => stream.end()),
      wb2.xlsx.read(stream),
    ]);
    const ws2 = read.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });
  it('types row and column values as arrays', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    const column: Column = ws.getColumn(1);
    column.values = [, 'a', 1, null];
    const columnValues: CellValue[] = column.values;
    expect(columnValues[1]).to.equal('a');
    expect(columnValues[2]).to.equal(1);

    const row: Row = ws.getRow(5);
    row.values = [, 'x', 'y'];
    expect((row.values as CellValue[])[2]).to.equal('y');
    ws.columns = [{key: 'id'}, {key: 'name'}];
    row.values = {id: 3, name: 'z'};
    expect(row.getCell(2).value).to.equal('z');
  });
  // issue #2877 / #1032: Buffer in index.d.ts must be Node's Buffer, not a local ArrayBuffer
  it('loads a node Buffer, Uint8Array or ArrayBuffer and writes a node Buffer', async () => {
    const fileBuffer: Buffer = readFileSync(join(__dirname, '../integration/data/gold.xlsx'));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(fileBuffer);
    await wb.xlsx.load(new Uint8Array(fileBuffer));
    await wb.xlsx.load(new Uint8Array(fileBuffer).buffer);
    expect(wb.worksheets.length).to.be.greaterThan(0);

    const xlsxBuffer: Buffer = await wb.xlsx.writeBuffer();
    const csvBuffer: Buffer = await wb.csv.writeBuffer();
    const bytes: Uint8Array = xlsxBuffer;
    expect(Buffer.isBuffer(xlsxBuffer)).to.equal(true);
    expect(Buffer.isBuffer(csvBuffer)).to.equal(true);
    expect(bytes.length).to.be.greaterThan(0);

    const imageId = wb.addImage({
      buffer: readFileSync(join(__dirname, '../integration/data/image.png')),
      extension: 'png',
    });
    expect(imageId).to.equal(0);
  });
});
