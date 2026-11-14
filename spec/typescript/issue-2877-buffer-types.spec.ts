import 'regenerator-runtime/runtime';

import { expect } from 'chai';
import { readFileSync } from 'fs';
import { join } from 'path';
import ExcelJS from '../../index';
// type-only import so the hand-written typings in index.d.ts are checked
import type { ExcelBuffer, Workbook } from '../../index.d';

// exceljs#2877, exceljs#1032: index.d.ts used to declare its own `Buffer extends ArrayBuffer`,
// so Node's Buffer (e.g. from readFileSync) could not be passed to load() and writeBuffer()
// could not be assigned to a Node Buffer. Code typed against the old ArrayBuffer must still compile.
describe('typescript: issue 2877 - buffer types', () => {
  const fixture = join(__dirname, '../integration/data/test-issue-1669.xlsx');

  it('loads a Node Buffer from readFileSync', async () => {
    const wb: Workbook = new ExcelJS.Workbook();
    await wb.xlsx.load(readFileSync(fixture));
    expect(wb.worksheets.length).to.be.greaterThan(0);
  });

  it('loads a Uint8Array and an ArrayBuffer', async () => {
    const data = readFileSync(fixture);
    const u8 = new Uint8Array(data);
    await (new ExcelJS.Workbook() as Workbook).xlsx.load(u8);
    const ab: ArrayBuffer = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
    await (new ExcelJS.Workbook() as Workbook).xlsx.load(ab);
  });

  it('types writeBuffer() as a Node Buffer and as an ArrayBuffer', async () => {
    const wb: Workbook = new ExcelJS.Workbook();
    wb.addWorksheet('blort').getCell('A1').value = 7;

    const buffer: Buffer = await wb.xlsx.writeBuffer();
    const arrayBuffer: ArrayBuffer = await wb.xlsx.writeBuffer();
    const bytes: Uint8Array = await wb.xlsx.writeBuffer();
    const excelBuffer: ExcelBuffer = await wb.xlsx.writeBuffer();
    const csv: Buffer = await wb.csv.writeBuffer();
    const csvArrayBuffer: ArrayBuffer = await wb.csv.writeBuffer();

    expect(Buffer.isBuffer(buffer)).to.equal(true);
    expect(buffer.toString('base64').length).to.be.greaterThan(0);
    expect(arrayBuffer.byteLength).to.equal(buffer.length);
    expect(bytes.length).to.equal(buffer.length);
    expect(Buffer.concat([excelBuffer]).length).to.equal(buffer.length);
    expect(csv.toString()).to.equal('7');
    expect(csvArrayBuffer.byteLength).to.equal(1);

    const wb2: Workbook = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    expect(wb2.getWorksheet('blort').getCell('A1').value).to.equal(7);
  });

  it('accepts a Node Buffer in addImage() and returns it from getImage()', () => {
    const wb: Workbook = new ExcelJS.Workbook();
    const png = readFileSync(join(__dirname, '../integration/data/image.png'));
    const imageId = wb.addImage({buffer: png, extension: 'png'});
    const image: Buffer | undefined = wb.getImage(imageId).buffer;
    const imageArrayBuffer: ArrayBuffer | undefined = wb.getImage(imageId).buffer;
    expect(Buffer.isBuffer(image)).to.equal(true);
    expect(imageArrayBuffer).to.equal(image);
  });
});
