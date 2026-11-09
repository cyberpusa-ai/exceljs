# Changelog

## 4.5.0 (first release as `exceljs-lts`)
Published as `exceljs-lts` on npm. Install it as `exceljs@npm:exceljs-lts` to keep `require('exceljs')` working.
Pull request numbers without a prefix refer to https://github.com/cyberpusa-ai/exceljs.

### Files Excel had to repair
Excel showed "We found a problem with some content in '….xlsx'. Do you want us to try to recover as much as we can?" for these files:
- A sheet with a note (comment) plus a table, background picture or `extLst` (e.g. a dataBar): `legacyDrawing` is now written before `picture`, `tableParts` and `extLst` (#4, exceljs#1369)
- A streamed sheet (`stream.xlsx.WorkbookWriter`) with both sheet protection and an autoFilter (#3, exceljs#2685)
- A sheet with both `pageSetup.fitToPage` and `properties.outlineProperties`: `outlinePr` is now written before `pageSetUpPr` (#1, exceljs#1348)
- Conditional formatting with a number format after a read/write round trip: `styles.xml` contained `formatCode="[object Object]"`, Excel repaired the file and dropped the rule (#15)

### Crashes
- `TypeError: Cannot create property 'richText' on string ''` when reading shared strings that have plain text before rich-text runs (`<si><t></t><r>…</r></si>`), written by some third-party exporters (#18)
- `Worksheet name already exists: …` when opening a file whose sheet name is longer than 31 characters (#16, exceljs#3080)
- `TypeError: Cannot read properties of undefined` when chaining `workbook.stream.pipe(dest).on('finish', …)`: `StreamBuf.pipe()` now returns the destination (#10, exceljs#1595)

### Wrong data, no error
- Formula results `0`, `false` and `''` were dropped: `{formula: 'B1', result: 0}` read back as `{formula: 'B1'}` (#14, exceljs#2943)
- Styling one cell restyled every cell sharing its style (often the whole sheet) after loading a workbook (#17, exceljs#879, exceljs#791)
- `<b val="0"/>`, `<i val="0"/>`, `<strike val="0"/>` were read as on: loading and saving turned on bold, italic and strikethrough in fonts and conditional formats (#5)
- With `useSharedStrings: true`, the streaming writer gave every rich-text cell the text of the first one (#11, exceljs#2267)
- CJK, accented and emoji text could turn into `�` in the streaming reader when a character fell across a chunk boundary (#8)
- The streaming `WorkbookReader` sometimes returned `{sharedString: n}` instead of the cell text (#6, exceljs#683)
- `spliceRows()` left deleted rows in place when deleting near the end of a sheet (#19)
- Fractional image anchors (`tl: {col: 1.5, row: 1.5}`) used made-up cell sizes, so images landed in the wrong place (#13)

### Features
- Conditional formatting keeps Excel's "Stop If True" (`stopIfTrue`) through read/write; typings list all `cellIs` operators and make `priority` optional (#12, exceljs#2233)
- Notes can size their box to fit the text with `autoShape: true` (#20, exceljs#981)

### Types
- `Column.values` is typed `CellValue[]`; before, reading or assigning an array failed with TS2322 (#9)

### Security
- `npm audit` no longer flags `uuid` (GHSA-w5hq-g745-h8pq, dependency removed) or `tmp` (GHSA-9r2w-394v-53qc, now `^0.2.4`) (#2, exceljs#3055)

### Merged upstream after 4.4.0 but never released
- Pivot tables, with limitations (exceljs#2551, Michael)
- Table column style fix and `addTable` column type (exceljs#2649, Jamie Buck)
- `Column.values` no longer `readonly` (exceljs#2571, Siemienik Pawel)
- Docs typo fix (exceljs#2566, Mohammad Ahmadi); CI action updates (Christian Clauss)

For releases up to 4.4.0, see the Release History section of the README.
