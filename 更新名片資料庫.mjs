import fs from 'node:fs/promises';
import path from 'node:path';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const projectDir = path.dirname(decodeURIComponent(new URL(import.meta.url).pathname.replace(/^\//, '')));
const dataPath = path.join(projectDir, '名片資料庫.json');
const outputPath = path.join(projectDir, '名片資料庫.xlsx');
const data = JSON.parse(await fs.readFile(dataPath, 'utf8'));
const records = [...(data.records || [])].sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.name).localeCompare(String(b.name), 'zh-Hant'));

function toDate(value) {
  const [year, month, day] = String(value).split('-').map(Number);
  return new Date(year, month - 1, day);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add('名片資料庫');
sheet.showGridLines = false;
sheet.tabColor = '#2E4057';
sheet.getRange('A1:L1').values = [[
  '編號', '日期', '人脈分類', '姓名', '公司／機構', '職位', '電話', '傳真', 'Email', '照片檔名', '照片路徑', '備註'
]];
const rows = records.map((record, index) => [
  String(index + 1).padStart(3, '0'),
  toDate(record.date),
  record.category,
  record.name,
  record.company,
  record.position,
  record.phone,
  record.fax,
  record.email,
  record.source,
  `照片/${record.category}/已歸檔/${record.source}`,
  record.notes || '',
]);
if (rows.length) sheet.getRange(`A2:L${rows.length + 1}`).values = rows;

sheet.getRange('A1:L1').format = {
  fill: '#2E4057',
  font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' },
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
  wrapText: true,
  borders: { preset: 'all', style: 'thin', color: '#BFBFBF' },
};
if (rows.length) {
  const body = sheet.getRange(`A2:L${rows.length + 1}`);
  body.format = {
    font: { name: 'Arial', size: 10, color: '#111827' },
    verticalAlignment: 'center',
    wrapText: false,
    borders: { preset: 'all', style: 'thin', color: '#BFBFBF' },
  };
  sheet.getRange(`A2:C${rows.length + 1}`).format.horizontalAlignment = 'center';
  sheet.getRange(`B2:B${rows.length + 1}`).format.numberFormat = 'yyyy-mm-dd';
  for (let row = 2; row <= rows.length + 1; row += 1) {
    if (row % 2 === 0) sheet.getRange(`A${row}:L${row}`).format.fill = '#D6E4F0';
    if (sheet.getRange(`D${row}`).values[0][0] === '（待查）') sheet.getRange(`D${row}`).format.font = { name: 'Arial', size: 10, color: '#FF0000', bold: true };
  }
  sheet.tables.add(`A1:L${rows.length + 1}`, true, '名片資料庫');
}
const widths = { A: 8, B: 13, C: 12, D: 14, E: 34, F: 24, G: 25, H: 22, I: 32, J: 30, K: 38, L: 42 };
for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}:${column}`).format.columnWidth = width;
sheet.freezePanes.freezeRows(1);

workbook.recalculate();
const inspect = await workbook.inspect({ kind: 'table', range: `名片資料庫!A1:L${Math.max(2, rows.length + 1)}`, include: 'values,formulas', tableMaxRows: 20, tableMaxCols: 12 });
console.log(inspect.ndjson);
const preview = await workbook.render({ sheetName: '名片資料庫', range: `A1:L${Math.max(2, rows.length + 1)}`, scale: 1, format: 'png' });
await fs.mkdir(path.join(projectDir, '處理紀錄'), { recursive: true });
await fs.writeFile(path.join(projectDir, '處理紀錄', '名片資料庫預覽.png'), new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`已更新名片資料庫：${outputPath}`);
