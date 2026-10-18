import fs from 'node:fs/promises';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

const projectDir = path.dirname(decodeURIComponent(new URL(import.meta.url).pathname.replace(/^\//, '')));
const sourceDir = path.join(projectDir, '照片', '00-散落的名片');
const dataPath = path.join(projectDir, '名片資料庫.json');
const batchPath = path.join(projectDir, '名片待歸檔.json');
const batch = JSON.parse(await fs.readFile(batchPath, 'utf8'));
const existing = await fs.access(dataPath).then(async () => JSON.parse(await fs.readFile(dataPath, 'utf8'))).catch(() => ({ records: [] }));
const existingBySource = new Map((existing.records || []).map(record => [record.source, record]));

for (const record of batch.records || []) {
  if (!['建築師', '業主'].includes(record.category)) throw new Error(`分類無效：${record.category}`);
  const source = path.join(sourceDir, record.source);
  const destinationDir = path.join(projectDir, '照片', record.category, '已歸檔');
  const destination = path.join(destinationDir, record.source);
  await fs.mkdir(destinationDir, { recursive: true });
  if (!(await fs.access(source).then(() => true).catch(() => false)) && !(await fs.access(destination).then(() => true).catch(() => false))) {
    throw new Error(`找不到照片：${record.source}`);
  }
  if (await fs.access(source).then(() => true).catch(() => false)) await fs.rename(source, destination);
  existingBySource.set(record.source, record);
}

const merged = { records: [...existingBySource.values()] };
await fs.writeFile(dataPath, JSON.stringify(merged, null, 2), 'utf8');
const result = spawnSync(process.execPath, [path.join(projectDir, '更新名片資料庫.mjs')], { cwd: projectDir, encoding: 'utf8' });
if (result.status !== 0) throw new Error(`名片資料庫更新失敗：${result.stderr || result.stdout}`);
console.log(JSON.stringify({ archived: (batch.records || []).map(record => ({ source: record.source, category: record.category })), database: path.join(projectDir, '名片資料庫.xlsx') }, null, 2));
