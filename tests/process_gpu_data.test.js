import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { parseXlsx } from '../process_gpu_data.js';

test('parseXlsx reads sheet rows from an Excel workbook', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'gpu-market-'));

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('202512');
    worksheet.addRow(['显卡型号', 'Time spy平均跑分', '显存容量', 'TDP功耗', '12月价格', '涨跌幅', '翻车概率']);
    worksheet.addRow(['RTX 4060', 12345, '8GB', 115, 2999, 5, '★']);

    const filePath = join(dir, 'sample.xlsx');
    await workbook.xlsx.writeFile(filePath);

    const parsed = await parseXlsx(filePath);

    assert.deepEqual(parsed['202512'].headers, ['显卡型号', 'Time spy平均跑分', '显存容量', 'TDP功耗', '12月价格', '涨跌幅', '翻车概率']);
    assert.equal(parsed['202512'].rows[0][0], 'RTX 4060');
    assert.equal(parsed['202512'].rows[0][1], '12345');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
