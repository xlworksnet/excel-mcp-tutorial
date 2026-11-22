// backup.js — 엑셀을 고치기 전에 통합문서 사본 만들기
import { readdir, unlink, mkdir } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { runExcel } from './excel.js';

// .mcpb 설정 화면에서 넘어오는 값, 비어 있으면 기본값 사용
const BACKUP_DIR = process.env.BACKUP_DIR || path.join(os.tmpdir(), 'excel-mcp-backup');
const KEEP_COUNT = Number(process.env.KEEP_COUNT) || 20;
const BACKUP_ON = process.env.BACKUP_ENABLED !== 'false';

function stamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

/** 변경 직전에 호출, 만든 백업 파일 경로를 돌려준다. */
export async function backupWorkbook(workbook) {
  if (!BACKUP_ON) return null;
  await mkdir(BACKUP_DIR, { recursive: true });

  const result = await runExcel(
    String.raw`
      $wb = Get-ScheduleBook $arg.workbook
      $base = [IO.Path]::GetFileNameWithoutExtension($wb.Name)
      $ext = [IO.Path]::GetExtension($wb.Name)
      if (-not $ext) { $ext = '.xlsx' }
      $file = Join-Path $arg.dir ($base + '_' + $arg.stamp + $ext)
      $wb.SaveCopyAs($file)
      @{ file = $file; base = $base }
    `,
    { workbook, dir: BACKUP_DIR, stamp: stamp() }
  );

  await prune(result.base);
  return result.file;
}

// 같은 파일의 백업이 KEEP_COUNT개를 넘으면 오래된 것부터 삭제
async function prune(base) {
  const files = (await readdir(BACKUP_DIR))
    .filter((f) => f.startsWith(base + '_'))
    .sort(); // 파일 이름에 날짜·시간이 들어 있어 이름순 = 시간순
  const extra = files.length - KEEP_COUNT;
  for (let i = 0; i < extra; i++) {
    await unlink(path.join(BACKUP_DIR, files[i]));
  }
}
