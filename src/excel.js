// excel.js — Node.js에서 PowerShell을 실행해 열려 있는 엑셀을 COM으로 조작
import { spawn } from 'node:child_process';

// 모든 PowerShell 스크립트 앞에 붙는 공통 부분
const PRELUDE = String.raw`
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 실행 중인 엑셀에 연결, 엑셀이 없으면 EXCEL_NOT_RUNNING 오류를 낸다.
function Get-Excel {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
  catch { throw 'EXCEL_NOT_RUNNING' }
}

# '일정' 시트가 있는 통합문서 찾기, 이름을 주면 그 파일만 찾는다.
function Get-ScheduleBook($name) {
  $xl = Get-Excel
  $books = @()
  foreach ($wb in $xl.Workbooks) {
    foreach ($ws in $wb.Worksheets) {
      if ($ws.Name -eq '일정') { $books += $wb; break }
    }
  }
  if ($name) { $books = @($books | Where-Object { $_.Name -eq $name }) }
  if ($books.Count -eq 0) { throw 'NO_SCHEDULE_WORKBOOK' }
  if ($books.Count -gt 1) {
    throw ('MULTIPLE_WORKBOOKS:' + (($books | ForEach-Object { $_.Name }) -join ', '))
  }
  return ,$books[0]
}

# B열(업무명)에서 마지막으로 값이 있는 행 번호
function Get-LastRow($ws) {
  return $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row   # -4162 = xlUp
}
`;

/**
 * PowerShell 스크립트를 실행하고 결과(JSON)를 돌려받는다.
 * @param {string} body   실행할 PowerShell 코드, 마지막에 해시테이블 하나를 돌려줘야 한다.
 * @param {object} input  스크립트에 넘길 값, 스크립트 안에서 $arg 로 읽는다.
 */
export function runExcel(body, input = {}, { timeoutMs = 30000 } = {}) {
  // 작은따옴표 문자열 안에 넣으므로 ' 를 '' 로 바꾼다.
  const json = JSON.stringify(input).replace(/'/g, "''");
  const script = `${PRELUDE}
$arg = ConvertFrom-Json '${json}'
try {
  $result = & { ${body} }
  @{ ok = $true; data = $result } | ConvertTo-Json -Depth 6 -Compress
} catch {
  $code = $_.Exception.HResult
  $msg = $_.Exception.Message
  # 셀 편집 중이거나 대화상자가 열려 있으면 엑셀이 호출을 거절한다.
  if ($code -eq -2146777998 -or $code -eq -2147418111) { $msg = 'EXCEL_BUSY' }
  @{ ok = $false; error = $msg } | ConvertTo-Json -Compress
}`;

  // 따옴표 문제를 피하려고 스크립트를 Base64로 바꿔 넘긴다(UTF-16LE).
  const encoded = Buffer.from(script, 'utf16le').toString('base64');

  return new Promise((resolve, reject) => {
    const ps = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded],
      { windowsHide: true }
    );
    const out = [];
    const err = [];
    ps.stdout.on('data', (d) => out.push(d));
    ps.stderr.on('data', (d) => err.push(d));

    const timer = setTimeout(() => {
      ps.kill();
      reject(new Error('TIMEOUT'));
    }, timeoutMs);

    ps.on('error', (e) => {
      clearTimeout(timer);
      reject(new Error('POWERSHELL_FAILED:' + e.message));
    });

    ps.on('close', () => {
      clearTimeout(timer);
      const text = Buffer.concat(out).toString('utf8').trim();
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch {
        const detail = Buffer.concat(err).toString('utf8').trim() || text || '응답 없음';
        return reject(new Error('POWERSHELL_FAILED:' + detail));
      }
      if (parsed.ok) resolve(parsed.data);
      else reject(new Error(parsed.error));
    });
  });
}
