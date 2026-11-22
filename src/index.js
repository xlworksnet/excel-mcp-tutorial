// index.js — Claude가 엑셀 '업무 일정표'를 읽고 고칠 수 있게 하는 MCP 서버
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import { z } from 'zod';
import { runExcel } from './excel.js';
import { backupWorkbook } from './backup.js';

// ⚠️ stdio 서버에서는 console.log 사용 금지(Claude와 주고받는 통로가 깨진다).
//    확인용 출력은 console.error로.

// 오류 코드를 Claude가 사용자에게 전할 수 있는 문장으로 변환
const MESSAGES = {
  EXCEL_NOT_RUNNING: '엑셀이 실행 중이 아닙니다. 업무 일정 파일을 엑셀에서 먼저 열어 주세요.',
  NO_SCHEDULE_WORKBOOK: "'일정' 시트가 있는 파일이 열려 있지 않습니다. 업무 일정 파일을 열어 주세요.",
  MULTIPLE_WORKBOOKS: '업무 일정 파일이 여러 개 열려 있습니다. workbook에 파일 이름을 지정해 주세요. 열린 파일: ',
  EXCEL_BUSY: '엑셀이 바쁩니다. 셀 편집 중이거나 대화상자가 열려 있으면 닫고 다시 시도해 주세요.',
  TIMEOUT: '엑셀이 30초 안에 응답하지 않았습니다. 엑셀에 메시지 창이 떠 있는지 확인해 주세요.',
  ROW_MISMATCH: '행 번호와 업무명이 맞지 않아 아무것도 바꾸지 않았습니다. 목록을 다시 읽어 주세요. 그 행의 실제 업무명: ',
  FORMULA_CELL: '수식이 들어 있는 셀이라 덮어쓰지 않았습니다: ',
  DATE_ORDER: '완료일이 시작일보다 빠릅니다. 날짜를 다시 확인해 주세요.',
  MACRO_FAILED: '매크로를 실행하지 못했습니다. 파일이 .xlsm인지, 매크로(콘텐츠 사용)를 허용했는지 확인해 주세요. 상세: ',
  POWERSHELL_FAILED: 'PowerShell을 실행하지 못했습니다. 회사 PC라면 보안 정책으로 막혔을 수 있습니다. 상세: '
};

function friendly(error) {
  const [code, ...rest] = String(error.message).split(':');
  const detail = rest.join(':');
  return MESSAGES[code] ? MESSAGES[code] + detail : error.message;
}

// 도구 결과를 MCP 형식으로 감싼다, 오류는 isError: true 로 돌려준다.
async function respond(work) {
  try {
    const data = await work();
    return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
  } catch (error) {
    return { content: [{ type: 'text', text: friendly(error) }], isError: true };
  }
}

const workbookArg = z
  .string()
  .optional()
  .describe('대상 파일 이름(예: 업무일정.xlsm). 업무 일정 파일이 하나만 열려 있으면 비워 둡니다.');

const dateArg = z.string().regex(/^\d{4}-\d{2}-\d{2}$/);

function createServer() {
  const server = new McpServer({ name: 'excel-schedule', version: '1.0.0' });

  // 6강: 열린 통합문서 정보 읽기
  server.registerTool(
    'get_workbook_info',
    {
      title: '열린 엑셀 파일 정보',
      description: '엑셀에 열려 있는 파일 목록과 각 파일의 시트 이름을 알려 줍니다. 다른 도구를 쓰기 전에 어떤 파일이 열려 있는지 확인할 때 씁니다.',
      inputSchema: z.object({}),
      annotations: { readOnlyHint: true }
    },
    async () =>
      respond(() =>
        runExcel(String.raw`
          $xl = Get-Excel
          $list = @()
          foreach ($wb in $xl.Workbooks) {
            $sheets = @()
            foreach ($ws in $wb.Worksheets) { $sheets += $ws.Name }
            $list += [ordered]@{ name = $wb.Name; path = $wb.FullName; sheets = $sheets }
          }
          @{ count = $list.Count; workbooks = $list }
        `)
      )
  );

  // 7강: 업무 목록 읽기
  server.registerTool(
    'list_tasks',
    {
      title: '업무 목록 읽기',
      description:
        "'일정' 시트의 업무를 읽습니다. 계획진척률은 오늘 날짜 기준으로 엑셀 수식이 계산한 값이고, 실적진척률은 사람이 입력한 값입니다. 실적이 계획보다 낮으면 delayed가 true입니다. 각 업무의 row(행 번호)는 수정·삭제할 때 그대로 넘겨야 합니다. 숨긴 행은 빼고 읽습니다.",
      inputSchema: z.object({
        workbook: workbookArg,
        owner: z.string().optional().describe('이 담당자의 업무만 읽습니다.'),
        delayed_only: z.boolean().optional().describe('true면 지연된 업무만 읽습니다.')
      }),
      annotations: { readOnlyHint: true }
    },
    async (args) =>
      respond(() =>
        runExcel(
          String.raw`
            $wb = Get-ScheduleBook $arg.workbook
            $ws = $wb.Worksheets.Item('일정')
            $last = Get-LastRow $ws
            $list = @()
            for ($r = 2; $r -le $last; $r++) {
              if ($ws.Rows.Item($r).Hidden) { continue }
              $task = $ws.Cells.Item($r, 2).Text
              if ($task -eq '') { continue }
              $owner = $ws.Cells.Item($r, 3).Text
              $delayed = ($ws.Cells.Item($r, 8).Text -eq '지연')
              if ($arg.owner -and $owner -ne $arg.owner) { continue }
              if ($arg.delayed_only -and -not $delayed) { continue }
              $list += [ordered]@{
                row = $r; no = $ws.Cells.Item($r, 1).Text; task = $task; owner = $owner
                start = $ws.Cells.Item($r, 4).Text; end = $ws.Cells.Item($r, 5).Text
                planned = $ws.Cells.Item($r, 6).Text; actual = $ws.Cells.Item($r, 7).Text; delayed = $delayed
              }
            }
            @{ workbook = $wb.Name; count = $list.Count; tasks = $list }
          `,
          args
        )
      )
  );

  // 8강: 업무 추가
  server.registerTool(
    'add_task',
    {
      title: '업무 추가',
      description: "'일정' 시트 맨 아래에 업무를 한 줄 추가하고 담당자를 할당합니다. 실적진척률은 0%로 시작합니다.",
      inputSchema: z.object({
        workbook: workbookArg,
        task: z.string().min(1).describe('업무명'),
        owner: z.string().optional().describe('담당자'),
        start: dateArg.describe('시작일, YYYY-MM-DD 형식'),
        end: dateArg.describe('완료일, YYYY-MM-DD 형식')
      })
    },
    async (args) =>
      respond(async () => {
        if (args.end < args.start) throw new Error('DATE_ORDER');
        const backup = await backupWorkbook(args.workbook);
        const result = await runExcel(
          String.raw`
            $wb = Get-ScheduleBook $arg.workbook
            $ws = $wb.Worksheets.Item('일정')
            $r = (Get-LastRow $ws) + 1
            $no = [int]$ws.Application.WorksheetFunction.Max($ws.Range('A:A')) + 1
            $ws.Cells.Item($r, 1).Value2 = $no
            $ws.Cells.Item($r, 2).Value2 = $arg.task
            if ($arg.owner) { $ws.Cells.Item($r, 3).Value2 = $arg.owner }
            $ws.Cells.Item($r, 4).Value2 = ([datetime]::ParseExact($arg.start, 'yyyy-MM-dd', $null)).ToOADate()
            $ws.Cells.Item($r, 5).Value2 = ([datetime]::ParseExact($arg.end, 'yyyy-MM-dd', $null)).ToOADate()
            $ws.Range("D$($r):E$($r)").NumberFormat = 'yyyy-mm-dd'
            $ws.Cells.Item($r, 6).Formula = "=IF(OR(D$($r)="""",E$($r)=""""),"""",MAX(0,MIN(1,NETWORKDAYS(D$($r),MIN(TODAY(),E$($r)))/NETWORKDAYS(D$($r),E$($r)))))"
            $ws.Cells.Item($r, 7).Value2 = 0
            $ws.Range("F$($r):G$($r)").NumberFormat = '0%'
            $ws.Cells.Item($r, 8).Formula = "=IF(OR(F$($r)="""",G$($r)=""""),"""",IF(G$($r)<F$($r),""지연"",""""))"
            @{ row = $r; no = $no; task = $arg.task }
          `,
          args
        );
        return { ...result, backup };
      })
  );

  // 8강: 업무 수정 (행 번호 + 업무명 이중 확인, 수식 셀 보호)
  server.registerTool(
    'update_task',
    {
      title: '업무 수정',
      description:
        '업무의 담당자, 시작일, 완료일, 실적진척률을 바꿉니다. 실적진척률은 0~100 사이 숫자(%)로 넘깁니다. row와 task는 list_tasks에서 받은 값을 그대로 넘깁니다. 둘이 맞지 않으면 아무것도 바꾸지 않습니다.',
      inputSchema: z.object({
        workbook: workbookArg,
        row: z.number().int().min(2).describe('list_tasks에서 받은 행 번호'),
        task: z.string().describe('그 행의 업무명(확인용)'),
        owner: z.string().optional(),
        start: dateArg.optional().describe('YYYY-MM-DD'),
        end: dateArg.optional().describe('YYYY-MM-DD'),
        actual: z.number().min(0).max(100).optional().describe('실적진척률(%), 예: 50')
      }),
      annotations: { idempotentHint: true }
    },
    async (args) =>
      respond(async () => {
        const backup = await backupWorkbook(args.workbook);
        const result = await runExcel(
          String.raw`
            $wb = Get-ScheduleBook $arg.workbook
            $ws = $wb.Worksheets.Item('일정')
            $current = $ws.Cells.Item($arg.row, 2).Text
            if ($current -ne $arg.task) { throw ('ROW_MISMATCH:' + $current) }

            $changes = @{}
            if ($arg.owner) { $changes[3] = $arg.owner }
            if ($arg.start) { $changes[4] = ([datetime]::ParseExact($arg.start, 'yyyy-MM-dd', $null)).ToOADate() }
            if ($arg.end)   { $changes[5] = ([datetime]::ParseExact($arg.end, 'yyyy-MM-dd', $null)).ToOADate() }
            if ($null -ne $arg.actual) { $changes[7] = $arg.actual / 100 }

            # 먼저 모두 검사한 뒤에 쓴다, 하나라도 수식 셀이면 아무것도 바꾸지 않는다.
            foreach ($col in $changes.Keys) {
              $cell = $ws.Cells.Item($arg.row, $col)
              if ($cell.HasFormula) { throw ('FORMULA_CELL:' + $cell.Address($false, $false)) }
            }
            foreach ($col in $changes.Keys) { $ws.Cells.Item($arg.row, $col).Value2 = $changes[$col] }
            @{ row = $arg.row; task = $current; changed = $changes.Count }
          `,
          args
        );
        return { ...result, backup };
      })
  );

  // 11강: 업무 삭제 (사용자 확인 필수)
  server.registerTool(
    'delete_task',
    {
      title: '업무 삭제',
      description:
        '업무 한 줄을 삭제합니다. 되돌릴 수 없으므로, 부르기 전에 반드시 삭제할 업무를 사용자에게 보여 주고 삭제해도 되는지 물어본 뒤 사용자가 동의했을 때만 부르세요. 여러 개를 삭제할 때는 목록을 먼저 보여 주세요.',
      inputSchema: z.object({
        workbook: workbookArg,
        row: z.number().int().min(2),
        task: z.string().describe('그 행의 업무명(확인용)')
      }),
      annotations: { destructiveHint: true }
    },
    async (args) =>
      respond(async () => {
        const backup = await backupWorkbook(args.workbook);
        const result = await runExcel(
          String.raw`
            $wb = Get-ScheduleBook $arg.workbook
            $ws = $wb.Worksheets.Item('일정')
            $current = $ws.Cells.Item($arg.row, 2).Text
            if ($current -ne $arg.task) { throw ('ROW_MISMATCH:' + $current) }
            [void]$ws.Rows.Item($arg.row).Delete()
            @{ deleted = $current }
          `,
          args
        );
        return { ...result, backup };
      })
  );

  // 9강: VBA 매크로 실행
  server.registerTool(
    'update_schedule',
    {
      title: '일정 업데이트',
      description: "UpdateSchedule 매크로를 실행합니다. 진척률을 다시 계산하고 완료된 업무는 회색, 지연된 업무는 빨간색으로 표시한 뒤 '완료 n건, 지연 n건'을 돌려줍니다.",
      inputSchema: z.object({ workbook: workbookArg })
    },
    async (args) =>
      respond(async () => {
        const backup = await backupWorkbook(args.workbook);
        const result = await runExcel(
          String.raw`
            $wb = Get-ScheduleBook $arg.workbook
            try { $summary = $wb.Application.Run("'" + $wb.Name + "'!UpdateSchedule") }
            catch { throw ('MACRO_FAILED:' + $_.Exception.Message) }
            @{ summary = $summary }
          `,
          args
        );
        return { ...result, backup };
      })
  );

  return server;
}

serveStdio(createServer);
console.error('excel-schedule MCP 서버가 시작되었습니다.');
