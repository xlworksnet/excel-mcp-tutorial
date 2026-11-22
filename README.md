# excel-mcp-tutorial

Claude와 엑셀 업무 일정표를 연결하는 MCP 서버 예제입니다.
[엑셀웍스](https://xlworks.net) "AI-Excel 연동" 강좌에서 쓰는 코드입니다.

## 할 수 있는 것

엑셀에 `업무일정.xlsm`을 열어 놓고 Claude Desktop에서 다음과 같이 시킬 수 있습니다.

- 지연된 업무 알려줘
- 'OO사 견적서 발송' 업무를 추가하고 홍길동에게 할당해줘. 기간은 다음 주 월요일부터 금요일까지
- 매출 보고서 작성 실적 70%로 입력해줘
- 일정 업데이트해줘 (VBA 매크로 실행)

## 필요한 것

- Windows 10 또는 11
- 엑셀 데스크톱(Microsoft 365 또는 설치형). 웹용 엑셀은 안 됩니다.
- Claude Desktop(Windows)
- Node.js 20 이상

## 폴더 구성

- `steps/hello.js` : 연결 확인용 서버(5강)
- `src/index.js` : MCP 서버. 도구 6개를 등록합니다.
- `src/excel.js` : PowerShell을 실행해 엑셀을 COM으로 조작합니다.
- `src/backup.js` : 고치기 전에 통합문서 사본을 만듭니다.
- `examples/업무일정.xlsx` : 예제 파일. 매크로를 넣어 `.xlsm`으로 저장해서 씁니다.
- `examples/UpdateSchedule.bas` : 예제 파일에 넣을 VBA 모듈
- `manifest.json` : `.mcpb`로 묶을 때 쓰는 설정

## 예제 파일 구성 ('일정' 시트)

번호, 업무명, 담당, 시작일, 완료일, 계획진척률(수식), 실적진척률(입력), 지연(수식)

- 계획진척률 = 시작일부터 오늘까지 작업일수 ÷ 전체 작업일수
- 실적진척률이 계획진척률보다 낮으면 '지연'으로 표시됩니다.

## 설치 (개발용)

```
npm install
```

`claude_desktop_config.json`에 추가합니다. 파일 위치는 Claude Desktop 버전과 설치 방법에 따라 다릅니다.

- `%LOCALAPPDATA%\Claude-Data\claude_xxxxxxxx\claude_desktop_config.json` (최신 버전)
- `%APPDATA%\Claude\claude_desktop_config.json` (예전 버전)

찾지 못하면 PowerShell에서 다음 명령으로 위치를 확인합니다.

```
Get-ChildItem $env:APPDATA, $env:LOCALAPPDATA -Recurse -Filter claude_desktop_config.json -ErrorAction SilentlyContinue | Select-Object FullName
```

```json
{
  "mcpServers": {
    "excel-schedule": {
      "command": "node",
      "args": ["C:\\excel-mcp-tutorial\\src\\index.js"]
    }
  }
}
```

Claude Desktop을 완전히 종료(트레이 아이콘 → 종료)한 뒤 다시 실행합니다.

## .mcpb로 묶기

```
npm install --omit=dev
npx @anthropic-ai/mcpb pack . excel-schedule-1.0.0.mcpb
```

## 도구 목록

- `get_workbook_info` : 열린 파일과 시트 목록
- `list_tasks` : 업무 목록 읽기(담당자, 지연 업무만)
- `add_task` : 업무 추가하고 담당자 할당
- `update_task` : 담당, 시작일, 완료일, 실적진척률 수정(행 번호 + 업무명 확인, 수식 셀 보호)
- `delete_task` : 업무 삭제(사용자 확인 후)
- `update_schedule` : `UpdateSchedule` 매크로 실행

## 라이선스

MIT
