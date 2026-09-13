# 제피로스 일정관리

텔레그램, Google Sheets, Google Calendar를 연결한 업무 일정관리 자동화입니다.

## 현재 운영 소스

현재 배포에 맞춘 공개용 소스는 `current/`에 있습니다.

```
current/
  apps-script/Code.gs          Google Apps Script 본문
  apps-script/appsscript.json  Apps Script 설정
  cloudflare-worker/worker.js  Telegram 즉시 응답용 Cloudflare Worker
  docs/직원_사용설명서.md        직원용 텔레그램 사용법
examples/
  script-properties.example.json  비밀값을 제외한 설정 항목 예시
```

루트의 `apps-script/`, `docs/`, `tests/`는 이전 개발 버전 기록으로 보존합니다. 새 배포나 복구에는 `current/` 파일을 사용합니다.

## 공개 저장소 보안 원칙

- 봇 토큰, 웹훅 키, 실제 시트 ID, 실제 캘린더 ID는 이 저장소에 넣지 않습니다.
- 실제 값은 Google Apps Script **스크립트 속성**과 Cloudflare **Secret**에만 넣습니다.
- `current/apps-script/Code.gs`에는 실제 운영 식별값이 없습니다.

## 설치 요약

1. 운영용 Google Sheets에서 **확장 프로그램 → Apps Script**를 엽니다.
2. `current/apps-script/Code.gs`와 `appsscript.json`을 반영합니다.
3. 스크립트 속성에 `BOT_TOKEN`, `TEAM_CAL_ID`, `WEB_APP_URL`, `WEBHOOK_KEY`를 넣습니다. 독립형 Apps Script라면 `SPREADSHEET_ID`도 넣습니다.
4. `current/cloudflare-worker/worker.js`를 Cloudflare Worker에 배포합니다. Worker에는 `APPS_SCRIPT_URL`, `WEBHOOK_KEY`를 Secret으로 설정합니다.
5. Apps Script 속성 `TELEGRAM_WEBHOOK_URL`에 Worker 주소를 넣고 `zRegisterWebhook`을 실행합니다.

직원별 업무 캘린더는 구성원 시트에 이름과 Gmail을 입력한 뒤, 최종 운영자 계정으로 `직원 전용 캘린더 만들기/공유`를 실행합니다. 실행한 계정이 새 업무 캘린더의 소유자가 됩니다.

## 사용

- 일정 등록: `9/15 오후 2시 팀 회의`
- 일정 등록: `일정 2026-09-15 14:00 팀 회의`
- 오늘 조회: `오늘`

텔레그램 조회와 등록 답장은 보통 수 초 안에 도착하며, 팀 캘린더 반영은 최대 약 1분 걸릴 수 있습니다.
