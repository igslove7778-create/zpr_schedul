# 현재 운영 버전

이 폴더는 Google Apps Script와 Cloudflare Worker의 현재 운영 소스입니다.

- `apps-script/Code.gs`를 Google Sheets에 연결된 Apps Script 프로젝트에 붙여 넣습니다.
- `apps-script/appsscript.json`은 프로젝트 설정입니다.
- `cloudflare-worker/worker.js`는 Telegram 웹훅을 Apps Script로 안전하게 전달합니다.

실제 운영값은 소스 코드에 넣지 않습니다. 저장소 루트의 `examples/script-properties.example.json`을 보고 Apps Script 속성과 Cloudflare Secret에만 설정합니다.
