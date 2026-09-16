# 현재 운영 버전

이 폴더는 Google Apps Script와 Cloudflare Worker의 현재 운영 소스입니다.

## Apps Script 소스

최신 Apps Script 전체 소스는 `apps-script/` 아래의 여러 `.gs` 파일로 나눠 저장했습니다.

- `Code.gs`
- `Zephyrus_02.gs` ~ `Zephyrus_15.gs`
- `appsscript.json`

Google Apps Script에서는 같은 프로젝트 안의 여러 `.gs` 파일이 하나의 전역 코드처럼 동작합니다. **한 파일만 복사하면 안 되고, 위 `.gs` 파일 전체가 현재 운영 코드입니다.**

GitHub 연동의 파일 크기 제한 때문에 기존 단일 `Code.gs`를 함수 경계에서만 안전하게 분할했으며, 기능 구조 자체를 분리하거나 변경한 것은 아닙니다.

`cloudflare-worker/worker.js`는 Telegram 웹훅을 Apps Script로 안전하게 전달합니다.

실제 운영값(계정, 캘린더 ID, 토큰 등)은 소스 코드에 새로 넣지 말고 Script Properties/Secret에서 관리합니다.

## 2026-09-17 운영 기준 변경

Google Calendar의 설명란에 `대상:` 접두어 없이 구성원 이름만 적어도 대상자로 인식하도록 개선했습니다.

예:

```text
김형원
김형원, 김명재
김형원, 김명재, 성도형
김형원 김명재 성도형
전체
전부
```

기존 `대상: 김형원, 김명재` / `대상자: 성도형` 형식도 계속 지원합니다. 일반 문장이 섞인 설명은 대상자로 오인하지 않도록 구성원 명단과 전체 일치 검증을 합니다.

## 인수인계 문서

- `docs/HANDOFF.md` — 현재 구조, 권한, 동기화, 최근 이슈
- `docs/TEST_CHECKLIST.md` — 수정 후 회귀 테스트
- `docs/CHANGELOG.md` — 최근 변경 기록

현재 저장소의 `current/`를 운영 기준으로 보고 작업합니다.
