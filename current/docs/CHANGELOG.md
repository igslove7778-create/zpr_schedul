# Changelog

## 2026-09-17

### Changed

- Google Calendar 설명란에 `대상:` 접두어 없이 구성원 이름만 입력해도 대상자로 인식
- 쉼표/공백 구분 모두 지원
- 1명, 2명, 3명 이상 지원
- `전체`, `전부` 지원
- 기존 `대상:` / `대상자:` 방식 하위 호환 유지
- 일반 문장이 섞이면 대상자 오인식하지 않도록 구성원 명단 전체 일치 검증

### Verified

- `Code.gs` JavaScript 문법 검사 통과
- Google Calendar 서버에 `구글위젯테스트` 이벤트 저장 확인

### To verify

- 캘린더 변경감지 트리거가 시트에 자동 반영하는지
- Script Properties `TEAM_CAL_ID`가 현재 `제피로스 팀 일정 (관리자)`를 가리키는지
