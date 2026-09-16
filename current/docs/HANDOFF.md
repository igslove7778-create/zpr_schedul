# 제피로스 일정관리 인수인계

기준일: 2026-09-17

## 1. 목적

직원들이 Google Calendar, Google Sheet, Telegram 중 편한 채널에서 일정을 등록/수정하고, 중앙 일정 시트와 팀/개인 캘린더가 최대한 자동으로 일치하도록 만드는 시스템입니다.

## 2. 현재 운영 구조

- 관리자: 김형원
- 협업/인수인계 대상: 성도형
- GitHub: `sdh0817/zpr_schedule`
- 팀 캘린더: `제피로스 팀 일정 (관리자)`
- 직원별 전용 캘린더: `제피로스 일정 - <이름>`
- 중앙 원장: Google Sheets의 `일정` 시트

소유권/공유 권한은 Google Calendar와 Google Sheets에서 별도로 관리합니다. 저장소에는 실제 운영 계정, 캘린더 ID, 토큰을 새로 하드코딩하지 않습니다.

## 3. 일정 시트 주요 열

`날짜 / 시작시간 / 종료시간 / 일정 / 장소 / 대상자 / 알람 / 파일 / 일정ID / 등록자 / 등록채널 / 캘린더ID / 수정일시 / 상태 / 리마인드`

구성원 시트 핵심 헤더:

`이름 / 구글 계정 / 텔레그램 ID / 권한 / 사용여부 / 인증코드 / 캘린더ID`

## 4. Google Calendar → Sheet 필드 대응

- 제목 → `일정`
- 날짜/시간 → `날짜`, `시작시간`, `종료시간`
- 위치 → `장소`
- 설명의 대상자 표기 → `대상자`

## 5. 대상자 입력 규칙

캘린더에는 별도 대상자 필드가 없으므로 설명란을 사용합니다.

### 새 입력 방식

설명란 전체가 등록된 구성원 이름들만이면 대상자로 읽습니다.

```text
김형원
김형원, 김명재
김형원 김명재 성도형
전체
전부
```

1명/2명/3명 이상 모두 가능하며 쉼표와 공백을 모두 지원합니다.

### 기존 방식도 유지

```text
대상: 김형원, 김명재
대상자: 성도형
```

### 오인식 방지

`김명재 부장과 회의 후 자료 전달`처럼 구성원 이름 외의 일반 문장이 섞이면 대상자로 처리하지 않습니다.

## 6. 동기화 방향

### Sheet → Calendar

`syncCalendar_()` 중심입니다.

- 팀 캘린더 생성/수정
- 직원별 복사본 생성/수정
- 삭제 상태 처리
- 1분 주기 트리거 가능

### Calendar → Sheet

`onCalendarChange_()` 중심입니다.

- Google Calendar Advanced API sync token 사용
- 토큰 만료 시 제한된 전체 재읽기로 복구
- 변경이 있으면 시트 반영 후 정렬

주요 함수:

- `zRunCalendarSyncNow` — 즉시 동기화
- `zInstallCalendarReverseTrigger` — Calendar → Sheet 변경감지 트리거 재설치
- `zInstallCalendarSyncTrigger` — Sheet → Calendar 1분 동기화 트리거
- `zInstallCoreTriggers` — 핵심 자동 실행 트리거 재설치
- `zProvisionMemberCalendars` — 직원 전용 캘린더 생성/공유
- `zEnableMemberCalendarSync` — 직원 캘린더 양방향 동기화 활성화
- `zRemoveDeletedScheduleCopies` — 삭제 일정의 개인 캘린더 잔여 복사본 정리

## 7. 삭제 원칙

삭제는 reverse sync보다 먼저 처리합니다. 시트에서 삭제된 일정/삭제 상태 일정의 팀 캘린더 및 직원별 복사본을 먼저 정리하여 남은 캘린더 이벤트가 시트에 다시 살아나는 것을 막습니다.

## 8. 정렬

시트 날짜 정렬이 느리면 여러 사용자가 동시에 입력할 때 다른 행을 건드릴 수 있어 빠른 정렬 로직을 적용했습니다. 현재 `Code.gs`는 빠른 정렬/빈 행 정리 버전을 베이스로 대상자 설명 파서를 추가 수정한 버전입니다.

## 9. 최근 확인 사항

- 휴대폰 Google Calendar 앱에서 일부 캘린더가 안 보였으나 앱 재설치 후 정상 표시됨. 삼성 캘린더와 서버에는 정상 존재했음.
- 팀 캘린더에 테스트 일정 생성 시 Google Calendar 서버에는 제목/위치/설명이 정상 저장됨.
- 당시 Sheet 자동 반영이 즉시 일어나지 않아 Calendar → Sheet 변경감지 트리거 상태를 확인 중.

확인 순서:

1. Apps Script에서 `onCalendarChange_` 수동 실행
2. 시트에 들어오면 reverse trigger 문제로 판단
3. `zInstallCalendarReverseTrigger` 실행
4. 그래도 안 되면 Script Properties의 `TEAM_CAL_ID`가 현재 팀 캘린더를 가리키는지 확인

## 10. 협업자가 이어서 작업하려면

1. Google Sheet 편집 권한을 부여받기
2. 필요한 캘린더에 일정 변경 권한을 부여받기
3. 이 저장소를 pull/clone
4. `current/apps-script/Code.gs`를 기준으로 수정
5. Apps Script에 반영 후 저장
6. 자기 계정으로 수동 실행/새 트리거 생성이 필요하면 Google 권한 승인

설치형 트리거는 **트리거를 만든 계정 권한으로 실행**됩니다. 코드가 수정돼도 기존 트리거가 살아 있으면 그 트리거 소유 계정 권한으로 최신 저장 코드를 실행합니다.

## 11. 수정 원칙

- 전체 코드 교체를 반복하기 전에 원인 확인
- 기존 `일정ID`/`캘린더ID` 연결을 함부로 초기화하지 않기
- 새로운 팀 캘린더를 자동으로 만들지 않기
- 운영값은 Script Properties/Secret에서 관리
- 테스트 후 로그와 실제 Calendar 서버 상태를 함께 확인
- 사용자가 직접 해야 할 클릭/실행 횟수를 최소화
