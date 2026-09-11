# 제피로스 통합 일정관리 시스템 (ZPR Schedule)

텔레그램 · 구글시트 · 구글캘린더 · 제피로스 웹사이트 일정 연동 시스템.

## 문서

- `docs/제피로스_통합일정관리_기능명세서_v0.2.docx` — 현행 기능명세서 (구글시트는 담당자 1명 편집, 전 구성원 열람 전용으로 확정)
- `docs/제피로스_통합일정관리_기능명세서_v0.1.docx` — 초안
- `docs/기능명세서_v0.2.txt` — v0.2 텍스트 추출본 (변경 내역 확인용)

## 구성 (제안)

- 원장: Google Sheets
- 서버 로직: Google Apps Script (웹앱 + 설치형 트리거)
- 텔레그램: Bot API (Webhook)
- 캘린더: Google Calendar API
