# GlobalMM 실시간 모니터링 시스템

Excel(`.xlsb`) 파일을 실시간으로 읽어 웹 브라우저로 제공하는 사내 모니터링 시스템.  
**Windows 전용** (win32com 의존), 관리자 권한 불필요, 인터넷 차단 환경에서 동작.

---

## 목차

1. [전체 아키텍처](#1-전체-아키텍처)
2. [파일 구조](#2-파일-구조)
3. [실행 방법](#3-실행-방법)
4. [데이터 흐름](#4-데이터-흐름)
5. [모듈별 상세](#5-모듈별-상세)
   - [serve.py](#servepy)
   - [reader_common.py](#reader_commonpy)
   - [reader_ssf.py](#reader_ssfpy)
   - [reader_sso.py](#reader_ssopy)
6. [프론트엔드 구조](#6-프론트엔드-구조)
   - [index.html](#indexhtml)
   - [ssfo_pnl.html](#ssfo_pnlhtml)
   - [ssfo_idx.html](#ssfo_idxhtml)
   - [sso_report.html](#sso_reporthtml)
7. [컬럼 정의 레퍼런스](#7-컬럼-정의-레퍼런스)
8. [주요 설계 결정](#8-주요-설계-결정)
9. [유지보수 가이드](#9-유지보수-가이드)
10. [패키지 설치 (폐쇄망)](#10-패키지-설치-폐쇄망)

---

## 1. 전체 아키텍처

```
GlobalMM_Realtime_Monitoring_V2_20260520.xlsb  (반드시 Excel에서 열려있어야 함)
        │
        │  win32com.client.GetObject() — ROT 직접 참조
        │
        ├── DashboardMonitor (reader_ssf.py)   ← 백그라운드 스레드, 1초 폴링
        │       B5:AA36  → section: main  (SSF MM 개별 종목)
        │       B38:AA40 → section: sum1  (소계)
        │       B42:AA44 → section: sum2  (합계)
        │       AC5:AI36 → section: idx   (SSF MM2 지수)
        │
        └── SsoMonitor (reader_sso.py)         ← 백그라운드 스레드, 1초 폴링
                B3:M9    → 요약 테이블 (SSO MM 종목별 합계)
                B11:X350 → 상세 테이블 (종목별 옵션 rows)
                │
serve.py (Flask, 포트 8080)
        ├── GET /                  → index.html
        ├── GET /api/dashboard     → SSF 전체 JSON (초기 로드)
        ├── GET /api/stream        → SSF SSE 스트림 (diff only)
        ├── GET /api/sso           → SSO 전체 JSON (초기 로드)
        ├── GET /api/sso/stream    → SSO SSE 스트림 (diff only)
        └── GET /<path>            → 정적 파일 서빙
                │
브라우저 (index.html iframe SPA)
        ├── ssfo_pnl.html   — SSF MM 테이블
        ├── ssfo_idx.html   — SSF MM2 테이블
        └── sso_report.html — SSO MM 테이블 + 종목별 상세
```

---

## 2. 파일 구조

```
server_hana/
├── serve.py             # Flask 서버 진입점
├── reader_ssf.py        # SSF(DashBoard 탭) 모니터 + 데이터 파싱
├── reader_sso.py        # SSO(Option_DashBoard 탭) 모니터 + 데이터 파싱
├── reader_common.py     # 공유 포매터 함수 (_pct, _num, _date, _rgb)
├── database_setting.py  # MySQL DB 초기화 (Position / T-1_Position 탭)
├── templates/
│   ├── index.html       # 로그인 + iframe SPA 껍데기
│   ├── ssfo_pnl.html    # SSF MM 실시간 테이블
│   ├── ssfo_idx.html    # SSF MM2 실시간 테이블 (지수)
│   └── sso_report.html  # SSO MM 실시간 테이블 + 종목 상세
├── archive/             # 날짜별 스냅샷 HTML (날짜 picker에서 로드)
├── extra/               # 구버전/테스트용 파일 (현재 미사용)
└── wheels/              # 오프라인 pip 설치용 wheel 파일
```

---

## 3. 실행 방법

```bash
# 웹 서버 시작 (포트 8080)
python serve.py

# MySQL 수동 시작 (서비스 등록 불가 환경)
cd C:/mysql-8.0.45-winx64/bin && ./mysqld --console

# DB 초기화 (최초 1회, MySQL 실행 중이어야 함)
python database_setting.py
```

**전제 조건**:
- `GlobalMM_Realtime_Monitoring_V2_20260520.xlsb` 파일이 Excel에서 열려 있어야 함
- Python 패키지: `flask`, `pywin32` 설치 필요

브라우저에서 `http://localhost:8080` 접속.

---

## 4. 데이터 흐름

### 신규 브라우저 접속 시

```
브라우저 → GET /api/dashboard → DashboardMonitor.get_full() → 전체 JSON 반환
브라우저 → GET /api/stream    → SSE 연결 수립 (큐 등록)
이후 변경 발생 시 → SSE로 diff cells push
```

### 실시간 업데이트 (SSE diff)

```
1초마다:
  _read_full() → 현재 전체 데이터 읽기
  이전 데이터와 비교 → 변경된 셀만 changes 리스트에 수집
  항상 push: {"cells": [...changes], "updated": "HH:MM:SS"}
  (변경 없을 때는 cells = [], updated만 갱신됨)
```

### SSE 메시지 형식

```json
// 연결 확인 (최초 1회)
{"type": "connected"}

// 데이터 갱신 (매 1초 전송, 변경 없으면 cells = [])
{
  "cells": [
    {"key": "main_3_15", "v": "1,234", "color": "green", "bg": "rgb(224,102,102)"}
  ],
  "updated": "11:46:22"
}
```

**key 형식 (SSF)**: `"{sid}_{ri}_{cidx}"` — 예: `"main_3_15"`, `"idx_0_6"`  
**key 형식 (SSO)**: `"{ri}_{cidx}"` — 예: `"0_8"`

---

## 5. 모듈별 상세

### serve.py

Flask 서버 진입점. 모니터 스레드 시작 및 HTTP 엔드포인트 정의.

| 함수 | 역할 |
|------|------|
| `_sse_response(monitor)` | SSE 응답 헬퍼. 클라이언트 큐를 monitor에 등록하고 스트림 생성. 20초 무응답 시 keepalive 전송 |
| `api_ssf_full()` | `GET /api/dashboard`. monitor.get_full() 반환. 미준비 시 503 |
| `api_ssf_stream()` | `GET /api/stream`. SSF SSE 스트림 |
| `api_sso_full()` | `GET /api/sso`. SSO 전체 데이터 반환 |
| `api_sso_stream()` | `GET /api/sso/stream`. SSO SSE 스트림 |
| `static_files(filename)` | `.html`은 templates/, 나머지는 BASE_DIR에서 서빙. no-cache 헤더 적용 |
| `start()` | 두 모니터 daemon 스레드 시작 후 Flask 실행 (use_reloader=False) |

---

### reader_common.py

공유 포매터 함수 모음. reader_ssf / reader_sso에서 import.

| 함수 | 역할 | 예시 |
|------|------|------|
| `_pct(v)` | 소수 → 퍼센트 문자열 | `0.0123` → `"1.23%"` |
| `_num(v, d=0)` | 숫자 → 천단위 콤마 + 소수점 d자리 | `1234567` → `"1,234,567"` |
| `_date(v)` | datetime → `"YYYY-MM-DD"` 문자열 | |
| `_rgb(color)` | xlwings color tuple → CSS `rgb()` | `(255,0,0)` → `"rgb(255,0,0)"` |

---

### reader_ssf.py

Excel `DashBoard` 탭을 1초마다 폴링하여 SSF MM / SSF MM2 데이터를 제공.

#### 상수 / 컬럼 정의

| 상수 | 설명 |
|------|------|
| `BOOK_NAME` | 대상 Excel 파일명. 파일 변경 시 이 값만 수정 |
| `MM_COLS` | SSF MM 컬럼 정의. `(cidx, 헤더, 포매터)` 튜플 리스트. cidx는 B열=0 기준 |
| `MM_PNL_CIDX` | 글자색(green/red) 적용 PnL 컬럼 set. MTM PnL(24), Theo PnL(25) 제외 |
| `MM_NUM_CIDX` | 숫자 우측정렬 컬럼 set. Stock(0), StockName(1), MM Spread(10) 제외 |
| `MM_COL_SUBS` | 헤더 두 번째 줄 텍스트 (단축명과 다를 때만) |
| `IDX_COLS` | SSF MM2 컬럼 정의. cidx는 AC열=0 기준 |
| `IDX_PNL_CIDX` | MM2 PnL 글자색 적용 컬럼 set |
| `COLORSCALE_CIDX` | ColorScale 그라디언트 배경 적용 컬럼 `{4, 15, 16, 25}` |
| `TOTAL_VALS` | 합계 행 판별 문자열 set (`"합계"`, `"소계"` 등) |

#### 함수

| 함수 | 역할 |
|------|------|
| `_pnl_color(cidx, raw, pnl_cidx)` | PnL 컬럼 값에 따라 `"green"` / `"red"` / `""` 반환 |
| `_colorscale_bg(val, vmin, vmax)` | 음수→파랑, 양수→빨강 그라디언트 배경색 계산. 범위 내 상대 위치로 보간 |
| `_read_section(all_vals, cols, pnl_cidx)` | Excel 2D 튜플을 `rows` / `row_meta` 리스트로 변환. ColorScale min/max 사전 계산 포함 |

#### DashboardMonitor 클래스

| 메서드 | 역할 |
|--------|------|
| `add_client()` | SSE 클라이언트 큐 생성 및 등록. maxsize=200 |
| `remove_client(q)` | 클라이언트 큐 제거 (접속 끊김 시) |
| `get_full()` | 최신 전체 데이터 반환 (스레드 안전) |
| `_push(payload_str)` | 모든 클라이언트 큐에 메시지 전송. 꽉 찬 큐는 dead로 제거 |
| `_get_ws()` | win32com ROT에서 DashBoard 시트 참조 획득. 실패 시 Excel 직접 실행 |
| `_read_full(ws)` | 헤더 셀 + 4개 범위 bulk read → sections 구조 반환 |
| `run()` | 모니터 루프. `CoInitialize()` → 1초마다 `_read_full()` → diff 계산 → 항상 push |

#### _read_full() 반환 구조

```python
{
  "header": {
    "base_date": "2026-05-25",   # ws.Cells(2,3)
    "ytd":       "2026-05-22",   # ws.Cells(3,3)
    "ir":        "0.03",         # ws.Cells(2,6)
    "exp1":      "2026-06-11",   # ws.Cells(2,9)
    "exp2":      "2026-07-09"    # ws.Cells(3,9)
  },
  "sections": [
    {
      "id": "main",              # "main" | "sum1" | "sum2" | "idx"
      "cols": [(cidx, hdr, sub), ...],
      "num_cidx": [2, 3, ...],
      "rows": [
        {"0": {"v": "005930", "color": "", "bg": ""}, ...},
      ],
      "row_meta": [{"sep": False, "total": False}, ...]
    },
    ...
  ],
  "updated": "11:46:22"
}
```

---

### reader_sso.py

Excel `Option_DashBoard` 탭을 1초마다 폴링하여 SSO MM 데이터를 제공.

#### 상수 / 컬럼 정의

| 상수 | 설명 |
|------|------|
| `COLS` | SSO 요약 컬럼 정의. B3:M9 기준, cidx는 B열=0 기준 |
| `PNL_CIDX` | PnL 컬럼 set (글자색 미적용, 현재 미사용) |
| `NUM_CIDX` | 숫자 우측정렬 컬럼 set |
| `COLORSCALE_CIDX` | ColorScale 그라디언트 적용 컬럼 `{8, 10}` (Theo PnL, MTM PnL) |
| `ALL_DETAIL_FMTS` | 상세 테이블 전체 컬럼 포매터 dict (B11:X350 기준) |
| `DETAIL_PNL_CIDX` | 상세 테이블 PnL 글자색 적용 컬럼 `{7, 8, 21, 22}` |

#### SsoMonitor 클래스

| 메서드 | 역할 |
|--------|------|
| `add_client()` / `remove_client(q)` / `get_full()` / `_push()` | DashboardMonitor와 동일한 SSE 클라이언트 관리 패턴 |
| `_get_ws()` | win32com ROT에서 Option_DashBoard 시트 참조 |
| `_parse_detail(ws)` | B11:X350 범위를 읽어 종목별 flat rows 파싱. 행 타입: `summary` / `futures` / `option` |
| `_read_full(ws)` | B3:M9 요약 읽기 + ColorScale bg 계산 + `_parse_detail()` 호출 |
| `run()` | 1초마다 `_read_full()` → diff 계산 → 항상 push (updated 포함) |

#### _parse_detail() 행 타입 판별 로직

```
B열(col0) 있고 C열(col1) 있음  → "summary" (종목 소계 행, current_stock 갱신)
B열 없고  C열이 "A"로 시작     → "futures" (선물 행)
B열 있고  C열 없음             → "option"  (옵션 행, Strike(col12) 없으면 스킵)
그 외                          → 스킵
```

---

## 6. 프론트엔드 구조

### index.html

로그인 화면 + iframe 기반 SPA.

- 로그인 후 사이드바 메뉴에서 각 페이지를 iframe으로 로드
- **날짜 picker**: `archive/<page>_<yyyymmdd>.html` 로드 (과거 스냅샷)
- 현재 메뉴: SSF MM (`ssfo_pnl.html`), SSF MM2 (`ssfo_idx.html`), SSO MM (`sso_report.html`)

---

### ssfo_pnl.html

SSF MM 실시간 테이블 (DashBoard 탭 main/sum1/sum2 섹션).

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `loadFull()` | `/api/dashboard` 호출 → `buildTable()` → `connectStream()` |
| `buildTable(d)` | thead/tbody 렌더링. DIM_STOCKS 행 배경 처리. 스크롤 위치 보존 |
| `makeTd(sid, ri, cidx, cell)` | td 생성. ColorScale bg, MM Spread Alert, 글자색 적용 |
| `connectStream()` | SSE 연결. 재연결 시 전체 재로드. liveDot 상태 표시 |
| `applyChanges(cells, updated)` | SSE diff 적용. ColorScale bg / MM Spread Alert / 글자색 / flash 애니메이션 |

#### 특수 처리

| 조건 | 처리 |
|------|------|
| `DIM_STOCKS` 종목 | 행 배경 연파랑(`#d8e6f5`), 좌측 테두리 강조 |
| `cidx === 10`, `v === "Alert"` | 노란 배경 + 빨간 굵은 글자 (MM Spread Alert) |
| `CS_CIDX = {4, 15, 16, 25}` | `cell.bg` 값으로 ColorScale 배경 적용 |
| `sid === "idx"` 셀 | `applyChanges`에서 무시 (MM2는 ssfo_idx.html에서 처리) |

---

### ssfo_idx.html

SSF MM2 실시간 테이블 (DashBoard 탭 idx 섹션).

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `buildGroups(rows, rowMeta)` | SECTION_STARTERS 기반으로 rows를 그룹(테이블 단위)으로 분할 |
| `buildTables(d)` | 그룹별 `<table>` 렌더링. 첫/마지막 테이블에만 thead 표시 |
| `applyChanges(cells, updated)` | idx 섹션만 처리. flash 애니메이션 적용 |

#### SECTION_STARTERS

`"Q150 Roll"`, `"K2G01P"`, `"091160"`, `"Udly"`, `"IndexCode"` — IdxCode(cidx 0) 값이 이 중 하나이면 새 테이블 시작.  
단, `"IndexCode"` 이후에 나타나는 `"K2G01P"`는 새 섹션으로 처리하지 않음.

---

### sso_report.html

SSO MM 요약 테이블 + 종목별 상세 패널.

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `loadFull()` | `/api/sso` 호출 → `buildTable()` + `buildButtons()` → `connectStream()` |
| `buildTable(d)` | 요약 테이블 렌더링 |
| `buildButtons(d)` | 종목 코드별 버튼 생성 (클릭 시 상세 패널 오픈) |
| `makeTd(ri, cidx, cell)` | td 생성. `CS_CIDX = {8, 10}` ColorScale bg 적용. 글자색 미적용 |
| `selectStock(code)` | 버튼 클릭 → `fetchAndRender()` 최초 호출 + 3초 interval 시작 |
| `fetchAndRender()` | `/api/sso` 재호출 → `renderDetail()` |
| `renderDetail(code)` | 종목 상세 테이블 렌더링 (DETAIL_HDRS 기준) |
| `applyChanges(cells, updated)` | SSE diff 적용. ColorScale bg 갱신. 글자색 미적용 |
| `closeDetail()` | 상세 패널 닫기 + interval 정리 |

#### 행 타입별 스타일

| 타입 | 스타일 |
|------|--------|
| `row-summary` | 보라빛 배경(`#e8eaf6`), bold |
| `row-futures` | 회색 배경(`#f3f4f6`), 연한 글자 |
| `row-atm` (OTM === "ATM") | 노란 배경(`#fff9c4`), bold |

---

## 7. 컬럼 정의 레퍼런스

### SSF MM (MM_COLS) — 범위: B5:AA36 / B38:AA40 / B42:AA44, cidx 기준: B=0

| cidx | 헤더 | 포맷 | 비고 |
|------|------|------|------|
| 0 | Stock | 문자열 | |
| 1 | StockName | 문자열 | |
| 2 | LastPrice | 정수 | |
| 3 | Change(%) | `_pct` | 글자색 적용 |
| 4 | Amt(Mil) | 정수 | ColorScale 배경 |
| 5 | Amt(Shr) | 정수 | |
| 6 | MMQty | 정수 | |
| 7 | (skip) | — | I열, 사용 안 함 |
| 8 | Ask | 정수 | |
| 9 | Bid | 정수 | |
| 10 | MM Spread | 문자열 | Alert 시 노란/빨강 |
| 11 | Shares | 정수 | |
| 12 | Diff | 정수 | |
| 13 | Vol(Lots) | 정수 | |
| 14 | Vol(Mil) | 정수 | |
| 15 | Delta(Shr) | 정수 | ColorScale 배경 |
| 16 | Delta(Mil) | 정수 | ColorScale 배경 |
| 17 | Theo Price | 정수 | |
| 18 | Theo Basis | 정수 | |
| 19 | Ex1 B-Sp | 정수 | |
| 20 | Ex1 S-Sp | 정수 | |
| 21 | Ex2 Basis | 정수 | |
| 22 | Ex2 B-Sp | 정수 | |
| 23 | Ex2 S-Sp | 정수 | |
| 24 | MTM PnL | 정수 | 글자색 미적용 |
| 25 | Theo PnL | 정수 | ColorScale 배경, 글자색 미적용 |

### SSF MM2 (IDX_COLS) — 범위: AC5:AI36, cidx 기준: AC=0

| cidx | 헤더 | 포맷 | 비고 |
|------|------|------|------|
| 0 | IdxCode | 문자열 | SECTION_STARTERS 판별 기준 |
| 1 | IdxName | 문자열 | |
| 2 | Fund | 정수 | |
| 3 | Delta | 정수 | |
| 4 | Volume | 정수 | |
| 5 | MTM PnL(Idx) | 정수 | 글자색 적용 |
| 6 | Theo PnL(Idx) | 정수 | 글자색 적용 |

### SSO MM (COLS) — 범위: B3:M9, cidx 기준: B=0

| cidx | 헤더 | 포맷 | 비고 |
|------|------|------|------|
| 0 | StockCode | 문자열 | |
| 1 | StockName | 문자열 | |
| 2 | Delta | 정수 | |
| 4 | %Gamma | 소수 2자리 | |
| 6 | Volume | 정수 | |
| 8 | Theo PnL | 정수 | ColorScale 배경, 글자색 미적용 |
| 10 | MTM PnL | 정수 | ColorScale 배경, 글자색 미적용 |

---

## 8. 주요 설계 결정

### win32com GetObject (ROT 직접 참조)

```python
win32com.client.GetObject(path)  # Excel이 이미 열고 있는 파일을 ROT에서 참조
```

`xw.apps` 열거 방식(`get_xl_app_from_hwnd`)은 access violation이 발생하여 완전히 우회.  
Excel이 닫혀있으면 `Dispatch("Excel.Application")`으로 새로 열음.

### COM 스레드 초기화

```python
def run(self):
    pythoncom.CoInitialize()  # 비메인 스레드에서 COM 사용 시 필수
```

### SSE diff 방식 + 항상 push

전체 재전송 대신 변경 셀만 push. 변경 없어도 매 1초 `updated` 타임스탬프는 항상 전송.  
신규 접속 → `/api/dashboard` 전체 스냅샷 → SSE로 이후 diff 적용.

### cidx = 범위 시작열 기준 0-based

각 섹션의 cidx는 해당 범위 첫 열을 0으로 하는 인덱스.  
`row_data[cidx]`로 offset 연산 없이 바로 접근.

- MM_COLS: B열=0 기준
- IDX_COLS: AC열=0 기준
- SSO COLS: B열=0 기준

### ColorScale 그라디언트

컬럼 내 전체 값의 min/max를 기준으로 상대 위치 계산:

- `v < 0`: `vmin` ~ `0` 구간 → 파랑(`#6fa8dc`) → 흰(`#ffffff`)
- `v > 0`: `0` ~ `vmax` 구간 → 흰(`#ffffff`) → 빨강(`#e06666`)

min/max는 매 폴링 시 해당 섹션 전체 행에서 재계산.

---

## 9. 유지보수 가이드

### Excel 파일명 변경 시

`reader_ssf.py` 상단의 `BOOK_NAME` 만 수정:

```python
BOOK_NAME = "GlobalMM_Realtime_Monitoring_V2_20260520.xlsb"
```

### 컬럼 추가/제거 시

**SSF MM**: `reader_ssf.py`의 `MM_COLS` 리스트에 `(cidx, "헤더", 포매터)` 튜플 추가/제거.  
**SSF MM2**: `IDX_COLS` 동일하게 수정.  
**SSO**: `reader_sso.py`의 `COLS` 수정.

cidx는 각 범위의 시작열(B열 또는 AC열)을 0으로 하는 인덱스.

### ColorScale 컬럼 추가/제거 시

서버와 클라이언트 양쪽을 함께 수정해야 함:

- **SSF**: `reader_ssf.py`의 `COLORSCALE_CIDX` + `ssfo_pnl.html`의 `CS_CIDX`
- **SSO**: `reader_sso.py`의 `COLORSCALE_CIDX` + `sso_report.html`의 `CS_CIDX`

### PnL 글자색(green/red) 적용 컬럼 변경 시

- **SSF MM**: `reader_ssf.py`의 `MM_PNL_CIDX` set에서 추가/제거 (현재 24, 25 제외)
- **SSF MM2**: `IDX_PNL_CIDX` 수정
- **SSO 요약**: 현재 글자색 미적용. `sso_report.html`의 `makeTd` / `applyChanges` 수정 필요
- **SSO 상세**: `reader_sso.py`의 `DETAIL_PNL_CIDX` 수정

### DIM_STOCKS (배경 강조 종목) 변경 시

`ssfo_pnl.html` 상단 `DIM_STOCKS` Set에 종목코드 추가/제거:

```javascript
const DIM_STOCKS = new Set(["010950", "086280", ...]);
```

### SECTION_STARTERS (MM2 테이블 분할 기준) 변경 시

`ssfo_idx.html`의 `SECTION_STARTERS` Set 수정:

```javascript
const SECTION_STARTERS = new Set(["Q150 Roll", "K2G01P", "091160", "Udly", "IndexCode"]);
```

### 헤더 정보 셀 위치 변경 시

`reader_ssf.py`의 `_read_full()` 내 `header` 딕셔너리 (Cells는 1-based):

```python
header = {
    "base_date": _date(ws.Cells(2, 3).Value),  # C2
    "ytd":       _date(ws.Cells(3, 3).Value),  # C3
    "ir":        str(ws.Cells(2, 6).Value),    # F2
    "exp1":      _date(ws.Cells(2, 9).Value),  # I2
    "exp2":      _date(ws.Cells(3, 9).Value),  # I3
}
```

### 폴링 주기 변경 시

각 reader의 `run()` 마지막 줄 `time.sleep(초)` 수정. 현재 SSF/SSO 모두 1.0초.

### 아카이브(스냅샷) 생성

`reader_ssf.py` 또는 `reader_sso.py`를 직접 실행하면 스냅샷 HTML 생성:

```bash
python reader_ssf.py   # → ssf_snapshot.html, ssf_idx_snapshot.html
python reader_sso.py   # → sso_snapshot.html
```

---

## 10. 패키지 설치 (폐쇄망)

```bash
# 인터넷 가능한 PC에서 wheel 다운로드
pip download flask pywin32 -d ./wheels/flask
pip download mysql-connector-python -d ./wheels/mysql

# 대상 서버에서 오프라인 설치
pip install --no-index --find-links=./wheels/flask flask pywin32
pip install --no-index --find-links=./wheels/mysql mysql-connector-python
```
