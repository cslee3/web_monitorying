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
   - [common.py](#commonpy)
   - [reader_ssf.py](#reader_ssfpy)
   - [reader_sso.py](#reader_ssopy)
   - [serve.py](#servepy)
   - [database_setting.py](#database_settingpy)
6. [프론트엔드 구조](#6-프론트엔드-구조)
   - [index.html](#indexhtml)
   - [ssfo_pnl.html](#ssfo_pnlhtml)
   - [ssfo_idx.html](#ssfo_idxhtml)
   - [sso_report.html](#sso_reporthtml)
   - [welcome.html](#welcomehtml)
7. [컬럼 정의 레퍼런스](#7-컬럼-정의-레퍼런스)
8. [주요 설계 결정](#8-주요-설계-결정)
9. [유지보수 가이드](#9-유지보수-가이드)
10. [패키지 설치 (폐쇄망)](#10-패키지-설치-폐쇄망)

---

## 1. 전체 아키텍처

```
GlobalMM_Realtime_Monitoring_V2.xlsb  (반드시 Excel에서 열려있어야 함)
        │
        │  win32com.client.GetObject() — ROT 직접 참조
        │
        ├── DashboardMonitor (reader_ssf.py)   ← 백그라운드 스레드, 1초 폴링
        │       B2:AK44 단일 bulk read (COM 호출 1회)
        │         ├─ rows 0~1   : 헤더 정보 (BaseDate, Ytd, I/R, Expiry, 업데이트 시각)
        │         ├─ rows 3~34  : section: main  (SSF MM 개별 종목)
        │         ├─ rows 36~38 : section: sum1  (소계)
        │         ├─ rows 40~42 : section: sum2  (합계)
        │         └─ rows 3~42  : section: idx1~idx7 (SSF MM2, 7개 섹션)
        │
        └── SsoMonitor (reader_sso.py)         ← 백그라운드 스레드, 1초 폴링
                B3:M9    → 요약 테이블 (SSO MM 종목별 합계)
                B11:X350 → 상세 테이블 (종목별 선물/옵션 rows)

serve.py (Flask, 포트 8080)
        ├── GET /                  → index.html
        ├── GET /api/dashboard     → SSF 전체 JSON (초기 로드)
        ├── GET /api/stream        → SSF SSE 스트림 (diff only)
        ├── GET /api/sso           → SSO 전체 JSON (초기 로드)
        ├── GET /api/sso/stream    → SSO SSE 스트림 (diff only)
        └── GET /<path>            → 정적 파일 서빙

브라우저 (index.html iframe SPA)
        ├── welcome.html    — 홈 화면
        ├── ssfo_pnl.html   — SSF MM 테이블
        ├── ssfo_idx.html   — SSF MM2 테이블 (7개 섹션)
        └── sso_report.html — SSO MM 테이블 + 종목별 상세
```

---

## 2. 파일 구조

```
web_monitor/
├── serve.py              # Flask 서버 진입점
├── reader_ssf.py         # SSF(DashBoard 탭) 모니터 + 데이터 파싱
├── reader_sso.py         # SSO(Option_DashBoard 탭) 모니터 + 데이터 파싱
├── common.py             # 공유 상수 + 유틸리티 함수
├── database_setting.py   # MySQL DB 초기화 (Position / T-1_Position 탭)
├── templates/
│   ├── index.html        # 로그인 + iframe SPA 껍데기
│   ├── welcome.html      # 홈 화면 (로고 + 이미지)
│   ├── ssfo_pnl.html     # SSF MM 실시간 테이블
│   ├── ssfo_idx.html     # SSF MM2 실시간 테이블 (7개 섹션)
│   └── sso_report.html   # SSO MM 실시간 테이블 + 종목 상세
├── static/               # 정적 이미지 파일 (logo.png, elon.jpg 등)
├── archive/              # 날짜별 스냅샷 HTML (날짜 picker에서 로드)
├── extra/                # 구버전/테스트용 파일 (현재 미사용)
└── wheels/               # 오프라인 pip 설치용 wheel 파일
```

---

## 3. 실행 방법

```bash
# 웹 서버 시작 (포트 8080)
python serve.py

# MySQL 수동 시작 (서비스 등록 불가 환경) — 프로젝트와 동등한 위치
cd D:/OneDrive/mysql/bin && ./mysqld --console

# DB 초기화 (최초 1회, MySQL 실행 중이어야 함)
python database_setting.py

# 스냅샷 HTML 단독 생성 (서버 없이 확인할 때)
python reader_ssf.py   # → ssf_snapshot.html, ssf_idx_snapshot.html
python reader_sso.py   # → templates/sso_snapshot.html
```

**전제 조건**:
- `GlobalMM_Realtime_Monitoring_V2.xlsb` 파일이 Excel에서 열려 있어야 함
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
  _read_full() → B2:AK44 단일 COM 호출로 전체 데이터 읽기
  이전 데이터와 비교 → 변경된 셀만 changes 리스트에 수집
  항상 push: {"cells": [...changes], "updated": "HH:MM:SS"}
  변경 없을 때는 cells = [], updated만 갱신
```

### SSE 메시지 형식

```json
// 연결 확인 (최초 1회)
{"type": "connected"}

// 데이터 갱신 (매 1초 전송)
{
  "cells": [
    {"key": "main_3_15", "v": "1,234", "color": "green", "bg": "rgb(224,102,102)"},
    {"key": "idx1_0_2",  "v": "51003", "color": "",      "bg": ""}
  ],
  "updated": "11:46:22"
}
```

**key 형식 — MM 섹션**: `"{sid}_{ri}_{cidx}"` — 예: `"main_3_15"`, `"sum1_0_24"`  
**key 형식 — MM2 섹션**: `"{sid}_{ri}_{ci}"` — 예: `"idx1_2_3"`, `"idx6_0_4"` (ci는 섹션 내 0-based)  
**key 형식 — SSO**: `"{ri}_{cidx}"` — 예: `"0_8"`

---

## 5. 모듈별 상세

### common.py

공유 상수 및 유틸리티 함수. `reader_ssf.py` / `reader_sso.py`에서 import.

#### 상수

| 상수 | 설명 |
|------|------|
| `BASE_DIR` | 스크립트 기준 절대 디렉터리. Spyder 환경 대응 (`__file__` 미정의 시 `os.getcwd()`) |
| `BOOK_NAME` | 대상 Excel 파일명. **파일 변경 시 여기만 수정** |

#### 함수

| 함수 | 역할 | 예시 |
|------|------|------|
| `_is_xl_error(v)` | Excel COM 에러코드 여부 판별 (`0x800A0xxx` 패턴, 약 `-2,147,500,000 ~ -2,146,000,000`) | `#BLOCKED!`, `#VALUE!` 등 |
| `_pct(v)` | 소수 → 퍼센트 문자열. 에러값 → `""` | `0.0123` → `"1.23%"` |
| `_num(v, d=0)` | 숫자 → 천단위 콤마 + 소수점 d자리. 에러값 → `""` | `1234567` → `"1,234,567"` |
| `_date(v)` | datetime → `"YYYY-MM-DD"` 문자열 | |
| `_rgb(color)` | xlwings color tuple → CSS `rgb()` | `(255,0,0)` → `"rgb(255,0,0)"` |
| `_colorscale_bg(val, vmin, vmax)` | 음수→파랑(`#6fa8dc`), 양수→빨강(`#e06666`) 그라디언트 배경색 계산 | |

---

### reader_ssf.py

Excel `DashBoard` 탭을 1초마다 폴링하여 SSF MM / SSF MM2 데이터를 제공.  
**핵심: `B2:AK44` 단일 COM 호출**로 헤더·MM·IDX 전 영역을 읽고 Python에서 슬라이싱.

#### 섹션 설정

##### MM_SECTIONS — B:AA (cols 0~25), B2 기준 row 오프셋

```python
MM_SECTIONS = [
    {"id": "main", "row_s":  3, "row_e": 34},  # B5:AA36
    {"id": "sum1", "row_s": 36, "row_e": 38},  # B38:AA40
    {"id": "sum2", "row_s": 40, "row_e": 42},  # B42:AA44
]
```

범위 변경 시 `row_s` / `row_e`만 수정. (B2=row0, B3=row1, B5=row3, ...)

##### IDX2_SECTIONS — B2 기준 row/col 오프셋 (AE=col29, AK=col35)

| id | Excel 범위 | has_header | header |
|----|-----------|------------|--------|
| idx1 | AE5:AK11 | True (Excel 첫 행) | — |
| idx2 | AE13:AF13 | False | — |
| idx3 | AE14:AK15 | False | IdxCode, IndexName, Fund, Delta, Volume, MTM PnL, Theo PnL |
| idx4 | AE17:AK18 | False | IdxCode, IndexName, Fund, Delta, Volume, MTM PnL, Theo PnL |
| idx5 | AE20:AK25 | True (Excel 첫 행) | — |
| idx6 | AE29:AK38 | True (Excel 첫 행) | — |
| idx7 | AH43:AJ44 | False | — |

- `has_header: True` → Excel 데이터의 첫 행을 thead로 사용
- `header: [...]` → 커스텀 헤더 문자열 지정 (Excel에 헤더 행 없을 때)
- IDX2 섹션은 값을 **그대로(raw)** 표시하고 모두 정수로 변환

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `_is_xl_error(v)` | → `common._is_xl_error()` 재사용 |
| `_raw(v)` | Excel raw 값 → 문자열. 에러코드 `""`, datetime → `"HH:MM:SS"` (오늘) / `"YYYY-MM-DD HH:MM:SS"` (과거), float → 정수 |
| `_pnl_color(cidx, raw, pnl_cidx)` | PnL 컬럼 값에 따라 `"green"` / `"red"` / `""` 반환 |
| `_read_section(all_vals, cols, pnl_cidx)` | Excel 2D 튜플을 `rows` / `row_meta` 리스트로 변환. ColorScale min/max 사전 계산 포함 |
| `_read_mm(bulk)` | bulk 데이터 → `MM_SECTIONS` 슬라이싱 → MM 섹션 리스트 반환 |
| `_read_idx2(bulk)` | bulk 데이터 → `IDX2_SECTIONS` 슬라이싱 → raw 섹션 리스트 반환 |

#### DashboardMonitor 클래스

| 메서드 | 역할 |
|--------|------|
| `add_client()` | SSE 클라이언트 큐 생성 및 등록 (maxsize=200) |
| `remove_client(q)` | 클라이언트 큐 제거 (접속 끊김 시) |
| `get_full()` | 최신 전체 데이터 반환 (스레드 안전) |
| `_push(payload_str)` | 모든 클라이언트 큐에 메시지 전송. 꽉 찬 큐는 자동 제거 |
| `_get_ws()` | win32com ROT에서 DashBoard 시트 참조 획득. 실패 시 Excel 직접 실행 |
| `_read_full(ws)` | `B2:AK44` 단일 호출 → header + MM + IDX2 전체 파싱 |
| `run()` | `CoInitialize()` → 1초마다 `_read_full()` → diff 계산 → push |

#### _read_full() 반환 구조

```python
{
  "header": {
    "base_date": "2026-05-26",   # bulk[0][1]  = C2
    "ytd":       "2026-05-22",   # bulk[1][1]  = C3
    "ir":        "3.50",         # bulk[0][4]  = F2
    "exp1":      "2026-06-11",   # bulk[0][7]  = I2
    "exp2":      "2026-09-10",   # bulk[1][7]  = I3
    "u2": "...", "v2": "...",    # bulk[0][19,20] = U2, V2 (Excel 업데이트 시각 등)
    "u3": "...", "v3": "..."     # bulk[1][19,20] = U3, V3
  },
  "sections": [
    # MM 섹션 (raw: False)
    {
      "id": "main",              # "main" | "sum1" | "sum2"
      "cols": [(cidx, hdr, sub), ...],
      "num_cidx": [2, 3, ...],
      "rows":     [{"0": {"v": "005930", "color": "", "bg": ""}, ...}, ...],
      "row_meta": [{"sep": False, "total": False}, ...]
    },
    # IDX2 섹션 (raw: True)
    {
      "id": "idx1",
      "has_header": True,
      "header": None,            # 커스텀 헤더 없을 때 None
      "raw": True,
      "rows": [["코스닥150", "51003", "-400", ...], ...]
    },
    ...
  ],
  "updated": "11:46:22"
}
```

#### diff 처리 (run() 내부)

```python
# MM 섹션 (raw: False)
for cidx, *_ in sec_new["cols"]:
    if new_v != old_v or new_bg != old_bg:
        changes.append({"key": f"{sid}_{ri}_{cidx}", ...})

# IDX2 섹션 (raw: True)
for ci, (nv, ov) in enumerate(zip(new_row, old_row)):
    if nv != ov:
        changes.append({"key": f"{sid}_{ri}_{ci}", "v": nv, ...})
```

---

### reader_sso.py

Excel `Option_DashBoard` 탭을 1초마다 폴링하여 SSO MM 데이터를 제공.

#### 상수 / 컬럼 정의

| 상수 | 설명 |
|------|------|
| `COLS` | SSO 요약 컬럼 정의. B3:M9 기준, cidx는 B열=0 기준 |
| `NUM_CIDX` | 숫자 우측정렬 컬럼 set |
| `COLORSCALE_CIDX` | ColorScale 그라디언트 적용 컬럼 `{8, 10}` (Theo PnL, MTM PnL) |
| `ALL_DETAIL_FMTS` | 상세 테이블 전체 컬럼 포매터 dict (B11:X350 기준, cidx 0~22) |
| `DETAIL_PNL_CIDX` | 상세 테이블 PnL 글자색 적용 컬럼 `{7, 8, 21, 22}` |

#### SsoMonitor 클래스

| 메서드 | 역할 |
|--------|------|
| `add_client()` / `remove_client(q)` / `get_full()` / `_push()` | DashboardMonitor와 동일한 SSE 클라이언트 관리 패턴 |
| `_get_ws()` | win32com ROT에서 Option_DashBoard 시트 참조 |
| `_parse_detail(ws)` | B11:X350 범위를 읽어 종목별 flat rows 파싱. 행 타입: `summary` / `futures` / `option` |
| `_read_full(ws)` | B3:M9 요약 읽기 + ColorScale bg 계산 + `_parse_detail()` 호출 |
| `run()` | 1초마다 `_read_full()` → diff 계산 → push |

#### _parse_detail() 행 타입 판별 로직

```
B열(col0) 있고 C열(col1) 있음  → "summary" (종목 소계 행, current_stock 갱신)
B열 없고  C열이 "A"로 시작     → "futures" (선물 행)
B열 있고  C열 없음             → "option"  (옵션 행, Strike(col12) 없으면 스킵)
그 외                          → 스킵
```

---

### serve.py

Flask 서버 진입점. 모니터 스레드 시작 및 HTTP 엔드포인트 정의.

| 함수 / 엔드포인트 | 역할 |
|------|------|
| `_sse_response(monitor)` | SSE 응답 헬퍼. 클라이언트 큐를 monitor에 등록하고 스트림 생성. 20초 무응답 시 keepalive |
| `GET /` | `templates/index.html` 반환. no-cache 헤더 적용 |
| `GET /api/dashboard` | `DashboardMonitor.get_full()` 반환. 미준비 시 503 |
| `GET /api/stream` | SSF SSE 스트림 |
| `GET /api/sso` | `SsoMonitor.get_full()` 반환. 미준비 시 503 |
| `GET /api/sso/stream` | SSO SSE 스트림 |
| `GET /<path>` | `.html` → `templates/`에서, 그 외 → `BASE_DIR`에서 서빙 |
| `start()` | 두 모니터 daemon 스레드 시작 후 Flask 실행 (`debug=False`, `use_reloader=False`) |

---

### database_setting.py

MySQL DB 초기화 스크립트. **최초 1회만 실행**.

- `market_making` 데이터베이스에 `POSITION`, `PREV_POSITION` 테이블 생성
- Excel `Position` / `T-1_Position` 시트에서 데이터 읽어 INSERT
- MySQL이 실행 중이지 않으면 자동으로 `mysqld.exe --console` 기동 시도

---

## 6. 프론트엔드 구조

### index.html

로그인 화면 + iframe 기반 SPA.

- 로그인 후 사이드바 메뉴에서 각 페이지를 `<iframe id="mainFrame">`으로 로드
- 기본 src: `welcome.html`
- **날짜 picker**: `archive/<page>_<yyyymmdd>.html` 로드 (과거 스냅샷 열람)
- URL 캐시 방지: `iframe.src = page + "?_=" + Date.now()`
- Ctrl+Shift+R(하드 리프레시) 시 `welcome.html`이 없으면 404 — `static/welcome.html` 확인

---

### ssfo_pnl.html

SSF MM 실시간 테이블 (DashBoard 탭 main / sum1 / sum2 섹션).

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `loadFull()` | `/api/dashboard` 호출 → `buildTable()` → `connectStream()` |
| `buildTable(d)` | `d.sections.filter(s => !s.raw)`로 MM 섹션만 필터링. thead/tbody 렌더링. 스크롤 위치 보존 |
| `makeTd(sid, ri, cidx, cell)` | td 생성. ColorScale bg, MM Spread Alert, 글자색 적용 |
| `connectStream()` | SSE 연결. 재연결 시 전체 재로드. liveDot 상태 표시 |
| `applyChanges(cells, updated)` | `sid.startsWith("idx")` 셀은 무시. ColorScale bg / MM Spread Alert / 글자색 / flash 애니메이션 |

#### 특수 처리

| 조건 | 처리 |
|------|------|
| `cidx === 10`, `v === "Alert"` | 노란 배경 + 빨간 굵은 글자 (MM Spread Alert) |
| `CS_CIDX = {4, 15, 16, 25}` | `cell.bg` 값으로 ColorScale 배경 적용 |
| `sid.startsWith("idx")` | `applyChanges`에서 무시 (MM2는 ssfo_idx.html에서 처리) |
| `s.raw === true` 섹션 | `buildTable` 필터에서 제외 |

---

### ssfo_idx.html

SSF MM2 실시간 테이블 — `IDX2_SECTIONS` 7개를 각각 별도 테이블로 렌더링.

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `buildTables(d)` | `d.sections.filter(s => s.raw)`로 IDX2 섹션만 필터링. 섹션별 `<table>` 렌더링 |
| `connectStream()` | SSE 연결 및 재연결 처리 |
| `applyChanges(cells, updated)` | `key.startsWith("idx")` 인 셀만 처리. `c-{sid}-{ri}-{ci}` ID로 td 찾아 업데이트. flash 애니메이션 |

#### 섹션 렌더링 규칙

| 조건 | 처리 |
|------|------|
| `sec.has_header === true` | `rows[0]`을 `<thead>`로 사용, 데이터는 `rows[1]`부터 |
| `sec.header !== null` | 커스텀 헤더 배열을 `<thead>`로 사용, `rows[0]`부터 데이터 |
| 둘 다 없음 | `<thead>` 없이 데이터 rows만 렌더링 |

- 모든 셀: `text-align: right`
- 셀 ID: `c-{sid}-{ri}-{ci}` (ci = 섹션 내 0-based 컬럼 인덱스)

---

### sso_report.html

SSO MM 요약 테이블 + 종목별 상세 패널.

#### 주요 함수

| 함수 | 역할 |
|------|------|
| `loadFull()` | `/api/sso` 호출 → `buildTable()` + `buildButtons()` → `connectStream()` → **첫 번째 종목 자동 선택** |
| `buildTable(d)` | 요약 테이블 렌더링 |
| `buildButtons(d)` | 종목 코드별 버튼 생성 |
| `selectStock(code)` | 버튼 클릭 → `fetchAndRender()` 최초 호출 + 3초 interval 시작. 같은 버튼 재클릭 시 닫기 |
| `fetchAndRender()` | `/api/sso` 재호출 → `renderDetail()` |
| `renderDetail(code)` | 종목 상세 테이블 렌더링 (`DETAIL_HDRS` 기준 23컬럼) |
| `applyChanges(cells, updated)` | SSE diff 적용. `CS_CIDX = {8, 10}` ColorScale bg 갱신 |
| `closeDetail()` | 상세 패널 닫기 + 3초 interval 정리 |

#### 행 타입별 스타일

| 타입 | 스타일 |
|------|--------|
| `row-summary` | 보라빛 배경 (`#e8eaf6`), bold |
| `row-futures` | 회색 배경 (`#f3f4f6`), 연한 글자 |
| `row-atm` (OTM === "ATM") | 노란 배경 (`#fff9c4`), bold |

---

### welcome.html

홈 화면. 로그인 후 기본 표시 페이지.

- 흰 배경
- 상단 중앙: `static/logo.png` (width: 180px)
- 중앙: `static/elon.jpg`

이미지 파일은 `static/` 폴더에 위치해야 함.

---

## 7. 컬럼 정의 레퍼런스

### SSF MM (MM_COLS) — cidx 기준: B=0

| cidx | 헤더 | 포맷 | 비고 |
|------|------|------|------|
| 0 | Stock | 문자열 | |
| 1 | StockName | 문자열 | |
| 2 | LastPrice | 정수 | |
| 3 | Change(%) | `_pct` | |
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

### SSF MM2 (IDX2_SECTIONS) — 값 그대로(raw), 정수 표시

모든 섹션의 값은 `_raw()`를 통해 처리됨:
- 숫자 → `int(round(v))`
- datetime → `"HH:MM:SS"` (오늘) / `"YYYY-MM-DD HH:MM:SS"` (과거)
- Excel 에러코드 (`#BLOCKED!` 등) → `""`

### SSO MM (COLS) — cidx 기준: B=0

| cidx | 헤더 | 포맷 | 비고 |
|------|------|------|------|
| 0 | StockCode | 문자열 | |
| 1 | StockName | 문자열 | |
| 2 | Delta | 정수 | |
| 4 | %Gamma | 소수 2자리 | |
| 6 | Volume | 정수 | |
| 8 | Theo PnL | 정수 | ColorScale 배경 |
| 10 | MTM PnL | 정수 | ColorScale 배경 |

---

## 8. 주요 설계 결정

### B2:AK44 단일 COM 호출

```python
bulk = ws.Range("B2:AK44").Value  # 헤더 + MM 3섹션 + IDX2 7섹션 전부
```

기존 8회 COM 호출(셀 개별 5회 + Range 3회) → 1회로 통합.  
Python 슬라이싱으로 각 섹션 분리. COM 호출 비용(수십ms)이 슬라이싱(수μs)보다 압도적으로 큼.

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

### IDX2 섹션 raw 처리

MM2(IDX2) 섹션은 컬럼 포매터 없이 값 그대로 전달. 이유:
- 섹션마다 컬럼 구조가 달라 통일된 포매터 정의 불가
- 유지보수 단순화: `IDX2_SECTIONS` 설정만으로 범위/헤더 관리

### Excel 에러값 필터링

```python
def _is_xl_error(v):
    iv = int(v)
    return -2_147_500_000 <= iv <= -2_146_000_000  # 0x800A0xxx 패턴
```

`#BLOCKED!`, `#VALUE!`, `#REF!` 등 Excel COM 에러코드를 `""` 로 처리.  
`_num()`, `_pct()`, `_raw()` 세 곳에서 공통 사용.

### cidx = 범위 시작열 기준 0-based

각 섹션의 cidx는 해당 범위 첫 열을 0으로 하는 인덱스.  
`row_data[cidx]`로 offset 연산 없이 바로 접근.

- `MM_COLS`: B열=0 기준
- `SSO COLS`: B열=0 기준
- `IDX2_SECTIONS`: B열=0 기준 (AE=29)

### ColorScale 그라디언트

컬럼 내 전체 값의 min/max를 기준으로 상대 위치 계산:

- `v < 0`: `vmin` ~ `0` 구간 → 파랑(`#6fa8dc`) → 흰(`#ffffff`)
- `v > 0`: `0` ~ `vmax` 구간 → 흰(`#ffffff`) → 빨강(`#e06666`)

min/max는 매 폴링 시 해당 섹션 전체 행에서 재계산.

---

## 9. 유지보수 가이드

### Excel 파일명 변경 시

`common.py` 상단의 `BOOK_NAME`만 수정:

```python
BOOK_NAME = "GlobalMM_Realtime_Monitoring_V2.xlsb"
```

### MM 섹션 범위 변경 시

`reader_ssf.py`의 `MM_SECTIONS` 수정:

```python
MM_SECTIONS = [
    {"id": "main", "row_s":  3, "row_e": 34},  # B5:AA36  (row = Excel행 - 2)
    {"id": "sum1", "row_s": 36, "row_e": 38},  # B38:AA40
    {"id": "sum2", "row_s": 40, "row_e": 42},  # B42:AA44
]
```

### MM2 섹션 범위/헤더 변경 시

`reader_ssf.py`의 `IDX2_SECTIONS` 수정:

```python
# row_s/row_e: Excel 행번호 - 2  (row2=0, row5=3, row13=11, ...)
# col_s/col_e: Excel 열번호 - 2  (B=0, AE=29, AH=32, AK=35)
IDX2_SECTIONS = [
    {"id": "idx3", ..., "header": ["col1", "col2", ...]},  # 커스텀 헤더
    {"id": "idx5", ..., "has_header": True},               # Excel 첫 행 사용
]
```

단, 전체 bulk read 범위 `B2:AK44`를 벗어나는 셀을 추가할 경우 `_read_full()` 내 Range 수정 필요.

### MM 컬럼 추가/제거 시

`reader_ssf.py`의 `MM_COLS` 리스트에 `(cidx, "헤더", 포매터)` 튜플 추가/제거.  
cidx는 B열=0 기준. 같이 관리해야 할 항목:

- `MM_PNL_CIDX`: 글자색 적용 PnL 컬럼 set
- `MM_NUM_CIDX`: 숫자 우측정렬 컬럼 set
- `MM_COL_SUBS`: 헤더 두 번째 줄 텍스트
- `COLORSCALE_CIDX`: ColorScale 배경 적용 컬럼 (서버 + `ssfo_pnl.html`의 `CS_CIDX` 동시 수정)

### 헤더 정보 셀 위치 변경 시

`reader_ssf.py`의 `_read_full()` 내 `header` 딕셔너리 수정 (B2 기준 0-based):

```python
header = {
    "base_date": _date(bulk[0][1]),   # C2  = (row2=0, col C-B=1)
    "ytd":       _date(bulk[1][1]),   # C3  = (row3=1, col C-B=1)
    "ir":        _raw(bulk[0][4]),    # F2  = (row2=0, col F-B=4)
    "exp1":      _date(bulk[0][7]),   # I2  = (row2=0, col I-B=7)
    "exp2":      _date(bulk[1][7]),   # I3  = (row3=1, col I-B=7)
    "u2": _raw(bulk[0][19]),          # U2  = (row2=0, col U-B=19)
    "v2": _raw(bulk[0][20]),          # V2  = (row2=0, col V-B=20)
    ...
}
```

### 폴링 주기 변경 시

각 reader의 `run()` 마지막 줄 `time.sleep(초)` 수정. 현재 SSF/SSO 모두 `1.0`초.

### 스냅샷 생성

```bash
python reader_ssf.py   # → ssf_snapshot.html, ssf_idx_snapshot.html
python reader_sso.py   # → templates/sso_snapshot.html
```

### 홈 화면 이미지 변경 시

`static/` 폴더에 이미지 파일 교체 후 `templates/welcome.html`의 `<img src>` 경로 수정.

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
