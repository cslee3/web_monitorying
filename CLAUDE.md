# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

**GlobalMM 실시간 모니터링 시스템** — 회사 내부망(폐쇄망) Windows 환경에서 Excel 파일을 실시간으로 읽어 웹 브라우저로 제공하는 시스템.

- **Excel → COM → Python → SSE → 브라우저** 파이프라인
- Windows 전용 (win32com COM 방식 의존)
- 관리자 권한 없음, 인터넷 차단 환경

## 실행 방법

```bash
# 웹 서버 시작 (포트 8080)
python serve.py

# 스냅샷 HTML 단독 생성 (SSE 없이 확인할 때)
python reader_ssf.py     # → ssf_snapshot.html + ssf_idx_snapshot.html
python reader_sso.py     # → sso_snapshot.html

# DB 초기화 (최초 1회, MySQL이 실행 중이어야 함)
python database_setting.py

# MySQL 수동 시작 (서비스 등록 불가) — 프로젝트와 동등한 위치
cd D:/OneDrive/mysql/bin && ./mysqld --console
```

## 아키텍처

### 실시간 데이터 흐름

```
GlobalMM_Realtime_Monitoring_V2_20260520.xlsb (Excel, 반드시 열려있어야 함)
  └─ win32com.client.GetObject()  ← ROT에서 직접 참조 (xw.apps 열거 우회)
       ├─ DashboardMonitor.run()  ← 백그라운드 스레드, 0.3초 폴링 (reader_ssf.py)
       │    └─ _read_full() → 4개 범위 bulk read → sections 캐시 → SSE diff push
       └─ SsoMonitor.run()        ← 백그라운드 스레드, 1초 폴링 (reader_sso.py)
            └─ _read_full() → B3:M9 bulk read → rows 캐시 → SSE diff push
```

### Flask 엔드포인트 (serve.py)

| 경로 | 역할 |
|---|---|
| `GET /` | `index.html` 반환 |
| `GET /api/dashboard` | SSF 전체 JSON (초기 로드, 미준비 시 503) |
| `GET /api/stream` | SSF SSE 스트림 — 변경 셀만 push, 20초 keepalive |
| `GET /api/sso` | SSO 전체 JSON (초기 로드) |
| `GET /api/sso/stream` | SSO SSE 스트림 |
| `GET /<path>` | 정적 파일 서빙 |

### 프론트엔드 구조 (index.html)

- 로그인 → iframe 기반 SPA
- 사이드바 메뉴에서 각 HTML 페이지를 iframe으로 로드
- `ssfo_pnl.html` / `ssfo_idx.html` / `sso_report.html` 이 iframe 안에서 동작
- 날짜 picker → `archive/<page>_<yyyymmdd>.html` 로드 (아카이브 뷰)

### Excel 탭 → 코드 매핑

| Excel 탭 | 코드 | 읽는 범위 |
|---|---|---|
| `DashBoard` | `reader_ssf.py` | B5:AA36, B38:AA40, B42:AA44, AC5:AI36 (4개) |
| `Option_DashBoard` | `reader_sso.py` | B3:M9 |
| `Position` | `database_setting.py` | skiprows=2 |
| `T-1_Position` | `database_setting.py` | skiprows=2 |

### reader_ssf.py 섹션 구성

`_read_full()`이 반환하는 `result["sections"]` 리스트:

| id | 범위 | COLS | 설명 |
|---|---|---|---|
| `main` | B5:AA36 | `MM_COLS` | SSF MM 개별 종목 |
| `sum1` | B38:AA40 | `MM_COLS` | SSF MM 소계 |
| `sum2` | B42:AA44 | `MM_COLS` | SSF MM 합계 |
| `idx`  | AC5:AI36 | `IDX_COLS` | SSF MM2 지수 |

- `MM_COLS` cidx: B=0 기준 (B=0 … AA=25, I=cidx7 skip)
- `IDX_COLS` cidx: AC=0 기준 (AC=0 … AI=6)
- SSE diff key 형식: `"{sid}_{ri}_{cidx}"` (예: `main_3_15`)
- HTML 셀 ID: `c-{sid}-{ri}-{cidx}`

### ssfo_pnl.html 특수 처리

- MM Spread Alert: `cidx === 10` (노란 배경 + 빨간 글자)
- Delta(Shr)=0 파란 배경: `cidx === 15`
- Delta(Mil)=0.00 파란 배경: `cidx === 16`
- DIM_STOCKS 체크: `row["0"]?.v` (Stock = cidx 0)
- `sid === "idx"` 셀은 applyChanges에서 무시

### ssfo_idx.html 특수 처리

- idx 섹션만 렌더링 (MM 섹션 무시)
- IdxCode = cidx 0, IdxName = cidx 1
- SECTION_STARTERS로 그룹 분할 (IdxCode 값 기준)

### 컬럼 정의 위치

- **`reader_ssf.py`** `MM_COLS` / `IDX_COLS` — cidx는 각 범위 시작열 기준 0-based
- **`reader_sso.py`** `COLS` — cidx는 B열 기준 0-based (B=0)
- **`reader_common.py`** — 공유 포매터: `_pct`, `_num`, `_date`, `_rgb`

## 핵심 설계 결정

**win32com vs xlwings**: `_get_ws()`는 `win32com.client.GetObject(path)`로 ROT에서 직접 참조. `xw.apps` 열거를 사용하면 `get_xl_app_from_hwnd` access violation이 발생하므로 완전히 우회함.

**COM 스레드 초기화**: `run()` 시작 시 반드시 `pythoncom.CoInitialize()` 호출 (비메인 스레드에서 COM 사용 필수).

**SSE 방식**: 전체 재전송 대신 변경된 셀만 push. 신규 브라우저 접속 시 `/api/dashboard`로 전체 스냅샷 먼저 받고 이후 SSE로 diff 적용.

**cidx = 0-based 범위 시작열 기준**: 각 섹션의 cidx는 해당 범위의 첫 열을 0으로 하는 인덱스. `row_data[cidx]`로 바로 접근 가능 (offset 연산 없음).

**Spyder/IPython `__file__` 미정의 대응**: `try: BASE_DIR = os.path.dirname(os.path.abspath(__file__)) except NameError: BASE_DIR = os.getcwd()`

**`_pct` 예외처리**: Excel 셀에 텍스트("합계" 등)가 있을 때 `float(v)*100` 변환 실패 → `except: return str(v)`로 처리.

## 현재 타겟 파일

`BOOK_NAME = "GlobalMM_Realtime_Monitoring_V2_20260520.xlsb"` — `reader_ssf.py` 상단에서 변경.

## 패키지 설치 (폐쇄망)

```bash
# 인터넷 가능한 PC에서 wheel 다운로드
pip download flask pywin32 -d ./wheels/flask
pip download mysql-connector-python -d ./wheels/mysql

# 대상 서버에서 오프라인 설치
pip install --no-index --find-links=./wheels/flask flask pywin32
pip install --no-index --find-links=./wheels/mysql mysql-connector-python
```

## `before/` 폴더

구버전 Streamlit 기반 앱 (`app.py`, `reader.py`, `renderer.py`). 현재 메인 시스템에서 사용하지 않음.
