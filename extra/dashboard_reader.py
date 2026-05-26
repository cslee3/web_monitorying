# -*- coding: utf-8 -*-
"""
dashboard_reader.py — 하위 호환 shim (test_reader.py 등에서 import 유지용)
실제 구현은 reader_ssf.py / reader_sso.py 로 이동됨.
"""
# 새 코드는 reader_ssf / reader_sso 를 직접 import하세요.

import datetime, threading, json, os
from queue import Queue, Empty
import pythoncom
import win32com.client
import xlwings as xw

# 읽어올 Excel 파일명 (serve.py와 같은 디렉토리에 위치해야 함)
BOOK_NAME = "GlobalMM_Realtime_Monitoring_V2_20260520.xlsb"


# ── 값 포매터 ──────────────────────────────────────────────────────────────

def _pct(v):
    """소수 → 퍼센트 문자열. ex) 0.0123 → '1.23%'"""
    if v is None: return ""
    return f"{v*100:.2f}%"

def _num(v, d=0):
    """숫자 → 천단위 콤마 + 소수점 d자리. ex) 1234567.8 → '1,234,568'"""
    if v is None or v == "": return ""
    try: return f"{v:,.{d}f}"
    except: return str(v)

def _date(v):
    """datetime → 'YYYY-MM-DD' 문자열."""
    if v is None: return ""
    if isinstance(v, datetime.datetime): return v.strftime("%Y-%m-%d")
    return str(v)

def _rgb(color):
    """xlwings color tuple (r,g,b) → CSS rgb() 문자열."""
    if isinstance(color, tuple) and len(color) == 3:
        return f"rgb({color[0]},{color[1]},{color[2]})"
    return ""


# ── 컬럼 정의 ──────────────────────────────────────────────────────────────
# (Excel 1-based 열 번호, 화면 헤더 텍스트, 값 포매터 함수)
COLS = [
    (1,  "ID",            lambda v: str(v) if v else ""),
    (2,  "Stock",         lambda v: str(v) if v else ""),
    (3,  "StockName",     lambda v: str(v) if v else ""),
    (4,  "LastPrice",     lambda v: _num(v, 0)),
    (5,  "Change(%)",     _pct),
    (6,  "Amt(Mil)",      lambda v: _num(v, 1)),
    (7,  "Amt(Shr)",      lambda v: _num(v, 0)),
    (8,  "MMQty",         lambda v: _num(v, 0)),
    (10, "Ask",           lambda v: _num(v, 0)),
    (11, "Bid",           lambda v: _num(v, 0)),
    (12, "MM Spread",     lambda v: str(v) if v else ""),
    (13, "Shares",        lambda v: _num(v, 0)),
    (14, "Diff",          lambda v: _num(v, 0)),
    (15, "Vol(Lots)",     lambda v: _num(v, 0)),
    (16, "Vol(Mil)",      lambda v: _num(v, 2)),
    (17, "Delta(Shr)",    lambda v: _num(v, 0)),
    (18, "Delta(Mil)",    lambda v: _num(v, 2)),
    (19, "Theo Price",    lambda v: _num(v, 0)),
    (20, "Theo Basis",    lambda v: _num(v, 2)),
    (21, "Ex1 B-Sp",      lambda v: _num(v, 2)),
    (22, "Ex1 S-Sp",      lambda v: _num(v, 2)),
    (23, "Ex2 Basis",     lambda v: _num(v, 2)),
    (24, "Ex2 B-Sp",      lambda v: _num(v, 2)),
    (25, "Ex2 S-Sp",      lambda v: _num(v, 2)),
    (26, "MTM PnL",       lambda v: _num(v, 0)),
    (27, "Theo PnL",      lambda v: _num(v, 0)),
    # SSF MM2 전용 컬럼 (AC~AI열, ssfo_idx.html에서만 표시)
    (29, "IdxCode",       lambda v: str(v) if v else ""),
    (30, "IdxName",       lambda v: str(v) if v else ""),
    (31, "Fund",          lambda v: _num(v, 0)),
    (32, "Delta",         lambda v: _num(v, 0)),
    (33, "Volume",        lambda v: _num(v, 0)),
    (34, "MTM PnL(Idx)",  lambda v: _num(v, 0)),
    (35, "Theo PnL(Idx)", lambda v: _num(v, 0)),
]

# PnL 관련 컬럼 인덱스 집합 (녹색/적색 텍스트 적용 대상)
PNL_CIDX = {c[0] for c in COLS if "PnL" in c[1]}

# 숫자 정렬 대상 컬럼 인덱스 (텍스트/코드 컬럼 제외)
NUM_CIDX  = {c[0] for c in COLS if c[0] not in (1, 2, 3, 12, 29, 30)}


def _pnl_color(cidx, raw):
    """PnL 컬럼의 값에 따라 'green' / 'red' / '' 반환."""
    if cidx not in PNL_CIDX or raw is None: return ""
    try: return "green" if float(raw) > 0 else ("red" if float(raw) < 0 else "")
    except: return ""


class DashboardMonitor:
    """
    Excel DashBoard 탭을 주기적으로 읽어 변경분을 SSE 클라이언트에 push하는 모니터.
    - _full_data  : 최신 전체 테이블 데이터 (신규 브라우저 접속 시 제공)
    - _clients    : 연결된 SSE 클라이언트 큐 목록
    """

    def __init__(self):
        self._clients_lock = threading.Lock()
        self._clients      = []          # 연결 중인 클라이언트 큐 목록
        self._full_lock    = threading.Lock()
        self._full_data    = None        # 가장 최근 전체 데이터 (색상 포함)

    # ── 클라이언트 큐 관리 ────────────────────────────────────────────────

    def add_client(self):
        """새 SSE 클라이언트 큐를 생성하고 목록에 추가. 큐 반환."""
        q = Queue(maxsize=200)
        with self._clients_lock:
            self._clients.append(q)
        return q

    def remove_client(self, q):
        """클라이언트 연결 종료 시 큐를 목록에서 제거."""
        with self._clients_lock:
            try: self._clients.remove(q)
            except ValueError: pass

    def get_full(self):
        """현재 캐시된 전체 데이터 반환. 아직 초기화 전이면 None."""
        with self._full_lock:
            return self._full_data

    def _push(self, payload_str):
        """
        모든 클라이언트 큐에 변경 데이터 문자열을 전송.
        큐가 가득 찬 클라이언트(응답 없음)는 자동 제거.
        """
        with self._clients_lock:
            dead = []
            for q in self._clients:
                try: q.put_nowait(payload_str)
                except: dead.append(q)
            for q in dead:
                self._clients.remove(q)

    # ── 초기 전체 데이터 읽기 ─────────────────────────────────────────────

    def _read_full(self, ws):
        """
        DashBoard 탭 전체를 한 번에 읽어 JSON-직렬화 가능한 딕셔너리로 반환.
        - 헤더 정보: 기준일, YTD, I/R, 만기일 1·2 (COM 5회)
        - 데이터 범위: A6:AK45 (COM 1회 bulk read)
        - 빈 행은 sep=True 구분선으로 처리
        - 합계 행(col2='합계' 등)은 total=True 처리
        """
        # 헤더 셀 개별 읽기 (win32com, 5회 COM)
        header = {
            "base_date": _date(ws.Cells(2, 3).Value),
            "ytd":       _date(ws.Cells(3, 3).Value),
            "ir":        str(ws.Cells(2, 6).Value),
            "exp1":      _date(ws.Cells(2, 9).Value),
            "exp2":      _date(ws.Cells(3, 9).Value),
        }

        # 데이터 전체를 1회 COM 호출로 읽기 (40행 × 37열, win32com)
        all_vals = ws.Range("A6:AK45").Value

        rows, row_meta = [], []
        for row_data in all_vals:
            if row_data is None:
                row_data = [None] * 37

            cells = {}
            any_val = False
            for cidx, _, fmt in COLS:
                raw = row_data[cidx - 1]  # Excel 1-based → Python 0-based
                fval = ""
                if raw is not None and raw != "":
                    try: fval = fmt(raw)
                    except: fval = str(raw)
                    any_val = True
                cells[str(cidx)] = {"v": fval, "bg": "", "bold": False,
                                    "color": _pnl_color(cidx, raw)}

            if not any_val:
                # 모든 컬럼이 빈 행 → 구분선으로 처리
                row_meta.append({"sep": True,  "total": False})
                rows.append({})
                continue

            is_total = cells.get("2", {}).get("v", "") in ("합계", "합 계", "Total", "소계")
            row_meta.append({"sep": False, "total": is_total})
            rows.append(cells)

        return {
            "header":   header,
            "cols":     [(c[0], c[1]) for c in COLS],  # [(cidx, 헤더), ...]
            "num_cidx": list(NUM_CIDX),                 # 숫자 정렬 컬럼 목록
            "rows":     rows,
            "row_meta": row_meta,
            "updated":  datetime.datetime.now().strftime("%H:%M:%S"),
        }

    # ── Excel 워크시트 참조 획득 ──────────────────────────────────────────

    def _get_ws(self):
        """
        win32com.client.GetObject으로 워크북을 직접 참조하여 DashBoard 시트 반환.
        xw.books/xw.apps의 get_xl_app_from_hwnd 열거를 완전히 우회
        → access violation 방지.
        파일이 열려있지 않으면 Excel로 직접 열기.
        """
        path = os.path.join(os.path.dirname(__file__), BOOK_NAME)
        try:
            # 이미 열린 파일을 ROT(Running Object Table)에서 직접 참조
            com_wb = win32com.client.GetObject(path)
            return com_wb.Sheets("DashBoard")
        except Exception:
            pass
        # 열려있지 않으면 직접 열기
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        com_wb = xl.Workbooks.Open(path)
        return com_wb.Sheets("DashBoard")

    # ── 메인 모니터 루프 ──────────────────────────────────────────────────

    def run(self):
        """
        백그라운드 스레드에서 실행되는 메인 루프.
        1) CoInitialize(): 비메인 스레드에서 COM 사용 시 필수 초기화
        2) 0.3초마다 A6:AK45 범위 bulk read (COM 1회)
        3) 첫 실행: _read_full()로 전체 데이터 캐시
        4) 이후: 이전 값과 diff 비교 → 변경 셀만 클라이언트에 push
        5) 예외 발생 시 prev_vals 초기화 후 다음 루프에서 재시도
        """
        pythoncom.CoInitialize()   # 백그라운드 스레드 COM 초기화 (필수)
        prev_vals = None

        while True:
            try:
                ws = self._get_ws()

                # 전체 데이터 범위를 단일 COM 호출로 읽기 (40행 × 37열, win32com)
                all_vals = ws.Range("A6:AK45").Value

                # ── 첫 실행: 색상 포함 전체 데이터 캐시 ──────────────────
                if prev_vals is None:
                    full = self._read_full(ws)
                    with self._full_lock:
                        self._full_data = full
                    prev_vals = [list(row) if row else [None]*37 for row in all_vals]
                    print("[monitor] initial load done")
                    continue

                # ── 이후 실행: 이전 값과 비교하여 변경 셀 추출 ───────────
                changes = []
                now_str = datetime.datetime.now().strftime("%H:%M:%S")

                for ri, (new_row, old_row) in enumerate(zip(all_vals, prev_vals)):
                    if new_row is None: new_row = [None] * 37
                    for cidx, _, fmt in COLS:
                        col_i   = cidx - 1
                        new_raw = new_row[col_i]
                        old_raw = old_row[col_i]
                        if new_raw == old_raw:
                            continue  # 변경 없으면 skip
                        fval = ""
                        if new_raw is not None and new_raw != "":
                            try: fval = fmt(new_raw)
                            except: fval = str(new_raw)
                        changes.append({
                            "key":   f"{ri}_{cidx}",  # "행인덱스_열인덱스"
                            "v":     fval,
                            "color": _pnl_color(cidx, new_raw),
                        })

                if changes:
                    # 캐시(_full_data) 업데이트 후 클라이언트에 push
                    with self._full_lock:
                        if self._full_data:
                            for ch in changes:
                                ri, cidx = map(int, ch["key"].split("_"))
                                row = self._full_data["rows"][ri]
                                if str(cidx) in row:
                                    row[str(cidx)]["v"]     = ch["v"]
                                    row[str(cidx)]["color"] = ch["color"]
                            self._full_data["updated"] = now_str
                    self._push(json.dumps({"cells": changes, "updated": now_str}))
                    print(f"[monitor] {now_str} — pushed {len(changes)} changes")

                # 현재 값을 다음 루프의 비교 기준으로 저장
                prev_vals = [list(row) if row else [None]*37 for row in all_vals]

            except Exception as e:
                print(f"[monitor] error: {e}")
                prev_vals = None  # 예외 발생 시 다음 루프에서 전체 재초기화

            import time; time.sleep(0.3)  # 0.3초 대기 후 다음 폴링


monitor = DashboardMonitor()


# ── SSO Report 생성 (Option_DashBoard → sso_report.html) ──────────────────
# generate_sso_report()는 standalone 실행용으로만 정의.
# serve.py의 monitor 루프에서는 호출하지 않음 (COM 충돌 방지).

OUT_PATH = os.path.join(os.path.dirname(__file__), "sso_report.html")

# SSO 표시 컬럼 정의 (Excel 1-based 열 번호, 헤더, 포매터)
SSO_COLS = [
    (2,  "StockCode",  lambda v: str(v) if v else ""),
    (3,  "StockName",  lambda v: str(v) if v else ""),
    (4,  "Delta",      lambda v: _num(v, 0)),
    (6,  "%Gamma",     lambda v: _num(v, 2)),
    (8,  "Volume",     lambda v: _num(v, 0)),
    (10, "Theo PnL",   lambda v: _num(v, 0)),
    (12, "MTM PnL",    lambda v: _num(v, 0)),
]
_SSO_PNL = {c[0] for c in SSO_COLS if "PnL" in c[1]}


def _sso_cell_style(bg, bold, val, cidx):
    """SSO 셀 인라인 스타일 문자열 생성. PnL 컬럼은 값에 따라 색상 적용."""
    styles = []
    if bg: styles.append(f"background:{bg}")
    if bold: styles.append("font-weight:bold")
    if cidx in _SSO_PNL:
        try:
            f = float(str(val).replace(",", ""))
            styles.append("color:#1a7c34" if f > 0 else ("color:#cc0000" if f < 0 else ""))
        except Exception:
            pass
    return "; ".join(s for s in styles if s)


def generate_sso_report(ws_opt):
    """
    Option_DashBoard 워크시트를 읽어 sso_report.html을 생성.
    A3:L235 범위를 1회 bulk read하여 COM 부하 최소화.
    StockCode 컬럼(B열)이 빈 행 또는 헤더 행은 skip.
    """
    # A3:L235 전체를 1회 COM 호출로 읽기
    raw = ws_opt.range("A3:L235").value

    rows_html = []
    for row in raw:
        if row is None:
            continue
        stock_code = row[1]  # B열 (0-based index 1)
        if not isinstance(stock_code, str) or stock_code in ("", "StockCode"):
            continue

        tds = []
        for cidx, _, fmt in SSO_COLS:
            v = row[cidx - 1]  # Excel 1-based → 0-based
            formatted = ""
            if v is not None and v != "":
                try: formatted = fmt(v)
                except: formatted = str(v)
            style = _sso_cell_style("", False, formatted, cidx)
            align = "right" if cidx not in (2, 3) else "left"
            tds.append(f'<td style="{style}; text-align:{align}">{formatted}</td>')
        rows_html.append(f'<tr>{"".join(tds)}</tr>')

    headers = "".join(f"<th>{c[1]}</th>" for c in SSO_COLS)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8"/>
<meta http-equiv="refresh" content="10"/>
<title>SSO MM Dashboard</title>
<style>
  body {{ margin:0; padding:8px; font-family:'Segoe UI',sans-serif; font-size:12px; background:#fff; }}
  .refresh {{ text-align:right; color:#888; font-size:11px; margin-bottom:6px; }}
  table {{ border-collapse:collapse; width:100%; white-space:nowrap; }}
  th {{ background:#202123; color:#fff; padding:4px 8px; position:sticky; top:0; z-index:1;
        font-size:11px; text-align:center; border:1px solid #444; }}
  td {{ padding:3px 8px; border:1px solid #ddd; font-size:12px; }}
  tr:hover td {{ background:#fffbe6 !important; }}
  .table-wrap {{ overflow-x:auto; overflow-y:auto; max-height:calc(100vh - 40px); }}
</style></head><body>
<div class="refresh">Updated: {now}</div>
<div class="table-wrap"><table>
<thead><tr>{headers}</tr></thead>
<tbody>{"".join(rows_html)}</tbody>
</table></div></body></html>"""

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[sso_report] Generated at {now}")


# ── SSO 실시간 모니터 (Option_DashBoard 탭) ──────────────────────────────

_SSO_NUM_CIDX = {c[0] for c in SSO_COLS if c[0] not in (2, 3)}


def _sso_pnl_color(cidx, raw):
    """SSO PnL 컬럼 값에 따라 'green' / 'red' / '' 반환."""
    if cidx not in _SSO_PNL or raw is None: return ""
    try: return "green" if float(raw) > 0 else ("red" if float(raw) < 0 else "")
    except: return ""


class SsoMonitor:
    """
    Option_DashBoard 탭을 주기적으로 읽어 변경분을 SSE 클라이언트에 push.
    DashboardMonitor와 동일한 구조, 별도 스레드에서 1초 폴링.
    """

    def __init__(self):
        self._clients_lock = threading.Lock()
        self._clients      = []
        self._full_lock    = threading.Lock()
        self._full_data    = None

    def add_client(self):
        q = Queue(maxsize=200)
        with self._clients_lock:
            self._clients.append(q)
        return q

    def remove_client(self, q):
        with self._clients_lock:
            try: self._clients.remove(q)
            except ValueError: pass

    def get_full(self):
        with self._full_lock:
            return self._full_data

    def _push(self, payload_str):
        with self._clients_lock:
            dead = []
            for q in self._clients:
                try: q.put_nowait(payload_str)
                except: dead.append(q)
            for q in dead:
                self._clients.remove(q)

    def _get_ws(self):
        """win32com ROT에서 Option_DashBoard 시트 직접 참조."""
        path = os.path.join(os.path.dirname(__file__), BOOK_NAME)
        try:
            com_wb = win32com.client.GetObject(path)
            return com_wb.Sheets("Option_DashBoard")
        except Exception:
            pass
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        com_wb = xl.Workbooks.Open(path)
        return com_wb.Sheets("Option_DashBoard")

    def _read_full(self, ws):
        """A3:L235 전체를 1회 COM 호출로 읽어 JSON-직렬화 가능한 딕셔너리 반환."""
        all_vals = ws.Range("A3:L235").Value  # 233행 × 12열
        rows, row_meta = [], []
        for row_data in all_vals:
            if row_data is None:
                row_data = [None] * 12
            stock_code = row_data[1]  # B열 (0-based)
            if not isinstance(stock_code, str) or stock_code in ("", "StockCode"):
                row_meta.append({"skip": True})
                rows.append({})
                continue
            cells = {}
            for cidx, _, fmt in SSO_COLS:
                raw = row_data[cidx - 1]
                fval = ""
                if raw is not None and raw != "":
                    try: fval = fmt(raw)
                    except: fval = str(raw)
                cells[str(cidx)] = {"v": fval, "color": _sso_pnl_color(cidx, raw)}
            row_meta.append({"skip": False})
            rows.append(cells)
        return {
            "cols":     [(c[0], c[1]) for c in SSO_COLS],
            "num_cidx": list(_SSO_NUM_CIDX),
            "rows":     rows,
            "row_meta": row_meta,
            "updated":  datetime.datetime.now().strftime("%H:%M:%S"),
        }

    def run(self):
        """
        백그라운드 스레드에서 실행. 1초 폴링.
        DashboardMonitor.run()과 동일한 패턴.
        """
        pythoncom.CoInitialize()
        prev_vals = None

        while True:
            try:
                ws = self._get_ws()
                all_vals = ws.Range("A3:L235").Value

                if prev_vals is None:
                    full = self._read_full(ws)
                    with self._full_lock:
                        self._full_data = full
                    prev_vals = [list(row) if row else [None]*12 for row in all_vals]
                    print("[sso_monitor] initial load done")
                    import time; time.sleep(1.0)
                    continue

                changes = []
                now_str = datetime.datetime.now().strftime("%H:%M:%S")

                for ri, (new_row, old_row) in enumerate(zip(all_vals, prev_vals)):
                    if new_row is None: new_row = [None] * 12
                    stock_code = new_row[1]
                    if not isinstance(stock_code, str) or stock_code in ("", "StockCode"):
                        continue  # 스킵 행은 비교 안 함
                    for cidx, _, fmt in SSO_COLS:
                        col_i   = cidx - 1
                        new_raw = new_row[col_i]
                        old_raw = old_row[col_i]
                        if new_raw == old_raw:
                            continue
                        fval = ""
                        if new_raw is not None and new_raw != "":
                            try: fval = fmt(new_raw)
                            except: fval = str(new_raw)
                        changes.append({
                            "key":   f"{ri}_{cidx}",
                            "v":     fval,
                            "color": _sso_pnl_color(cidx, new_raw),
                        })

                if changes:
                    with self._full_lock:
                        if self._full_data:
                            for ch in changes:
                                ri, cidx = map(int, ch["key"].split("_"))
                                row = self._full_data["rows"][ri]
                                if str(cidx) in row:
                                    row[str(cidx)]["v"]     = ch["v"]
                                    row[str(cidx)]["color"] = ch["color"]
                            self._full_data["updated"] = now_str
                    self._push(json.dumps({"cells": changes, "updated": now_str}))
                    print(f"[sso_monitor] {now_str} — pushed {len(changes)} changes")

                prev_vals = [list(row) if row else [None]*12 for row in all_vals]

            except Exception as e:
                print(f"[sso_monitor] error: {e}")
                prev_vals = None

            import time; time.sleep(1.0)


sso_monitor = SsoMonitor()
