# -*- coding: utf-8 -*-
"""
reader_sso.py — Option_DashBoard 탭 실시간 모니터 (SSO MM)
- SsoMonitor.run() 을 별도 스레드에서 실행
- 1초마다 A3:L235 범위를 bulk read하여 이전 값과 비교
- 변경된 셀만 SSE 클라이언트에 push
- serve.py에서 import하여 사용
"""

import datetime, threading, json, os, webbrowser, time
from queue import Queue
import pythoncom
import win32com.client
import pywintypes

_COM_BUSY = {-2147418111, -2147417846}

from common import BASE_DIR, BOOK_NAME, _num, _colorscale_bg
from logger import get_logger
log = get_logger(__name__)

# ── 컬럼 정의 ──────────────────────────────────────────────────────────────
# (Excel 1-based 열 번호, 화면 헤더 텍스트, 값 포매터 함수)
# cidx = row_data 0-based 인덱스 (범위 시작열 기준)
# B3:M9 기준: B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7, J=8, K=9, L=10, M=11
COLS = [
    (0,  "StockCode",  lambda v: str(v) if v else ""),
    (1,  "StockName",  lambda v: str(v) if v else ""),
    (2,  "Delta",      lambda v: _num(v, 2)),
    (4,  "%Gamma",     lambda v: _num(v, 2)),
    (6,  "Vol Cnt",      lambda v: _num(v, 0)),
    (7,  "Vol Amt",      lambda v: _num(v, 0)),
    (8,  "Theo PnL",   lambda v: _num(v, 0)),
    (10, "MTM PnL",    lambda v: _num(v, 0)),
]

NUM_CIDX  = {c[0] for c in COLS if c[0] not in (0, 1)}
COLORSCALE_CIDX = {8, 10}  # Theo PnL, MTM PnL

# ── 상세 컬럼 정의 (B11:Q350 기준, cidx = B열 0-based) ──────────────────
# 종목 소계 / 선물 / 옵션 행을 동일 컬럼 구조로 flat하게 처리
ALL_DETAIL_FMTS = {
    0:  lambda v: str(v).strip() if v else "",
    1:  lambda v: str(v).strip() if v else "",
    2:  lambda v: _num(v, 0),
    3:  lambda v: _num(v, 0),
    4:  lambda v: _num(v, 0),
    5:  lambda v: _num(v, 2),
    6:  lambda v: _num(v, 3),
    7:  lambda v: _num(v, 0),
    8:  lambda v: _num(v, 0),
    9:  lambda v: _num(v, 2),
    10: lambda v: str(v) if v else "",
    11: lambda v: str(v) if v else "",
    12: lambda v: _num(v, 0),
    13: lambda v: str(v) if v else "",
    14: lambda v: _num(v, 2),
    15: lambda v: str(v) if v else "",
    16: lambda v: _num(v, 0),
    17: lambda v: _num(v, 0),
    18: lambda v: _num(v, 0),
    19: lambda v: _num(v, 2),
    20: lambda v: _num(v, 3),
    21: lambda v: _num(v, 0),
    22: lambda v: _num(v, 0),
}

class SsoMonitor:
    """
    Option_DashBoard 탭을 주기적으로 읽어 변경분을 SSE 클라이언트에 push.
    - _full_data  : 최신 전체 테이블 데이터 (신규 브라우저 접속 시 제공)
    - _clients    : 연결된 SSE 클라이언트 큐 목록
    """

    def __init__(self):
        self._clients_lock  = threading.Lock()
        self._clients       = []
        self._full_lock     = threading.Lock()
        self._full_data     = None
        self._running        = False  # 기본값: 중지 (터미널에서 r 입력으로 시작)
        self._read_ok        = False  # 마지막 Excel 읽기 성공 여부
        self._last_status_ts = 0
        self._changes_since  = 0
        self._wake           = threading.Event()

    # ── 모니터 시작/중지 ──────────────────────────────────────────────────

    def pause(self):
        self._running = False
        self._push(json.dumps({"type": "monitor_paused", "mon": "sso"}))

    def resume(self):
        self._running = True
        self._wake.set()
        self._push(json.dumps({"type": "monitor_resumed", "mon": "sso"}))

    def is_running(self):
        return self._running

    def _print_status(self):
        now_str  = datetime.datetime.now().strftime("%H:%M:%S")
        run_str  = "SSO:실행중" if self._running else "SSO:중지"
        xl_str   = "Excel:OK" if self._read_ok else "Excel:대기"
        with self._clients_lock:
            n_clients = len(self._clients)
        changes = self._changes_since
        self._changes_since = 0
        log.info(f"{run_str} | {xl_str} | 클라이언트:{n_clients} | 변경:{changes}건/30s")

    # ── 클라이언트 큐 관리 ────────────────────────────────────────────────

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

    # ── Excel 워크시트 참조 획득 ──────────────────────────────────────────

    def _get_ws(self):
        """실행 중인 Excel에서 BOOK_NAME 워크북을 이름으로 찾아 Option_DashBoard 시트 반환.
        GetObject(path) 대신 GetActiveObject 사용 — OneDrive 경로 불일치 문제 회피.
        """
        xl = win32com.client.GetActiveObject("Excel.Application")
        for i in range(1, xl.Workbooks.Count + 1):
            wb = xl.Workbooks.Item(i)
            if wb.Name == BOOK_NAME:
                return wb.Sheets("Option_DashBoard")
        raise RuntimeError(f"'{BOOK_NAME}' 파일이 Excel에서 열려있지 않음")

    # ── 상세 데이터 파싱 (3초 캐시) ──────────────────────────────────────

    def _parse_detail(self, ws):
        """B11:Q350 범위를 읽어 종목별 flat rows 반환.
        각 row: {"type": "summary"|"futures"|"option", "cells": {cidx: {v, color}}}
        """
        try:
            raw = ws.Range("B11:X350").Value
        except Exception:
            return {}

        detail = {}
        current_stock = None

        for row_data in raw:
            b = row_data[0]
            c = row_data[1]
            b_str = str(b).strip() if b is not None else ""
            c_str = str(c).strip() if c is not None else ""

            if not b_str and not c_str:
                continue

            # 행 타입 판별
            if b_str and c_str:
                row_type = "summary"
                current_stock = b_str
                detail[current_stock] = {"name": c_str, "rows": []}
            elif not b_str and c_str.startswith("A"):
                row_type = "futures"
                if not current_stock or current_stock not in detail:
                    continue
            elif b_str and not c_str:
                row_type = "option"
                # Strike 없으면 filler 행
                if not current_stock or current_stock not in detail:
                    continue
                if row_data[12] is None:
                    continue
            else:
                continue

            cells = {}
            for cidx, fmt in ALL_DETAIL_FMTS.items():
                raw_v = row_data[cidx] if cidx < len(row_data) else None
                fval  = fmt(raw_v) if (raw_v is not None and raw_v != "") else ""
                cells[str(cidx)] = {"v": fval}

            detail[current_stock]["rows"].append({"type": row_type, "cells": cells})

        if not detail:
            log.warning("_parse_detail: no rows parsed (check B11:Q350 range)")
        return detail

    # ── 전체 데이터 읽기 ──────────────────────────────────────────────────

    def _read_full(self, ws, with_detail=True):
        all_vals = ws.Range("B3:M9").Value

        # ColorScale 컬럼별 min/max 사전 계산
        cs_vals = {cidx: [] for cidx in COLORSCALE_CIDX}
        for row_data in all_vals:
            for cidx in COLORSCALE_CIDX:
                v = row_data[cidx] if cidx < len(row_data) else None
                try: cs_vals[cidx].append(float(v))
                except (TypeError, ValueError): pass

        cs_min, cs_max = {}, {}
        for cidx, vals in cs_vals.items():
            neg = [v for v in vals if v < 0]
            pos = [v for v in vals if v > 0]
            cs_min[cidx] = min(neg) if neg else None
            cs_max[cidx] = max(pos) if pos else None

        rows = []
        for row_data in all_vals:
            cells = {}
            for cidx, _, fmt in COLS:
                raw  = row_data[cidx]
                fval = fmt(raw) if (raw is not None and raw != "") else ""
                bg   = _colorscale_bg(raw, cs_min.get(cidx), cs_max.get(cidx)) if cidx in COLORSCALE_CIDX else ""
                cells[str(cidx)] = {"v": fval, "bg": bg}
            rows.append(cells)

        result = {
            "cols":     [(c[0], c[1]) for c in COLS],
            "num_cidx": list(NUM_CIDX),
            "rows":     rows,
            "updated":  datetime.datetime.now().strftime("%H:%M:%S"),
            "detail":   self._parse_detail(ws) if with_detail else None,
        }
        return result

    # ── 메인 모니터 루프 ──────────────────────────────────────────────────

    def run(self):
        pythoncom.CoInitialize()
        prev_data = None
        ws        = None
        _last_detail_ts = 0.0   # 상세범위(무거움)는 ~1초마다만 읽음
        _cached_detail  = {}

        while True:
            if not self._running:
                if ws is not None:
                    ws        = None   # COM 참조 해제
                    prev_data = None
                    self._read_ok = False
                self._wake.wait(timeout=0.5)
                self._wake.clear()
                continue

            try:
                if ws is None:
                    ws = self._get_ws()
                _now = time.time()
                _read_detail = (_now - _last_detail_ts) >= 1.0
                data    = self._read_full(ws, with_detail=_read_detail)
                if _read_detail:
                    _last_detail_ts = _now
                    _cached_detail  = data["detail"]
                else:
                    data["detail"] = _cached_detail
                self._read_ok = True
                now_str = datetime.datetime.now().strftime("%H:%M:%S")

                if prev_data is None:
                    with self._full_lock:
                        self._full_data = data
                    prev_data = data
                    log.warning("initial load done")
                    self._wake.wait(timeout=1.0)
                    self._wake.clear()
                    continue

                changes = []
                for ri, (new_row, old_row) in enumerate(zip(data["rows"], prev_data["rows"])):
                    for cidx, _ in data["cols"]:
                        new_cell = new_row.get(str(cidx), {})
                        old_cell = old_row.get(str(cidx), {})
                        new_v = new_cell.get("v", "")
                        old_v = old_cell.get("v", "")
                        new_bg = new_cell.get("bg", "")
                        old_bg = old_cell.get("bg", "")
                        if new_v == old_v and new_bg == old_bg:
                            continue
                        changes.append({
                            "key": f"{ri}_{cidx}",
                            "v":   new_v,
                            "bg":  new_bg,
                        })

                # _full_data는 detail을 포함하므로 항상 갱신
                data["updated"] = now_str
                with self._full_lock:
                    self._full_data = data
                self._push(json.dumps({"cells": changes, "updated": now_str}))
                self._changes_since += len(changes)
                if changes:
                    log.info(f"{now_str} — pushed {len(changes)} changes")

                if time.time() - self._last_status_ts >= 30:
                    self._print_status()
                    self._last_status_ts = time.time()

                prev_data = data

            except pywintypes.com_error as e:
                hresult = e.args[0] if e.args else 0
                if hresult in _COM_BUSY:
                    self._wake.wait(timeout=0.5)
                    self._wake.clear()
                    continue
                log.error(f"com_error: {e}")
                self._read_ok = False
                prev_data = None
                ws        = None
            except Exception as e:
                log.error(f"error: {e}")
                self._read_ok = False
                prev_data = None
                ws        = None

            self._wake.wait(timeout=0.05)
            self._wake.clear()


monitor = SsoMonitor()
'''
self = monitor
'''


# ── 단독 실행: 스냅샷 HTML 생성 ──────────────────────────────────────────

_DETAIL_COLS = [
    (0,  "StockCode",   "left"),
    (1,  "Name/Code",   "left"),
    (2,  "C.Qty",       "right"),
    (3,  "C.△Qty",      "right"),
    (4,  "C.Vol",       "right"),
    (5,  "C.Delta",     "right"),
    (6,  "C.Gamma",     "right"),
    (7,  "C.Eval PnL",  "right"),
    (8,  "C.Trade PnL", "right"),
    (9,  "Call IV",     "right"),
    (10, "Call OTM",    "center"),
    (11, "Call Code",   "left"),
    (12, "Strike",      "right"),
    (13, "Put Code",    "left"),
    (14, "Put IV",      "right"),
    (15, "Put OTM",     "center"),
    (16, "P.Qty",       "right"),
    (17, "P.△Qty",      "right"),
    (18, "P.Vol",       "right"),
    (19, "P.Delta",     "right"),
    (20, "P.Gamma",     "right"),
    (21, "P.Eval PnL",  "right"),
    (22, "P.Trade PnL", "right"),
]

def generate_html(out_path=None, open_browser=True, _data=None):
    """
    Option_DashBoard 탭을 한 번 읽어 정적 HTML 스냅샷을 생성.
    요약 테이블 + 전 종목 상세 테이블 인라인 포함.
    _data: 이미 읽어둔 데이터 dict. None이면 Excel COM에서 직접 읽음.
    """
    if out_path is None:
        out_path = os.path.join(BASE_DIR, "templates", "sso_snapshot.html")

    if _data is not None:
        data = _data
    else:
        ws   = monitor._get_ws()
        data = monitor._read_full(ws)
    now  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── 요약 테이블 ──────────────────────────────────────────────────────
    headers = "".join(f"<th>{hdr}</th>" for _, hdr in data["cols"])
    num_set = set(data["num_cidx"])
    sum_rows = ""
    for row in data["rows"]:
        tds = ""
        for cidx, _ in data["cols"]:
            cell  = row.get(str(cidx), {"v": "", "bg": ""})
            align = "right" if cidx in num_set else "left"
            bg    = f"background:{cell['bg']};" if cell.get("bg") else ""
            tds  += f'<td style="text-align:{align};{bg}">{cell["v"]}</td>'
        sum_rows += f"<tr>{tds}</tr>"

    # ── 전 종목 상세 테이블 ───────────────────────────────────────────────
    det_thead = "".join(f'<th style="text-align:{a}">{h}</th>' for _, h, a in _DETAIL_COLS)
    det_body  = ""
    for code, stock in data.get("detail", {}).items():
        if not code.isdigit():
            continue
        rows_html = ""
        for row in stock["rows"]:
            t = row["type"]
            if   t == "summary": tr_style = 'style="background:#e8eaf6;font-weight:bold"'
            elif t == "futures": tr_style = 'style="background:#f3f4f6;color:#555"'
            elif row["cells"].get("10", {}).get("v") == "ATM":
                                  tr_style = 'style="background:#fff9c4;font-weight:bold"'
            else:                 tr_style = ""
            tds = "".join(
                f'<td style="text-align:{a}">{row["cells"].get(str(ci), {"v":""})["v"]}</td>'
                for ci, _, a in _DETAIL_COLS
            )
            rows_html += f"<tr {tr_style}>{tds}</tr>"
        det_body += (
            f'<div style="margin-top:20px;padding-bottom:6px;margin-bottom:6px;'
            f'border-bottom:2px solid #202123;font-size:13px;font-weight:bold;">'
            f'{code}&nbsp;&nbsp;{stock["name"]}</div>'
            f'<div style="overflow-x:auto">'
            f'<table><thead><tr>{det_thead}</tr></thead><tbody>{rows_html}</tbody></table>'
            f'</div>'
        )

    html = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8"/>
<title>SSO MM Snapshot</title>
<style>
  body {{ margin:0; padding:8px; font-family:'Segoe UI',sans-serif; font-size:12px; }}
  .ts {{ text-align:right; color:#888; font-size:11px; margin-bottom:6px; }}
  table {{ border-collapse:collapse; white-space:nowrap; outline:1px solid #000; outline-offset:-1px; }}
  th {{ background:#DEE7ED; color:#000; font-weight:bold; padding:4px 8px; font-size:11px;
        text-align:center; border:1px solid #b6c2cc; border-bottom:1px solid #000; }}
  td {{ padding:2px 8px; border:1px solid #e0e0e0; }}
  tr:hover td {{ background:#fffde7 !important; }}
</style></head><body>
<div class="ts">Snapshot: {now}</div>
<table><thead><tr>{headers}</tr></thead><tbody>{sum_rows}</tbody></table>
{det_body}
</body></html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    log.info(f"saved → {out_path}")
    if open_browser:
        webbrowser.open("file:///" + out_path.replace("\\", "/"))


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()
    generate_html()
