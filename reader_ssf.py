# -*- coding: utf-8 -*-
"""
reader_ssf.py — DashBoard 탭 실시간 모니터 (SSF MM / SSF MM2)
- DashboardMonitor.run() 을 별도 스레드에서 실행
- 0.3초마다 4개 범위를 bulk read하여 이전 값과 비교
- 변경된 셀만 SSE 클라이언트에 push
- serve.py에서 import하여 사용

섹션 구성:
  main : B5:AA36  — SSF MM 개별 종목
  sum1 : B38:AA40 — SSF MM 소계/합계 (1단)
  sum2 : B42:AA43 — SSF MM 합계 (2단)
  idx  : AC5:AI36 — SSF MM2 지수 정보
"""

import datetime, threading, json, os
from queue import Queue
import pythoncom
import win32com.client

from reader_common import _pct, _num, _date

try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

BOOK_NAME = "GlobalMM_Realtime_Monitoring_V2_20260520.xlsb"

# ── 컬럼 정의 ──────────────────────────────────────────────────────────────
# MM sections: B5:AA36, B38:AA40, B42:AA43
# cidx = 0-based from B  (B=0, C=1, ..., H=6, I=7 skip, J=8, ..., AA=25)
MM_COLS = [
    (0,  "Stock",         lambda v: str(v) if v else ""),
    (1,  "StockName",     lambda v: str(v) if v else ""),
    (2,  "LastPrice",     lambda v: _num(v, 0)),
    (3,  "Change(%)",     _pct),
    (4,  "Amt(Mil)",      lambda v: _num(v, 0)),
    (5,  "Amt(Shr)",      lambda v: _num(v, 0)),
    (6,  "MMQty",         lambda v: _num(v, 0)),
    # cidx 7 = I열 (skip)
    (8,  "Ask",           lambda v: _num(v, 0)),
    (9,  "Bid",           lambda v: _num(v, 0)),
    (10, "MM Spread",     lambda v: str(v) if v else ""),
    (11, "Shares",        lambda v: _num(v, 0)),
    (12, "Diff",          lambda v: _num(v, 0)),
    (13, "Vol(Lots)",     lambda v: _num(v, 0)),
    (14, "Vol(Mil)",      lambda v: _num(v, 0)),
    (15, "Delta(Shr)",    lambda v: _num(v, 0)),
    (16, "Delta(Mil)",    lambda v: _num(v, 0)),
    (17, "Theo Price",    lambda v: _num(v, 0)),
    (18, "Theo Basis",    lambda v: _num(v, 0)),
    (19, "Ex1 B-Sp",      lambda v: _num(v, 0)),
    (20, "Ex1 S-Sp",      lambda v: _num(v, 0)),
    (21, "Ex2 Basis",     lambda v: _num(v, 0)),
    (22, "Ex2 B-Sp",      lambda v: _num(v, 0)),
    (23, "Ex2 S-Sp",      lambda v: _num(v, 0)),
    (24, "MTM PnL",       lambda v: _num(v, 0)),
    (25, "Theo PnL",      lambda v: _num(v, 0)),
]
MM_PNL_CIDX = {c[0] for c in MM_COLS if "PnL" in c[1]} - {24, 25}  # MTM PnL, Theo PnL 글자색 제외
MM_NUM_CIDX  = {c[0] for c in MM_COLS if c[0] not in (0, 1, 10)}

# 컬럼 헤더 두 번째 줄 (단축명과 다른 경우만)
MM_COL_SUBS = {
    3:  "Change (%)",
    4:  "Amount (Mil KRW)",
    5:  "Amount (Shares)",
    11: "△Shares",
    13: "Volume (Lots)",
    14: "Volume (Mkrw)",
    15: "Delta (Shares)",
    16: "Delta (Mil KRW)",
    19: "Ex1 B-Spread",
    20: "Ex1 S-Spread",
    21: "Theo Basis",
    22: "Ex2 B-Spread",
    23: "Ex2 S-Spread",
}

# IDX section: AC5:AI36
# cidx = 0-based from AC  (AC=0, AD=1, AE=2, AF=3, AG=4, AH=5, AI=6)
IDX_COLS = [
    (0, "IdxCode",       lambda v: str(v) if v else ""),
    (1, "IdxName",       lambda v: str(v) if v else ""),
    (2, "Fund",          lambda v: _num(v, 0)),
    (3, "Delta",         lambda v: _num(v, 0)),
    (4, "Volume",        lambda v: _num(v, 0)),
    (5, "MTM PnL(Idx)",  lambda v: _num(v, 0)),
    (6, "Theo PnL(Idx)", lambda v: _num(v, 0)),
]
IDX_PNL_CIDX = {c[0] for c in IDX_COLS if "PnL" in c[1]}
IDX_NUM_CIDX  = {c[0] for c in IDX_COLS if c[0] not in (0, 1)}

TOTAL_VALS = {"합계", "합 계", "Total", "소계"}

# ColorScale 그라디언트 적용 컬럼 (cidx 기준)
# F=4(Amt Mil), Q=15(Delta Shr), R=16(Delta Mil), AA=25(Theo PnL)
COLORSCALE_CIDX = {4, 15, 16, 25}


def _pnl_color(cidx, raw, pnl_cidx):
    if cidx not in pnl_cidx or raw is None: return ""
    try: return "green" if float(raw) > 0 else ("red" if float(raw) < 0 else "")
    except: return ""


def _colorscale_bg(val, vmin, vmax):
    """
    val의 min/max 대비 상대 위치로 파랑→흰→빨강 배경색 계산.
    음수 구간: blue(#6fa8dc) → white(#ffffff)
    양수 구간: white(#ffffff) → red(#e06666)
    """
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""

    if vmin is None or vmax is None or vmin == vmax:
        return ""

    if v < 0:
        # 음수: vmin~0 구간에서 파랑→흰 보간
        if vmin >= 0:
            return ""
        t = max(0.0, min(1.0, v / vmin))  # vmin일 때 t=1(파랑), 0일 때 t=0(흰)
        r = int(255 - t * (255 - 111))
        g = int(255 - t * (255 - 168))
        b = int(255 - t * (255 - 220))
        return f"rgb({r},{g},{b})"
    elif v > 0:
        # 양수: 0~vmax 구간에서 흰→빨강 보간
        if vmax <= 0:
            return ""
        t = max(0.0, min(1.0, v / vmax))  # 0일 때 t=0(흰), vmax일 때 t=1(빨강)
        r = int(255 - t * (255 - 224))
        g = int(255 - t * (255 - 102))
        b = int(255 - t * (255 - 102))
        return f"rgb({r},{g},{b})"
    else:
        return ""


def _read_section(all_vals, cols, pnl_cidx):
    """raw Excel 2D 튜플을 섹션 rows / row_meta 로 변환."""
    rows, row_meta = [], []
    ncols = len(all_vals[0]) if all_vals else 0

    first_cidx = cols[0][0]
    first_hdr  = cols[0][1]

    # ColorScale 컬럼별 min/max 사전 계산
    cs_vals = {cidx: [] for cidx in COLORSCALE_CIDX}
    for row_data in all_vals:
        if row_data is None:
            continue
        for cidx in COLORSCALE_CIDX:
            if cidx < len(row_data):
                v = row_data[cidx]
                try:
                    cs_vals[cidx].append(float(v))
                except (TypeError, ValueError):
                    pass

    cs_min = {}
    cs_max = {}
    for cidx, vals in cs_vals.items():
        neg = [v for v in vals if v < 0]
        pos = [v for v in vals if v > 0]
        cs_min[cidx] = min(neg) if neg else None
        cs_max[cidx] = max(pos) if pos else None

    for row_data in all_vals:
        if row_data is None:
            row_data = [None] * ncols

        # 헤더 행 건너뜀
        if str(row_data[first_cidx] or "").strip() == first_hdr:
            row_meta.append({"sep": True, "total": False})
            rows.append({})
            continue

        cells = {}
        any_val = False
        for cidx, _, fmt in cols:
            raw  = row_data[cidx]
            fval = fmt(raw) if (raw is not None and raw != "") else ""
            if fval: any_val = True

            # ColorScale 배경색 계산
            bg = ""
            if cidx in COLORSCALE_CIDX:
                bg = _colorscale_bg(raw, cs_min.get(cidx), cs_max.get(cidx))

            cells[str(cidx)] = {
                "v":     fval,
                "color": _pnl_color(cidx, raw, pnl_cidx),
                "bg":    bg,
            }

        if not any_val:
            row_meta.append({"sep": True, "total": False})
            rows.append({})
            continue

        is_total = (
            cells.get("0", {}).get("v", "") in TOTAL_VALS or
            cells.get("1", {}).get("v", "") in TOTAL_VALS
        )
        row_meta.append({"sep": False, "total": is_total})
        rows.append(cells)

    return rows, row_meta


class DashboardMonitor:
    """
    Excel DashBoard 탭을 주기적으로 읽어 변경분을 SSE 클라이언트에 push하는 모니터.
    - _full_data : 최신 전체 섹션 데이터 (신규 브라우저 접속 시 제공)
    - _clients   : 연결된 SSE 클라이언트 큐 목록
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
        path = os.path.join(BASE_DIR, BOOK_NAME)
        try:
            com_wb = win32com.client.GetObject(path)
            return com_wb.Sheets("DashBoard")
        except Exception:
            pass
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        com_wb = xl.Workbooks.Open(path)
        return com_wb.Sheets("DashBoard")

    def _read_full(self, ws):
        header = {
            "base_date": _date(ws.Cells(2, 3).Value),
            "ytd":       _date(ws.Cells(3, 3).Value),
            "ir":        str(ws.Cells(2, 6).Value),
            "exp1":      _date(ws.Cells(2, 9).Value),
            "exp2":      _date(ws.Cells(3, 9).Value),
        }

        main_rows, main_meta = _read_section(ws.Range("B5:AA36").Value,  MM_COLS,  MM_PNL_CIDX)
        sum1_rows, sum1_meta = _read_section(ws.Range("B38:AA40").Value, MM_COLS,  MM_PNL_CIDX)
        sum2_rows, sum2_meta = _read_section(ws.Range("B42:AA44").Value, MM_COLS,  MM_PNL_CIDX)
        idx_rows,  idx_meta  = _read_section(ws.Range("AC5:AI36").Value, IDX_COLS, IDX_PNL_CIDX)

        result = {
            "header": header,
            "sections": [
                {"id": "main", "cols": [(c[0], c[1], MM_COL_SUBS.get(c[0], "")) for c in MM_COLS],  "num_cidx": list(MM_NUM_CIDX),  "rows": main_rows, "row_meta": main_meta},
                {"id": "sum1", "cols": [(c[0], c[1], MM_COL_SUBS.get(c[0], "")) for c in MM_COLS],  "num_cidx": list(MM_NUM_CIDX),  "rows": sum1_rows, "row_meta": sum1_meta},
                {"id": "sum2", "cols": [(c[0], c[1], MM_COL_SUBS.get(c[0], "")) for c in MM_COLS],  "num_cidx": list(MM_NUM_CIDX),  "rows": sum2_rows, "row_meta": sum2_meta},
                {"id": "idx",  "cols": [(c[0], c[1], "") for c in IDX_COLS],              "num_cidx": list(IDX_NUM_CIDX), "rows": idx_rows,  "row_meta": idx_meta},
            ],
            "updated": datetime.datetime.now().strftime("%H:%M:%S"),
        }
        return result

    def run(self):
        pythoncom.CoInitialize()
        prev_data = None

        while True:
            try:
                ws      = self._get_ws()
                data    = self._read_full(ws)
                now_str = datetime.datetime.now().strftime("%H:%M:%S")

                if prev_data is None:
                    with self._full_lock:
                        self._full_data = data
                    prev_data = data
                    print("[ssf_monitor] initial load done")
                    import time; time.sleep(1.0)
                    continue

                changes = []
                for sec_new, sec_old in zip(data["sections"], prev_data["sections"]):
                    sid = sec_new["id"]
                    for ri, (new_row, old_row) in enumerate(zip(sec_new["rows"], sec_old["rows"])):
                        for cidx, *_ in sec_new["cols"]:
                            new_cell = new_row.get(str(cidx), {})
                            old_cell = old_row.get(str(cidx), {})
                            new_v = new_cell.get("v", "")
                            old_v = old_cell.get("v", "")
                            new_bg = new_cell.get("bg", "")
                            old_bg = old_cell.get("bg", "")
                            if new_v == old_v and new_bg == old_bg: continue
                            changes.append({
                                "key":   f"{sid}_{ri}_{cidx}",
                                "v":     new_v,
                                "color": new_cell.get("color", ""),
                                "bg":    new_bg,
                            })

                data["updated"] = now_str
                with self._full_lock:
                    self._full_data = data
                self._push(json.dumps({"cells": changes, "updated": now_str}))
                if changes:
                    print(f"[ssf_monitor] {now_str} — pushed {len(changes)} changes")

                prev_data = data

            except Exception as e:
                print(f"[ssf_monitor] error: {e}")
                prev_data = None

            import time; time.sleep(1.0)


monitor = DashboardMonitor()


# ── 단독 실행: 스냅샷 HTML 생성 ──────────────────────────────────────────

def generate_html(out_path=None):
    """MM 섹션(main/sum1/sum2) 스냅샷 HTML 생성."""
    import pythoncom, webbrowser
    pythoncom.CoInitialize()

    if out_path is None:
        out_path = os.path.join(BASE_DIR, "ssf_snapshot.html")

    ws   = monitor._get_ws()
    data = monitor._read_full(ws)

    header      = data["header"]
    mm_sections = [s for s in data["sections"] if s["id"] != "idx"]
    cols        = mm_sections[0]["cols"]
    num_set     = set(mm_sections[0]["num_cidx"])
    headers     = "".join(f"<th>{hdr}</th>" for _, hdr in cols)
    now         = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tbody = ""
    for si, sec in enumerate(mm_sections):
        for row, meta in zip(sec["rows"], sec["row_meta"]):
            if meta["sep"]:
                tbody += f'<tr><td colspan="{len(cols)}" style="height:5px;background:#f0f0f0"></td></tr>'
                continue
            tds = ""
            for cidx, *_ in cols:
                cell  = row.get(str(cidx), {"v": "", "color": ""})
                align = "right" if cidx in num_set else "left"
                cs    = "color:#1a7c34;" if cell["color"] == "green" else ("color:#cc0000;" if cell["color"] == "red" else "")
                bold  = "font-weight:bold;" if meta.get("total") else ""
                tds  += f'<td style="{cs}{bold}text-align:{align}">{cell["v"]}</td>'
            tbody += f"<tr>{tds}</tr>"
        if si < len(mm_sections) - 1:
            tbody += f'<tr><td colspan="{len(cols)}" style="height:5px;background:#f0f0f0"></td></tr>'

    html = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8"/>
<title>SSF MM Snapshot</title>
<style>
  body {{ margin:0; padding:8px; font-family:'Segoe UI',sans-serif; font-size:12px; }}
  .info {{ display:flex; gap:20px; padding:5px 10px; background:#eef2f6; margin-bottom:6px; font-size:11px; }}
  .info span {{ font-weight:bold; }}
  .ts {{ margin-left:auto; color:#888; }}
  table {{ border-collapse:collapse; white-space:nowrap; }}
  th {{ background:#202123; color:#fff; padding:4px 7px; font-size:11px; text-align:center; border:1px solid #444; }}
  td {{ padding:2px 6px; border:1px solid #e0e0e0; }}
  tr:hover td {{ background:#fffde7 !important; }}
</style></head><body>
<div class="info">
  <div>BaseDate: <span>{header["base_date"]}</span></div>
  <div>Ytd: <span>{header["ytd"]}</span></div>
  <div>I/R: <span>{header["ir"]}</span></div>
  <div>Expiry1: <span>{header["exp1"]}</span></div>
  <div>Expiry2: <span>{header["exp2"]}</span></div>
  <div class="ts">Snapshot: {now}</div>
</div>
<table><thead><tr>{headers}</tr></thead><tbody>{tbody}</tbody></table>
</body></html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[reader_ssf] saved → {out_path}")
    webbrowser.open("file:///" + out_path.replace("\\", "/"))


def generate_html_idx(out_path=None):
    """IDX 섹션(AC5:AI36) 스냅샷 HTML 생성."""
    import pythoncom, webbrowser
    pythoncom.CoInitialize()

    if out_path is None:
        out_path = os.path.join(BASE_DIR, "ssf_idx_snapshot.html")

    ws   = monitor._get_ws()
    data = monitor._read_full(ws)

    idx_sec = next(s for s in data["sections"] if s["id"] == "idx")
    cols    = idx_sec["cols"]
    num_set = set(idx_sec["num_cidx"])
    headers = "".join(f"<th>{hdr}</th>" for _, hdr in cols)
    now     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tbody = ""
    for row, meta in zip(idx_sec["rows"], idx_sec["row_meta"]):
        if meta["sep"]:
            tbody += f'<tr><td colspan="{len(cols)}" style="height:5px;background:#f0f0f0"></td></tr>'
            continue
        tds = ""
        for cidx, *_ in cols:
            cell  = row.get(str(cidx), {"v": "", "color": ""})
            align = "right" if cidx in num_set else "left"
            cs    = "color:#1a7c34;" if cell["color"] == "green" else ("color:#cc0000;" if cell["color"] == "red" else "")
            tds  += f'<td style="{cs}text-align:{align}">{cell["v"]}</td>'
        tbody += f"<tr>{tds}</tr>"

    html = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8"/>
<title>SSF MM2 Snapshot</title>
<style>
  body {{ margin:0; padding:8px; font-family:'Segoe UI',sans-serif; font-size:12px; }}
  .ts {{ text-align:right; color:#888; font-size:11px; margin-bottom:6px; }}
  table {{ border-collapse:collapse; white-space:nowrap; }}
  th {{ background:#202123; color:#fff; padding:4px 7px; font-size:11px; text-align:center; border:1px solid #444; }}
  td {{ padding:2px 6px; border:1px solid #e0e0e0; }}
  tr:hover td {{ background:#fffde7 !important; }}
</style></head><body>
<div class="ts">Snapshot: {now}</div>
<table><thead><tr>{headers}</tr></thead><tbody>{tbody}</tbody></table>
</body></html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[reader_ssf] saved → {out_path}")
    webbrowser.open("file:///" + out_path.replace("\\", "/"))


if __name__ == "__main__":
    generate_html()
    generate_html_idx()