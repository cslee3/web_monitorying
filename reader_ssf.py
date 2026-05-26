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

import datetime, threading, json, os, time
from queue import Queue
import pythoncom
import win32com.client
import pywintypes

# COM 바쁨 에러 코드 (DDE/VBA 실행 중 등)
_COM_BUSY = {-2147418111, -2147417846}   # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER

from common import BASE_DIR, BOOK_NAME, _pct, _num, _date, _colorscale_bg, _is_xl_error

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

# ── 섹션 정의 — B2:AK44 단일 bulk read 기준 (COM 호출 통합) ──────────
# row_s/row_e : B2 기준 0-based 행 오프셋 (row2=0, row3=1, row5=3, ...)
# col_s/col_e : B  기준 0-based 열 오프셋 (B=0 … AA=25 … AE=29 … AK=35)

# MM 섹션 (B:AA = cols 0~25)
MM_SECTIONS = [
    {"id": "main", "row_s":  3, "row_e": 34},  # B5:AA36
    {"id": "sum1", "row_s": 36, "row_e": 38},  # B38:AA40
    {"id": "sum2", "row_s": 40, "row_e": 42},  # B42:AA44
]

# IDX2 섹션 (has_header: True면 첫 번째 행을 헤더로 처리)
IDX2_SECTIONS = [
    {"id": "idx1", "row_s":  3, "row_e":  9, "col_s": 29, "col_e": 35, "has_header": True},   # AE5:AK11
    {"id": "idx2", "row_s": 11, "row_e": 11, "col_s": 29, "col_e": 30, "has_header": False},  # AE13:AF13
    {"id": "idx3", "row_s": 12, "row_e": 13, "col_s": 29, "col_e": 35, "has_header": False,
     "header": ["IdxCode", "IndexName", "Fund", "Delta", "Volume", "MTM PnL", "Theo PnL"]},  # AE14:AK15
    {"id": "idx4", "row_s": 15, "row_e": 16, "col_s": 29, "col_e": 35, "has_header": False,
     "header": ["IdxCode", "IndexName", "Fund", "Delta", "Volume", "MTM PnL", "Theo PnL"]},  # AE17:AK18
    {"id": "idx5", "row_s": 18, "row_e": 23, "col_s": 29, "col_e": 35, "has_header": True},   # AE20:AK25
    {"id": "idx6", "row_s": 27, "row_e": 36, "col_s": 29, "col_e": 35, "has_header": True},   # AE29:AK38
    {"id": "idx7", "row_s": 41, "row_e": 42, "col_s": 32, "col_e": 34, "has_header": False},  # AH43:AJ44
]

TOTAL_VALS = {"합계", "합 계", "Total", "소계"}

# DB 저장 허용 시간대 (HH, MM)
DB_SAVE_START = (9, 5)
DB_SAVE_END   = (15, 20)

# ColorScale 그라디언트 적용 컬럼 (cidx 기준)
# F=4(Amt Mil), Q=15(Delta Shr), R=16(Delta Mil), AA=25(Theo PnL)
COLORSCALE_CIDX = {4, 15, 16, 25}


def _raw(v):
    """Excel raw 값 → 문자열. 그대로 표시용.
    #BLOCKED! / #VALUE! 등 Excel 에러값은 빈 문자열로 처리.
    """
    if v is None or v == "": return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%H:%M:%S") if v.date() == datetime.date.today() else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if _is_xl_error(v): return ""
        return str(int(round(v)))
    s = str(v)
    if s.startswith("#"): return ""   # #BLOCKED!, #VALUE!, #REF!, #N/A 등
    return s


def _read_mm(bulk):
    """B2:AK44 bulk 데이터 → MM_SECTIONS 슬라이싱 → 섹션 리스트 반환."""
    cols = [(c[0], c[1], MM_COL_SUBS.get(c[0], "")) for c in MM_COLS]
    sections = []
    for sec in MM_SECTIONS:
        rows, meta = _read_section(bulk[sec["row_s"]: sec["row_e"] + 1], MM_COLS, MM_PNL_CIDX)
        sections.append({
            "id":       sec["id"],
            "cols":     cols,
            "num_cidx": list(MM_NUM_CIDX),
            "rows":     rows,
            "row_meta": meta,
        })
    return sections


def _read_idx2(bulk):
    """B2:AK44 bulk 데이터 → IDX2_SECTIONS 슬라이싱 → raw 2D 배열 반환."""
    sections = []
    ncols = len(bulk[0]) if bulk else 0
    for sec in IDX2_SECTIONS:
        rows = []
        for row_data in bulk[sec["row_s"]: sec["row_e"] + 1]:
            if row_data is None:
                row_data = [None] * ncols
            row = [_raw(row_data[ci]) if ci < len(row_data) else ""
                   for ci in range(sec["col_s"], sec["col_e"] + 1)]
            rows.append(row)
        sections.append({
            "id":         sec["id"],
            "has_header": sec["has_header"],
            "header":     sec.get("header"),
            "raw":        True,
            "rows":       rows,
        })
    return sections


def _pnl_color(cidx, raw, pnl_cidx):
    if cidx not in pnl_cidx or raw is None: return ""
    try: return "green" if float(raw) > 0 else ("red" if float(raw) < 0 else "")
    except: return ""


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
        self.db_save_enabled = False
        self._db_conn        = None

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

    def toggle_db_save(self):
        self.db_save_enabled = not self.db_save_enabled
        return self.db_save_enabled

    def _get_db_conn(self):
        import mysql.connector
        from db_manager import DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME
        try:
            if self._db_conn and self._db_conn.is_connected():
                return self._db_conn
        except Exception:
            pass
        self._db_conn = mysql.connector.connect(
            host=DB_HOST, port=DB_PORT,
            user=DB_USER, password=DB_PASSWORD,
            database=DB_NAME, charset="utf8mb4",
        )
        return self._db_conn

    def _save_to_db(self, data, ts):
        def _parse(v):
            if v is None or v == "": return None
            try: return float(str(v).replace(",", "").replace("%", "").strip())
            except: return None

        main_sec = next((s for s in data["sections"] if s["id"] == "main"), None)
        if main_sec is None:
            return

        rows_to_insert = []
        for row in main_sec["rows"]:
            stock = row.get("0", {}).get("v", "")
            if not stock or stock in TOTAL_VALS:
                continue
            rows_to_insert.append((
                ts,
                stock,
                row.get("1",  {}).get("v") or None,
                _parse(row.get("2",  {}).get("v")),
                _parse(row.get("3",  {}).get("v")),
                _parse(row.get("4",  {}).get("v")),
                _parse(row.get("5",  {}).get("v")),
                _parse(row.get("6",  {}).get("v")),
                _parse(row.get("8",  {}).get("v")),
                _parse(row.get("9",  {}).get("v")),
                row.get("10", {}).get("v") or None,
                _parse(row.get("11", {}).get("v")),
                _parse(row.get("12", {}).get("v")),
                _parse(row.get("13", {}).get("v")),
                _parse(row.get("14", {}).get("v")),
                _parse(row.get("15", {}).get("v")),
                _parse(row.get("16", {}).get("v")),
                _parse(row.get("17", {}).get("v")),
                _parse(row.get("18", {}).get("v")),
                _parse(row.get("19", {}).get("v")),
                _parse(row.get("20", {}).get("v")),
                _parse(row.get("21", {}).get("v")),
                _parse(row.get("22", {}).get("v")),
                _parse(row.get("23", {}).get("v")),
                _parse(row.get("24", {}).get("v")),
                _parse(row.get("25", {}).get("v")),
            ))

        if not rows_to_insert:
            return
        try:
            conn = self._get_db_conn()
            cur  = conn.cursor()
            cur.executemany("""
                INSERT IGNORE INTO ssf_history
                (ts, stock, stock_name, last_price, change_pct, amt_mil, amt_shr,
                 mmqty, ask, bid, mm_spread, shares, diff, vol_lots, vol_mil,
                 delta_shr, delta_mil, theo_price, theo_basis,
                 ex1_bsp, ex1_ssp, ex2_basis, ex2_bsp, ex2_ssp, mtm_pnl, theo_pnl)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, rows_to_insert)
            conn.commit()
            cur.close()
            print(f"[ssf_monitor] db saved {len(rows_to_insert)} rows  ts={ts:.3f}")
        except Exception as e:
            print(f"[ssf_monitor] db save error: {e}")
            self._db_conn = None

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
        # B2:AK44 단일 COM 호출 — 헤더·MM·IDX 전 영역 커버
        bulk = ws.Range("B2:AK44").Value

        header = {
            "base_date": _date(bulk[0][1]),   # C2
            "ytd":       _date(bulk[1][1]),   # C3
            "ir":        _raw(bulk[0][4]),    # F2
            "exp1":      _date(bulk[0][7]),   # I2
            "exp2":      _date(bulk[1][7]),   # I3
            "u2": _raw(bulk[0][19]), "v2": _raw(bulk[0][20]),  # U2, V2
            "u3": _raw(bulk[1][19]), "v3": _raw(bulk[1][20]),  # U3, V3
        }

        mm_sections  = _read_mm(bulk)
        idx_sections = _read_idx2(bulk)

        result = {
            "header": header,
            "sections": mm_sections + idx_sections,
            "updated": datetime.datetime.now().strftime("%H:%M:%S"),
        }
        return result

    def run(self):
        pythoncom.CoInitialize()
        prev_data = None
        ws        = None

        while True:
            try:
                if ws is None:
                    ws = self._get_ws()
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
                    if sec_new.get("raw"):
                        for ri, (new_row, old_row) in enumerate(zip(sec_new["rows"], sec_old["rows"])):
                            for ci, (nv, ov) in enumerate(zip(new_row, old_row)):
                                if nv != ov:
                                    changes.append({"key": f"{sid}_{ri}_{ci}", "v": nv, "color": "", "bg": ""})
                    else:
                        for ri, (new_row, old_row) in enumerate(zip(sec_new["rows"], sec_old["rows"])):
                            for cidx, *_ in sec_new["cols"]:
                                new_cell = new_row.get(str(cidx), {})
                                old_cell = old_row.get(str(cidx), {})
                                new_v  = new_cell.get("v", "")
                                old_v  = old_cell.get("v", "")
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

                if self.db_save_enabled:
                    now = datetime.datetime.now()
                    in_window = DB_SAVE_START <= (now.hour, now.minute) <= DB_SAVE_END
                    if in_window:
                        self._save_to_db(data, time.time())

                prev_data = data

            except pywintypes.com_error as e:
                hresult = e.args[0] if e.args else 0
                if hresult in _COM_BUSY:
                    # DDE/VBA 실행 중 — ws 유지하고 잠시 대기
                    time.sleep(0.5)
                    continue
                print(f"[ssf_monitor] com_error: {e}")
                prev_data = None
                ws        = None
            except Exception as e:
                print(f"[ssf_monitor] error: {e}")
                prev_data = None
                ws        = None

            time.sleep(1.0)


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
    mm_sections = [s for s in data["sections"] if not s.get("raw")]
    cols        = mm_sections[0]["cols"]
    num_set     = set(mm_sections[0]["num_cidx"])
    headers     = "".join(f"<th>{hdr}</th>" for _, hdr, *_ in cols)
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
    """IDX2 섹션 스냅샷 HTML 생성."""
    import webbrowser
    pythoncom.CoInitialize()

    if out_path is None:
        out_path = os.path.join(BASE_DIR, "ssf_idx_snapshot.html")

    ws       = monitor._get_ws()
    bulk     = ws.Range("B2:AK44").Value
    sections = _read_idx2(bulk)
    now      = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    body = ""
    for sec in sections:
        rows = sec["rows"]
        if not rows:
            continue
        start_row  = 0
        thead_html = ""
        if sec["has_header"]:
            thead_html = "<thead><tr>" + "".join(f"<th>{v}</th>" for v in rows[0]) + "</tr></thead>"
            start_row  = 1
        tbody = "".join(
            "<tr>" + "".join(f"<td>{v}</td>" for v in rows[ri]) + "</tr>"
            for ri in range(start_row, len(rows))
        )
        body += f'<div style="margin-bottom:10px"><table>{thead_html}<tbody>{tbody}</tbody></table></div>'

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
{body}
</body></html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[reader_ssf] saved → {out_path}")
    webbrowser.open("file:///" + out_path.replace("\\", "/"))


if __name__ == "__main__":
    generate_html()
    generate_html_idx()
