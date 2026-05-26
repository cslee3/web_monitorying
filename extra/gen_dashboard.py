# -*- coding: utf-8 -*-
"""
gen_dashboard.py — SSF MM 대시보드 HTML 정적 생성기 (구버전)
DashBoard 탭을 읽어 ssfo_pnl.html로 저장.
현재는 serve.py + dashboard_reader.py의 실시간 SSE 방식으로 대체됨.
단독 실행 시 Excel을 별도 App으로 열고 종료 (새 인스턴스 생성).
"""

import os
import datetime
import xlwings as xw

# 대상 Excel 파일 경로
EXCEL_PATH = os.path.join(os.path.dirname(__file__),
                          "GlobalMM_Realtime_Monitoring_V2_20260520.xlsb")
# 생성될 HTML 파일 경로
OUT_PATH   = os.path.join(os.path.dirname(__file__), "ssfo_pnl.html")


# ── 값 포매터 ────────────────────────────────────────────────────────────

def pct(v):
    """소수 → 퍼센트 문자열. ex) 0.0123 → '1.23%'"""
    if v is None: return ""
    return f"{v*100:.2f}%"

def num(v, d=0):
    """숫자 → 천단위 콤마 + 소수점 d자리 문자열."""
    if v is None or v == "": return ""
    try:
        return f"{v:,.{d}f}"
    except Exception:
        return str(v)

def date_fmt(v):
    """datetime → 'YYYY-MM-DD' 문자열."""
    if v is None: return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    return str(v)

def to_rgb(color):
    """xlwings color tuple (r,g,b) → CSS rgb() 문자열."""
    if color is None:
        return ""
    if isinstance(color, tuple) and len(color) == 3:
        return f"rgb({color[0]},{color[1]},{color[2]})"
    return ""


# ── 컬럼 정의 (Excel 1-based 열 번호, 헤더, 포매터) ─────────────────────
SSF_COLS = [
    (1,  "ID",              str),
    (2,  "Stock",           str),
    (3,  "StockName",       str),
    (4,  "LastPrice",       lambda v: num(v, 0)),
    (5,  "Change(%)",       pct),
    (6,  "Amt(Mil)",        lambda v: num(v, 1)),
    (7,  "Amt(Shr)",        lambda v: num(v, 0)),
    (8,  "MMQty",           lambda v: num(v, 0)),
    (10, "Ask",             lambda v: num(v, 0)),
    (11, "Bid",             lambda v: num(v, 0)),
    (12, "MM Spread",       str),
    (13, "Shares",          lambda v: num(v, 0)),
    (14, "Diff",            lambda v: num(v, 0)),
    (15, "Vol(Lots)",       lambda v: num(v, 0)),
    (16, "Vol(Mil)",        lambda v: num(v, 2)),
    (17, "Delta(Shr)",      lambda v: num(v, 0)),
    (18, "Delta(Mil)",      lambda v: num(v, 2)),
    (19, "Theo Price",      lambda v: num(v, 0)),
    (20, "Theo Basis",      lambda v: num(v, 2)),
    (21, "Ex1 B-Sp",        lambda v: num(v, 2)),
    (22, "Ex1 S-Sp",        lambda v: num(v, 2)),
    (23, "Ex2 Basis",       lambda v: num(v, 2)),
    (24, "Ex2 B-Sp",        lambda v: num(v, 2)),
    (25, "Ex2 S-Sp",        lambda v: num(v, 2)),
    (26, "MTM PnL",         lambda v: num(v, 0)),
    (27, "Theo PnL",        lambda v: num(v, 0)),
    (29, "IdxCode",         str),
    (30, "IdxName",         str),
    (31, "Fund",            lambda v: num(v, 0)),
    (32, "Delta",           lambda v: num(v, 0)),
    (33, "Volume",          lambda v: num(v, 0)),
    (34, "MTM PnL(Idx)",    lambda v: num(v, 0)),
    (35, "Theo PnL(Idx)",   lambda v: num(v, 0)),
]

# PnL 컬럼 인덱스 집합 (녹색/적색 텍스트 적용 대상)
PNL_COLS = {c[0] for c in SSF_COLS if "PnL" in c[1]}


def cell_style(bg, bold, val, col_idx):
    """
    셀 인라인 CSS 스타일 문자열 생성.
    - bg: 배경색
    - bold: 굵은 글씨 여부
    - PnL 컬럼은 값 부호에 따라 녹색/적색 텍스트 적용
    """
    styles = []
    if bg:
        styles.append(f"background:{bg}")
    if bold:
        styles.append("font-weight:bold")
    if col_idx in PNL_COLS:
        try:
            fval = float(str(val).replace(",", "").replace("%", ""))
            if fval > 0:
                styles.append("color:#1a7c34")
            elif fval < 0:
                styles.append("color:#cc0000")
        except Exception:
            pass
    return "; ".join(styles)


def generate():
    """
    Excel을 새 App으로 열어 DashBoard 탭(6~45행)을 읽고 ssfo_pnl.html 생성.
    - 헤더 정보(기준일, YTD, I/R, 만기일) 상단 info-bar에 표시
    - 빈 행은 구분선(높이 6px 회색 행)으로 처리
    - 합계 행(col2='합계' 등)은 total-row 클래스로 녹색 배경 처리
    - 실행 후 Excel App 종료 (새 인스턴스였으므로)
    """
    app = xw.App(visible=False)
    try:
        wb = app.books.open(EXCEL_PATH)
        ws = wb.sheets["DashBoard"]

        # 헤더 정보 읽기 (개별 셀 COM 호출)
        base_date = date_fmt(ws.cells(2, 3).value)
        ytd_date  = date_fmt(ws.cells(3, 3).value)
        ir        = ws.cells(2, 6).value
        exp1      = date_fmt(ws.cells(2, 9).value)
        exp2      = date_fmt(ws.cells(3, 9).value)

        # 데이터 행 읽기 (6~45행, 각 셀 개별 COM 호출)
        rows_html = []
        for r in range(6, 46):
            cells = {}
            any_val = False
            for col_def in SSF_COLS:
                cidx = col_def[0]
                cell = ws.cells(r, cidx)
                v    = cell.value
                bg   = to_rgb(cell.color)
                bold = cell.font.bold
                cells[cidx] = (v, bg, bold)
                if v is not None and v != "":
                    any_val = True

            if not any_val:
                rows_html.append("")   # 빈 행 → 구분선용 빈 문자열
                continue

            # col2 값으로 합계 행 판별
            is_total = (cells[2][0] in ("합계", "합 계", "Total", "소계"))

            tds = []
            for col_def in SSF_COLS:
                cidx, hdr, fmt = col_def
                v, bg, bold = cells[cidx]
                formatted = ""
                if v is not None and v != "":
                    try:
                        formatted = fmt(v)
                    except Exception:
                        formatted = str(v)
                style = cell_style(bg, bold, formatted, cidx)
                align = "right" if cidx not in (1, 2, 3, 12, 29, 30) else "left"
                tds.append(f'<td style="{style}; text-align:{align}">{formatted}</td>')

            row_class = "total-row" if is_total else ""
            rows_html.append(f'<tr class="{row_class}">{"".join(tds)}</tr>')

        wb.close()
    finally:
        app.quit()  # 새로 열었던 Excel 인스턴스 종료

    # HTML 조립
    headers = "".join(f"<th>{c[1]}</th>" for c in SSF_COLS)
    body_rows = []
    for rh in rows_html:
        if rh == "":
            # 빈 행 → 회색 구분선 행으로 변환
            body_rows.append(f'<tr><td colspan="{len(SSF_COLS)}" style="height:6px;background:#f0f0f0"></td></tr>')
        else:
            body_rows.append(rh)

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8"/>
<meta http-equiv="refresh" content="10"/>
<title>SSF MM Dashboard</title>
<style>
  body {{ margin:0; padding:8px; font-family:'Segoe UI',sans-serif; font-size:12px; background:#fff; }}
  .info-bar {{ display:flex; gap:24px; margin-bottom:8px; padding:6px 10px;
               background:#eef2f6; border-radius:4px; font-size:12px; }}
  .info-bar span {{ font-weight:bold; }}
  .refresh {{ margin-left:auto; color:#888; font-size:11px; }}
  table {{ border-collapse:collapse; width:100%; white-space:nowrap; }}
  th {{ background:#202123; color:#fff; padding:4px 6px; position:sticky; top:0; z-index:1;
        font-size:11px; text-align:center; border:1px solid #444; }}
  td {{ padding:3px 6px; border:1px solid #ddd; font-size:12px; }}
  tr:hover td {{ background:#fffbe6 !important; }}
  .total-row td {{ background:#e2efda !important; font-weight:bold; }}
  .table-wrap {{ overflow-x:auto; overflow-y:auto; max-height:calc(100vh - 80px); }}
</style>
</head>
<body>
<div class="info-bar">
  <div>BaseDate: <span>{base_date}</span></div>
  <div>Ytd: <span>{ytd_date}</span></div>
  <div>I/R: <span>{ir}</span></div>
  <div>Expiry1: <span>{exp1}</span></div>
  <div>Expiry2: <span>{exp2}</span></div>
  <div class="refresh">Updated: {now}</div>
</div>
<div class="table-wrap">
<table>
<thead><tr>{headers}</tr></thead>
<tbody>
{"".join(body_rows)}
</tbody>
</table>
</div>
</body>
</html>"""

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[gen_dashboard] Generated at {now}")


if __name__ == "__main__":
    generate()
