# -*- coding: utf-8 -*-
"""
gen_sso_report.py — SSO MM 리포트 HTML 생성기
Option_DashBoard 탭을 읽어 sso_report.html로 저장.
serve.py와 독립적으로 단독 실행 가능 (python gen_sso_report.py).
"""

import os
import datetime
import xlwings as xw

# 대상 Excel 파일 경로 (이 스크립트와 같은 디렉토리에 위치해야 함)
EXCEL_PATH = os.path.join(os.path.dirname(__file__),
                          "GlobalMM_Realtime_Monitoring_V2_20260520.xlsb")
# 생성될 HTML 파일 경로
OUT_PATH   = os.path.join(os.path.dirname(__file__), "sso_report.html")


# ── 값 포매터 ────────────────────────────────────────────────────────────

def num(v, d=0):
    """숫자 → 천단위 콤마 + 소수점 d자리 문자열."""
    if v is None or v == "": return ""
    try:
        return f"{v:,.{d}f}"
    except Exception:
        return str(v)

def to_rgb(color):
    """xlwings color tuple (r,g,b) → CSS rgb() 문자열."""
    if isinstance(color, tuple) and len(color) == 3:
        return f"rgb({color[0]},{color[1]},{color[2]})"
    return ""


# ── 컬럼 정의 (Excel 1-based 열 번호, 헤더, 포매터) ─────────────────────
SSO_COLS = [
    (2,  "StockCode",  lambda v: str(v) if v else ""),
    (3,  "StockName",  lambda v: str(v) if v else ""),
    (4,  "Delta",      lambda v: num(v, 0)),
    (6,  "%Gamma",     lambda v: num(v, 2)),
    (8,  "Volume",     lambda v: num(v, 0)),
    (10, "Theo PnL",   lambda v: num(v, 0)),
    (12, "MTM PnL",    lambda v: num(v, 0)),
]

# PnL 컬럼 인덱스 집합 (녹색/적색 텍스트 적용 대상)
PNL_COLS = {c[0] for c in SSO_COLS if "PnL" in c[1]}


def cell_style(bg, bold, val, col_idx):
    """
    셀 인라인 CSS 스타일 문자열 생성.
    - bg: 배경색 (Excel 셀 색상에서 읽음)
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
            fval = float(str(val).replace(",", ""))
            if fval > 0:
                styles.append("color:#1a7c34")
            elif fval < 0:
                styles.append("color:#cc0000")
        except Exception:
            pass
    return "; ".join(styles)


def _get_ws():
    """
    이미 열린 Excel 워크북을 재사용하여 Option_DashBoard 시트 반환.
    열려있지 않으면 파일을 직접 열기.
    새 xw.App을 생성하지 않아 불필요한 Excel 인스턴스 생성 방지.
    """
    fname = os.path.basename(EXCEL_PATH)
    try:
        # 방법 1: 이름으로 직접 조회
        return xw.books[fname].sheets["Option_DashBoard"]
    except Exception:
        pass
    try:
        # 방법 2: 활성 App에서 조회
        app = xw.apps.active
        if app is not None:
            return app.books[fname].sheets["Option_DashBoard"]
    except Exception:
        pass
    # 방법 3: 파일 직접 열기
    return xw.Book(EXCEL_PATH).sheets["Option_DashBoard"]


def generate():
    """
    Option_DashBoard 탭(3~235행)을 행 단위로 읽어 sso_report.html 생성.
    StockCode(B열)가 비어있거나 헤더 텍스트인 행은 skip.
    각 셀의 배경색·굵기를 읽어 인라인 스타일로 적용.
    """
    ws = _get_ws()

    rows_html = []
    for r in range(3, 235):  # 3행부터 235행까지 (헤더 2행 제외)
        stock_code = ws.cells(r, 2).value
        # StockCode 컬럼이 문자열이 아니거나 비어있으면 skip
        if not isinstance(stock_code, str) or stock_code in ("", "StockCode"):
            continue

        tds = []
        for cidx, hdr, fmt in SSO_COLS:
            cell      = ws.cells(r, cidx)
            v         = cell.value
            bg        = to_rgb(cell.color)       # 셀 배경색 (COM 호출)
            bold      = bool(cell.font.bold)     # 굵은 글씨 여부 (COM 호출)
            formatted = ""
            if v is not None and v != "":
                try:
                    formatted = fmt(v)
                except Exception:
                    formatted = str(v)
            style = cell_style(bg, bold, formatted, cidx)
            align = "right" if cidx not in (2, 3) else "left"
            tds.append(f'<td style="{style}; text-align:{align}">{formatted}</td>')

        rows_html.append(f'<tr>{"".join(tds)}</tr>')

    headers = "".join(f"<th>{c[1]}</th>" for c in SSO_COLS)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8"/>
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
</style>
</head>
<body>
<div class="refresh">Updated: {now}</div>
<div class="table-wrap">
<table>
<thead><tr>{headers}</tr></thead>
<tbody>
{"".join(rows_html)}
</tbody>
</table>
</div>
</body>
</html>"""

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[gen_sso_report] Generated at {now}")


if __name__ == "__main__":
    generate()
