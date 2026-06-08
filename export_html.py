# -*- coding: utf-8 -*-
"""
export_html.py — Excel 범위 → HTML 파일 변환 (독립 실행 스크립트)

열려있는 Excel 파일의 지정 탭/범위를 스타일 보존 HTML 테이블로 저장.
셀 배경색 · 글자색 · 볼드 · 정렬 · 병합 유지.
"""

import os
import datetime
import pythoncom
import win32com.client

# ── INPUT VARIABLES ──────────────────────────────────────────────────────────
BOOK_NAME  = "MM파트_재무손익.xlsb"  # 열려있는 Excel 파일명 (확장자 포함)
SHEET_NAME = "PnL"                                    # 탭(시트) 이름
CELL_RANGE = "W3:AA24"                                 # 읽을 범위
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archive")
OUTPUT_NAME = ""  # 직접 지정 시 여기에. 비우면 PnL_Report_YYYYMMDD.html 자동 생성
# ─────────────────────────────────────────────────────────────────────────────


def _get_ws(book_name, sheet_name):
    """열려있는 Excel에서 sheet_name 시트 반환. 없으면 자동으로 열기."""
    xl = win32com.client.GetActiveObject("Excel.Application")
    for i in range(1, xl.Workbooks.Count + 1):
        wb = xl.Workbooks.Item(i)
        if (book_name.lower() in wb.Name.lower()
                or book_name.lower() in wb.FullName.lower()):
            return wb.Sheets(sheet_name)

    # 열려있지 않으면 파일 경로로 직접 열기
    book_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), book_name)
    if not os.path.exists(book_path):
        raise RuntimeError(f"파일 없음: {book_path}")
    print(f"[export] '{book_name}' 열는 중...")
    wb = xl.Workbooks.Open(book_path)
    return wb.Sheets(sheet_name)


def _xl_color(xl_color):
    """Excel COM BGR int → CSS #rrggbb. 변환 불가 시 빈 문자열."""
    try:
        c = int(xl_color)
        if c < 0:
            return ""
        r = c & 0xFF
        g = (c >> 8) & 0xFF
        b = (c >> 16) & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return ""


def _cell_style(cell):
    """cell 서식 → CSS style 문자열."""
    styles = []

    # 배경색 (ColorIndex -4142 = 채우기 없음)
    try:
        if cell.Interior.ColorIndex != -4142:
            bg = _xl_color(cell.Interior.Color)
            if bg and bg.lower() != "#ffffff":
                styles.append(f"background:{bg}")
    except Exception:
        pass

    # 글자색
    try:
        fg = _xl_color(cell.Font.Color)
        if fg and fg.lower() != "#000000":
            styles.append(f"color:{fg}")
    except Exception:
        pass

    # 볼드
    try:
        if cell.Font.Bold:
            styles.append("font-weight:bold")
    except Exception:
        pass

    # 수평 정렬 (-4152=center, -4161=right, -4131=left, 1=general)
    try:
        align_map = {-4152: "center", -4161: "right", -4131: "left"}
        align = align_map.get(cell.HorizontalAlignment, "")
        if align:
            styles.append(f"text-align:{align}")
    except Exception:
        pass

    return "; ".join(styles)


def _merge_info(cell):
    """(is_master, colspan, rowspan). 병합 슬레이브 셀이면 is_master=False."""
    try:
        if not cell.MergeCells:
            return True, 1, 1
        area = cell.MergeArea
        master = area.Cells(1, 1)
        if cell.Row == master.Row and cell.Column == master.Column:
            return True, area.Columns.Count, area.Rows.Count
        return False, 0, 0
    except Exception:
        return True, 1, 1


def _prev_business_day(d):
    """d 기준 직전 영업일 (토→금, 일→금)."""
    d -= datetime.timedelta(days=1)
    while d.weekday() >= 5:  # 5=토, 6=일
        d -= datetime.timedelta(days=1)
    return d


def _ask_date():
    """오늘 날짜를 기본값으로 제안, 엔터 시 사용 / 직접 입력 시 파싱."""
    default = datetime.date.today()
    raw = input(f"날짜 [{default.strftime('%Y%m%d')}]: ").strip()
    if not raw:
        return default
    for fmt in ("%Y%m%d", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    print(f"[export] 날짜 형식 오류 ({raw}), 기본값 사용")
    return default


def export(book_name=BOOK_NAME, sheet_name=SHEET_NAME, cell_range=CELL_RANGE,
           output_dir=OUTPUT_DIR, output_name=OUTPUT_NAME):

    pythoncom.CoInitialize()

    try:
        ws = _get_ws(book_name, sheet_name)
    except Exception as e:
        print(f"[export] Excel 연결 실패: {e}")
        return

    rng = ws.Range(cell_range)
    rows_count = rng.Rows.Count
    cols_count = rng.Columns.Count

    if not output_name:
        target_date = _ask_date()
        output_name = f"PnL_Report_{target_date.strftime('%Y%m%d')}.html"
    else:
        target_date = datetime.date.today()

    os.makedirs(output_dir, exist_ok=True)
    print(f"[export] 읽는 중: {sheet_name} {cell_range} ({rows_count}행 × {cols_count}열) → {output_name}")

    html_rows = []
    for r in range(1, rows_count + 1):
        html_cells = []
        for c in range(1, cols_count + 1):
            cell = rng.Cells(r, c)
            is_master, colspan, rowspan = _merge_info(cell)
            if not is_master:
                continue  # 병합 슬레이브 셀 스킵

            try:
                text = str(cell.Text).strip()
            except Exception:
                try:
                    v = cell.Value
                    text = "" if v is None else str(v).strip()
                except Exception:
                    text = ""

            # 숫자 값이면 우정렬
            is_num = False
            try:
                v = cell.Value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    is_num = True
            except Exception:
                pass

            style_parts = _cell_style(cell)
            if is_num and "text-align" not in style_parts:
                style_parts = (style_parts + "; text-align:right").lstrip("; ")

            attrs = f' style="{style_parts}"' if style_parts else ""
            if colspan > 1:
                attrs += f' colspan="{colspan}"'
            if rowspan > 1:
                attrs += f' rowspan="{rowspan}"'

            html_cells.append(f"      <td{attrs}>{text}</td>")

        html_rows.append("    <tr>\n" + "\n".join(html_cells) + "\n    </tr>")

    table_html = "\n".join(html_rows)
    generated_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <title>PnL Report {target_date.strftime('%Y-%m-%d')}</title>
  <style>
    body {{
      font-family: 'Segoe UI', sans-serif;
      font-size: 16px;
      padding: 8px 36px;
      background: #fff;
    }}
    .meta {{
      color: #aaa;
      font-size: 13px;
      margin-bottom: 10px;
    }}
    table {{
      border-collapse: collapse;
      white-space: nowrap;
      min-width: 780px;
    }}
    td {{
      border: 1px solid #d0d0d0;
      padding: 6px 24px;
      min-width: 150px;
    }}
    td:first-child {{
      min-width: 90px;
    }}
  </style>
</head>
<body>
  <div class="meta">PnL Report &nbsp;|&nbsp; {target_date.strftime('%Y-%m-%d')} &nbsp;|&nbsp; Generated: {generated_at}</div>
  <table>
{table_html}
  </table>
</body>
</html>
"""

    out_path = os.path.join(output_dir, output_name)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[export] 저장 완료: {out_path}")


if __name__ == "__main__":
    export()
