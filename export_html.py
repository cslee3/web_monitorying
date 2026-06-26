# -*- coding: utf-8 -*-
"""
export_html.py — MM파트_재무손익.xlsb → 웹용 PnL 리포트 2종 생성 (일회성)

매일 아침 한 번 실행하면 같은 워크북에서 둘 다 생성:
  1) PnL Trend    : 'Report' 탭 월별 손익/장부금액 → templates/pnl_trend.html (날짜 없음)
  2) Daily Report : 'PnL' 탭 W3:AA24 셀 스타일 덤프 → archive/PnL_Report_<날짜>.html (날짜 입력)
기준일은 'PnL' 탭 C2. 라이브 아님(serve.py와 별개 일회성).

차트는 폐쇄망용 로컬 chart.umd.js 사용 → 이 스크립트 폴더(BASE_DIR)에 chart.umd.js 를 둘 것.
실행:  python export_html.py
"""

import os
import json
import datetime
import pythoncom
import win32com.client

# ── 설정 ─────────────────────────────────────────────────────────────────────
BOOK_NAME   = "MM파트_재무손익.xlsb"   # 열려있는 Excel 파일명 (부분일치, '0. ' 접두어 등 무관)
SHEET_NAME  = "Report"                  # 월별 트렌드 표(주식/지수/차익/합계/장부금액 × 1~12월)가 있는 탭
DATE_SHEET  = "PnL"                     # 기준일이 있는 탭
DATE_CELL   = "C2"                      # 기준일 셀
ANCHOR      = "장부금액"                # 월별 블록의 기준 라벨(시트에서 유일)
ROW_LABELS  = ["주식파생", "지수파생", "차익거래", "합계", "장부금액"]
PNL_ROWS    = ["주식파생", "지수파생", "차익거래", "합계"]   # 손익(좌축)
NOTIONAL    = "장부금액"                # 자산(우축)
MONTHS      = 12
UNIT_LABEL  = "단위 백만원"             # PnL Trend
OUT_UNIT_LABEL = "단위 억"              # Outstanding Trend (숫자는 억 단위로 맞춰둠)
SHOW_BASEDATE = True                    # 헤더에 기준일(C2) 표기 여부

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
OUTPUT      = os.path.join(BASE_DIR, "templates", "pnl_trend.html")
CHART_JS    = "chart.umd.js"            # BASE_DIR 에 위치, /chart.umd.js 로 서빙됨

# Daily Report (옛 보고서): PnL 시트 셀을 스타일 보존 덤프 → archive/PnL_Report_<날짜>.html
DAILY_SHEET = "PnL"
DAILY_RANGE = "W3:AA29"
ARCHIVE_DIR = os.path.join(BASE_DIR, "archive")

# Outstanding Trend: Report 탭의 장부금액·평가금액 월별 → templates/outstanding_trend.html
OUT_LABELS = ["장부금액", "평가금액"]
OUT_OUTPUT = os.path.join(BASE_DIR, "templates", "outstanding_trend.html")
# ─────────────────────────────────────────────────────────────────────────────


def _get_ws():
    """열려있는 모든 Excel 인스턴스에서 BOOK_NAME 워크북의 SHEET_NAME 시트를 찾는다.
    못 찾으면 파일을 읽기전용으로 직접 연다."""
    # 1) Running Object Table 전수 검색 (인스턴스 여러 개여도 탐색)
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for mon in rot:
            try:
                name = mon.GetDisplayName(ctx, None)
            except Exception:
                continue
            if name and BOOK_NAME.lower() in name.lower():
                try:
                    obj = rot.GetObject(mon)
                    wb = win32com.client.Dispatch(obj)
                    return wb.Sheets(SHEET_NAME)
                except Exception:
                    continue
    except Exception:
        pass

    # 2) 활성 인스턴스의 Workbooks 검색
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        for i in range(1, xl.Workbooks.Count + 1):
            wb = xl.Workbooks.Item(i)
            if (BOOK_NAME.lower() in wb.Name.lower()
                    or BOOK_NAME.lower() in wb.FullName.lower()):
                return wb.Sheets(SHEET_NAME)
    except Exception:
        xl = None

    # 3) 파일을 읽기전용으로 직접 열기 (BASE_DIR에서 이름 맞는 .xlsb 중 가장 최근 파일)
    import glob
    stem = os.path.splitext(BOOK_NAME)[0]
    cands = glob.glob(os.path.join(BASE_DIR, f"*{stem}*.xlsb"))
    if not cands:
        raise RuntimeError(f"열려있는 '{BOOK_NAME}'를 못 찾고, BASE_DIR에 파일도 없음")
    book_path = max(cands, key=os.path.getmtime)
    if xl is None:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
    print(f"[export] '{BOOK_NAME}' 읽기전용으로 여는 중...")
    wb = xl.Workbooks.Open(book_path, ReadOnly=True, UpdateLinks=0)
    return wb.Sheets(SHEET_NAME)


def _num(v):
    """셀 값 → 숫자 또는 None (빈칸/대시는 None)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        n = round(float(v), 2)
        return int(n) if float(n).is_integer() else n
    s = str(v).strip().replace(",", "")
    if s in ("", "-"):
        return None
    try:
        n = round(float(s), 2)
        return int(n) if float(n).is_integer() else n
    except ValueError:
        return None


def _read_report_rows(ws):
    """Report 탭의 데이터 행을 시트 순서 그대로 반환(하위행 합치지 않음).
    각 행: {'label':주라벨 or '', 'sub':하위라벨 or '', 'vals':[12개월], 'ytd':연간합계}
    - 라벨 열: '장부금액' 앵커가 있는 열, 하위라벨 열: 그 다음 열
    - 월 시작 열: '1월' 헤더 열(라벨↔월 사이 하위열이 끼어도 자동 대응), 연간합계: 12월 다음 열
    예) 차익거래 행(sub=주식) + 다음 행(label='', sub=ELW) 으로 따로 보존.
    """
    used = ws.UsedRange
    vals = used.Value
    if vals is None:
        raise RuntimeError(f"'{SHEET_NAME}' 시트 UsedRange가 비어있음")
    lab_col = None
    for row in vals:
        for j, v in enumerate(row):
            if isinstance(v, str) and v.strip() == ANCHOR:
                lab_col = j; break
        if lab_col is not None:
            break
    if lab_col is None:
        raise RuntimeError(f"'{ANCHOR}' 라벨을 '{SHEET_NAME}' 시트에서 못 찾음")
    month_col = None
    for row in vals:
        for j, v in enumerate(row):
            if isinstance(v, str) and v.strip() == "1월":
                month_col = j; break
        if month_col is not None:
            break
    if month_col is None:
        month_col = lab_col + 1
    sub_col = lab_col + 1
    yc = month_col + MONTHS
    rows = []
    for row in vals:
        lab = row[lab_col] if lab_col < len(row) else None
        labs = lab.strip() if isinstance(lab, str) else ""
        sub = ""
        if sub_col < month_col:                       # 하위라벨 열이 따로 있을 때만
            sv = row[sub_col] if sub_col < len(row) else None
            sub = sv.strip() if isinstance(sv, str) else ""
        mvals = []
        for k in range(MONTHS):
            cidx = month_col + k
            mvals.append(_num(row[cidx]) if cidx < len(row) else None)
        yv = _num(row[yc]) if yc < len(row) else None
        if (labs or sub) and any(v is not None for v in mvals):
            rows.append({"label": labs, "sub": sub, "vals": mvals, "ytd": yv})
    return rows


def _extract(ws):
    """Report 탭 → (차트용 series, 표용 tableRows[하위행 분리], 기준일).
    - series: 라벨 있는 행만 {라벨:[12]} (차트는 합계·장부금액만 사용)
    - tableRows: 처음~장부금액(평가금액 제외), 차익거래의 주식/ELW 하위행 각각 보존."""
    rows = _read_report_rows(ws)
    series = {}
    for r in rows:
        if r["label"]:
            series[r["label"]] = r["vals"]
    table_rows = []
    for r in rows:
        if r["label"] == "평가금액":
            break
        table_rows.append({"label": r["label"], "sub": r["sub"], "vals": r["vals"], "ytd": r["ytd"]})
    for need in ("합계", NOTIONAL):
        if need not in series:
            print(f"[export] 경고: 차트용 '{need}' 행 없음")
    basedate = ""
    if SHOW_BASEDATE:
        try:
            ds = ws.Parent.Sheets(DATE_SHEET)
            basedate = str(ds.Range(DATE_CELL).Text).strip()
        except Exception:
            basedate = ""
    return series, table_rows, basedate


def _render(series, table_rows, basedate):
    payload = {
        "months":    [f"{m}월" for m in range(1, MONTHS + 1)],
        "notional":  NOTIONAL,
        "series":    series,
        "tableRows": table_rows,
        "basedate":  basedate,
        "unit":      UNIT_LABEL,
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return _HTML_HEAD + f"<script>window.PNL_DATA = {data_json};</script>\n" + _HTML_BODY


def _extract_outstanding(ws):
    """Report 탭 → ({장부금액, 평가금액: [12개월]}, {라벨: 연간합계}) (Outstanding Trend용)."""
    rows = _read_report_rows(ws)
    data = {l: [None] * MONTHS for l in OUT_LABELS}
    ytd  = {l: None for l in OUT_LABELS}
    for r in rows:
        if r["label"] in OUT_LABELS:
            data[r["label"]] = r["vals"]
            ytd[r["label"]]  = r["ytd"]
    return data, ytd


def _render_outstanding(data, ytd, basedate):
    payload = {
        "months":   [f"{m}월" for m in range(1, MONTHS + 1)],
        "rows":     OUT_LABELS,
        "series":   {k: data[k] for k in OUT_LABELS},
        "ytd":      {k: ytd[k] for k in OUT_LABELS},
        "basedate": basedate,
        "unit":     OUT_UNIT_LABEL,
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return _HTML_HEAD + f"<script>window.OUT_DATA = {data_json};</script>\n" + _OUT_BODY


def export():
    pythoncom.CoInitialize()
    try:
        ws = _get_ws()
    except Exception as e:
        print(f"[export] Excel 연결 실패: {e}")
        return
    series, table_rows, basedate = _extract(ws)
    html = _render(series, table_rows, basedate)
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    if not os.path.exists(os.path.join(BASE_DIR, CHART_JS)):
        print(f"[export] 경고: {CHART_JS} 가 {BASE_DIR} 에 없음 → 차트가 안 그려집니다. "
              f"chart.umd.js(Chart.js v4 UMD)를 이 폴더에 두세요.")
    print(f"[export] 저장 완료: {OUTPUT}  (기준일 {basedate or '미상'})")

    # Outstanding Trend (장부금액·평가금액)
    try:
        odata, oytd = _extract_outstanding(ws)
        with open(OUT_OUTPUT, "w", encoding="utf-8") as f:
            f.write(_render_outstanding(odata, oytd, basedate))
        print(f"[out] 저장 완료: {OUT_OUTPUT}")
    except Exception as e:
        print(f"[out] Outstanding Trend 생성 실패: {e}")

    # Daily Report (옛 보고서)도 같은 워크북에서 함께 생성 (기준일자로 파일명)
    try:
        _export_daily_report(ws.Parent, basedate)
    except Exception as e:
        print(f"[daily] Daily Report 생성 실패: {e}")


# ── Daily Report : PnL 시트 W3:AA24 셀 스타일 보존 덤프 ───────────────────────
def _xl_color(xl_color):
    """Excel COM BGR int → CSS #rrggbb."""
    try:
        c = int(xl_color)
        if c < 0:
            return ""
        r = c & 0xFF; g = (c >> 8) & 0xFF; b = (c >> 16) & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return ""


def _cell_style(cell):
    """cell 서식 → CSS style 문자열 (배경/글자색/볼드/정렬)."""
    styles = []
    try:
        if cell.Interior.ColorIndex != -4142:
            bg = _xl_color(cell.Interior.Color)
            if bg and bg.lower() != "#ffffff":
                styles.append(f"background:{bg}")
    except Exception:
        pass
    try:
        fg = _xl_color(cell.Font.Color)
        if fg and fg.lower() != "#000000":
            styles.append(f"color:{fg}")
    except Exception:
        pass
    try:
        if cell.Font.Bold:
            styles.append("font-weight:bold")
    except Exception:
        pass
    try:
        align_map = {-4152: "center", -4161: "right", -4131: "left"}
        align = align_map.get(cell.HorizontalAlignment, "")
        if align:
            styles.append(f"text-align:{align}")
    except Exception:
        pass
    return "; ".join(styles)


def _merge_info(cell):
    """(is_master, colspan, rowspan). 병합 슬레이브면 is_master=False."""
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


def _export_daily_report(wb, basedate):
    try:
        ws = wb.Sheets(DAILY_SHEET)
    except Exception as e:
        print(f"[daily] '{DAILY_SHEET}' 시트 접근 실패: {e}")
        return
    rng = ws.Range(DAILY_RANGE)
    rows_count = rng.Rows.Count
    cols_count = rng.Columns.Count

    # 기준일자(PnL!C2)에서 숫자만 뽑아 YYYYMMDD; 못 읽으면 오늘
    digits = "".join(ch for ch in str(basedate) if ch.isdigit())
    if len(digits) >= 8:
        ymd = digits[:8]
        date_disp = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:8]}"
    else:
        _t = datetime.date.today()
        ymd = _t.strftime("%Y%m%d"); date_disp = _t.strftime("%Y-%m-%d")
    output_name = f"PnL_Report_{ymd}.html"
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    print(f"[daily] 읽는 중: {DAILY_SHEET} {DAILY_RANGE} → {output_name}")

    html_rows = []
    for r in range(1, rows_count + 1):
        if r in [1, 4, 5, 6, 7]:   # 1행(제목)·2행(Report 줄) + 4~6 빈행 제외
            continue
        # 요약 3줄 제외: 1열=주식/지수/차익파생 이고 2열(시장/괴리/자금/이론 라벨) 비어있는 행
        try:
            _a = str(rng.Cells(r, 1).Text).strip()
            _b = str(rng.Cells(r, 2).Text).strip()
            if _a in ("주식파생", "지수파생", "차익거래") and _b == "":
                continue
        except Exception:
            pass
        html_cells = []
        for c in range(1, cols_count + 1):
            cell = rng.Cells(r, c)
            is_master, colspan, rowspan = _merge_info(cell)
            if not is_master:
                continue
            try:
                text = str(cell.Text).strip()
            except Exception:
                try:
                    v = cell.Value
                    text = "" if v is None else str(v).strip()
                except Exception:
                    text = ""
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
    html = (_DAILY_HTML
            .replace("__DATE__", date_disp)
            .replace("__GEN__", generated_at)
            .replace("__TABLE__", table_html))
    out_path = os.path.join(ARCHIVE_DIR, output_name)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[daily] 저장 완료: {out_path}")


_DAILY_HTML = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <title>Daily Report __DATE__</title>
  <style>
    body { font-family: 'Segoe UI', sans-serif; font-size: 16px; padding: 8px 36px; background: #fff; }
    .meta { color: #aaa; font-size: 13px; margin-bottom: 10px; }
    table { border-collapse: collapse; white-space: nowrap; min-width: 780px; }
    td { border: 1px solid #d0d0d0; padding: 6px 24px; min-width: 150px; }
    td:first-child { min-width: 90px; }
  </style>
</head>
<body>
  <div class="meta">__DATE__ &nbsp;|&nbsp; Generated: __GEN__</div>
  <table>
__TABLE__
  </table>
</body>
</html>
"""


# ── 출력 HTML 템플릿 (정적·독립 실행, 폐쇄망 로컬 Chart.js) ──────────────────────
_HTML_HEAD = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8" />
<title>PnL Trend</title>
<style>
  :root{ --pos:#A32D2D; --neg:#185FA5; --ink:#1f2328; --muted:#5F5E5A; --hint:#8a8a85;
         --line:#e3e3df; --line2:#cfcfca; --bg:#fff; --emph:#f4f6f9;
         --c-stk:#378ADD; --c-idx:#EF9F27; --c-arb:#9A9892; --c-book:#0C447C; --c-sum:#E0A82E; }
  *{ box-sizing:border-box; }
  body{ font-family:'Segoe UI','Malgun Gothic',sans-serif; color:var(--ink); background:var(--bg);
        margin:0; padding:18px 28px; }
  .lgrow{ display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; margin:2px 0 12px; }
  .s{ font-size:13px; color:var(--hint); }
  .lg{ display:flex; flex-wrap:wrap; gap:16px; font-size:13px; color:var(--muted); }
  .lg span{ display:inline-flex; align-items:center; gap:5px; }
  .sw{ width:10px; height:10px; border-radius:2px; display:inline-block; }
  .ln{ width:16px; height:3px; border-radius:2px; display:inline-block; }
  .chart-wrap{ position:relative; width:100%; height:480px; margin-bottom:30px; }
  table{ border-collapse:collapse; width:100%; min-width:760px; }
  thead th{ font-size:13px; font-weight:600; color:var(--muted); text-align:right; padding:9px 13px; border-bottom:1.5px solid var(--line2); }
  thead th.lab{ text-align:center; border-right:1px solid var(--line2); }
  tbody td{ font-size:15px; text-align:right; padding:9px 13px; font-variant-numeric:tabular-nums; white-space:nowrap; border-bottom:1px solid var(--line); }
  tbody tr:last-child td{ border-bottom:none; }
  tbody td.lab{ text-align:center; color:var(--muted); white-space:nowrap; font-size:13px; padding-left:5px; padding-right:5px; border-right:1px solid var(--line2); }
  tr.sum td{ border-top:1.5px solid var(--line2); border-bottom-color:var(--line2); font-weight:600; }
  tr.sum td.lab{ color:var(--ink); }
  tr.book td{ color:var(--muted); }
  .ytd{ border-left:1px solid var(--line2); font-weight:600; }
  thead th.ytd{ color:var(--ink); }
  td.lab.parent{ vertical-align:middle; color:var(--ink); border-right:none; padding-left:3px; padding-right:3px; }
  td.lab.sub{ color:var(--muted); padding-left:3px; padding-right:5px; }
</style>
</head>
<body>
"""

_HTML_BODY = """<div class="lgrow">
  <div class="lg" id="legend"></div>
  <span class="s" id="sub"></span>
</div>
<div class="chart-wrap"><canvas id="pnlChart" role="img" aria-label="월별 누적손익(좌축)과 장부금액(우축) 추세"></canvas></div>
<div id="tbl" style="overflow-x:auto;"></div>

<script src="chart.umd.js"></script>
<script>
(function(){
  var D = window.PNL_DATA || {months:[],series:{},pnlRows:[],notional:'',basedate:'',unit:''};
  var C = { '주식파생':'#378ADD','지수파생':'#EF9F27','차익거래':'#9A9892','장부금액':'#0C447C','합계':'#E0A82E' };
  var POS='#A32D2D', NEG='#185FA5', MUTED='#5F5E5A';
  function fmt(v){ return (v===null||v===undefined) ? '' : Math.round(v).toLocaleString('en-US'); }
  function col(v){ return (v===null||v===undefined||v===0) ? '#8a8a85' : (v>0?POS:NEG); }

  document.getElementById('sub').textContent =
    (D.basedate ? '기준일 '+D.basedate+' · ' : '') + (D.unit||'');

  var lg = document.getElementById('legend'); var html='';
  html += '<span><span class="ln" style="background:'+C['합계']+'"></span>누적손익 <span style="color:#8a8a85">(좌축)</span></span>';
  html += '<span style="color:#d4d4d0">│</span>';
  html += '<span><span class="sw" style="background:'+C[D.notional]+'"></span>'+D.notional+' <span style="color:#8a8a85">(우축)</span></span>';
  lg.innerHTML = html;

  var colg = '<colgroup><col style="width:70px"><col style="width:46px">' + D.months.map(function(){ return '<col>'; }).join('') + '<col style="width:84px"><col style="width:64px"></colgroup>';
  var thM = D.months.map(function(m){ return '<th>'+m+'</th>'; }).join('');
  function vcells(r){
    var pnl = (r.label!==D.notional);
    var cs = (r.vals||[]).map(function(v){
      var c = pnl ? col(v) : (v==null?'#8a8a85':MUTED);
      return '<td style="color:'+c+'">'+fmt(v)+'</td>';
    }).join('');
    var yv = (r.ytd!=null) ? r.ytd : null;
    var yc = pnl ? col(yv) : (yv==null?'#8a8a85':MUTED);
    return cs + '<td class="ytd" style="color:'+yc+'">'+fmt(yv)+'</td><td></td>';
  }
  var rows = D.tableRows || [];
  var body = '';
  for(var i=0;i<rows.length;i++){
    var r = rows[i];
    if(r.sub){                                  // 하위행 그룹(예: 차익거래 = 주식+ELW)
      var grp=[r], j=i+1;
      while(j<rows.length && !rows[j].label && rows[j].sub){ grp.push(rows[j]); j++; }
      for(var g=0; g<grp.length; g++){
        var gr=grp[g];
        var ptd = (g===0) ? '<td class="lab parent" rowspan="'+grp.length+'">'+r.label+'</td>' : '';
        body += '<tr>'+ptd+'<td class="lab sub">'+gr.sub+'</td>'+vcells(gr)+'</tr>';
      }
      i = j-1;
    } else {                                    // 단일행
      var n=r.label, isSum=(n==='합계'), isBook=(n===D.notional);
      var cls=isSum?'sum':(isBook?'book':'');
      var disp=(n==='합계')?'손익':n;
      body += '<tr class="'+cls+'"><td class="lab" colspan="2">'+disp+'</td>'+vcells(r)+'</tr>';
    }
  }
  document.getElementById('tbl').innerHTML =
    '<table style="table-layout:fixed;width:100%">'+colg+'<thead><tr><th class="lab" colspan="2"></th>'+thM+'<th class="ytd">연간합계</th><th></th></tr></thead><tbody>'+body+'</tbody></table>';

  if(typeof Chart==='undefined'){ return; }
  if(Chart.Tooltip && Chart.Tooltip.positioners){
    Chart.Tooltip.positioners.bookRight = function(items){
      if(!items.length) return false;
      var chart=this.chart, it=null;
      for(var i=0;i<items.length;i++){
        var di=items[i].datasetIndex, dsx=chart.data.datasets[di];
        if(dsx && dsx.label==='장부금액'){ it=items[i]; break; }
      }
      if(!it) it=items[0];
      var el=it.element; var w=(el && el.width)?el.width:0;
      var ymid=(el.y!=null && el.base!=null) ? (el.y+el.base)/2 : el.y;
      return { x: el.x + w/2, y: ymid };
    };
  }
  function hexA(hex,a){ var h=(hex||'#888888').replace('#',''); if(h.length===3){h=h[0]+h[0]+h[1]+h[1]+h[2]+h[2];}
    return 'rgba('+parseInt(h.substr(0,2),16)+','+parseInt(h.substr(2,2),16)+','+parseInt(h.substr(4,2),16)+','+a+')'; }
  function bar(n,axis){ return {type:'bar',label:n,data:D.series[n],backgroundColor:hexA(C[n],0.65),
                               yAxisID:axis,borderRadius:2,categoryPercentage:0.8,barPercentage:0.9,maxBarThickness:40}; }
  var ds = [];
  ds.push(bar(D.notional,'y1'));
  // 손익선 = 월별 합계의 누적합(cumsum), 데이터 없는 월(null)은 끊음
  function cumsum(arr){ var out=[], run=0; (arr||[]).forEach(function(v){ if(v==null){ out.push(null); } else { run+=v; out.push(Math.round(run*100)/100); } }); return out; }
  var cum = cumsum(D.series['합계']);
  ds.push({type:'line',label:'누적손익',data:cum,yAxisID:'y',
           borderColor:C['합계'],backgroundColor:C['합계'],borderWidth:3.75,pointRadius:2.5,tension:0.3});
  // 두 축의 0(영점) 위치를 일치시킴: min/max 비율을 동일하게 맞춤
  var _L=cum.filter(function(v){ return v!=null; }), _R=(D.series[D.notional]||[]).filter(function(v){ return v!=null; });
  var Lmax=Math.max.apply(null,[0].concat(_L))*1.08, Lmin=Math.min.apply(null,[0].concat(_L))*1.08;
  var Rmax=Math.max.apply(null,[0].concat(_R))*1.08, Rmin=Math.min.apply(null,[0].concat(_R))*1.08;
  var _lo=Lmax?Lmin/Lmax:0, _ro=Rmax?Rmin/Rmax:0;
  if(_lo<_ro){ Rmin=Rmax*_lo; } else { Lmin=Lmax*_ro; }
  var _chart = new Chart(document.getElementById('pnlChart'),{
    data:{ labels:D.months, datasets:ds },
    options:{ responsive:true, maintainAspectRatio:false,
      interaction:{ mode:'index', intersect:false },
      plugins:{ legend:{display:false},
        tooltip:{ position:'bookRight', xAlign:'left',
          callbacks:{ label:function(c){ return c.dataset.label+': '+(c.parsed.y==null?'–':Math.round(c.parsed.y).toLocaleString()); } } } },
      scales:{
        x:{ grid:{display:false}, ticks:{ color:MUTED, font:{size:12} } },
        y:{ position:'left',  min:Lmin, max:Lmax, afterFit:function(a){ a.width=88; }, title:{display:true,text:'누적손익',color:MUTED,font:{size:12}},
            grid:{ color:'rgba(0,0,0,0.07)' }, ticks:{ color:MUTED, font:{size:12}, callback:function(v){ return Math.round(v).toLocaleString(); } } },
        y1:{ position:'right', min:Rmin, max:Rmax, afterFit:function(a){ a.width=64; }, title:{display:true,text:'장부금액',color:MUTED,font:{size:12}},
            grid:{ drawOnChartArea:false }, ticks:{ color:MUTED, font:{size:12}, callback:function(v){ return Math.round(v).toLocaleString(); } } }
      } }
  });
  var YTDW = 84, LABELW = 116, SPACERW = 58;   // 연간합계 폭 / 라벨 2칸 합(70+46) / 우측여백
  function _syncMonths(){
    var cg=document.querySelector('#tbl colgroup'), tbl=document.querySelector('#tbl table'), cont=document.getElementById('tbl');
    if(!cg || cg.children.length<16 || !tbl || !cont) return;
    // 라벨은 고정, 월 12칸+연간합계를 컨테이너 가용폭에 정확히 맞춤(가로 스크롤바 방지)
    var tw=cont.clientWidth - 2;
    var monthW=Math.max(26,(tw - LABELW - YTDW - SPACERW)/12);
    for(var k=0;k<12;k++){ cg.children[2+k].style.width=monthW+'px'; }
    cg.children[14].style.width=YTDW+'px';
    cg.children[15].style.width=SPACERW+'px';
  }
  requestAnimationFrame(function(){ _syncMonths(); requestAnimationFrame(_syncMonths); });
  window.addEventListener('resize', function(){ requestAnimationFrame(_syncMonths); });
})();
</script>
</body>
</html>
"""


_OUT_BODY = """<div class="lgrow">
  <div class="lg" id="legend"></div>
  <span class="s" id="sub"></span>
</div>
<div class="chart-wrap"><canvas id="outChart" role="img" aria-label="월별 장부금액·평가금액 추세"></canvas></div>
<div id="outtbl" style="overflow-x:auto;"></div>

<script src="chart.umd.js"></script>
<script>
(function(){
  var D = window.OUT_DATA || {months:[],series:{},rows:[],basedate:'',unit:''};
  var C = { '장부금액':'#0C447C', '평가금액':'#1D9E75' };
  var MUTED='#5F5E5A';
  function fmt(v){ return (v===null||v===undefined) ? '' : Math.round(v).toLocaleString('en-US'); }

  document.getElementById('sub').textContent =
    (D.basedate ? '기준일 '+D.basedate+' · ' : '') + (D.unit||'');

  var lg = document.getElementById('legend'); var html='';
  D.rows.forEach(function(n){
    html += '<span><span class="sw" style="background:'+(C[n]||'#888')+'"></span>'+n+'</span>';
  });
  lg.innerHTML = html;

  // 평가금액의 가장 최근월(마지막 non-null) 인덱스 — AL43 실시간 갱신 대상
  var EVAL='평가금액', evalArr=D.series[EVAL]||[], latestIdx=-1;
  for(var _i=evalArr.length-1;_i>=0;_i--){ if(evalArr[_i]!=null){ latestIdx=_i; break; } }

  var colg = '<colgroup><col style="width:88px">' + D.months.map(function(){ return '<col>'; }).join('') + '<col style="width:64px"></colgroup>';
  var thM = D.months.map(function(m){ return '<th>'+m+'</th>'; }).join('');
  var body = D.rows.map(function(n){
    var arr = D.series[n] || [];
    var cells = arr.map(function(v,i){
      var idAttr = (n===EVAL && i===latestIdx) ? ' id="evalCell"' : '';
      return '<td'+idAttr+'>'+fmt(v)+'</td>';
    }).join('');
    return '<tr><td class="lab" style="color:'+(C[n]||MUTED)+';font-weight:600">'+n+'</td>'+cells+'<td></td></tr>';
  }).join('');
  document.getElementById('outtbl').innerHTML =
    '<table style="table-layout:fixed;width:100%">'+colg+'<thead><tr><th class="lab"></th>'+thM+'<th></th></tr></thead><tbody>'+body+'</tbody></table>';

  if(typeof Chart==='undefined'){ return; }
  if(Chart.Tooltip && Chart.Tooltip.positioners){
    Chart.Tooltip.positioners.rightBar = function(items){
      if(!items.length) return false;
      var it=items[0];
      for(var i=1;i<items.length;i++){ if(items[i].element.x > it.element.x) it=items[i]; }
      var el=it.element; var w=(el && el.width)?el.width:0;
      var ymid=(el.y!=null && el.base!=null) ? (el.y+el.base)/2 : el.y;
      return { x: el.x + w/2, y: ymid };
    };
  }
  var ds = D.rows.map(function(n){
    return {type:'bar',label:n,data:D.series[n],backgroundColor:C[n]||'#888',
            borderRadius:2,categoryPercentage:0.7,barPercentage:0.9,maxBarThickness:40};
  });
  var _chart = new Chart(document.getElementById('outChart'),{
    data:{ labels:D.months, datasets:ds },
    options:{ responsive:true, maintainAspectRatio:false,
      interaction:{ mode:'index', intersect:false },
      plugins:{ legend:{display:false},
        tooltip:{ position:'rightBar', xAlign:'left',
          callbacks:{ label:function(c){ return c.dataset.label+': '+(c.parsed.y==null?'–':Math.round(c.parsed.y).toLocaleString()); } } } },
      scales:{
        x:{ grid:{display:false}, ticks:{ color:MUTED, font:{size:12} } },
        y:{ position:'left', afterFit:function(a){ a.width=88; }, grid:{ color:'rgba(0,0,0,0.07)' },
            ticks:{ color:MUTED, font:{size:12}, callback:function(v){ return Math.round(v).toLocaleString(); } } }
      } }
  });

  // ── AL43(최근월 평가금액) — SSF MM 스트림에 얹어 실시간 갱신 ──
  var _lastEval;
  function applyEval(v){
    if(v===null || v===undefined || latestIdx<0 || v===_lastEval) return;
    _lastEval = v;
    var di = D.rows.indexOf(EVAL);
    if(di>=0 && _chart.data.datasets[di]){ _chart.data.datasets[di].data[latestIdx] = v; _chart.update('none'); }
    var cell = document.getElementById('evalCell');
    if(cell) cell.textContent = fmt(v);
  }
  fetch('/api/dashboard').then(function(r){ return r.ok?r.json():null; })
    .then(function(d){ if(d && d.header) applyEval(d.header.eval_latest); }).catch(function(){});
  try{
    var _es = new EventSource('/api/stream');
    _es.onmessage = function(e){ try{ var d=JSON.parse(e.data); if(d && d.header) applyEval(d.header.eval_latest); }catch(_e){} };
  }catch(_e){}

  function _syncMonths(){
    var ca=_chart.chartArea, cg=document.querySelector('#outtbl colgroup'), tbl=document.querySelector('#outtbl table');
    if(!ca || !cg || !cg.children.length || !tbl) return;
    var tw=tbl.getBoundingClientRect().width;
    cg.children[0].style.width = Math.round(ca.left)+'px';
    cg.children[cg.children.length-1].style.width = Math.max(0, Math.round(tw-ca.right))+'px';
  }
  requestAnimationFrame(function(){ _syncMonths(); requestAnimationFrame(_syncMonths); });
  window.addEventListener('resize', function(){ requestAnimationFrame(_syncMonths); });
})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    export()
