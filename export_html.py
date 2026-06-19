# -*- coding: utf-8 -*-
"""
export_html.py — MM파트_재무손익.xlsb 'Report' 탭 → 웹용 PnL Trend HTML 생성 (일회성)

매일 아침 한 번 실행. 'Report' 시트의 월별 손익/장부금액(MM PnL Trend)을 읽어
(기준일은 'PnL' 시트 C2) templates/pnl_trend.html 하나를 덮어쓴다.
  - 날짜 파일명 없음 / 날짜 picker 없음 (고정 파일 1개)
  - 라이브 아님: serve.py(모니터링 Excel 점유)와 별개의 일회성 스크립트

차트는 폐쇄망용 로컬 chart.umd.js 사용 → 이 스크립트 폴더(BASE_DIR)에 chart.umd.js 를 둘 것.
실행:  python export_html.py
"""

import os
import json
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
UNIT_LABEL  = "단위 백만원"
SHOW_BASEDATE = True                    # 헤더에 기준일(C2) 표기 여부

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
OUTPUT      = os.path.join(BASE_DIR, "templates", "pnl_trend.html")
CHART_JS    = "chart.umd.js"            # BASE_DIR 에 위치, /chart.umd.js 로 서빙됨
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


def _extract(ws):
    """'Report' 시트의 월별 블록을 라벨 기준으로 추출. {라벨: [12개월 값]} 와 기준일 반환."""
    used = ws.UsedRange
    vals = used.Value  # 2D 튜플 (row-major), bulk read 1회
    if vals is None:
        raise RuntimeError(f"'{SHEET_NAME}' 시트 UsedRange가 비어있음")

    # 1) 앵커('장부금액') 위치 → 라벨 열 확정
    anchor = None
    for i, row in enumerate(vals):
        for j, v in enumerate(row):
            if isinstance(v, str) and v.strip() == ANCHOR:
                anchor = (i, j)
                break
        if anchor:
            break
    if not anchor:
        raise RuntimeError(f"'{ANCHOR}' 라벨을 '{SHEET_NAME}' 시트에서 못 찾음")
    anchor_i, lab_col = anchor

    # 2) 같은 열에서 라벨 행 찾고(앵커 ±8행), 오른쪽 12칸 읽기
    data = {}
    for i, row in enumerate(vals):
        if abs(i - anchor_i) > 8:
            continue
        cell = row[lab_col] if lab_col < len(row) else None
        if isinstance(cell, str) and cell.strip() in ROW_LABELS:
            nums = []
            for k in range(1, MONTHS + 1):
                cidx = lab_col + k
                nums.append(_num(row[cidx]) if cidx < len(row) else None)
            data[cell.strip()] = nums

    missing = [l for l in ROW_LABELS if l not in data]
    if missing:
        print(f"[export] 경고: 라벨 누락 {missing} (해당 행은 빈 값으로 처리)")
    for l in missing:
        data[l] = [None] * MONTHS

    basedate = ""
    if SHOW_BASEDATE:
        try:
            ds = ws.Parent.Sheets(DATE_SHEET)
            basedate = str(ds.Range(DATE_CELL).Text).strip()
        except Exception:
            basedate = ""
    return data, basedate


def _render(data, basedate):
    payload = {
        "months":   [f"{m}월" for m in range(1, MONTHS + 1)],
        "pnlRows":  PNL_ROWS,
        "notional": NOTIONAL,
        "series":   {k: data[k] for k in ROW_LABELS},
        "basedate": basedate,
        "unit":     UNIT_LABEL,
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return _HTML_HEAD + f"<script>window.PNL_DATA = {data_json};</script>\n" + _HTML_BODY


def export():
    pythoncom.CoInitialize()
    try:
        ws = _get_ws()
    except Exception as e:
        print(f"[export] Excel 연결 실패: {e}")
        return
    data, basedate = _extract(ws)
    html = _render(data, basedate)
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    if not os.path.exists(os.path.join(BASE_DIR, CHART_JS)):
        print(f"[export] 경고: {CHART_JS} 가 {BASE_DIR} 에 없음 → 차트가 안 그려집니다. "
              f"chart.umd.js(Chart.js v4 UMD)를 이 폴더에 두세요.")
    print(f"[export] 저장 완료: {OUTPUT}  (기준일 {basedate or '미상'})")


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
  thead th{ font-size:14px; font-weight:600; color:var(--muted); text-align:right; padding:10px 14px; }
  thead th.lab{ text-align:left; }
  tbody td{ font-size:15px; text-align:right; padding:10px 14px; font-variant-numeric:tabular-nums; white-space:nowrap; }
  tbody td.lab{ text-align:left; color:var(--muted); white-space:nowrap; }
  tr.sum td{ border-top:1px solid var(--line2); font-weight:600; }
  tr.sum td.lab{ color:var(--ink); }
  tr.book td{ border-top:1px solid var(--line); color:var(--muted); }
</style>
</head>
<body>
"""

_HTML_BODY = """<div class="lgrow">
  <div class="lg" id="legend"></div>
  <span class="s" id="sub"></span>
</div>
<div class="chart-wrap"><canvas id="pnlChart" role="img" aria-label="월별 MM 손익(좌축)과 장부금액(우축) 추세"></canvas></div>
<div id="tbl" style="overflow-x:auto;"></div>

<script src="chart.umd.js"></script>
<script>
(function(){
  var D = window.PNL_DATA || {months:[],series:{},pnlRows:[],notional:'',basedate:'',unit:''};
  var C = { '주식파생':'#378ADD','지수파생':'#EF9F27','차익거래':'#9A9892','장부금액':'#0C447C','합계':'#E0A82E' };
  var POS='#A32D2D', NEG='#185FA5', MUTED='#5F5E5A';
  function fmt(v){ return (v===null||v===undefined) ? '–' : Math.round(v).toLocaleString('en-US'); }
  function col(v){ return (v===null||v===undefined||v===0) ? '#8a8a85' : (v>0?POS:NEG); }

  document.getElementById('sub').textContent =
    (D.basedate ? '기준일 '+D.basedate+' · ' : '') + (D.unit||'');

  var lg = document.getElementById('legend'); var html='';
  html += '<span><span class="ln" style="background:'+C['합계']+'"></span>손익합계 <span style="color:#8a8a85">(좌축)</span></span>';
  html += '<span style="color:#d4d4d0">│</span>';
  html += '<span><span class="sw" style="background:'+C[D.notional]+'"></span>'+D.notional+' <span style="color:#8a8a85">(우축)</span></span>';
  lg.innerHTML = html;

  var order = D.pnlRows.concat([D.notional]);
  var colg = '<colgroup><col style="width:88px">' + D.months.map(function(){ return '<col>'; }).join('') + '<col style="width:64px"></colgroup>';
  var thM = D.months.map(function(m){ return '<th>'+m+'</th>'; }).join('');
  var body = order.map(function(n){
    var isSum = (n==='합계'), isBook = (n===D.notional), pnl = !isBook;
    var cls = isSum ? 'sum' : (isBook ? 'book' : '');
    var arr = D.series[n] || [];
    var cells = arr.map(function(v){
      var c = pnl ? col(v) : (v==null?'#8a8a85':MUTED);
      return '<td style="color:'+c+'">'+fmt(v)+'</td>';
    }).join('');
    return '<tr class="'+cls+'"><td class="lab">'+(n==='합계'?'손익합계':n)+'</td>'+cells+'<td></td></tr>';
  }).join('');
  document.getElementById('tbl').innerHTML =
    '<table style="table-layout:fixed;width:100%">'+colg+'<thead><tr><th class="lab"></th>'+thM+'<th></th></tr></thead><tbody>'+body+'</tbody></table>';

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
  ds.push({type:'line',label:'손익합계',data:D.series['합계'],yAxisID:'y',
           borderColor:C['합계'],backgroundColor:C['합계'],borderWidth:3.75,pointRadius:2.5,tension:0.3});
  // 두 축의 0(영점) 위치를 일치시킴: min/max 비율을 동일하게 맞춤
  function _flat(keys){ var a=[]; keys.forEach(function(k){ (D.series[k]||[]).forEach(function(v){ if(v!=null)a.push(v); }); }); return a; }
  var _L=_flat(['합계']), _R=(D.series[D.notional]||[]).filter(function(v){ return v!=null; });
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
        y:{ position:'left',  min:Lmin, max:Lmax, afterFit:function(a){ a.width=88; }, title:{display:true,text:'손익',color:MUTED,font:{size:12}},
            grid:{ color:'rgba(0,0,0,0.07)' }, ticks:{ color:MUTED, font:{size:12}, callback:function(v){ return Math.round(v).toLocaleString(); } } },
        y1:{ position:'right', min:Rmin, max:Rmax, afterFit:function(a){ a.width=64; }, title:{display:true,text:'장부금액',color:MUTED,font:{size:12}},
            grid:{ drawOnChartArea:false }, ticks:{ color:MUTED, font:{size:12}, callback:function(v){ return Math.round(v).toLocaleString(); } } }
      } }
  });
  var NUDGE = 14;  // 표 월 위치 미세조정(px, 양수=오른쪽으로)
  function _syncMonths(){
    var ca=_chart.chartArea, cg=document.querySelector('#tbl colgroup'), tbl=document.querySelector('#tbl table');
    if(!ca || !cg || !cg.children.length || !tbl) return;
    var tw=tbl.getBoundingClientRect().width;
    cg.children[0].style.width = Math.round(ca.left+NUDGE)+'px';
    cg.children[cg.children.length-1].style.width = Math.max(0, Math.round(tw-ca.right-NUDGE))+'px';
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
