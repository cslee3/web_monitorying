#%%
import xlwings as xw
import pandas as pd
import threading
import time
from config import EXCEL_PATH, REFRESH_SEC
#%%
data = {
    "dashboard_MM":     None,
    "dashboard_arb":    None,
    "option_dashboard": None,
    "live_orders":      None,
    "position":         None,
    "daily":            None,
}
_lock = threading.Lock()

#%%
def _get_wb():
    import os
    fname = os.path.basename(EXCEL_PATH)
    try:
        app = xw.apps.active
        if app is None:
            raise Exception("no active excel")
        return app.books[fname]
    except Exception:
        return xw.Book(EXCEL_PATH)

#%%
def _sht_to_raw(sht, nrows, ncols, pivot_pt = "A1"):
    
    raw = sht.range(pivot_pt).resize(nrows, ncols).options(pd.DataFrame, header=False, index=False).value
    raw = pd.DataFrame(raw)
    return raw

#%% ── DashBoard stocks ───────────────────────────────────────────
_DASHBOARD_HEADERS = [
    "Type","Stock","StockName","LastPrice","Change(%)",
    "Amount(MilKRW)","Amount(Shares1)","Amount(Shares2)","MMQty",
    "Ask","Bid","MMSpread","△Shares","Diff",
    "Volume(Lots)","Volume(Mil)","Delta(Shares)","Delta(MilKRW)",
    "TheoPrice","TheoBasis","Ex1 B","Ex1 S",
    "TheoBasis2","Ex2 B","Ex2 S","MTM PnL","Theo PnL"
]

def _parse_dashboard_stocks(wb, pivot_pt, nrows, ncols, nrows_slice):
    df = _sht_to_raw(sht=wb.sheets["DashBoard"], nrows=nrows, ncols=ncols, pivot_pt=pivot_pt)

    rows = df.iloc[:nrows_slice, :].copy()
    rows.columns = _DASHBOARD_HEADERS

    for col in ["LastPrice","Volume(Lots)","Volume(Mil)","MTM_PnL","Theo_PnL"]:
        if col in rows.columns:
            rows[col] = pd.to_numeric(rows[col], errors="coerce")
    
    for col_name in ["LastPrice", "Amount(MilKRW)","Amount(Shares1)","Amount(Shares2)","Volume(Mil)","Ask","Bid",
                     "△Shares",
                     "Delta(Shares)","Delta(MilKRW)","TheoPrice","TheoBasis","Ex1 B","Ex1 S",
                     "TheoBasis2","Ex2 B","Ex2 S","MTM PnL","Theo PnL"]:
        rows[col_name] = pd.to_numeric(rows[col_name], errors="coerce").round(0).astype("Int64")


    col_name = "Type"
    rows[col_name] = rows[col_name].apply(
        lambda v: str(int(v)) if isinstance(v, float) and pd.notna(v) else (str(v) if v is not None else "")
    )
        
    
    rows["Change(%)"] = (pd.to_numeric(rows["Change(%)"], errors="coerce") * 100).round(2).astype(str) + "%"
    
    return rows.reset_index(drop=True)

def _parse_dashboard_MM(wb):
    return _parse_dashboard_stocks(wb, pivot_pt="A6",  nrows=45, ncols=27, nrows_slice=31)

def _parse_dashboard_arb(wb):
    return _parse_dashboard_stocks(wb, pivot_pt="A38", nrows=45, ncols=27, nrows_slice=12)

def _parse_dashboard_MM_sector(wb):
    headers = ["IndexCode","IndexName","Fund","Delta","Volume","MTM PnL","THeoPnL"]
    
    df = _sht_to_raw(sht = wb.sheets["DashBoard"], nrows = 7, ncols = 10, pivot_pt="AC6")

    n = min(len(df.columns), len(headers))
    rows = df.iloc[:7, :n].copy()

    rows.columns = headers[:n]

    for col_name in ["MTM PnL","THeoPnL"]:
        rows[col_name] = pd.to_numeric(rows[col_name], errors="coerce").round(0).astype("Int64")

    col_name = "Fund"
    rows[col_name] = rows[col_name].apply(
        lambda v: str(int(v)) if isinstance(v, float) and pd.notna(v) else (str(v) if v is not None else "")
    )
    
    for col_name in ["Delta", "MTM PnL", "THeoPnL"]:
        rows[col_name] = pd.to_numeric(rows[col_name], errors="coerce").round(0).astype("Int64")

    col_name = "Volume"
    rows[col_name] = (pd.to_numeric(rows[col_name], errors="coerce") * 100).round(0).astype("Int64").astype(str) + "%"

    return rows.reset_index(drop=True)

#%% ── Option_DashBoard: row2~, 요약 행만 ──────────────────────────
def _parse_option_dashboard(wb):
    df = _sht_to_raw(wb.sheets["Option_DashBoard"], 233, 12)
    rows = df.iloc[2:].copy()
    rows = rows[rows.iloc[:, 1].apply(lambda v: isinstance(v, str) and v not in ("StockCode", ""))].copy()
    result = pd.DataFrame({
        "StockCode": rows.iloc[:, 1].values,
        "StockName": rows.iloc[:, 2].values,
        "Delta":     pd.to_numeric(rows.iloc[:, 3].values, errors="coerce"),
        "%Gamma":    pd.to_numeric(rows.iloc[:, 5].values, errors="coerce"),
        "Volume":    pd.to_numeric(rows.iloc[:, 7].values, errors="coerce"),
        "Theo_PnL":  pd.to_numeric(rows.iloc[:, 9].values, errors="coerce"),
        "MTM_PnL":   pd.to_numeric(rows.iloc[:, 11].values, errors="coerce"),
    })
    return result.reset_index(drop=True)


#%% ── LiveOrders: row1 헤더, row2~ 데이터 ─────────────────────────
def _parse_live_orders(wb):
    df = _sht_to_raw(wb.sheets["LiveOrders"], 500, 7)
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    return df[df.iloc[:, 0].notna()].reset_index(drop=True)


#%% ── Position: row1 헤더, row3~ 데이터 ───────────────────────────
def _parse_position(wb):
    df = _sht_to_raw(wb.sheets["Position"], 1750, 13)
    headers = df.iloc[0].tolist()
    clean = []
    seen = {}
    for h in headers:
        h = str(h) if h is not None and not (isinstance(h, float) and pd.isna(h)) else ""
        seen[h] = seen.get(h, 0) + 1
        clean.append(f"{h}_{seen[h]}" if seen[h] > 1 else h)
    df = df.iloc[2:].reset_index(drop=True)
    df.columns = clean
    return df[df.iloc[:, 0].notna()].reset_index(drop=True)

#%% ── Daily: 분류별 요약 (왼쪽 테이블) ────────────────────────────
def _parse_daily(wb):
    df = _sht_to_raw(wb.sheets["Daily"], 31, 15)
    left = df.iloc[2:8, 1:5].copy()
    left.columns = ["분류", "MTM", "THEO", "누적괴리"]
    return left[left["분류"].notna()].reset_index(drop=True)


#%% ── 루프 ────────────────────────────────────────────────────────
def _read_all():
    try:
        wb = _get_wb()
    except Exception as e:
        print(f"[wb error] {e}")
        return

    parsers = {
        "dashboard_MM":     _parse_dashboard_MM,
        "dashboard_arb":    _parse_dashboard_arb,
        "option_dashboard": _parse_option_dashboard,
        "live_orders":      _parse_live_orders,
        "position":         _parse_position,
        "daily":            _parse_daily,
    }
    for key, fn in parsers.items():
        try:
            new_df = fn(wb)
            with _lock:
                old = data[key]
                if old is None or not old.equals(new_df):
                    print(f"[updated] {key}")
                    data[key] = new_df
        except Exception as e:
            print(f"[parse error] {key}: {e}")


def start_loop():
    def loop():
        while True:
            _read_all()
            time.sleep(REFRESH_SEC)
    threading.Thread(target=loop, daemon=True).start()


def get(key):
    with _lock:
        return data.get(key)
    
    
#%% end