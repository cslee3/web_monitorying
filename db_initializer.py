# -*- coding: utf-8 -*-
"""
db_initializer.py — MySQL DB 초기 테이블 생성 및 초기 데이터 적재
market_making 데이터베이스에 아래 두 테이블을 생성하고 Excel 파일에서 데이터를 INSERT.
  - POSITION      : 현재 포지션 (Position 시트)
  - PREV_POSITION : 전일 포지션 (T-1_Position 시트)

단독 실행용 스크립트 (python db_initializer.py).
MySQL 기동/종료는 db_manager.py 의 DBManager 사용.
"""
#%% header
import os
import mysql.connector
import xlwings as xw
import pandas as pd
import numpy as np

from db_manager import DBManager

try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

EXCEL_FILE = "GlobalMM_Realtime_Monitoring_V2.xlsb"


# ── Excel 시트 읽기 ────────────────────────────────────────────────────────
def _read_sheet(sheet_name, skiprows=2):
    fname = os.path.basename(EXCEL_FILE)
    try:
        wb = xw.books[fname]
    except Exception:
        wb = xw.Book(os.path.join(BASE_DIR, EXCEL_FILE))
    ws = wb.sheets[sheet_name]
    data = ws.used_range.value
    data = data[skiprows:]
    return pd.DataFrame(data)


# ── 값 변환 유틸리티 ───────────────────────────────────────────────────────
def to_val(v):
    """None, NaN, '-', '' → None. float → int."""
    if v is None:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if str(v).strip() in ('-', ''):
        return None
    if isinstance(v, float):
        return int(v)
    return v

def to_date(v):
    """None, float, NaN, '--', '-', '' → None. datetime/date → .date()."""
    if v is None:
        return None
    if isinstance(v, float):
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    if str(v).strip() in ('--', '-', ''):
        return None
    if hasattr(v, 'date'):
        return v.date()
    return v

def to_float(v):
    """None, NaN, '-', '' → None. 그 외 float() 시도."""
    if v is None:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if str(v).strip() in ('-', ''):
        return None
    try:
        return float(v)
    except Exception:
        return None


# ── POSITION 테이블 ────────────────────────────────────────────────────────
def init_position(cur, conn):
    cur.execute("""
    CREATE TABLE IF NOT EXISTS POSITION (
        id              INT AUTO_INCREMENT PRIMARY KEY,
        base_date       DATE,
        fund            VARCHAR(10),
        code            VARCHAR(10),
        stock_code      VARCHAR(10),
        prev_qty        BIGINT,
        curr_qty        BIGINT,
        buy_qty         BIGINT,
        buy_amt         BIGINT,
        sell_qty        BIGINT,
        sell_amt        BIGINT,
        stock_price     BIGINT,
        dividend        DOUBLE,
        expiry          DATE,
        strike          BIGINT,
        multiplier      INT,
        fin_prev_price  DOUBLE,
        fin_curr_price  DOUBLE,
        fin_eval        DOUBLE,
        fin_eval_pnl    DOUBLE,
        fin_trade_pnl   DOUBLE,
        theo_prev_price DOUBLE,
        theo_curr_price DOUBLE,
        theo_eval       DOUBLE,
        theo_eval_pnl   DOUBLE,
        theo_trade_pnl  DOUBLE,
        funding_cost    DOUBLE,
        delta           DOUBLE,
        gamma           DOUBLE,
        theta           DOUBLE,
        distortion      DOUBLE,
        created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
        INDEX idx_date      (base_date),
        INDEX idx_date_fund (base_date, fund),
        INDEX idx_date_isu  (base_date, code)
    )
    """)

    df = _read_sheet("Position", skiprows=2)
    rows = []
    for _, row in df.iterrows():
        if pd.isna(row.iloc[0]) or to_date(row.iloc[0]) is None:
            continue
        rows.append((
            to_date(row.iloc[0]),    # base_date
            to_val(row.iloc[1]),     # fund
            to_val(row.iloc[2]),     # code
            to_val(row.iloc[3]),     # stock_code
            to_val(row.iloc[4]),     # prev_qty
            to_val(row.iloc[5]),     # curr_qty
            to_val(row.iloc[6]),     # buy_qty
            to_val(row.iloc[7]),     # buy_amt
            to_val(row.iloc[8]),     # sell_qty
            to_val(row.iloc[9]),     # sell_amt
            to_val(row.iloc[10]),    # stock_price
            to_float(row.iloc[11]),  # dividend
            to_date(row.iloc[12]),   # expiry
            to_val(row.iloc[13]),    # strike
            to_val(row.iloc[14]),    # multiplier
            to_float(row.iloc[15]),  # fin_prev_price
            to_float(row.iloc[16]),  # fin_curr_price
            to_float(row.iloc[17]),  # fin_eval
            to_float(row.iloc[18]),  # fin_eval_pnl
            to_float(row.iloc[19]),  # fin_trade_pnl
            to_float(row.iloc[20]),  # theo_prev_price
            to_float(row.iloc[21]),  # theo_curr_price
            to_float(row.iloc[22]),  # theo_eval
            to_float(row.iloc[23]),  # theo_eval_pnl
            to_float(row.iloc[24]),  # theo_trade_pnl
            to_float(row.iloc[25]),  # funding_cost
            to_float(row.iloc[26]),  # delta
            to_float(row.iloc[27]),  # gamma
            to_float(row.iloc[28]),  # theta
            to_float(row.iloc[29]),  # distortion
        ))

    cur.executemany("""
    INSERT INTO POSITION
        (base_date, fund, code, stock_code,
         prev_qty, curr_qty, buy_qty, buy_amt, sell_qty, sell_amt,
         stock_price, dividend, expiry, strike, multiplier,
         fin_prev_price, fin_curr_price, fin_eval, fin_eval_pnl, fin_trade_pnl,
         theo_prev_price, theo_curr_price, theo_eval, theo_eval_pnl, theo_trade_pnl,
         funding_cost, delta, gamma, theta, distortion)
    VALUES (%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s)
    """, rows)
    conn.commit()
    print(f"POSITION INSERT 완료: {len(rows)}행")


# ── PREV_POSITION 테이블 ───────────────────────────────────────────────────
def init_prev_position(cur, conn):
    cur.execute("""
    CREATE TABLE IF NOT EXISTS PREV_POSITION (
        id              INT AUTO_INCREMENT PRIMARY KEY,
        base_date       DATE,
        fund            VARCHAR(10),
        code            VARCHAR(10),
        stock_code      VARCHAR(10),
        prev_qty        BIGINT,
        curr_qty        BIGINT,
        buy_qty         BIGINT,
        buy_amt         BIGINT,
        sell_qty        BIGINT,
        sell_amt        BIGINT,
        stock_price     BIGINT,
        dividend        DOUBLE,
        expiry          DATE,
        strike          BIGINT,
        multiplier      INT,
        fin_prev_price  DOUBLE,
        fin_curr_price  DOUBLE,
        fin_eval        DOUBLE,
        fin_eval_pnl    DOUBLE,
        fin_trade_pnl   DOUBLE,
        theo_prev_price DOUBLE,
        theo_curr_price DOUBLE,
        theo_eval       DOUBLE,
        theo_eval_pnl   DOUBLE,
        theo_trade_pnl  DOUBLE,
        funding_cost    DOUBLE,
        delta           DOUBLE,
        gamma           DOUBLE,
        month           DOUBLE,
        created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
        INDEX idx_date      (base_date),
        INDEX idx_date_fund (base_date, fund),
        INDEX idx_date_isu  (base_date, code)
    )
    """)

    df = _read_sheet("T-1_Position", skiprows=2)
    rows = []
    for _, row in df.iterrows():
        if pd.isna(row.iloc[0]) or to_date(row.iloc[0]) is None:
            continue
        rows.append((
            to_date(row.iloc[0]),    # base_date
            to_val(row.iloc[1]),     # fund
            to_val(row.iloc[2]),     # code
            to_val(row.iloc[3]),     # stock_code
            to_val(row.iloc[4]),     # prev_qty
            to_val(row.iloc[5]),     # curr_qty
            to_val(row.iloc[6]),     # buy_qty
            to_val(row.iloc[7]),     # buy_amt
            to_val(row.iloc[8]),     # sell_qty
            to_val(row.iloc[9]),     # sell_amt
            to_val(row.iloc[10]),    # stock_price
            to_float(row.iloc[11]),  # dividend
            to_date(row.iloc[12]),   # expiry
            to_val(row.iloc[13]),    # strike
            to_val(row.iloc[14]),    # multiplier
            to_float(row.iloc[15]),  # fin_prev_price
            to_float(row.iloc[16]),  # fin_curr_price
            to_float(row.iloc[17]),  # fin_eval
            to_float(row.iloc[18]),  # fin_eval_pnl
            to_float(row.iloc[19]),  # fin_trade_pnl
            to_float(row.iloc[20]),  # theo_prev_price
            to_float(row.iloc[21]),  # theo_curr_price
            to_float(row.iloc[22]),  # theo_eval
            to_float(row.iloc[23]),  # theo_eval_pnl
            to_float(row.iloc[24]),  # theo_trade_pnl
            to_float(row.iloc[25]),  # funding_cost
            to_float(row.iloc[26]),  # delta
            to_float(row.iloc[27]),  # gamma
            to_float(row.iloc[28]),  # month
        ))

    cur.executemany("""
    INSERT INTO PREV_POSITION
        (base_date, fund, code, stock_code,
         prev_qty, curr_qty, buy_qty, buy_amt, sell_qty, sell_amt,
         stock_price, dividend, expiry, strike, multiplier,
         fin_prev_price, fin_curr_price, fin_eval, fin_eval_pnl, fin_trade_pnl,
         theo_prev_price, theo_curr_price, theo_eval, theo_eval_pnl, theo_trade_pnl,
         funding_cost, delta, gamma, month)
    VALUES (%s,%s,%s,%s,
            %s,%s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s,%s,
            %s,%s,%s,%s)
    """, rows)
    conn.commit()
    print(f"PREV_POSITION INSERT 완료: {len(rows)}행")


# ── SSF_HISTORY 테이블 ────────────────────────────────────────────────────
def init_ssf_history(cur):
    cur.execute("""
    CREATE TABLE IF NOT EXISTS ssf_history (
        ts         DOUBLE       NOT NULL,
        stock      VARCHAR(20)  NOT NULL,
        stock_name VARCHAR(50),
        last_price DOUBLE,
        change_pct DOUBLE,
        amt_mil    DOUBLE,
        amt_shr    DOUBLE,
        sprd_score DOUBLE,
        mm_spread  VARCHAR(30),
        shares     DOUBLE,
        diff       DOUBLE,
        vol_mil    DOUBLE,
        delta_shr  DOUBLE,
        delta_mil  DOUBLE,
        theo_price DOUBLE,
        theo_basis DOUBLE,
        ex1_bsp    DOUBLE,
        ex1_ssp    DOUBLE,
        ex2_basis  DOUBLE,
        ex2_bsp    DOUBLE,
        ex2_ssp    DOUBLE,
        mtm_pnl    DOUBLE,
        theo_pnl   DOUBLE,
        PRIMARY KEY (ts, stock),
        INDEX idx_stock (stock),
        INDEX idx_ts    (ts)
    )
    """)
    print("ssf_history 테이블 준비 완료")


# ── 메인 ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    db = DBManager()
    conn = db.connect(database="market_making")
    cur  = conn.cursor()

    init_position(cur, conn)
    init_prev_position(cur, conn)
    init_ssf_history(cur)

    cur.close()
    conn.close()
    print("전체 완료")
