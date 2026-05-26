# -*- coding: utf-8 -*-
"""
database_setting.py — MySQL DB 초기 테이블 생성 및 초기 데이터 적재
market_making 데이터베이스에 아래 두 테이블을 생성하고 Excel 샘플 파일에서 데이터를 INSERT.
  - POSITION      : 현재 포지션 (Position 시트)
  - PREV_POSITION : 전일 포지션 (T-1_Position 시트)

단독 실행용 스크립트 (python database_setting.py).
"""

import os
import time
import subprocess
import mysql.connector
import pandas as pd
import numpy as np


# ── MySQL 경로 ───────────────────────────────────────────────────────────
# 프로젝트와 동등한 위치: D:\OneDrive\mysql
MYSQL_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mysql")
)
MYSQLD_BIN = os.path.join(MYSQL_DIR, "bin", "mysqld.exe")

def _ensure_mysql():
    """mysqld가 실행 중이지 않으면 백그라운드로 기동 후 대기."""
    import socket
    try:
        socket.create_connection(("localhost", 3306), timeout=1).close()
        return  # 이미 실행 중
    except OSError:
        pass

    if not os.path.exists(MYSQLD_BIN):
        print(f"[경고] mysqld.exe 를 찾을 수 없습니다: {MYSQLD_BIN}")
        return

    print(f"MySQL 기동 중 ... ({MYSQLD_BIN})")
    subprocess.Popen(
        [MYSQLD_BIN, "--console"],
        cwd=os.path.join(MYSQL_DIR, "bin"),
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )
    for _ in range(15):          # 최대 15초 대기
        time.sleep(1)
        try:
            socket.create_connection(("localhost", 3306), timeout=1).close()
            print("MySQL 기동 완료.")
            return
        except OSError:
            pass
    print("[경고] MySQL 기동 확인 실패 — 수동으로 확인하세요.")

_ensure_mysql()


# ── 값 변환 유틸리티 ────────────────────────────────────────────────────

def to_val(v):
    """
    일반 값 정제 함수.
    None, NaN, '-', '' → None 반환.
    float → int 변환 (Excel 숫자가 float으로 읽히는 경우 처리).
    """
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
    """
    날짜 값 정제 함수.
    None, float, NaN, '--', '-', '' → None 반환.
    datetime/date 객체 → .date() 반환.
    """
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
    """
    실수 값 정제 함수.
    None, NaN, '-', '' → None 반환.
    그 외 float() 변환 시도, 실패하면 None 반환.
    """
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


# ── DB 연결 ──────────────────────────────────────────────────────────────
conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="market_making"
)
cur = conn.cursor()


# ── POSITION 테이블 생성 및 데이터 적재 ────────────────────────────────

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

# Excel 샘플 파일의 Position 시트에서 데이터 읽기 (헤더 2행 skip)
position_data_init = pd.read_excel(
    "./GlobalMM_Realtime_Monitoring_V2 2026 04 30_sample.xlsx",
    sheet_name="Position", skiprows=2, header=None)

rows = []
for _, row in position_data_init.iterrows():
    if pd.isna(row.iloc[0]):   # base_date가 없는 행 skip
        continue
    if to_date(row.iloc[0]) is None:
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


# ── T-1_POSITION 테이블 생성 및 데이터 적재 ────────────────────────────

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
    month           DOUBLE,           -- T-1은 theta 대신 month 컬럼 사용
    created_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_date      (base_date),
    INDEX idx_date_fund (base_date, fund),
    INDEX idx_date_isu  (base_date, code)
)
""")

# Excel 샘플 파일의 T-1_Position 시트에서 데이터 읽기
position_data_init = pd.read_excel(
    "./GlobalMM_Realtime_Monitoring_V2 2026 04 30_sample.xlsx",
    sheet_name="T-1_Position", skiprows=2, header=None)

rows = []
for _, row in position_data_init.iterrows():
    if pd.isna(row.iloc[0]):   # base_date가 없는 행 skip
        continue
    if to_date(row.iloc[0]) is None:
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


# ── 종료 ────────────────────────────────────────────────────────────────
cur.close()
conn.close()
print("전체 완료")
