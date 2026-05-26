# -*- coding: utf-8 -*-
"""
db_query_example.py — MySQL DB 조회 예시
대상 서버: 10.115.31.99
데이터베이스: market_making
테이블: POSITION, PREV_POSITION
"""

import mysql.connector
from datetime import date

# ── 연결 설정 ──────────────────────────────────────────────────────────────
conn = mysql.connector.connect(
    host="10.115.31.99",
    port=3306,
    user="root",
    password="",
    database="market_making",
    charset="utf8mb4",
)
cur = conn.cursor(dictionary=True)   # dictionary=True → 결과를 dict로 반환


# ── 예시 1: POSITION 전체 조회 (최근 base_date 기준) ───────────────────────
print("=" * 60)
print("[예시 1] POSITION — 최근 base_date 전체 조회")
print("=" * 60)

cur.execute("""
    SELECT *
    FROM POSITION
    WHERE base_date = (SELECT MAX(base_date) FROM POSITION)
    ORDER BY fund, code
""")
rows = cur.fetchall()
for row in rows:
    print(row)
print(f"→ {len(rows)}행\n")


# ── 예시 2: 특정 날짜 + fund 필터 ──────────────────────────────────────────
print("=" * 60)
print("[예시 2] POSITION — 특정 날짜 + fund 필터")
print("=" * 60)

TARGET_DATE = date(2026, 5, 20)   # 조회할 날짜
TARGET_FUND = "SSF"               # 펀드명 (실제 값으로 변경)

cur.execute("""
    SELECT base_date, fund, code, stock_code,
           curr_qty, fin_eval, fin_eval_pnl, fin_trade_pnl,
           delta, gamma, theta
    FROM POSITION
    WHERE base_date = %s
      AND fund      = %s
    ORDER BY code
""", (TARGET_DATE, TARGET_FUND))
rows = cur.fetchall()
for row in rows:
    print(row)
print(f"→ {len(rows)}행\n")


# ── 예시 3: PREV_POSITION 전체 조회 ────────────────────────────────────────
print("=" * 60)
print("[예시 3] PREV_POSITION — 최근 base_date 전체 조회")
print("=" * 60)

cur.execute("""
    SELECT *
    FROM PREV_POSITION
    WHERE base_date = (SELECT MAX(base_date) FROM PREV_POSITION)
    ORDER BY fund, code
""")
rows = cur.fetchall()
for row in rows:
    print(row)
print(f"→ {len(rows)}행\n")


# ── 예시 4: POSITION vs PREV_POSITION 비교 (curr_qty 변동) ─────────────────
print("=" * 60)
print("[예시 4] 당일 vs 전일 curr_qty 변동")
print("=" * 60)

cur.execute("""
    SELECT p.fund,
           p.code,
           p.curr_qty            AS today_qty,
           pp.curr_qty           AS prev_qty,
           p.curr_qty - pp.curr_qty AS qty_change
    FROM POSITION p
    JOIN PREV_POSITION pp
      ON p.fund = pp.fund AND p.code = pp.code
    WHERE p.base_date  = (SELECT MAX(base_date) FROM POSITION)
      AND pp.base_date = (SELECT MAX(base_date) FROM PREV_POSITION)
    ORDER BY p.fund, p.code
""")
rows = cur.fetchall()
for row in rows:
    print(row)
print(f"→ {len(rows)}행\n")


# ── 예시 5: fund별 fin_eval_pnl 합계 ───────────────────────────────────────
print("=" * 60)
print("[예시 5] fund별 fin_eval_pnl 합계 (최근 base_date)")
print("=" * 60)

cur.execute("""
    SELECT fund,
           SUM(fin_eval_pnl)   AS total_eval_pnl,
           SUM(fin_trade_pnl)  AS total_trade_pnl,
           SUM(delta)          AS total_delta
    FROM POSITION
    WHERE base_date = (SELECT MAX(base_date) FROM POSITION)
    GROUP BY fund
    ORDER BY fund
""")
rows = cur.fetchall()
for row in rows:
    print(row)
print(f"→ {len(rows)}행\n")


# ── 종료 ───────────────────────────────────────────────────────────────────
cur.close()
conn.close()
print("조회 완료.")
