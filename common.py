# -*- coding: utf-8 -*-
"""
common.py — 공유 상수 및 유틸리티
reader_ssf.py / reader_sso.py 등에서 import하여 사용.
"""

import os
import datetime

try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

BOOK_NAME = "GlobalMM_Realtime_Monitoring_V2.xlsb"


# ── 주식 호가 틱 단위 (KRX 기준, STOCK 타입) ──────────────────────────────
# 출처: workspace_ref/XCommon/XLibraries.h getHoga()
# (price_from, price_to, tick) — price_to는 미만(exclusive)
STOCK_TICK_TABLE = [
    (        0,    2_000,    1),
    (    2_000,    5_000,    5),
    (    5_000,   20_000,   10),
    (   20_000,   50_000,   50),
    (   50_000,  200_000,  100),
    (  200_000,  500_000,  500),
    (  500_000, 9_999_999, 1_000),
]

def get_stock_tick(price: float) -> int:
    """주가에 해당하는 호가 틱 단위 반환."""
    for lo, hi, tick in STOCK_TICK_TABLE:
        if lo <= price < hi:
            return tick
    return 1_000


# ── Excel COM 에러값 판별 ──────────────────────────────────────────────────

def _is_xl_error(v):
    """Excel COM 에러코드 여부 (0x800A0xxx 패턴, 약 -2,147,500,000 ~ -2,146,000,000)."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        iv = int(v)
        return -2_147_500_000 <= iv <= -2_146_000_000
    return False


# ── 값 포매터 ──────────────────────────────────────────────────────────────

def _pct(v):
    """소수 → 퍼센트 문자열. ex) 0.0123 → '1.23%'"""
    if v is None: return ""
    if _is_xl_error(v): return ""
    try: return f"{float(v)*100:.2f}%"
    except: return str(v)


def _num(v, d=0):
    """숫자 → 천단위 콤마 + 소수점 d자리. ex) 1234567.8 → '1,234,568'"""
    if v is None or v == "": return ""
    if _is_xl_error(v): return ""
    try: return f"{v:,.{d}f}"
    except: return str(v)


def _date(v):
    """datetime → 'YYYY-MM-DD' 문자열."""
    if v is None: return ""
    if isinstance(v, datetime.datetime): return v.strftime("%Y-%m-%d")
    return str(v)


def _rgb(color):
    """xlwings color tuple (r,g,b) → CSS rgb() 문자열."""
    if isinstance(color, tuple) and len(color) == 3:
        return f"rgb({color[0]},{color[1]},{color[2]})"
    return ""


# ── ColorScale 배경색 계산 ─────────────────────────────────────────────────

def _colorscale_rank(vals: list) -> dict:
    """
    값 리스트를 순위 기반 배경색 dict로 변환. {float_val: rgb_str}
    음수: 가장 음수(1위) → 진한 파랑, 덜 음수 → 연한 파랑
    양수: 가장 양수(1위) → 진한 빨강, 덜 양수 → 연한 빨강
    동일 값은 동일 색.
    """
    negs = sorted(set(v for v in vals if v < 0))   # 오름차순: 가장 음수가 index 0
    pos  = sorted(set(v for v in vals if v > 0))   # 오름차순: 가장 양수가 index -1
    result = {}
    n = len(negs)
    for i, v in enumerate(negs):
        t = (n - i) / n   # index 0(최소) → t=1(진), index n-1(최대) → t=1/n(연)
        result[v] = f"rgb({int(255-t*(255-111))},{int(255-t*(255-168))},{int(255-t*(255-220))})"
    n = len(pos)
    for i, v in enumerate(pos):
        t = (i + 1) / n   # index 0(최소) → t=1/n(연), index n-1(최대) → t=1(진)
        result[v] = f"rgb({int(255-t*(255-224))},{int(255-t*(255-102))},{int(255-t*(255-102))})"
    return result


def _colorscale_bg(val, vmin, vmax):
    """
    val의 min/max 대비 상대 위치로 파랑→흰→빨강 배경색 계산.
    음수 구간: blue(#6fa8dc) → white(#ffffff)
    양수 구간: white(#ffffff) → red(#e06666)
    """
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""

    if v < 0:
        if vmin is None or vmin >= 0:
            return ""
        t = max(0.0, min(1.0, v / vmin))
        r = int(255 - t * (255 - 111))
        g = int(255 - t * (255 - 168))
        b = int(255 - t * (255 - 220))
        return f"rgb({r},{g},{b})"
    elif v > 0:
        if vmax is None or vmax <= 0:
            return ""
        t = max(0.0, min(1.0, v / vmax))
        r = int(255 - t * (255 - 224))
        g = int(255 - t * (255 - 102))
        b = int(255 - t * (255 - 102))
        return f"rgb({r},{g},{b})"
    else:
        return ""
