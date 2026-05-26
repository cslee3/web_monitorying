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

    if vmin is None or vmax is None or vmin == vmax:
        return ""

    if v < 0:
        if vmin >= 0:
            return ""
        t = max(0.0, min(1.0, v / vmin))
        r = int(255 - t * (255 - 111))
        g = int(255 - t * (255 - 168))
        b = int(255 - t * (255 - 220))
        return f"rgb({r},{g},{b})"
    elif v > 0:
        if vmax <= 0:
            return ""
        t = max(0.0, min(1.0, v / vmax))
        r = int(255 - t * (255 - 224))
        g = int(255 - t * (255 - 102))
        b = int(255 - t * (255 - 102))
        return f"rgb({r},{g},{b})"
    else:
        return ""
