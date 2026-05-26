# -*- coding: utf-8 -*-
"""
reader_common.py — 공유 값 포매터
reader_ssf.py / reader_sso.py 등에서 import하여 사용.
"""

import datetime


def _pct(v):
    """소수 → 퍼센트 문자열. ex) 0.0123 → '1.23%'"""
    if v is None: return ""
    try: return f"{float(v)*100:.2f}%"
    except: return str(v)


def _num(v, d=0):
    """숫자 → 천단위 콤마 + 소수점 d자리. ex) 1234567.8 → '1,234,568'"""
    if v is None or v == "": return ""
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
