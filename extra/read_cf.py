# -*- coding: utf-8 -*-
"""
read_cf.py — DashBoard / Option_DashBoard 탭의 조건부 서식 규칙 출력
실행: python read_cf.py
"""
import pythoncom
import win32com.client
import win32com.client.dynamic as _dyn
import os

from reader_ssf import BOOK_NAME, BASE_DIR

TYPE_NAMES = {
    1: "CellValue",
    2: "Expression",
    3: "ColorScale",
    4: "DataBar",
    5: "IconSet",
    6: "Top10",
    7: "UniqueValues",
    8: "Text",
    9: "Blanks",
    10: "NoBlanks",
    11: "Errors",
    12: "NoErrors",
    13: "TimePeriod",
    14: "AboveAverage",
}

OP_NAMES = {
    1: "Between", 2: "NotBetween", 3: "Equal", 4: "NotEqual",
    5: "GreaterThan", 6: "LessThan", 7: "GreaterEqual", 8: "LessEqual",
}

def rgb_from_color(color_val):
    if color_val is None:
        return ""
    color_val = int(color_val)   # COM이 float으로 반환하는 경우 대응
    if color_val < 0:
        return ""
    r = color_val & 0xFF
    g = (color_val >> 8) & 0xFF
    b = (color_val >> 16) & 0xFF
    return f"rgb({r},{g},{b})"

def read_cf(ws, sheet_name):
    print(f"\n{'='*60}")
    print(f"  {sheet_name}")
    print(f"{'='*60}")

    # FormatConditions는 Worksheet가 아닌 Range 객체의 속성
    # ws.Cells(전체 셀)를 dynamic dispatch로 감싸서 접근
    try:
        cells = _dyn.Dispatch(ws.Cells._oleobj_)
        fc = cells.FormatConditions
    except Exception as e:
        print(f"  FormatConditions 접근 실패: {e}")
        return

    if fc.Count == 0:
        print("  (조건부 서식 없음)")
        return

    for i in range(1, fc.Count + 1):
        rule_raw = fc.Item(i)
        # rule도 dynamic dispatch로 감싸서 서브 속성 접근
        try:
            rule = _dyn.Dispatch(rule_raw._oleobj_)
        except Exception:
            rule = rule_raw

        try:
            rng = rule.AppliesTo.Address
        except Exception as e:
            rng = f"(범위 읽기 실패: {e})"
        try:
            t = rule.Type
        except Exception:
            t = -1
        type_name = TYPE_NAMES.get(t, f"Type{t}")

        print(f"\n[{i}] 범위: {rng}  /  타입: {type_name}")

        try:
            if t == 1:  # CellValue
                op = OP_NAMES.get(rule.Operator, f"Op{rule.Operator}")
                f1 = rule.Formula1
                f2 = getattr(rule, "Formula2", "")
                print(f"     조건: {op}  F1={f1}  F2={f2}")
            elif t == 2:  # Expression
                print(f"     수식: {rule.Formula1}")
        except Exception as e:
            print(f"     조건 읽기 실패: {e}")

        # ColorScale/DataBar/IconSet는 Font/Interior/Borders가 없음
        if t in (3, 4, 5):
            try:
                csc = rule.ColorScaleCriteria
                for ci in range(1, csc.Count + 1):
                    c = csc.Item(ci)
                    ctype = c.Type
                    cval  = getattr(c, "Value", "")
                    color = rgb_from_color(c.FormatColor.Color)
                    print(f"     ColorScale[{ci}]: type={ctype}  val={cval}  color={color}")
            except Exception:
                pass
            continue

        try:
            fmt = rule.Font
            fc_color = fmt.Color
            bold     = fmt.Bold
            italic   = fmt.Italic
            print(f"     글꼴: color={rgb_from_color(fc_color)}  bold={bold}  italic={italic}")
        except Exception as e:
            print(f"     글꼴 읽기 실패: {e}")

        try:
            interior = rule.Interior
            bg       = interior.Color
            pattern  = interior.Pattern
            if int(pattern) != -4142:  # xlNone
                print(f"     배경: color={rgb_from_color(bg)}")
            else:
                print(f"     배경: (없음)")
        except Exception as e:
            print(f"     배경 읽기 실패: {e}")

        try:
            borders = rule.Borders
            for bi in range(1, 5):
                b = borders.Item(bi)
                if int(b.LineStyle) != -4142:
                    print(f"     테두리[{bi}]: color={rgb_from_color(b.Color)}  style={b.LineStyle}")
        except Exception as e:
            print(f"     테두리 읽기 실패: {e}")


def main():
    pythoncom.CoInitialize()
    path = os.path.join(BASE_DIR, BOOK_NAME)
    try:
        wb = win32com.client.GetObject(path)
    except Exception:
        xl = win32com.client.Dispatch("Excel.Application")
        wb = xl.Workbooks.Open(path)

    import io, sys
    buf = io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = buf

    read_cf(wb.Sheets("DashBoard"),        "DashBoard")
    read_cf(wb.Sheets("Option_DashBoard"), "Option_DashBoard")

    sys.stdout = old_stdout
    output = buf.getvalue()
    print(output)

    out_path = os.path.join(BASE_DIR, "cf_rules.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(output)
    print(f"\n→ 저장 완료: {out_path}")

if __name__ == "__main__":
    main()
