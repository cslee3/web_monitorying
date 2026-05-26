# -*- coding: utf-8 -*-
"""
test_tabs.py — Excel 탭별 파싱 결과를 HTML로 출력하는 개발용 테스트 도구 (구버전)
before/ 폴더의 reader.py, renderer.py에 의존하며 현재 메인 시스템에서는 사용하지 않음.
탭 선택 → 파싱 → HTML 생성 → 브라우저 자동 오픈.
"""

import os
import webbrowser
import renderer
from reader import (_get_wb, _parse_dashboard_MM, _parse_dashboard_arb,
                    _parse_option_dashboard, _parse_live_orders,
                    _parse_position, _parse_daily)


# ── 탭 정의 ──────────────────────────────────────────────────────────────
# key: 탭 식별자, value: (화면 라벨, [(파서함수, PnL색상 컬럼 목록), ...])
# 한 탭에 여러 파서를 지정하면 HTML이 순서대로 연결됨
TABS = {
    "dashboard": ("DashBoard", [
        (_parse_dashboard_MM,  ["MTM_PnL", "Theo_PnL"]),
        (_parse_dashboard_arb, ["MTM_PnL", "Theo_PnL"]),
    ]),
    "option_dashboard": ("Option DB", [
        (_parse_option_dashboard, ["Delta", "%Gamma", "Theo_PnL", "MTM_PnL"]),
    ]),
    "live_orders": ("Live Orders", [(_parse_live_orders, [])]),
    "position":   ("Position",    [(_parse_position,    [])]),
    "daily":      ("Daily",       [(_parse_daily,       [])]),
}


def select_tab():
    """
    콘솔에서 탭 번호 또는 이름을 입력받아 처리할 탭 키 목록 반환.
    'all' 또는 마지막 번호 입력 시 전체 탭 반환.
    """
    keys = list(TABS.keys())
    for i, key in enumerate(keys, 1):
        print(f"  {i}. {key}")
    print(f"  {len(keys)+1}. all")
    val = input("번호 또는 탭 이름: ").strip()
    if val == str(len(keys) + 1) or val == "all":
        return keys
    if val.isdigit() and 1 <= int(val) <= len(keys):
        return [keys[int(val) - 1]]
    if val in TABS:
        return [val]
    raise ValueError(f"알 수 없는 입력: {val}")


def main(keys):
    """
    선택된 탭 키 목록을 순서대로 파싱하여 test_output/ 폴더에 HTML 저장 후 브라우저로 오픈.
    파싱 실패 시 에러 출력 후 다음 탭 처리 계속.
    """
    wb = _get_wb()
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_output")
    os.makedirs(out_dir, exist_ok=True)

    for key in keys:
        label, parsers = TABS[key]
        print(f"[{key}] 파싱 중...")
        try:
            content = ""
            for parse_fn, color_cols in parsers:
                df = parse_fn(wb)
                print(f"  {parse_fn.__name__}: {len(df)}행 × {len(df.columns)}열")
                content += renderer.df_to_html(df, color_cols=color_cols)

            html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>{label}</title>
  {renderer.BASE_STYLE}
</head>
<body>
  <h1>{label}</h1>
  {content}
</body>
</html>"""
            path = os.path.join(out_dir, f"{key}.html")
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            webbrowser.open("file:///" + os.path.abspath(path).replace("\\", "/"))
            print(f"  열기: {path}")
        except Exception as e:
            print(f"  ERROR: {e}")


if __name__ == "__main__":
    keys = select_tab()
    main(keys)
