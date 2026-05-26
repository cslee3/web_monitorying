import time
import streamlit as st
import pandas as pd
import reader
from config import REFRESH_SEC

# ── 페이지 설정 ────────────────────────────────────────────────
st.set_page_config(
    page_title="GlobalMM Monitor",
    page_icon="📈",
    layout="wide",
)

def _safe_df(df):
    """Arrow 변환 실패 컬럼을 문자열로 변환"""
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).replace({"None": "", "nan": ""})
    return df

def show_df(df, color_cols=None, height=400):
    if df is None or df.empty:
        st.info("데이터 없음")
        return
    st.dataframe(_safe_df(df), use_container_width=True, height=height)


# ── 백그라운드 루프 시작 (최초 1회) ───────────────────────────
if "loop_started" not in st.session_state:
    reader.start_loop()
    st.session_state.loop_started = True

# ── 헤더 ──────────────────────────────────────────────────────
st.title("📈 GlobalMM Realtime Monitor")
st.caption(f"자동 새로고침: {REFRESH_SEC}초마다")

# ── 탭 구성 ───────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 DashBoard",
    "⚙️ Option DB",
    "📋 Live Orders",
    "📌 Position",
    "📅 Daily",
])

# ── DashBoard ─────────────────────────────────────────────────
with tab1:
    pnl_cols = ["MTM PnL", "Theo PnL"]

    st.subheader("MM")
    show_df(reader.get("dashboard_MM"), color_cols=pnl_cols)

    with st.expander("Option Dashboard"):
        show_df(
            reader.get("option_dashboard"),
            color_cols=["Delta", "%Gamma", "Theo_PnL", "MTM_PnL"],
            height=400,
        )

    st.subheader("Arb")
    show_df(reader.get("dashboard_arb"), color_cols=pnl_cols)

# ── Option Dashboard ──────────────────────────────────────────
with tab2:
    show_df(
        reader.get("option_dashboard"),
        color_cols=["Delta", "%Gamma", "Theo_PnL", "MTM_PnL"],
    )

# ── Live Orders ───────────────────────────────────────────────
with tab3:
    show_df(reader.get("live_orders"))

# ── Position ──────────────────────────────────────────────────
with tab4:
    show_df(reader.get("position"), height=600)

# ── Daily ─────────────────────────────────────────────────────
with tab5:
    show_df(reader.get("daily"))

# ── 자동 새로고침 ──────────────────────────────────────────────
time.sleep(REFRESH_SEC)
st.rerun()