import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime

st.set_page_config(
    page_title="Innovation_Pascal@",
    page_icon="📊",
    layout="wide"
)

# ─────────────────────────────────────────────
# 색상 테마
# ─────────────────────────────────────────────

COLOR_TITLE = "#111D37"
COLOR_BUDGET = "#223D72"
COLOR_BUDGET_SUB = "#345CAD"
COLOR_EXECUTION = "#5F84CF"
COLOR_EXECUTION_SUB = "#9AB1E0"
COLOR_COMPARE = "#7186A8"
COLOR_EXTRA = "#6FA6B8"
COLOR_BACKGROUND = "#D0D9E8"

BLUE_PALETTE = [
    COLOR_BUDGET,
    COLOR_BUDGET_SUB,
    COLOR_EXECUTION,
    COLOR_EXECUTION_SUB,
    COLOR_COMPARE,
    COLOR_EXTRA,
    COLOR_BACKGROUND,
]

st.markdown(
    f"""
    <style>
    h1, h2, h3 {{
        color: {COLOR_TITLE};
    }}

    [data-testid="stMetric"] {{
        background-color: {COLOR_BACKGROUND};
        padding: 14px;
        border-radius: 8px;
    }}

    [data-testid="stMetricLabel"] {{
        color: {COLOR_TITLE};
    }}

    [data-testid="stMetricValue"] {{
        color: {COLOR_TITLE};
        font-weight: 700;
    }}

    [data-testid="stSidebar"] {{
        border-right: 1px solid {COLOR_BACKGROUND};
    }}

    div[data-testid="stDataFrame"] {{
        border: 1px solid {COLOR_BACKGROUND};
        border-radius: 6px;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

st.title("혁신지원사업 예산 대시보드")
st.caption(
    "예산 파일과 집행 파일을 함께 올리면 "
    "월별 집행액·예산 대비 집행률을 확인할 수 있습니다."
)


# ─────────────────────────────────────────────
# 공통 함수
# ─────────────────────────────────────────────

def norm(x):
    return re.sub(
        r"\s+",
        "",
        str(x).replace("\n", "")
    ).strip().lower()


def find_col(columns, candidates):
    cmap = {norm(c): c for c in columns}

    for c in candidates:
        if norm(c) in cmap:
            return cmap[norm(c)]
