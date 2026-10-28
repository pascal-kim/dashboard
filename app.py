import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# =========================================================
# 1. 페이지 설정 및 색상
# =========================================================
st.set_page_config(
    page_title="사업예산 집행현황 대시보드",
    page_icon="📊",
    layout="wide",
)

# CMYK 기준 근사 HEX 색상
COLOR_MAIN = "#002C9E"       # C100 M72 Y0 K38
COLOR_BLUE = "#1766FF"       # C91 M60 Y0 K0
COLOR_DARK = "#363627"       # C0 M0 Y28 K79

BLUE_PALETTE = [
    "#002C9E",
    "#1766FF",
    "#356FC7",
    "#6795D8",
    "#9ABBE8",
    "#C5D8F2",
    "#363627",
]

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    [data-testid="stMetricValue"] {
        color: #002C9E;
        font-weight: 700;
    }
    h1, h2, h3 {
        color: #002C9E;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("사업예산 집행현황 대시보드")
st.caption("예산 대비 집행실적을 부서·과제·자금·비목별로 확인합니다.")

# =========================================================
# 2. 공통 함수
# =========================================================
def clean_text(value):
    """문자열 공백 및 결측값 정리."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key(series):
    """연결 키의 공백 및 숫자형 문자열 차이를 완화."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def find_column(df, candidates, required=True, label=""):
    """후보 목록에서 실제 열 이름 찾기."""
    columns = {str(c).strip(): c for c in df.columns}

    for candidate in candidates:
        if candidate in columns:
            return columns[candidate]

    # 공백을 제거한 이름으로도 검색
    compact_columns = {
        str(c).strip().replace(" ", ""): c for c in df.columns
    }
    for candidate in candidates:
        compact = candidate.replace(" ", "")
        if compact in compact_columns:
            return compact_columns[compact]

    if required:
        raise ValueError(
            f"'{label or candidates[0]}' 열을 찾을 수 없습니다.\n"
            f"현재 열 목록: {list(df.columns)}"
        )
    return None


def to_number(series):
    """쉼표와 통화기호가 포함된 금액도 숫자로 변환."""
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₩", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.strip()
        .replace({"": None, "nan": None, "None": None}),
        errors="coerce",
    ).fillna(0)


def format_money(value):
    """금액 표시용 천 단위 쉼표."""
    if pd.isna(value):
        return ""
    return f"{value:,.0f}"


def display_table(df, money_columns=None, height=400):
    """금액 열을 오른쪽 정렬하고 천 단위 쉼표 표시."""
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return

    money_columns = money_columns or []
    money_columns = [c for c in money_columns if c in df.columns]

    styler = df.style

    if money_columns:
        styler = styler.format(
            {col: format_money for col in money_columns},
            na_rep="",
        )
        styler = styler.set_properties(
            subset=money_columns,
            **{"text-align": "right"},
        )

    st.dataframe(
        styler,
        use_container_width=True,
        height=height,
        hide_index=True,
    )


def style_amount_axis(fig, horizontal=False):
    """그래프 금액 축 및 레이아웃 공통 설정."""
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Arial", size=12, color=COLOR_DARK),
        title_font=dict(size=17, color=COLOR_MAIN),
        legend_title_text="",
        margin=dict(l=20, r=20, t=55, b=20),
    )

    if horizontal:
        fig.update_xaxes(tickformat=",.0f", separatethousands=True)
    else:
        fig.update_yaxes(tickformat=",.0f", separatethousands=True)

    return fig


def style_bar_labels(fig):
    fig.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        cliponaxis=False,
    )
    return fig


def read_excel_auto_header(uploaded_file, required_headers):
    """집행 파일에서 필요한 열이 있는 헤더 행을 자동 탐색."""
    raw = pd.read_excel(uploaded_file, header=None)

    required_compact = {
        x.replace(" ", "") for x in required_headers
    }

    for row_idx in range(min(len(raw), 15)):
        values = [
            clean_text(v).replace(" ", "")
            for v in raw.iloc[row_idx].tolist()
        ]
        found = set(values)

        if required_compact.issubset(found):
            df = raw.iloc[row_idx + 1:].copy()
            df.columns = [
                clean_text(v) for v in raw.iloc[row_idx].tolist()
            ]
            df = df.loc[:, [c != "" for c in df.columns]]
            return df.reset_index(drop=True)

    # 헤더명이 약간 다를 때 첫 행을 헤더로 사용
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file)


def prepare_budget(df):
    """예산 파일 열을 표준 열 이름으로 변환."""
    dept_col = find_column(
        df,
        ["부서(팀명)", "부서명", "부서", "지출 부서"],
        label="예산 부서",
    )
    task_col = find_column(
        df,
        ["과제번호", "과제명", "사업명", "과제"],
        label="예산 과제",
    )
    fund_col = find_column(
        df,
        ["자금명", "자금", "자금코드", "자금명칭"],
        label="예산 자금",
    )
    nrf_col = find_column(
        df,
        [
            "한국연구재단비목(드롭박스)",
            "한국연구재단비목",
            "연구재단 비목",
            "연구재단비목",
        ],
        label="예산 연구재단 비목",
    )
    account_col = find_column(
        df,
        [
            "교내계정과목(드롭박스)",
            "교내계정과목",
            "계정과목",
            "계정과목명",
        ],
        label="예산 계정과목",
    )
    amount_col = find_column(
        df,
        ["최초예산", "예산액", "예산금액", "예산"],
        label="예산 금액",
    )

    # 부서코드는 상세 구분용으로 보존
    dept_code_col = find_column(
        df, ["부서코드"], required=False
    )
    account_code_col = find_column(
        df, ["계정코드"], required=False
    )
    fund_code_col = find_column(
        df, ["자금코드"], required=False
    )

    out = pd.DataFrame()
    out["부서"] = df[dept_col].map(clean_text)
    out["과제"] = df[task_col].map(clean_text)
    out["자금"] = df[fund_col].map(clean_text)
    out["연구재단 비목"] = df[nrf_col].map(clean_text)
    out["계정과목"] = df[account_col].map(clean_text)
    out["예산액"] = to_number(df[amount_col])

    out["부서코드"] = (
        df[dept_code_col].map(clean_text)
        if dept_code_col else ""
    )
    out["계정코드"] = (
        df[account_code_col].map(clean_text)
        if account_code_col else ""
    )
    out["자금코드"] = (
        df[fund_code_col].map(clean_text)
        if fund_code_col else ""
    )

    # 연결 키: 공백 및 숫자형 코드 표기 차이 완화
    for col in ["부서", "과제", "자금", "연구재단 비목", "계정과목"]:
        out[f"{col}_key"] = normalize_key(out[col])

    return out


def prepare_execution(df):
    """집행 파일 열을 표준 열 이름으로 변환."""
    date_col = find_column(
        df,
        ["전기일", "전기 일자", "회계일자", "집행일자"],
        label="집행 전기일",
    )
    amount_col = find_column(
        df,
        ["금액(현지 통화)", "금액 (현지 통화)", "집행금액", "금액"],
        label="집행 금액",
    )
    dept_col = find_column(
        df,
        ["지출 부서", "부서명", "부서", "부서(팀명)"],
        label="집행 부서",
    )
    task_col = find_column(
        df,
        ["과제번호", "과제명", "사업명", "과제"],
        label="집행 과제",
    )
    fund_col = find_column(
        df,
        ["자금", "자금명", "자금명칭", "자금코드"],
        label="집행 자금",
    )
    nrf_col = find_column(
        df,
        [
            "연구재단 비목",
            "연구재단비목",
            "한국연구재단비목",
            "한국연구재단 비목",
        ],
        label="집행 연구재단 비목",
    )
    account_col = find_column(
        df,
        ["계정과목", "교내계정과목", "계정과목명"],
        label="집행 계정과목",
    )

    out = pd.DataFrame()
    out["전기일"] = pd.to_datetime(df[date_col], errors="coerce")
    out["집행액"] = to_number(df[amount_col])
    out["부서"] = df[dept_col].map(clean_text)
    out["과제"] = df[task_col].map(clean_text)
    out["자금"] = df[fund_col].map(clean_text)
    out["연구재단 비목"] = df[nrf_col].map(clean_text)
    out["계정과목"] = df[account_col].map(clean_text)

    for col in ["부서", "과제", "자금", "연구재단 비목", "계정과목"]:
        out[f"{col}_key"] = normalize_key(out[col])

    out["집행월"] = out["전기일"].dt.to_period("M").astype(str)
    out.loc[out["전기일"].isna(), "집행월"] = "날짜 미확인"

    return out


# =========================================================
# 3. 파일 업로드
# =========================================================
st.sidebar.header("엑셀 파일 업로드")

budget_file = st.sidebar.file_uploader(
    "예산내역 엑셀 파일",
    type=["xlsx", "xls"],
    key="budget_file",
)

execution_file = st.sidebar.file_uploader(
    "집행내역 엑셀 파일",
    type=["xlsx", "xls"],
    key="execution_file",
)

st.sidebar.caption(
    "예산 파일은 제목행이 두 번째 행(header=1)인 형식을 기준으로 읽습니다. "
    "집행 파일은 열 제목을 자동 탐색합니다."
)

if not budget_file or not execution_file:
    st.info("왼쪽에서 예산내역과 집행내역 엑셀 파일을 모두 업로드해 주세요.")
    st.stop()

try:
    budget_raw = pd.read_excel(budget_file, header=1)
    execution_raw = read_excel_auto_header(
        execution_file,
        ["전기일", "금액(현지 통화)"],
    )

    budget = prepare_budget(budget_raw)
    execution = prepare_execution(execution_raw)

except Exception as e:
    st.error("엑셀 파일을 읽거나 열을 연결하는 중 오류가 발생했습니다.")
    st.exception(e)
    st.stop()


# =========================================================
# 4. 예산 집계 및 집행내역 연결
# =========================================================
key_cols = [
    "부서_key",
    "과제_key",
    "자금_key",
    "연구재단 비목_key",
    "계정과목_key",
]

budget_summary = (
    budget.groupby(key_cols, dropna=False, as_index=False)
    .agg(예산액=("예산액", "sum"))
)

# 사람이 읽을 수 있는 이름 열을 키별로 보존
label_cols = [
    "부서",
    "과제",
    "자금",
    "연구재단 비목",
    "계정과목",
]
budget_labels = (
    budget.groupby(key_cols, dropna=False, as_index=False)[label_cols]
    .first()
)

budget_summary = budget_summary.merge(
    budget_labels, on=key_cols, how="left"
)

execution_summary = (
    execution.groupby(key_cols, dropna=False, as_index=False)
    .agg(누적집행액=("집행액", "sum"))
)

execution_labels = (
    execution.groupby(key_cols, dropna=False, as_index=False)[label_cols]
    .first()
)

execution_summary = execution_summary.merge(
    execution_labels, on=key_cols, how="left"
)

# 예산을 기준으로 집행액 연결
summary = budget_summary.merge(
    execution_summary[key_cols + ["누적집행액"]],
    on=key_cols,
    how="left",
    indicator=True,
)

summary["누적집행액"] = summary["누적집행액"].fillna(0)
summary["잔액"] = summary["예산액"] - summary["누적집행액"]
summary["집행률"] = (
    summary["누적집행액"]
    .div(summary["예산액"].replace(0, pd.NA))
    .mul(100)
    .fillna(0)
)

# =========================================================
# 5. 필터
# =========================================================
st.sidebar.divider()
st.sidebar.header("조회 조건")

def filter_options(df, col):
    vals = sorted(
        x for x in df[col].dropna().astype(str).unique()
        if x.strip()
    )
    return ["전체"] + vals


dept_choice = st.sidebar.selectbox(
    "부서", filter_options(summary, "부서")
)
task_choice = st.sidebar.selectbox(
    "과제", filter_options(summary, "과제")
)
fund_choice = st.sidebar.selectbox(
    "자금", filter_options(summary, "자금")
)
nrf_choice = st.sidebar.selectbox(
    "연구재단 비목", filter_options(summary, "연구재단 비목")
)

filtered = summary.copy()

for col, choice in [
    ("부서", dept_choice),
    ("과제", task_choice),
    ("자금", fund_choice),
    ("연구재단 비목", nrf_choice),
]:
    if choice != "전체":
        filtered = filtered[filtered[col] == choice]

filtered_keys = filtered[key_cols].drop_duplicates()

filtered_execution = execution.merge(
    filtered_keys,
    on=key_cols,
    how="inner",
)

# =========================================================
# 6. KPI
# =========================================================
total_budget = filtered["예산액"].sum()
total_execution = filtered["누적집행액"].sum()
total_balance = filtered["잔액"].sum()
total_rate = (
    total_execution / total_budget * 100
    if total_budget else 0
)

k1, k2, k3, k4 = st.columns(4)

k1.metric("총예산", format_money(total_budget) + " 원")
k2.metric("누적 집행액", format_money(total_execution) + " 원")
k3.metric("예산 잔액", format_money(total_balance) + " 원")
k4.metric("집행률", f"{total_rate:.1f}%")

st.divider()

# =========================================================
# 7. 월별 집행 추이
# =========================================================
st.subheader("월별 집행 추이")

monthly = (
    filtered_execution[filtered_execution["집행월"] != "날짜 미확인"]
    .groupby("집행월", as_index=False)
    .agg(집행액=("집행액", "sum"))
    .sort_values("집행월")
)

if not monthly.empty:
    fig_month = px.bar(
        monthly,
        x="집행월",
        y="집행액",
        text="집행액",
        color_discrete_sequence=[COLOR_MAIN],
    )
    fig_month.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        cliponaxis=False,
    )
    fig_month.update_yaxes(
        tickformat=",.0f",
        separatethousands=True,
        title="집행액",
    )
    fig_month.update_xaxes(title="")
    style_amount_axis(fig_month)
    st.plotly_chart(fig_month, use_container_width=True)
else:
    st.info("월별로 집계할 수 있는 집행일 데이터가 없습니다.")

# =========================================================
# 8. 예산 구성: 연구재단 비목별 파이 차트
# =========================================================
st.subheader("예산 구성 — 연구재단 비목별")

nrf_budget = (
    filtered.groupby("연구재단 비목", as_index=False)
    .agg(예산액=("예산액", "sum"))
)
nrf_budget = nrf_budget[nrf_budget["예산액"] != 0]

if not nrf_budget.empty:
    fig_pie = px.pie(
        nrf_budget,
        names="연구재단 비목",
        values="예산액",
        color_discrete_sequence=BLUE_PALETTE,
        hole=0.35,
    )
    fig_pie.update_traces(
        textposition="inside",
        textinfo="label+percent",
        hovertemplate=(
            "%{label}<br>예산액: %{value:,.0f}원"
            "<br>비율: %{percent}<extra></extra>"
        ),
    )
    fig_pie.update_layout(
        template="plotly_white",
        font=dict(color=COLOR_DARK),
        legend_title_text="연구재단 비목",
    )
    st.plotly_chart(fig_pie, use_container_width=True)
else:
    st.info("표시할 비목별 예산 데이터가 없습니다.")

# =========================================================
# 9. 부서별 예산·집행액
# =========================================================
st.subheader("부서별 예산 및 집행액")

dept_chart = (
    filtered.groupby("부서", as_index=False)
    .agg(
        예산액=("예산액", "sum"),
        누적집행액=("누적집행액", "sum"),
    )
    .sort_values("예산액", ascending=False)
    .head(15)
)

if not dept_chart.empty:
    dept_long = dept_chart.melt(
        id_vars="부서",
        value_vars=["예산액", "누적집행액"],
        var_name="구분",
        value_name="금액",
    )

    fig_dept = px.bar(
        dept_long,
        x="부서",
        y="금액",
        color="구분",
        barmode="group",
        text="금액",
        color_discrete_map={
            "예산액": COLOR_MAIN,
            "누적집행액": COLOR_BLUE,
        },
    )
    fig_dept.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        cliponaxis=False,
    )
    fig_dept.update_yaxes(
        tickformat=",.0f",
        separatethousands=True,
        title="금액",
    )
    fig_dept.update_xaxes(title="")
    style_amount_axis(fig_dept)
    st.plotly_chart(fig_dept, use_container_width=True)
else:
    st.info("표시할 부서별 데이터가 없습니다.")

# =========================================================
# 10. 영역별 예산·집행액
# =========================================================
st.subheader("영역별 예산 및 집행액")

area_col = "연구재단 비목"
area_chart = (
    filtered.groupby(area_col, as_index=False)
    .agg(
        예산액=("예산액", "sum"),
        누적집행액=("누적집행액", "sum"),
    )
    .sort_values("예산액", ascending=False)
)

if not area_chart.empty:
    area_long = area_chart.melt(
        id_vars=area_col,
        value_vars=["예산액", "누적집행액"],
        var_name="구분",
        value_name="금액",
    )

    fig_area = px.bar(
        area_long,
        x=area_col,
        y="금액",
        color="구분",
        barmode="group",
        text="금액",
        color_discrete_map={
            "예산액": COLOR_MAIN,
            "누적집행액": COLOR_BLUE,
        },
    )
    fig_area.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        cliponaxis=False,
    )
    fig_area.update_yaxes(
        tickformat=",.0f",
        separatethousands=True,
        title="금액",
    )
    fig_area.update_xaxes(title="")
    style_amount_axis(fig_area)
    st.plotly_chart(fig_area, use_container_width=True)
else:
    st.info("표시할 영역별 데이터가 없습니다.")

# =========================================================
# 11. 상세 예산 집행내역
# =========================================================
st.subheader("예산별 상세 집행현황")

detail_cols = [
    "부서",
    "과제",
    "자금",
    "연구재단 비목",
    "계정과목",
    "예산액",
    "누적집행액",
    "잔액",
    "집행률",
]

detail = filtered[detail_cols].copy()
detail["집행률"] = detail["집행률"].map(lambda x: f"{x:.1f}%")

display_table(
    detail,
    money_columns=["예산액", "누적집행액", "잔액"],
    height=450,
)

# =========================================================
# 12. 집행 상세내역
# =========================================================
st.subheader("집행 상세내역")

execution_detail = filtered_execution[
    [
        "전기일",
        "부서",
        "과제",
        "자금",
        "연구재단 비목",
        "계정과목",
        "집행액",
    ]
].copy()

execution_detail["전기일"] = execution_detail["전기일"].dt.strftime(
    "%Y-%m-%d"
)
execution_detail = execution_detail.sort_values(
    "전기일", ascending=False, na_position="last"
)

display_table(
    execution_detail,
    money_columns=["집행액"],
    height=400,
)

# =========================================================
# 13. 미연결 집행내역 검증
# =========================================================
st.subheader("미연결 집행내역 검증")

budget_keys = budget_summary[key_cols].drop_duplicates()
execution_check = execution.merge(
    budget_keys.assign(_예산연결=True),
    on=key_cols,
    how="left",
)

unmatched_execution = execution_check[
    execution_check["_예산연결"].isna()
].copy()

if unmatched_execution.empty:
    st.success("예산과 연결되지 않은 집행내역이 없습니다.")
else:
    st.warning(
        f"예산과 연결되지 않은 집행내역이 "
        f"{len(unmatched_execution):,}건 있습니다."
    )
    unmatched_cols = [
        "전기일",
        "부서",
        "과제",
        "자금",
        "연구재단 비목",
        "계정과목",
        "집행액",
    ]
    display_table(
        unmatched_execution[unmatched_cols],
        money_columns=["집행액"],
        height=350,
    )

# =========================================================
# 14. CSV 다운로드
# =========================================================
st.subheader("자료 다운로드")

download_col1, download_col2 = st.columns(2)

with download_col1:
    csv_detail = filtered[detail_cols].to_csv(
        index=False,
        encoding="utf-8-sig",
    )
    st.download_button(
        label="예산·집행 상세현황 CSV 다운로드",
        data=csv_detail,
        file_name="예산_집행_상세현황.csv",
        mime="text/csv",
        use_container_width=True,
    )

with download_col2:
    csv_unmatched = unmatched_execution[
        [
            "전기일",
            "부서",
            "과제",
            "자금",
            "연구재단 비목",
            "계정과목",
            "집행액",
        ]
    ].to_csv(index=False, encoding="utf-8-sig")

    st.download_button(
        label="미연결 집행내역 CSV 다운로드",
        data=csv_unmatched,
        file_name="미연결_집행내역.csv",
        mime="text/csv",
        use_container_width=True,
    )
