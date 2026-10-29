    "date",
    "dept",
    "task",
    "fund",
    "nrf",
    "account_code",
    "amount",
    "월"
]

detail = (
    ef[detail_cols]
    .rename(
        columns={
            "date": "전기일",
            "dept": "지출 부서",
            "task": "과제번호",
            "fund": "자금",
            "nrf": "연구재단 비목",
            "account_code": "계정과목",
            "amount": "원본 금액",
            "월": "집행월"
        }
    )
    .sort_values("전기일", ascending=False)
    .copy()
)

detail = format_money_columns(
    detail,
    ["원본 금액"]
)

detail["전기일"] = detail["전기일"].dt.strftime("%Y-%m-%d")

st.dataframe(
    align_table(detail),
    use_container_width=True,
    hide_index=True
)


# ─────────────────────────────────────────────
# 7. 집행내역 검증
# ─────────────────────────────────────────────

matched_keys = exp_group.merge(
    budget_group[key],
    on=key,
    how="inner"
)

st.subheader("집행내역 검증")

c1, c2, c3 = st.columns(3)

c1.metric(
    "집행 원본 행 수",
    f"{len(exp):,}"
)

c2.metric(
    "예산과 연결된 집행 조합",
    f"{len(matched_keys):,}"
)

c3.metric(
    "예산과 미연결 집행 조합",
    f"{len(exp_group) - len(matched_keys):,}"
)

with st.expander("미연결 집행 내역 확인"):
    unmatched = exp_group.merge(
        budget_group[key],
        on=key,
        how="left",
        indicator=True
    )

    unmatched = unmatched[
        unmatched["_merge"] == "left_only"
    ].drop(columns="_merge")

    unmatched_display = format_money_columns(
        unmatched,
        ["누적집행액"]
    )

    st.dataframe(
        align_table(unmatched_display),
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        "미연결은 부서명·과제번호·자금코드·연구재단비목·"
        "계정코드가 예산 파일과 정확히 일치하지 않는 경우입니다. "
        "예산과 연결되지 않은 집행액도 월별 집행 그래프에는 포함됩니다."
    )


# ─────────────────────────────────────────────
# 8. 예산 대비 집행 집계 CSV 다운로드
# ─────────────────────────────────────────────

csv = summary[show_cols].to_csv(
    index=False
).encode("utf-8-sig")

st.download_button(
    "예산 대비 집행 집계 CSV 다운로드",
    data=csv,
    file_name="예산대비집행_집계.csv",
    mime="text/csv"
)
