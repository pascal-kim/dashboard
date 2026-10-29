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

    return None


def read_budget(upload):
    xls = pd.ExcelFile(upload)

    preferred = [
        s for s in xls.sheet_names
        if "예산내역-whatif" in s
    ]

    preferred += [
        s for s in xls.sheet_names
        if "예산내역" in s and s not in preferred
    ]

    preferred += [
        s for s in xls.sheet_names
        if s not in preferred
    ]

    selected = None

    for sh in preferred:
        try:
            test = pd.read_excel(
                xls,
                sheet_name=sh,
                header=1,
                nrows=3
            )

            cn = [norm(c) for c in test.columns]

            if (
                any("부서코드" in c for c in cn)
                and any("자금코드" in c for c in cn)
                and any("최초예산" in c for c in cn)
            ):
                selected = sh
                break

        except Exception:
            pass

    if selected is None:
        selected = xls.sheet_names[0]

    return xls, selected


def area_class(task, area=""):
    t = (
        str(task)
        .strip()
        .upper()
        .replace("[", "")
        .replace("]", "")
    )

    if "사업관리" in t:
        return "사업관리"

    m = re.match(r"^([A-F])[-－]", t)

    if m:
        return "교육영역" if m.group(1) in "ABCDE" else "기타영역"

    a = str(area).strip()

    if a in ("교육", "교육영역"):
        return "교육영역"

    if a in ("기타", "기타영역"):
        return "기타영역"

    if a == "사업관리":
        return "사업관리"

    return "기타/미분류"


def clean_task(x):
    s = str(x).strip().upper()
    s = s.replace("[", "").replace("]", "").replace(" ", "")

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    # SAP 집행 자료의 사업운영은 예산의 사업관리와 연결
    if s == "사업운영":
        s = "사업관리"

    return s


def extract_update_date(filename):
    """파일명에서 YYYYMMDD 형식의 날짜를 찾아 YYYY년 MM월 DD일로 반환."""
    if not filename:
        return None

    # 파일명 중 독립된 8자리 숫자 중 실제 날짜로 유효한 값을 사용
    for match in re.finditer(r"(?<!\d)(\d{8})(?!\d)", str(filename)):
        date_text = match.group(1)
        try:
            parsed = datetime.strptime(date_text, "%Y%m%d")
            return parsed.strftime("%Y년 %m월 %d일")
        except ValueError:
            continue

    return None


def money(x):
    try:
        return f"{float(x):,.0f}원"
    except (ValueError, TypeError):
        return "—"


def format_money_columns(df, columns):
    """표에 표시할 금액을 천 단위 콤마 문자열로 변환."""
    out = df.copy()

    for col in columns:
        if col in out.columns:
            out[col] = out[col].apply(
                lambda x: f"{x:,.0f}"
                if pd.notna(x) and isinstance(x, (int, float))
                else (
                    f"{float(x):,.0f}"
                    if pd.notna(x)
                    and str(x).strip() != ""
                    and is_numeric(x)
                    else x
                )
            )

    return out


def is_numeric(x):
    try:
        float(x)
        return True
    except (ValueError, TypeError):
        return False


def right_align_columns(df, columns):
    """지정한 금액 열을 오른쪽 정렬."""
    cols = [col for col in columns if col in df.columns]

    if not cols:
        return df

    return df.style.set_properties(
        subset=cols,
        **{"text-align": "right"}
    )
