import re
from pathlib import Path

import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from bs4 import BeautifulSoup
import streamlit as st


# =========================
# 0) 기본 설정/경로 (GitHub/Streamlit Cloud)
# =========================
# =========================
# 한글 폰트 설정 (OTF / Streamlit Cloud 대응)
# =========================
import matplotlib.font_manager as fm

mpl.rcParams["axes.unicode_minus"] = False

def set_korean_font():
    font_path = (
        Path(__file__).resolve().parent / "NotoSansKR-Regular.otf"
        if "__file__" in globals()
        else Path.cwd() / "NotoSansKR-Regular.otf"
    )

    if font_path.exists():
        fm.fontManager.addfont(str(font_path))
        font_prop = fm.FontProperties(fname=str(font_path))
        mpl.rcParams["font.family"] = font_prop.get_name()
    else:
        mpl.rcParams["font.family"] = "sans-serif"

set_korean_font()


# ✅ 레포 루트 기준 상대경로
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

FILE_2023 = BASE_DIR / "2023 고속성장.xlsx"
FILE_2024 = BASE_DIR / "202411고속성장분석기(실채점)20241230.xlsx"

# ✅ CSV는 레포 내부 data/에 저장
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
SCRAPED_CSV = DATA_DIR / "sky_scraped_dept_candidates_bs4.csv"

missing = [p.name for p in [FILE_2023, FILE_2024] if not p.exists()]
if missing:
    st.error(
        "❌ 엑셀 파일 누락: " + str(missing) +
        "\n\n※ streamlit_app.py와 엑셀 파일 2개가 같은 폴더(레포 루트)에 있어야 합니다."
    )
    st.stop()

SKY_UNI = ["서울대학교", "고려대학교", "연세대학교"]
UNI_SHORT = {"서울대학교": "서울대", "고려대학교": "고려대", "연세대학교": "연세대"}

TARGET_URLS = {
    "서울대학교": "https://eng.snu.ac.kr/snu/main/contents.do?menuNo=200055",
    "고려대학교": "https://eng.korea.ac.kr/education/college.html",
    "연세대학교": "https://engineering.yonsei.ac.kr/engineering/about/major_1_1.do",
}

MED_KEYWORDS = ["의학","의예","의과","치의","치의학","약학","약학과","수의","수의학","간호","간호학","보건"]

SCRAPE_KEYWORDS = [
    "공학","공학부","공학과","학부","학과",
    "기계","전기","전자","컴퓨터",
    "화공","화학생명","신소재","재료",
    "산업","건축","도시","건설","환경",
    "반도체","디스플레이","에너지",
]

GROUP_MAP = {
    "컴": "컴퓨터", "전산": "컴퓨터", "소프트웨어": "컴퓨터", "데이터": "컴퓨터",
    "전기": "전기전자", "전자": "전기전자", "전기전자": "전기전자", "전기정보": "전기전자",
    "화공": "화공", "화학생명": "화공", "화학생물": "화공",
    "신소재": "신소재", "재료": "신소재",
    "기계": "기계", "산업": "산업", "건축": "건축", "도시": "도시",
    "건설": "건설환경", "환경": "건설환경",
    "반도체": "반도체", "디스플레이": "디스플레이", "에너지": "에너지",
}


# =========================
# 1) 유틸
# =========================
def normalize(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip().lower()
    s = re.sub(r"[\s\-\(\)\[\]\{\}·•\.,/\\:;\"'“”‘’]", "", s)
    return s

def is_medical(major: str) -> bool:
    if not isinstance(major, str):
        return False
    return any(k in major for k in MED_KEYWORDS)

def get_major_group(name: str):
    if not isinstance(name, str):
        return None
    for key, val in GROUP_MAP.items():
        if key in name:
            return val
    return None

def extract_candidates_from_text(text: str):
    lines = [x.strip() for x in text.split("\n")]
    out = []
    for line in lines:
        if not line:
            continue
        if len(line) > 40:
            continue
        if any(k.lower() in line.lower() for k in SCRAPE_KEYWORDS):
            out.append(line)
    return sorted(set(out))


# =========================
# 2) BS4 수집
# =========================
def fetch_html(url: str) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    if r.encoding is None:
        r.encoding = "utf-8"
    return r.text

def scrape_departments_bs4(output_csv: Path) -> pd.DataFrame:
    rows = []
    for uni, url in TARGET_URLS.items():
        st.write(f"**[BS4] {uni}** → {url}")
        try:
            html = fetch_html(url)
        except Exception as e:
            st.warning(f"{uni} 요청 실패: {e}")
            continue

        soup = BeautifulSoup(html, "html.parser")
        text = soup.get_text("\n")
        candidates = extract_candidates_from_text(text)

        st.write(f"후보 **{len(candidates)}개** 추출")
        for c in candidates:
            rows.append({
                "대학교": uni,
                "candidate": c,
                "candidate_norm": normalize(c),
                "source_url": url
            })

    df = pd.DataFrame(rows).drop_duplicates()
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    st.success(f"저장 완료: {output_csv} ({len(df)} rows)")
    return df

def build_scraped_token_set(scraped_df: pd.DataFrame) -> set:
    if scraped_df is None or scraped_df.empty:
        return set()
    toks = set(scraped_df["candidate_norm"].astype(str).tolist())
    return {t for t in toks if len(t) >= 3}


# =========================
# 3) 엑셀 로드/정제
# =========================
def read_excel_year(path: Path, year: int) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="이과계열분석결과", header=4)
    df = df[df["대학교"].isin(SKY_UNI)].copy()
    df = df[~df["전공"].apply(is_medical)].copy()
    df["연도"] = year
    return df[["연도", "대학교", "전공", "적정점수"]]

def score_major_by_tokens(major: str, token_set: set) -> int:
    m = normalize(major)
    if not m or not token_set:
        return 0
    return sum(1 for tok in token_set if tok in m)

def select_top7_per_uni(all_df: pd.DataFrame, token_set: set, k: int = 7) -> pd.DataFrame:
    df = all_df.copy()
    df["scrape_score"] = df["전공"].astype(str).apply(lambda x: score_major_by_tokens(x, token_set))

    df24 = df[df["연도"] == 2024].copy()
    summary24 = (
        df24.groupby(["대학교", "전공"], as_index=False)
        .agg(avg_score=("적정점수", "mean"),
             scrape_score=("scrape_score", "max"))
    )

    chosen = []
    for uni in SKY_UNI:
        sub = summary24[summary24["대학교"] == uni].copy()
        if sub.empty:
            continue
        sub = sub.sort_values(["scrape_score", "avg_score"], ascending=[False, False])
        picked = sub.head(k)["전공"].tolist()
        chosen.extend([(uni, m) for m in picked])

    chosen_df = pd.DataFrame(chosen, columns=["대학교", "전공"]).drop_duplicates()
    filtered = df.merge(chosen_df, on=["대학교", "전공"], how="inner").drop(columns=["scrape_score"])
    return filtered


# =========================
# 4) 그래프 함수 (Streamlit용) - ✅ 원본 그대로
# =========================
def fig_uni_year_bar(filtered_df: pd.DataFrame, uni: str, year: int):
    df24 = filtered_df[filtered_df["연도"] == 2024].copy()
    order = (
        df24[df24["대학교"] == uni]
        .groupby("전공")["적정점수"].mean()
        .sort_values(ascending=False)
        .index.tolist()
    )

    school_df = filtered_df[filtered_df["대학교"] == uni].copy()
    sub = school_df[school_df["연도"] == year]
    grouped = sub.groupby("전공")["적정점수"].mean().reindex(order).dropna()

    y_min = school_df["적정점수"].min()
    y_max = school_df["적정점수"].max()
    margin = max(3, (y_max - y_min) * 0.40)
    lower, upper = y_min - margin, y_max + margin

    fig, ax = plt.subplots(figsize=(8.5, 4.5))
    palette = plt.cm.tab10.colors
    majors = grouped.index.tolist()
    scores = grouped.values
    colors = [palette[i % len(palette)] for i in range(len(majors))]

    ax.bar(majors, scores, color=colors)
    ax.set_title(f"{UNI_SHORT[uni]} {year} 상위 7개 전공")
    ax.set_ylabel("적정점수")
    ax.set_ylim(lower, upper)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.tick_params(axis="x", rotation=20)
    plt.tight_layout()
    return fig

def fig_top5_dot(filtered_df: pd.DataFrame):
    latest = filtered_df[filtered_df["연도"] == 2024].copy()
    latest["전공군"] = latest["전공"].apply(get_major_group)
    latest = latest.dropna(subset=["전공군"])

    latest["school_mean"] = latest.groupby("대학교")["적정점수"].transform("mean")
    latest["school_std"] = latest.groupby("대학교")["적정점수"].transform("std").replace(0, 1)
    latest["z"] = (latest["적정점수"] - latest["school_mean"]) / latest["school_std"]

    rank = (
        latest.groupby("전공군")["z"].mean()
        .sort_values(ascending=False)
        .reset_index()
        .head(5)
    )

    fig, ax = plt.subplots(figsize=(8, 4.8))
    colors = list(plt.cm.Set2.colors) + list(plt.cm.Pastel1.colors)
    colors = colors[:len(rank)]

    xmin = rank["z"].min() - 0.2
    xmax = rank["z"].max() + 0.2

    for idx, row in rank.reset_index(drop=True).iterrows():
        y = row["전공군"]
        x = row["z"]
        c = colors[idx]
        ax.hlines(y, xmin, x, colors=c, linestyles="--", alpha=0.75, linewidth=2)
        ax.scatter(x, y, s=180, color=c, edgecolor="black", zorder=3)

    ax.set_xlim(xmin, xmax)
    ax.set_title("SKY 통합 TOP5 전공군 (2024, 학교 내 정규화)")
    ax.set_xlabel("평균 z-score (각 학교 평균 대비 우위)")
    ax.grid(axis="x", linestyle="--", alpha=0.35)
    plt.tight_layout()
    return fig, rank


# =========================
# 5) Streamlit 앱 - ✅ 원본 그대로
# =========================
@st.cache_data(show_spinner=False)
def load_analysis_data():
    df23 = read_excel_year(FILE_2023, 2023)
    df24 = read_excel_year(FILE_2024, 2024)
    return pd.concat([df23, df24], ignore_index=True)

@st.cache_data(show_spinner=False)
def load_scraped_csv_if_exists():
    if SCRAPED_CSV.exists():
        return pd.read_csv(SCRAPED_CSV)
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def build_filtered(all_df, scraped_df):
    token_set = build_scraped_token_set(scraped_df)
    return select_top7_per_uni(all_df, token_set, k=7)

def main():
    st.set_page_config(page_title="SKY 공대 입결 분석", layout="wide")
    st.title("SKY 공대 상위 전공 적정점수 분석 (2023–2024)")
    st.caption("BeautifulSoup으로 공대 학과/학부 후보를 수집한 뒤, 엑셀 입결 데이터와 결합하여 학교별 상위 7개 전공 및 통합 TOP5 전공군을 시각화합니다.")

    with st.expander("1) 데이터/방법 설명", expanded=True):
        st.markdown(
            """
- **데이터(입결)**: `2023 고속성장.xlsx`, `202411고속성장분석기(실채점)20241230.xlsx`의 `이과계열분석결과` 시트에서 **적정점수** 사용  
- **대상**: SKY(서울대/고대/연대) + **메디컬 계열 제외**  
- **웹 수집(BS4)**: 각 대학 공대 페이지에서 “학부/학과/공학” 등 키워드가 포함된 텍스트를 추출하여 **전공 후보 토큰**을 생성  
- **연결(유기적 사용)**: 2024 기준 학교별 상위 7개 전공을 뽑을 때, **웹 수집 토큰과 전공명이 얼마나 겹치는지(scrape_score)**를 우선 반영
            """
        )

    colA, colB = st.columns([1, 2], gap="large")

    with colA:
        st.subheader("2) (선택) BS4 수집 실행")
        st.write("처음 1번만 눌러서 CSV 저장해두면, 이후엔 CSV를 읽어서 빠르게 동작해.")
        if st.button("웹에서 학과 후보 수집(BS4) → CSV 저장"):
            with st.spinner("수집 중..."):
                scraped_df = scrape_departments_bs4(SCRAPED_CSV)
                st.dataframe(scraped_df.head(20))

        st.subheader("3) 보기 선택")
        view = st.radio(
            "메뉴",
            ["서울대 2023", "서울대 2024", "고려대 2023", "고려대 2024", "연세대 2023", "연세대 2024", "통합 TOP5"],
            index=0
        )

    with colB:
        all_df = load_analysis_data()
        scraped_df = load_scraped_csv_if_exists()
        if scraped_df.empty:
            st.warning("아직 수집 CSV가 없어. 왼쪽 버튼으로 한번 수집하거나, data/ 폴더에 CSV가 있는지 확인해줘.")
        filtered = build_filtered(all_df, scraped_df)

        st.caption("학교별 2024 전공 개수 체크")
        chk = (
            filtered[filtered["연도"] == 2024]
            .groupby("대학교")["전공"].nunique()
            .reset_index(name="n_majors_2024")
        )
        st.dataframe(chk)

        if view == "통합 TOP5":
            fig, rank = fig_top5_dot(filtered)
            st.pyplot(fig, clear_figure=True)
            st.subheader("TOP5 전공군 테이블")
            st.dataframe(rank)
        else:
            uni_map = {"서울대": "서울대학교", "고려대": "고려대학교", "연세대": "연세대학교"}
            short = view.split()[0]
            year = int(view.split()[1])
            uni = uni_map[short]
            fig = fig_uni_year_bar(filtered, uni, year)
            st.pyplot(fig, clear_figure=True)

            st.subheader(f"{short} {year} 상위 전공(평균 적정점수)")
            table = (
                filtered[(filtered["대학교"] == uni) & (filtered["연도"] == year)]
                .groupby("전공")["적정점수"].mean()
                .sort_values(ascending=False)
                .reset_index()
            )
            st.dataframe(table)

if __name__ == "__main__":
    main()
