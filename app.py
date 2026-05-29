# ══════════════════════════════════════════════════════════════════════════════
# app.py  |  경상남도 고교 상담지원 인프라 수급 불균형 분석 대시보드
# ══════════════════════════════════════════════════════════════════════════════

import pathlib
import warnings
warnings.filterwarnings("ignore")

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pydeck as pdk
import streamlit as st

# ── 1. 페이지 설정 ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="경남 고교 상담지원 수급 불균형 분석",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 2. CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* 전체 배경 */
[data-testid="stAppViewContainer"] > .main { background-color: #F0F4F8; }
[data-testid="stHeader"] { background: transparent; }

/* 사이드바 — 밝은 배경 + 좁은 너비 */
[data-testid="stSidebar"] { background-color: #F7F9FC !important; }
[data-testid="stSidebar"] > div:first-child { width: 220px !important; }
section[data-testid="stSidebarContent"] {
    width: 220px !important;
    padding-left: 0.5rem !important;
    padding-right: 0.5rem !important;
}
/* 메뉴 라디오 항목 - 작게 */
[data-testid="stSidebar"] .stRadio label p { font-size: 0.72rem !important; }
/* 필터 selectbox 레이블(시군구, 우선지원등급) - 크게 */
[data-testid="stSidebar"] .stSelectbox label { font-size: 0.84rem !important; }
/* selectbox 선택값(전체 등) - 크게 */
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] span { font-size: 0.84rem !important; }
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] div { font-size: 0.84rem !important; }
/* 필터 초기화 버튼 - 작게 */
[data-testid="stSidebar"] .stButton > button { font-size: 0.70rem !important; }

/* 메인 콘텐츠 패딩 조정 */
.block-container {
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
    padding-top: 3.5rem !important;
    max-width: 100% !important;
}
[data-testid="stSidebar"] .stButton > button {
    border-radius: 6px; font-size: 0.8rem;
}

/* 지표 카드 */
.metric-card {
    background: white;
    border-radius: 10px;
    padding: 14px 14px;
    display: flex;
    align-items: center;
    gap: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    height: 90px;
}
.metric-icon {
    font-size: 1.4rem;
    width: 44px; height: 44px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
}
.metric-label { font-size: 0.73rem; color: #718096; margin-bottom: 2px; }
.metric-value { font-size: 1.55rem; font-weight: 700; line-height: 1.1; }
.metric-sub   { font-size: 0.68rem; color: #A0AEC0; }

/* 섹션 제목 */
.sec-title {
    font-size: 0.88rem; font-weight: 700; color: #1E3A5F;
    padding: 10px 14px 8px 14px;
    border-bottom: 1px solid #E8EEF6;
    margin-bottom: 2px;
    background: white;
    border-radius: 10px 10px 0 0;
}

/* 범례 */
.legend-row { display:flex; flex-wrap:wrap; gap:5px 12px;
              padding: 6px 14px 10px 14px; background:white; }
.legend-item { display:flex; align-items:center; gap:4px;
               font-size:0.74rem; color:#4A5568; }
.legend-dot  { width:10px; height:10px; border-radius:50%; flex-shrink:0; }

/* TOP10 테이블 */
.top10-table { width:100%; border-collapse:collapse; font-size:0.79rem; }
.top10-table th {
    background:#EBF2FF; color:#1E3A5F; font-weight:700;
    padding:6px 6px; text-align:center;
}
.top10-table td { padding:5px 6px; border-bottom:1px solid #F0F4F8; }
.top10-table tr:hover td { background:#F7FAFC; }

/* 배지 */
.badge {
    display:inline-block; padding:1px 7px; border-radius:8px;
    font-size:0.7rem; font-weight:700; color:white; white-space:nowrap;
}

/* 하단 안내 */
.footer-note {
    color:#A0AEC0; font-size:0.7rem; text-align:center;
    padding:12px 0 4px 0; margin-top:8px;
}

/* 섹션 간 수직 여백 통일 */
[data-testid="stVerticalBlock"] { gap: 0.75rem !important; }

</style>
""", unsafe_allow_html=True)

# ── 3. 경로 / 상수 ─────────────────────────────────────────────────────────────
ROOT      = pathlib.Path(__file__).parent
DATA_PATH = ROOT / "data" / "processed" / \
            "gyeongnam_high_schools_policy_feedback_refined.xlsx"
SHEET     = "refined_policy_feedback_table"

REQUIRED_COLS = [
    "school_name", "sigungu", "CSI", "CDI",
    "priority_score", "priority_level", "policy_strategy_group",
]

# 우선지원등급 → N등급 표시 매핑
PRIORITY_DISPLAY = {
    "최우선 지원": "1등급(최우선)",
    "우선 지원":   "2등급(우선)",
    "모니터링":    "3등급(모니터링)",
    "안정":        "4등급(안정)",
}
PRIORITY_ORDER = ["1등급(최우선)", "2등급(우선)", "3등급(모니터링)", "4등급(안정)"]
PRIORITY_COLORS = {
    "1등급(최우선)":   "#C0392B",
    "2등급(우선)":     "#E67E22",
    "3등급(모니터링)": "#F4D03F",
    "4등급(안정)":     "#27AE60",
}

STRATEGY_COLORS = {
    "최우선 개입형":      "#C0392B",
    "우선 보완형":        "#E67E22",
    "고수요 유지관리형":  "#F4D03F",
    "인력 취약형":        "#9B59B6",
    "접근성 보완형":      "#2980B9",
    "최소 인프라 보완형": "#1ABC9C",
    "안정형":             "#27AE60",
    "확인 필요형":        "#BDC3C7",
}
STRATEGY_ORDER = [
    "최우선 개입형", "우선 보완형", "고수요 유지관리형",
    "인력 취약형", "접근성 보완형", "최소 인프라 보완형",
    "안정형", "확인 필요형",
]

SIGUNGU_COORDS = {
    "창원시": (35.2278, 128.6817), "진주시": (35.1798, 128.1076),
    "통영시": (34.8544, 128.4334), "사천시": (35.0035, 128.0636),
    "김해시": (35.2281, 128.8891), "밀양시": (35.4956, 128.7483),
    "거제시": (34.8800, 128.6211), "양산시": (35.3350, 129.0379),
    "의령군": (35.3218, 128.2618), "함안군": (35.2725, 128.4063),
    "창녕군": (35.5443, 128.4924), "고성군": (34.9730, 128.3229),
    "남해군": (34.8378, 127.8927), "하동군": (35.0675, 127.7516),
    "산청군": (35.4156, 127.8732), "함양군": (35.5206, 127.7254),
    "거창군": (35.6868, 127.9095), "합천군": (35.5668, 128.1652),
}


def _hex_to_rgba(hex_color: str, alpha: int = 210) -> list:
    h = hex_color.lstrip("#")
    return [int(h[i:i+2], 16) for i in (0, 2, 4)] + [alpha]


PRIORITY_COLORS_PYDECK = {k: _hex_to_rgba(v) for k, v in PRIORITY_COLORS.items()}
DEFAULT_COLOR_PYDECK   = [149, 165, 166, 200]


# ── 4. 데이터 로드 ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="데이터를 불러오는 중...")
def load_data(path: pathlib.Path, sheet: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet, dtype={"school_code": str})


# ── 5. 사이드바 ────────────────────────────────────────────────────────────────
def render_sidebar(df: pd.DataFrame):
    sb = st.sidebar

    # 로고 + 제목
    sb.markdown("""
    <div style='padding:10px 4px 12px 4px;'>
      <div style='display:flex;align-items:center;gap:8px;margin-bottom:4px;'>
        <div style='background:#2E5FA3;border-radius:8px;width:30px;height:30px;
                    display:flex;align-items:center;justify-content:center;
                    font-size:0.95rem;flex-shrink:0;'>📊</div>
        <div>
          <div style='color:#1E3A5F;font-size:0.75rem;font-weight:700;line-height:1.4;'>
            경상남도 교육청</div>
          <div style='color:#718096;font-size:0.68rem;'>상담지원 인프라 분석</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # 메뉴
    tabs = [
        "📋 현황 개요", "🗺️ 지역별 분석", "🔍 학교 검색",
        "📊 유형 분석",  "⚙️ 시뮬레이션",  "ℹ️ 데이터 설명",
    ]
    selected = sb.radio("메뉴", tabs, label_visibility="collapsed")

    sb.markdown("<hr style='border-color:#E2E8F0;margin:10px 0;'>",
                unsafe_allow_html=True)

    # 필터
    sb.markdown(
        "<p style='color:#4A5568;font-size:0.78rem;font-weight:700;"
        "margin:0 0 8px 0;'>필터</p>",
        unsafe_allow_html=True)

    sigungu_list = ["전체"] + sorted(df["sigungu"].dropna().unique().tolist())
    sel_sigungu  = sb.selectbox("시군구", sigungu_list)

    priority_list = ["전체"] + PRIORITY_ORDER
    sel_priority  = sb.selectbox("우선지원등급", priority_list)

    if sb.button("↺ 필터 초기화", use_container_width=True):
        st.rerun()

    sb.markdown("<hr style='border-color:#E2E8F0;margin:10px 0;'>",
                unsafe_allow_html=True)
    sb.markdown(
        "<p style='color:#4A5568;font-size:0.7rem;margin:0;'>파일럿 v1.0 · 2025</p>",
        unsafe_allow_html=True)

    return selected, sel_sigungu, sel_priority


# ── 6. 현황 개요 탭 ────────────────────────────────────────────────────────────
def show_overview(df_all: pd.DataFrame, sigungu_f: str, priority_f: str):
    # 필터 적용
    df = df_all.copy()
    if sigungu_f  != "전체":
        df = df[df["sigungu"]          == sigungu_f]
    if priority_f != "전체":
        df = df[df["priority_display"] == priority_f]

    # ── 헤더 ─────────────────────────────────────────────────────────────────
    st.markdown(
        "<h1 style='font-size:1.3rem;color:#1E3A5F;margin:0 0 2px 0;font-weight:700;'>"
        "경상남도 고등학교 상담지원 인프라 수급 불균형 분석 및 의사결정 대시보드"
        "</h1>"
        "<p style='color:#718096;font-size:0.78rem;margin:0 0 14px 0;'>"
        "CSI(상담공급지수) · CDI(상담수요지수) · 우선지원점수 기반 학교별 우선순위 분석 "
        "| 데이터 기준일: 2026.01.01</p>",
        unsafe_allow_html=True)

    # ── 지표 카드 5개 ─────────────────────────────────────────────────────────
    _render_metric_cards(df, df_all)
    st.markdown("<div style='margin:25px 0;'></div>", unsafe_allow_html=True)

    # ── 2열 레이아웃 ─────────────────────────────────────────────────────────
    map_col, right_col = st.columns([3, 2], gap="small")

    with map_col:
        _render_map_section(df)

    with right_col:
        _render_infra_section(df)

        _render_top10(df)

    # ── 분포 히스토그램 ───────────────────────────────────────────────────────

    _render_distributions(df)

    # ── 하단 안내 ─────────────────────────────────────────────────────────────
    st.markdown(
        "<div class='footer-note'>"
        "※ CSI(상담지원 인프라 공급지수) · CDI(상담수요 지수) · 우선지원점수 = CDI − CSI "
        "| 우선지원등급: 1등급(최우선) → 4등급(안정) "
        "| 본 대시보드는 분석 결과를 기반으로 한 정책 의사결정 지원 도구이며, "
        "실제 지원 확정 기준은 아닙니다."
        "</div>",
        unsafe_allow_html=True)


# ── 지표 카드 ──────────────────────────────────────────────────────────────────
def _render_metric_cards(df: pd.DataFrame, df_all: pd.DataFrame):
    n       = len(df)
    n_all   = len(df_all)
    csi     = df["CSI"].mean() if n else 0.0
    cdi     = df["CDI"].mean() if n else 0.0
    ps      = df["priority_score"].mean() if n else 0.0
    n_top   = int((df["priority_display"] == "1등급(최우선)").sum())
    pct_top = n_top / n * 100 if n else 0.0

    defs = [
        ("분석 대상 학교 수",    f"{n}개교",       f"전체 {n_all}개교 중",
         "🏫", "#EBF2FF", "#2980B9"),
        ("평균 CSI",             f"{csi:.3f}",      "(0 ~ 1)",
         "📋", "#E8F8F5", "#1ABC9C"),
        ("평균 CDI",             f"{cdi:.3f}",      "(0 ~ 1)",
         "👥", "#FEF9E7", "#F39C12"),
        ("평균 우선지원점수",    f"{ps:+.3f}",     "CDI − CSI",
         "🎯", "#F3E8FF", "#8E44AD"),
        ("1등급(최우선) 학교 수", f"{n_top}개교",  f"({pct_top:.1f}%)",
         "⚠️", "#FDECEA", "#C0392B"),
    ]

    cols = st.columns(5, gap="small")
    for col, (label, value, sub, icon, bg, color) in zip(cols, defs):
        col.markdown(
            f'<div class="metric-card" style="border-left:4px solid {color};">'
            f'<div class="metric-icon" style="background:{bg};">{icon}</div>'
            f'<div>'
            f'<div class="metric-label">{label}</div>'
            f'<div class="metric-value" style="color:{color};">{value}</div>'
            f'<div class="metric-sub">{sub}</div>'
            f'</div></div>',
            unsafe_allow_html=True)


# ── 지도 섹션 ──────────────────────────────────────────────────────────────────
def _render_map_section(df: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:6px;'>📍 경상남도 학교 분포 현황</div>",
            unsafe_allow_html=True)

        # 범례
        legend_html = '<div class="legend-row">'
        for grade in PRIORITY_ORDER:
            if (df["priority_display"] == grade).sum() > 0:
                c = PRIORITY_COLORS.get(grade, "#95A5A6")
                legend_html += (
                    f'<div class="legend-item">'
                    f'<span class="legend-dot" style="background:{c};"></span>'
                    f'{grade}</div>')
        legend_html += '</div>'
        st.markdown(legend_html, unsafe_allow_html=True)

        has_latlon = (
            "school_latitude"  in df.columns
            and "school_longitude" in df.columns
        )
        n_valid = int(
            df[["school_latitude", "school_longitude"]].notna().all(axis=1).sum()
        ) if has_latlon else 0

        if has_latlon and n_valid >= 10:
            _render_school_pydeck_map(df)
            st.caption(f"학교별 위경도 기반 ({n_valid}개교) · 점 색상: 우선지원등급 | 마우스를 올리면 상세 정보 표시")
        else:
            _render_sigungu_map(df)


# ── pydeck 학교 산점도 ─────────────────────────────────────────────────────────
def _render_school_pydeck_map(df: pd.DataFrame):
    try:
        map_cols = [
            "school_name", "sigungu",
            "school_latitude", "school_longitude",
            "CSI", "CDI", "priority_score",
            "priority_display", "policy_strategy_group",
        ]
        map_df = df[[c for c in map_cols if c in df.columns]].copy()
        map_df = map_df.dropna(subset=["school_latitude", "school_longitude"])

        map_df["CSI_d"] = map_df["CSI"].round(3).astype(str)
        map_df["CDI_d"] = map_df["CDI"].round(3).astype(str)
        map_df["ps_d"]  = map_df["priority_score"].round(3).astype(str)
        map_df["color"] = map_df["priority_display"].map(
            lambda g: PRIORITY_COLORS_PYDECK.get(g, DEFAULT_COLOR_PYDECK))

        layer = pdk.Layer(
            "ScatterplotLayer",
            data=map_df,
            get_position=["school_longitude", "school_latitude"],
            get_color="color",
            get_radius=800,
            pickable=True,
            opacity=0.88,
            stroked=True,
            filled=True,
            line_width_min_pixels=1,
            get_line_color=[255, 255, 255, 160],
            auto_highlight=True,
        )
        view_state = pdk.ViewState(
            latitude=35.23, longitude=128.15, zoom=7.8, pitch=0, bearing=0)
        tooltip = {
            "html": (
                "<div style='font-family:sans-serif;font-size:13px;"
                "line-height:1.6;padding:4px;'>"
                "<b>{school_name}</b><br/>"
                "시군구: {sigungu}<br/>"
                "CSI: {CSI_d} &nbsp;|&nbsp; CDI: {CDI_d}<br/>"
                "우선지원점수: {ps_d}<br/>"
                "우선지원등급: {priority_display}<br/>"
                "정책전략: {policy_strategy_group}"
                "</div>"
            ),
            "style": {
                "backgroundColor": "#1E3A5F",
                "color": "white",
                "borderRadius": "6px",
                "padding": "8px",
            },
        }
        deck = pdk.Deck(
            layers=[layer],
            initial_view_state=view_state,
            tooltip=tooltip,
            map_style="https://basemaps.cartocdn.com/gl/voyager-gl-style/style.json",
        )
        st.pydeck_chart(deck, width="stretch")

    except Exception as e:
        st.warning(f"지도 오류: {e}")
        _render_sigungu_map(df)


# ── 시군구 버블맵 fallback ─────────────────────────────────────────────────────
def _render_sigungu_map(df: pd.DataFrame):
    try:
        agg = (
            df.groupby("sigungu")
            .agg(학교수=("school_name", "count"),
                 CSI_평균=("CSI", "mean"),
                 CDI_평균=("CDI", "mean"),
                 우선지원점수_평균=("priority_score", "mean"))
            .round(3).reset_index()
        )
        agg["위도"] = agg["sigungu"].map(
            lambda s: SIGUNGU_COORDS.get(s, (None, None))[0])
        agg["경도"] = agg["sigungu"].map(
            lambda s: SIGUNGU_COORDS.get(s, (None, None))[1])
        agg = agg.dropna(subset=["위도", "경도"])
        if agg.empty:
            st.info("지도 데이터 없음")
            return
        fig = px.scatter_mapbox(
            agg, lat="위도", lon="경도", size="학교수",
            color="우선지원점수_평균",
            color_continuous_scale=["#27AE60", "#F4D03F", "#E67E22", "#C0392B"],
            size_max=35, zoom=7.8,
            center={"lat": 35.23, "lon": 128.15},
            mapbox_style="carto-positron",
            hover_name="sigungu",
            hover_data={"학교수": True, "CSI_평균": True, "CDI_평균": True,
                        "위도": False, "경도": False},
        )
        fig.update_layout(height=430, margin={"r": 0, "t": 0, "l": 0, "b": 0})
        st.plotly_chart(fig, width="stretch")
    except Exception as e:
        st.warning(f"지도 오류: {e}")


# ── 인프라 현황 (우선지원등급 막대 + 전략 도넛) ──────────────────────────────
def _render_infra_section(df: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:8px;'>📊 상담 인프라 현황</div>",
            unsafe_allow_html=True)

        bar_col, donut_col = st.columns(2, gap="small")

        with bar_col:
            cnt = (
                df["priority_display"]
                .value_counts()
                .reindex(PRIORITY_ORDER)
                .fillna(0)
                .astype(int)
            )
            colors = [PRIORITY_COLORS.get(g, "#95A5A6") for g in cnt.index]
            fig = go.Figure(go.Bar(
                x=[g.split("(")[0] for g in cnt.index],
                y=cnt.values,
                marker_color=colors,
                text=cnt.values,
                textposition="outside",
                textfont=dict(size=10),
            ))
            fig.update_layout(
                title=dict(text="우선지원등급별 학교 수",
                           font=dict(size=10, color="#1E3A5F")),
                xaxis=dict(tickfont=dict(size=8.5)),
                yaxis=dict(range=[0, max(cnt.max() * 1.35, 1)],
                           showticklabels=False),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(t=32, b=4, l=4, r=4), height=210,
                showlegend=False,
            )
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(showgrid=True, gridcolor="#EEE")
            st.plotly_chart(fig, width="stretch")

        with donut_col:
            grp    = df["policy_strategy_group"].value_counts()
            labels = grp.index.tolist()
            vals   = grp.values.tolist()
            clrs   = [STRATEGY_COLORS.get(g, "#BDC3C7") for g in labels]
            fig = go.Figure(go.Pie(
                labels=labels, values=vals,
                hole=0.52,
                marker=dict(colors=clrs),
                textinfo="percent",
                textfont=dict(size=8.5),
                hovertemplate="<b>%{label}</b><br>%{value}개교 (%{percent})<extra></extra>",
                showlegend=False,
            ))
            fig.update_layout(
                title=dict(text="정책전략 그룹 비율",
                           font=dict(size=10, color="#1E3A5F")),
                margin=dict(t=32, b=4, l=4, r=4), height=210,
                annotations=[dict(
                    text=f"전체<br>{len(df)}개교",
                    x=0.5, y=0.5,
                    font=dict(size=10, color="#1E3A5F"),
                    showarrow=False,
                )],
            )
            st.plotly_chart(fig, width="stretch")


# ── 우선지원 TOP 10 ────────────────────────────────────────────────────────────
def _render_top10(df: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:8px;'>🏫 우선지원 추천 TOP 10"
            "<span style='font-size:0.7rem;color:#999;font-weight:400;"
            "margin-left:6px;'>* 우선지원점수 기준</span></div>",
            unsafe_allow_html=True)

        top10 = (
            df[df["priority_score"].notna()]
            .nlargest(10, "priority_score")
            [["school_name", "sigungu", "CSI", "CDI",
              "priority_score", "priority_display"]]
            .reset_index(drop=True)
        )
        top10.index += 1

        rows = ""
        for rank, row in top10.iterrows():
            grade = row["priority_display"]
            color = PRIORITY_COLORS.get(grade, "#95A5A6")
            rows += (
                f"<tr>"
                f"<td style='text-align:center;font-weight:700;"
                f"color:#1E3A5F;'>{rank}</td>"
                f"<td style='font-weight:500;'>{row['school_name']}</td>"
                f"<td style='text-align:center;color:#4A5568;'>{row['sigungu']}</td>"
                f"<td style='text-align:center;'>{row['CSI']:.3f}</td>"
                f"<td style='text-align:center;'>{row['CDI']:.3f}</td>"
                f"<td style='text-align:center;font-weight:700;"
                f"color:{color};'>{row['priority_score']:.3f}</td>"
                f"<td style='text-align:center;'>"
                f"<span class='badge' style='background:{color};'>{grade}</span>"
                f"</td>"
                f"</tr>"
            )

        st.markdown(
            f"<div style='overflow-x:auto;'>"
            f"<table class='top10-table'>"
            f"<thead><tr>"
            f"<th>순위</th><th style='text-align:left;'>학교명</th>"
            f"<th>시군구</th><th>CSI</th><th>CDI</th>"
            f"<th>점수</th><th>등급</th>"
            f"</tr></thead>"
            f"<tbody>{rows}</tbody>"
            f"</table></div>",
            unsafe_allow_html=True)


# ── 분포 히스토그램 3종 ────────────────────────────────────────────────────────
def _render_distributions(df: pd.DataFrame):
    c1, c2, c3 = st.columns(3, gap="small")

    cfg = [
        (c1, "CSI",            "CSI 분포",         "#2980B9", "(0 ~ 1)"),
        (c2, "CDI",            "CDI 분포",         "#E67E22", "(0 ~ 1)"),
        (c3, "priority_score", "우선지원점수 분포", "#8E44AD", "(CDI − CSI)"),
    ]

    for col, xcol, title, color, unit in cfg:
        with col:
            with st.container(border=True):
                st.markdown(
                    f"<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
                    f"margin-bottom:4px;'>{title} "
                    f"<span style='font-size:0.7rem;color:#999;"
                    f"font-weight:400;'>{unit}</span></div>",
                    unsafe_allow_html=True)

                mean_v = df[xcol].mean()
                med_v  = df[xcol].median()
                std_v  = df[xcol].std()

                fig = px.histogram(df, x=xcol, nbins=20,
                                   color_discrete_sequence=[color])
                fig.add_vline(x=mean_v, line_dash="dash",
                              line_color="#2C3E50", line_width=1.5)
                if xcol == "priority_score":
                    fig.add_vline(x=0, line_dash="solid",
                                  line_color="#C0392B", line_width=1.5)

                fig.add_annotation(
                    x=0.98, y=0.97, xref="paper", yref="paper",
                    text=(f"평균 &nbsp;&nbsp; {mean_v:.3f}<br>"
                          f"중앙값 {med_v:.3f}<br>"
                          f"표준편차 {std_v:.3f}"),
                    showarrow=False, align="right",
                    font=dict(size=9, color="#4A5568"),
                    bgcolor="rgba(255,255,255,0.9)",
                    bordercolor="#E2E8F0", borderwidth=1,
                )

                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    margin=dict(t=8, b=8, l=8, r=8),
                    height=200,
                    xaxis_title="", yaxis_title="학교 수",
                    showlegend=False,
                )
                fig.update_xaxes(showgrid=False, tickfont=dict(size=9))
                fig.update_yaxes(showgrid=True, gridcolor="#EEE",
                                 tickfont=dict(size=9))
                st.plotly_chart(fig, width="stretch")


# ══════════════════════════════════════════════════════════════════════════════
# 지역별 분석 탭
# ══════════════════════════════════════════════════════════════════════════════

# 정책전략 그룹 순서 (지역별 분석 탭 전용)
_REGIONAL_STRAT = [
    "최우선 개입형", "우선 보완형", "인력 취약형", "접근성 보완형", "안정형",
]
_MATRIX_TYPES = ["고수요 보완형", "핵심 불균형형", "잠재 취약형"]


@st.cache_data(show_spinner=False)
def _build_sigungu_agg(_df_hash: int, df: pd.DataFrame) -> pd.DataFrame:
    """원본 학교 데이터로부터 시군구별 집계 DataFrame을 생성한다."""
    work = df.copy()
    work["sigungu"] = work["sigungu"].fillna("시군구 미상")

    # 기본 수치 집계
    base = (
        work.groupby("sigungu", as_index=False)
        .agg(
            학교수    = ("school_name",    "count"),
            평균CSI   = ("CSI",            "mean"),
            평균CDI   = ("CDI",            "mean"),
            평균PS    = ("priority_score", "mean"),
        )
        .round({"평균CSI": 3, "평균CDI": 3, "평균PS": 3})
    )

    # 정책전략 그룹별 학교 수
    strat_pivot = (
        work.groupby(["sigungu", "policy_strategy_group"])
        .size().unstack(fill_value=0).reset_index()
    )
    for g in _REGIONAL_STRAT:
        if g not in strat_pivot.columns:
            strat_pivot[g] = 0
    base = base.merge(strat_pivot[["sigungu"] + _REGIONAL_STRAT], on="sigungu", how="left")

    # 3×3 매트릭스 유형별 학교 수
    mat_pivot = (
        work.groupby(["sigungu", "supply_demand_matrix_3x3"])
        .size().unstack(fill_value=0).reset_index()
    )
    for m in _MATRIX_TYPES:
        if m not in mat_pivot.columns:
            mat_pivot[m] = 0
    base = base.merge(mat_pivot[["sigungu"] + _MATRIX_TYPES], on="sigungu", how="left")
    base[_REGIONAL_STRAT + _MATRIX_TYPES] = base[_REGIONAL_STRAT + _MATRIX_TYPES].fillna(0).astype(int)

    # 시군구별 평균 위경도 (학교 위경도 평균 사용)
    if "school_latitude" in work.columns and "school_longitude" in work.columns:
        coord = (
            work.groupby("sigungu")
            .agg(위도=("school_latitude", "mean"), 경도=("school_longitude", "mean"))
            .reset_index()
        )
    else:
        coord = pd.DataFrame(
            [{"sigungu": s, "위도": c[0], "경도": c[1]} for s, c in SIGUNGU_COORDS.items()]
        )
    base = base.merge(coord, on="sigungu", how="left")

    return base.sort_values("평균PS", ascending=False).reset_index(drop=True)


# ── 지역별 분석 메인 함수 ──────────────────────────────────────────────────────
def show_regional(df: pd.DataFrame):
    agg = _build_sigungu_agg(id(df), df)

    # 헤더
    st.markdown(
        "<h1 style='font-size:1.3rem;color:#1E3A5F;margin:0 0 2px 0;font-weight:700;'>"
        "지역별 상담지원 수급 불균형 분석"
        "</h1>"
        "<p style='color:#718096;font-size:0.78rem;margin:0 0 6px 0;'>"
        "시군구별 CSI·CDI·우선지원점수·정책전략 그룹 분포를 비교하여 "
        "지역 단위의 우선지원 필요성을 파악합니다."
        "</p>"
        "<div style='background:#EBF2FF;border-left:3px solid #2980B9;"
        "padding:6px 10px;border-radius:4px;font-size:0.75rem;"
        "color:#2C3E50;margin-bottom:14px;'>"
        "ℹ️ 지역 평균값은 해당 시군구 내 분석 대상 일반고등학교의 평균을 기준으로 산출되었습니다."
        "</div>",
        unsafe_allow_html=True,
    )

    # KPI 카드 4개
    _render_regional_kpi(agg)
    st.markdown("<div style='margin:25px 0;'></div>", unsafe_allow_html=True)

    # 중단: 지도(좌) | 지수비교+표(우)
    map_col, right_col = st.columns([1, 1], gap="small")
    with map_col:
        _render_regional_map(agg)
    with right_col:
        _render_regional_index_bar(agg)

        _render_regional_table(agg)

    # 하단: 개입형 바(좌) | 누적 바(중) | 인사이트(우)

    bot_l, bot_m, bot_r = st.columns([1, 1.2, 0.9], gap="small")
    with bot_l:
        _render_top_intervention_bar(agg)
    with bot_m:
        _render_strategy_stacked_bar(agg)
    with bot_r:
        _render_regional_insight(agg)

    st.markdown(
        "<div class='footer-note'>"
        "※ 우선지원점수 = CDI − CSI (음수 가능) "
        "| 학교 수가 적은 시군구의 평균값 해석에 주의가 필요합니다."
        "</div>",
        unsafe_allow_html=True,
    )


# ── KPI 카드 ───────────────────────────────────────────────────────────────────
def _render_regional_kpi(agg: pd.DataFrame):
    n_sgg       = len(agg)
    top_ps_row  = agg.iloc[0]
    top_ps_sgg  = top_ps_row["sigungu"]
    top_ps_val  = top_ps_row["평균PS"]

    top_intv_idx = agg["최우선 개입형"].idxmax()
    top_intv_sgg = agg.loc[top_intv_idx, "sigungu"]
    top_intv_cnt = int(agg.loc[top_intv_idx, "최우선 개입형"])

    top_manp_idx = agg["인력 취약형"].idxmax() if "인력 취약형" in agg.columns else None
    top_manp_sgg = agg.loc[top_manp_idx, "sigungu"] if top_manp_idx is not None else "-"
    top_manp_cnt = int(agg.loc[top_manp_idx, "인력 취약형"]) if top_manp_idx is not None else 0

    defs = [
        ("분석 대상 시군구 수",     f"{n_sgg}개 시군구",  "경남 일반고 전수 대상",
         "🗺️", "#EBF2FF", "#2980B9"),
        ("우선지원점수 최고 지역",  top_ps_sgg,           f"평균 {top_ps_val:+.3f}",
         "📍", "#FEF9E7", "#F39C12"),
        ("최우선 개입형 최다 지역", top_intv_sgg,         f"{top_intv_cnt}개교",
         "⚠️", "#FDECEA", "#C0392B"),
        ("인력 취약형 최다 지역",   top_manp_sgg,         f"{top_manp_cnt}개교",
         "👥", "#F3E8FF", "#8E44AD"),
    ]

    cols = st.columns(4, gap="small")
    for col, (label, value, sub, icon, bg, color) in zip(cols, defs):
        col.markdown(
            f'<div class="metric-card" style="border-left:4px solid {color};">'
            f'<div class="metric-icon" style="background:{bg};">{icon}</div>'
            f'<div>'
            f'<div class="metric-label">{label}</div>'
            f'<div class="metric-value" style="color:{color};font-size:1.25rem;">{value}</div>'
            f'<div class="metric-sub">{sub}</div>'
            f'</div></div>',
            unsafe_allow_html=True,
        )


# ── 시군구별 버블맵 ────────────────────────────────────────────────────────────
def _render_regional_map(agg: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:6px;'>📍 시군구별 평균 우선지원점수 분포</div>",
            unsafe_allow_html=True,
        )
        valid = agg.dropna(subset=["위도", "경도"]).copy()
        if valid.empty:
            st.info("위경도 데이터가 없어 버블맵을 표시할 수 없습니다.")
            return

        valid["툴팁"] = valid.apply(
            lambda r: (
                f"<b>{r['sigungu']}</b><br>"
                f"학교 수: {int(r['학교수'])}개교<br>"
                f"평균 CSI: {r['평균CSI']:.3f}<br>"
                f"평균 CDI: {r['평균CDI']:.3f}<br>"
                f"평균 우선지원점수: {r['평균PS']:+.3f}<br>"
                f"최우선 개입형: {int(r['최우선 개입형'])}개교<br>"
                f"우선 보완형: {int(r['우선 보완형'])}개교<br>"
                f"고수요 보완형: {int(r['고수요 보완형'])}개교"
            ),
            axis=1,
        )

        fig = px.scatter_mapbox(
            valid,
            lat="위도", lon="경도",
            size="학교수",
            color="평균PS",
            color_continuous_scale=["#C0392B", "#E67E22", "#F4D03F", "#2980B9"],
            color_continuous_midpoint=valid["평균PS"].median(),
            size_max=42,
            zoom=7.6,
            center={"lat": 35.23, "lon": 128.15},
            mapbox_style="carto-positron",
            custom_data=["툴팁"],
            text="sigungu",
        )
        fig.update_traces(
            hovertemplate="%{customdata[0]}<extra></extra>",
            textfont=dict(size=10, color="#1E3A5F"),
            mode="markers+text",
            textposition="top center",
        )
        fig.update_layout(
            height=430,
            margin={"r": 0, "t": 0, "l": 0, "b": 0},
            coloraxis_colorbar=dict(
                title="평균<br>우선지원점수",
                thickness=12, len=0.55, x=1.01,
                tickfont=dict(size=9),
            ),
        )
        st.plotly_chart(fig, width="stretch")
        st.caption("버블 크기: 학교 수 · 색상: 평균 우선지원점수 (파랑→빨강: 낮음→높음)")


# ── 주요 시군별 지수 비교 grouped bar ─────────────────────────────────────────
def _render_regional_index_bar(agg: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:6px;'>📊 주요 시군별 지수 비교</div>",
            unsafe_allow_html=True,
        )
        # 우선지원점수 내림차순 상위 10개 시군구
        top = agg.head(10)

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="평균 CSI", x=top["sigungu"], y=top["평균CSI"],
            marker_color="#1ABC9C", yaxis="y1",
            text=top["평균CSI"].round(2), textposition="outside",
            textfont=dict(size=8),
        ))
        fig.add_trace(go.Bar(
            name="평균 CDI", x=top["sigungu"], y=top["평균CDI"],
            marker_color="#E67E22", yaxis="y1",
            text=top["평균CDI"].round(2), textposition="outside",
            textfont=dict(size=8),
        ))
        fig.add_trace(go.Scatter(
            name="평균 우선지원점수", x=top["sigungu"], y=top["평균PS"],
            mode="lines+markers+text",
            line=dict(color="#C0392B", width=2),
            marker=dict(size=6, color="#C0392B"),
            text=top["평균PS"].apply(lambda v: f"{v:.2f}"),
            textposition="top center",
            textfont=dict(size=8, color="#C0392B"),
            yaxis="y2",
        ))

        fig.update_layout(
            barmode="group",
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=12, b=4, l=4, r=40), height=245,
            legend=dict(
                orientation="h", yanchor="bottom", y=1.01,
                xanchor="right", x=1, font=dict(size=8),
            ),
            xaxis=dict(tickfont=dict(size=8.5), showgrid=False),
            yaxis=dict(
                title="지수 (0~1)", range=[0, 1.05],
                tickfont=dict(size=8), showgrid=True, gridcolor="#EEE",
            ),
            yaxis2=dict(
                title="우선지원점수", side="right", overlaying="y",
                range=[-0.75, 0.05],
                tickfont=dict(size=8), showgrid=False,
                zeroline=True, zerolinecolor="#DDD",
            ),
        )
        st.plotly_chart(fig, width="stretch")
        st.caption("우선지원점수(CDI−CSI)는 우측 축 기준 · 우선지원점수가 높을수록(덜 음수) 수요 대비 공급 부족 가능성이 상대적으로 높음")


# ── 시군별 우선지원 현황 표 ────────────────────────────────────────────────────
def _render_regional_table(agg: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:6px;'>📋 시군별 우선지원 현황"
            "<span style='font-size:0.7rem;color:#999;font-weight:400;"
            "margin-left:6px;'>* 우선지원점수 내림차순</span></div>",
            unsafe_allow_html=True,
        )
        rows = ""
        for i, r in agg.iterrows():
            ps_color = (
                "#C0392B" if r["평균PS"] >= agg["평균PS"].quantile(0.75)
                else "#E67E22" if r["평균PS"] >= agg["평균PS"].median()
                else "#27AE60"
            )
            rows += (
                f"<tr>"
                f"<td style='font-weight:600;color:#1E3A5F;'>{r['sigungu']}</td>"
                f"<td style='text-align:center;'>{int(r['학교수'])}</td>"
                f"<td style='text-align:center;'>{r['평균CSI']:.3f}</td>"
                f"<td style='text-align:center;'>{r['평균CDI']:.3f}</td>"
                f"<td style='text-align:center;font-weight:700;color:{ps_color};'>{r['평균PS']:+.3f}</td>"
                f"<td style='text-align:center;'>{int(r['최우선 개입형'])}</td>"
                f"<td style='text-align:center;'>{int(r['우선 보완형'])}</td>"
                f"<td style='text-align:center;'>{int(r['인력 취약형'])}</td>"
                f"<td style='text-align:center;'>{int(r['접근성 보완형'])}</td>"
                f"</tr>"
            )
        st.markdown(
            "<div style='overflow-y:auto;max-height:310px;'>"
            "<table class='top10-table'>"
            "<thead><tr style='position:sticky;top:0;background:#EBF2FF;'>"
            "<th style='text-align:left;'>시군구</th>"
            "<th>학교수</th><th>CSI</th><th>CDI</th>"
            "<th>우선지원점수</th><th>최우선개입</th>"
            "<th>우선보완</th><th>인력취약</th><th>접근성보완</th>"
            "</tr></thead>"
            f"<tbody>{rows}</tbody>"
            "</table></div>",
            unsafe_allow_html=True,
        )


# ── 최우선 개입형 horizontal bar ───────────────────────────────────────────────
def _render_top_intervention_bar(agg: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:6px;'>🔴 시군구별 최우선 개입형 학교 수</div>",
            unsafe_allow_html=True,
        )
        data = (
            agg[agg["최우선 개입형"] > 0]
            .nlargest(10, "최우선 개입형")
            .sort_values("최우선 개입형")
        )
        if data.empty:
            st.info("최우선 개입형 학교가 없습니다.")
            return

        fig = go.Figure(go.Bar(
            x=data["최우선 개입형"],
            y=data["sigungu"],
            orientation="h",
            marker_color="#C0392B",
            text=data["최우선 개입형"].astype(int),
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=8, b=8, l=4, r=30), height=280,
            xaxis=dict(
                title="학교 수 (개)", tickfont=dict(size=9),
                showgrid=True, gridcolor="#EEE",
                range=[0, data["최우선 개입형"].max() * 1.35],
            ),
            yaxis=dict(tickfont=dict(size=9.5)),
            showlegend=False,
        )
        st.plotly_chart(fig, width="stretch")


# ── 시군구별 정책전략 그룹 100% 누적 bar ──────────────────────────────────────
def _render_strategy_stacked_bar(agg: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:6px;'>📊 시군구별 정책전략 그룹 분포</div>",
            unsafe_allow_html=True,
        )
        # 최우선 개입형 비율 내림차순 정렬
        df_stk = agg.copy()
        df_stk["합계"] = df_stk[_REGIONAL_STRAT].sum(axis=1)
        for g in _REGIONAL_STRAT:
            df_stk[f"{g}_pct"] = (df_stk[g] / df_stk["합계"] * 100).round(1)
        df_stk = df_stk.sort_values("최우선 개입형_pct", ascending=True)

        fig = go.Figure()
        for g in _REGIONAL_STRAT:
            fig.add_trace(go.Bar(
                name=g,
                x=df_stk[f"{g}_pct"],
                y=df_stk["sigungu"],
                orientation="h",
                marker_color=STRATEGY_COLORS.get(g, "#BDC3C7"),
                text=df_stk[f"{g}_pct"].apply(lambda v: f"{v:.0f}%" if v >= 10 else ""),
                textposition="inside",
                textfont=dict(size=8, color="white"),
                hovertemplate="<b>%{y}</b><br>" + g + ": %{x:.1f}%<extra></extra>",
            ))

        fig.update_layout(
            barmode="stack",
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=8, b=8, l=4, r=4), height=320,
            xaxis=dict(
                title="%", range=[0, 100],
                tickfont=dict(size=8), showgrid=True, gridcolor="#EEE",
            ),
            yaxis=dict(tickfont=dict(size=9)),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.01,
                xanchor="left", x=0, font=dict(size=7.5),
            ),
        )
        st.plotly_chart(fig, width="stretch")
        st.caption("비율 그래프이므로 학교 수가 적은 지역의 해석에는 주의가 필요합니다.")


# ── 지역별 핵심 인사이트 박스 ──────────────────────────────────────────────────
def _render_regional_insight(agg: pd.DataFrame):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:10px;'>💡 지역별 핵심 인사이트</div>",
            unsafe_allow_html=True,
        )

        top_ps      = agg.iloc[0]
        top_intv    = agg.loc[agg["최우선 개입형"].idxmax()]
        top_high    = agg.loc[agg["고수요 보완형"].idxmax()]
        top_manpow  = agg.loc[agg["인력 취약형"].idxmax()]
        top_access  = agg.loc[agg["접근성 보완형"].idxmax()]

        insights = [
            (
                "⚠️",
                "#C0392B",
                f"<b>최우선 검토 필요 지역</b>: {top_intv['sigungu']}({int(top_intv['최우선 개입형'])}개교)은 "
                "최우선 개입형 학교가 가장 많아 지역 단위의 집중적인 상담지원 강화를 검토할 수 있습니다.",
            ),
            (
                "📍",
                "#E67E22",
                f"<b>수급 불균형 주의 지역</b>: 평균 우선지원점수가 가장 높은 지역은 {top_ps['sigungu']}"
                f"({top_ps['평균PS']:+.3f})이며, 수요 대비 공급 부족 가능성이 상대적으로 높습니다.",
            ),
            (
                "📊",
                "#8E44AD",
                f"<b>고수요 보완 필요</b>: {top_high['sigungu']}({int(top_high['고수요 보완형'])}개교)은 "
                "고수요 보완형 학교가 많아 상담수요 대비 공급 확충 방안 모색이 필요할 수 있습니다.",
            ),
            (
                "👥",
                "#2980B9",
                f"<b>인력 배치 검토 지역</b>: {top_manpow['sigungu']}({int(top_manpow['인력 취약형'])}개교)은 "
                "인력 취약형 학교가 많아 전문상담교사 배치 또는 순회상담 연계 강화를 검토해 볼 수 있습니다.",
            ),
            (
                "🔗",
                "#1ABC9C",
                f"<b>접근성 개선 검토</b>: {top_access['sigungu']}({int(top_access['접근성 보완형'])}개교)은 "
                "접근성 보완형 학교가 많아 Wee센터 접근성 개선 또는 원격상담 지원 연계가 필요할 수 있습니다.",
            ),
        ]

        for icon, color, text in insights:
            st.markdown(
                f"<div style='display:flex;gap:8px;margin-bottom:10px;"
                f"padding:8px 10px;background:#F7FAFC;"
                f"border-radius:6px;border-left:3px solid {color};'>"
                f"<span style='font-size:1rem;flex-shrink:0;'>{icon}</span>"
                f"<p style='font-size:0.75rem;color:#2D3748;margin:0;line-height:1.5;'>{text}</p>"
                f"</div>",
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# ── 8. 학교 검색 탭 ──────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

_CSI_SUB = ["counseling_staff_supply_score", "wee_class_score", "wee_center_access_score"]
_CDI_SUB = ["demand_size_score", "counseling_use_score", "school_violence_risk_score"]
_SUB_LABELS = {
    "counseling_staff_supply_score": "상담인력 공급",
    "wee_class_score":               "Wee클래스",
    "wee_center_access_score":       "Wee센터 접근성",
    "demand_size_score":             "수요 규모",
    "counseling_use_score":          "상담 이용률",
    "school_violence_risk_score":    "학교폭력 위험",
}
_DETAIL_COLS = {
    "school_code":                   "학교코드",
    "school_name":                   "학교명",
    "sido":                          "시도",
    "sigungu":                       "시군구",
    "counseling_staff_supply_score": "상담인력 공급점수",
    "wee_class_score":               "Wee클래스 점수",
    "wee_center_access_score":       "Wee센터 접근성",
    "demand_size_score":             "수요 규모",
    "counseling_use_score":          "상담 이용률",
    "school_violence_risk_score":    "학교폭력 위험",
    "CSI":                           "CSI",
    "CDI":                           "CDI",
    "priority_score":                "우선지원점수",
    "priority_level":                "우선지원등급",
    "policy_strategy_group":         "정책전략 그룹",
    "supply_demand_matrix_3x3":      "수요공급 매트릭스",
}


def show_school_search(df: pd.DataFrame):
    # ── 헤더 ─────────────────────────────────────────────────────────────────
    st.markdown(
        "<h1 style='font-size:1.3rem;color:#1E3A5F;margin:0 0 2px 0;font-weight:700;'>"
        "학교별 상담지원 인프라 진단"
        "</h1>"
        "<p style='color:#718096;font-size:0.78rem;margin:0 0 14px 0;'>"
        "개별 학교의 상담공급지수(CSI), 상담수요지수(CDI), 우선지원점수, "
        "세부 구성 점수, 정책전략 그룹과 정책 피드백을 확인합니다."
        "</p>",
        unsafe_allow_html=True,
    )

    # ── 학교 선택 selectbox ──────────────────────────────────────────────────
    df_sorted = df.sort_values("priority_score", ascending=False).reset_index(drop=True)

    # 동명 학교 구분: 중복 레이블에 school_code 접미 추가
    base_labels = df_sorted.apply(
        lambda r: f"{r['school_name']} ({r['sigungu']})", axis=1
    ).tolist()
    from collections import Counter
    cnt = Counter(base_labels)
    seen: dict = {}
    options = []
    for lbl, code in zip(base_labels, df_sorted["school_code"].tolist()):
        if cnt[lbl] > 1:
            seen[lbl] = seen.get(lbl, 0) + 1
            options.append(f"{lbl} [{code}]")
        else:
            options.append(lbl)
    codes = df_sorted["school_code"].tolist()

    sel_opt = st.selectbox(
        "🔍 학교 선택",
        options,
        index=0,
        help="우선지원점수 내림차순 정렬 | 동명 학교는 학교코드로 구분됩니다.",
    )
    sel_code = codes[options.index(sel_opt)]
    row_df = df[df["school_code"] == sel_code]
    if row_df.empty:
        st.warning("선택한 학교 데이터를 찾을 수 없습니다.")
        return
    row = row_df.iloc[0]


    # ── Row 1: 기본 정보 카드 | 학교 위치 지도(정사각) | KPI 카드 ────────────
    info_col, map_col, kpi_col = st.columns([1.3, 1.0, 1.7], gap="small")
    with info_col:
        _render_school_info_card(row)
    with map_col:
        _render_single_school_map(row, height=192)
    with kpi_col:
        _render_school_kpi_cards(row, df)

    # ── Row 2: 좌(하위점수+평균비교+유사학교) | 우(정책피드백+세부테이블+지도) ─
    left_col, right_col = st.columns([1, 1], gap="small")

    with left_col:
        _render_sub_scores_chart(row, df)
        _render_school_avg_comparison(row, df)
        _render_similar_schools(row, df)

    with right_col:
        _render_policy_feedback(row)
        _render_school_detail_table(row)

    # ── footer ────────────────────────────────────────────────────────────────
    st.markdown(
        "<div class='footer-note'>"
        "※ 본 분석 결과는 정책 검토용 참고 자료이며 실제 지원 확정 기준이 아닙니다. "
        "| CSI·CDI·우선지원점수는 2025년 기준 산출값입니다."
        "</div>",
        unsafe_allow_html=True,
    )


# ── 학교 기본 정보 카드 ────────────────────────────────────────────────────────
def _render_school_info_card(row: pd.Series):
    def _v(col, default="확인 필요"):
        val = row.get(col, None)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return default
        return str(val)

    pdisp  = PRIORITY_DISPLAY.get(_v("priority_level", ""), _v("priority_level"))
    pcolor = PRIORITY_COLORS.get(pdisp, "#718096")
    sgroup = _v("policy_strategy_group")
    scolor = STRATEGY_COLORS.get(sgroup, "#718096")

    info_rows = [
        ("학교코드",      _v("school_code")),
        ("시도",          _v("sido")),
        ("시군구",        _v("sigungu")),
        ("수요공급 유형", _v("supply_demand_type")),
        ("매트릭스 유형", _v("supply_demand_matrix_3x3")),
    ]
    rows_html = "".join(
        f"<div style='display:flex;justify-content:space-between;align-items:center;"
        f"padding:5px 0;border-bottom:1px solid #F0F4F8;font-size:0.77rem;'>"
        f"<span style='color:#718096;flex-shrink:0;margin-right:6px;'>{lbl}</span>"
        f"<span style='color:#2D3748;font-weight:600;text-align:right;'>{val}</span>"
        f"</div>"
        for lbl, val in info_rows
    )
    badge_html = (
        f"<div style='display:flex;gap:6px;margin-top:10px;flex-wrap:wrap;'>"
        f"<span style='background:{pcolor};color:white;padding:3px 10px;"
        f"border-radius:12px;font-size:0.72rem;font-weight:700;'>{pdisp}</span>"
        f"<span style='background:{scolor};color:white;padding:3px 10px;"
        f"border-radius:12px;font-size:0.72rem;font-weight:700;'>{sgroup}</span>"
        f"</div>"
    )
    with st.container(border=True):
        st.markdown(
            f"<div style='font-size:1.0rem;font-weight:700;color:#1E3A5F;"
            f"margin-bottom:10px;padding-bottom:8px;border-bottom:2px solid #2E5FA3;'>"
            f"🏫 {_v('school_name')}</div>"
            + rows_html + badge_html +
            "<div style='margin-bottom:10px;'></div>",
            unsafe_allow_html=True,
        )


# ── KPI 5개 카드 ───────────────────────────────────────────────────────────────
def _render_school_kpi_cards(row: pd.Series, df_all: pd.DataFrame):
    def _fv(col):
        try:
            return round(float(row.get(col, 0) or 0), 3)
        except (TypeError, ValueError):
            return 0.0

    csi   = _fv("CSI"); cdi = _fv("CDI"); ps = _fv("priority_score")
    pdisp = PRIORITY_DISPLAY.get(str(row.get("priority_level", "")),
                                 str(row.get("priority_level", "확인 필요")))
    sg    = str(row.get("policy_strategy_group", "확인 필요"))

    avg_csi = round(float(df_all["CSI"].mean()), 3)
    avg_cdi = round(float(df_all["CDI"].mean()), 3)
    avg_ps  = round(float(df_all["priority_score"].mean()), 3)

    pcolor = PRIORITY_COLORS.get(pdisp, "#718096")
    scolor = STRATEGY_COLORS.get(sg, "#718096")

    def _card(color, label, value, sub, min_h="95px"):
        return (
            f"<div style='background:white;border-radius:10px;padding:12px 10px;"
            f"box-shadow:0 2px 8px rgba(0,0,0,0.08);border-top:3px solid {color};"
            f"min-height:{min_h};'>"
            f"<div style='font-size:0.70rem;color:#718096;margin-bottom:4px;'>{label}</div>"
            f"<div style='font-size:1.1rem;font-weight:700;color:{color};"
            f"line-height:1.2;word-break:keep-all;'>{value}</div>"
            f"<div style='font-size:0.65rem;color:#A0AEC0;margin-top:2px;'>{sub}</div>"
            f"</div>"
        )

    # 외부 2열: 좌(2x2 그리드) | 우(정책전략 그룹 단일 카드)
    grid_col, sg_col = st.columns([3.6, 1.4], gap="small")

    with grid_col:
        # 2x2 그리드 — 1행
        r1c1, r1c2 = st.columns(2, gap="small")
        with r1c1:
            st.markdown(_card("#1ABC9C", "CSI", f"{csi:.3f}", f"경남 평균 {avg_csi:.3f}", "120px"),
                        unsafe_allow_html=True)
        with r1c2:
            st.markdown(_card("#E67E22", "CDI", f"{cdi:.3f}", f"경남 평균 {avg_cdi:.3f}", "120px"),
                        unsafe_allow_html=True)
        # 행 간격
        st.markdown("<div style='margin:22px 0;'></div>", unsafe_allow_html=True)
        # 2x2 그리드 — 2행
        r2c1, r2c2 = st.columns(2, gap="small")
        with r2c1:
            st.markdown(_card("#C0392B", "우선지원점수", f"{ps:.3f}", f"경남 평균 {avg_ps:.3f}", "120px"),
                        unsafe_allow_html=True)
        with r2c2:
            st.markdown(_card(pcolor, "우선지원등급", pdisp, "등급", "120px"),
                        unsafe_allow_html=True)

    with sg_col:
        # 2x2 전체 높이(115px × 2 + gap 12px + spacer 20px ≈ 262px)에 맞춘 단일 카드
        st.markdown(_card(scolor, "정책전략 그룹", sg, "", "254px"),
                    unsafe_allow_html=True)


# ── 6개 하위 점수 수평 바 차트 ────────────────────────────────────────────────
def _render_sub_scores_chart(row: pd.Series, df_all: pd.DataFrame):
    sub_cols = _CSI_SUB + _CDI_SUB
    labels   = [_SUB_LABELS[c] for c in sub_cols]
    school_v = [round(float(row.get(c, 0) or 0), 3) for c in sub_cols]
    avg_v    = [round(float(df_all[c].mean()) if c in df_all.columns else 0, 3)
                for c in sub_cols]
    bar_colors  = ["#1ABC9C"] * 3 + ["#E67E22"] * 3
    avg_colors  = ["rgba(26,188,156,0.3)"] * 3 + ["rgba(230,126,34,0.3)"] * 3

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="해당 학교", y=labels, x=school_v, orientation="h",
        marker_color=bar_colors,
        text=[f"{v:.3f}" for v in school_v], textposition="outside",
        textfont=dict(size=9),
    ))
    fig.add_trace(go.Bar(
        name="경남 평균", y=labels, x=avg_v, orientation="h",
        marker_color=avg_colors, marker_line_color="#AAAAAA", marker_line_width=0.5,
        text=[f"{v:.3f}" for v in avg_v], textposition="outside",
        textfont=dict(size=9),
    ))
    fig.update_layout(
        title=dict(text="📊 6개 하위 지표 프로필 (경남 평균 비교)",
                   font=dict(size=12, color="#1E3A5F"), x=0),
        barmode="group", height=340,
        margin=dict(l=10, r=50, t=40, b=30),
        xaxis=dict(range=[0, 1.2], title="점수 (0~1)", tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=10)),
        legend=dict(orientation="h", y=-0.10, x=0, font=dict(size=9)),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
        shapes=[dict(type="line", x0=0.5, x1=0.5, y0=-0.5, y1=5.5,
                     line=dict(color="#D5D8DC", width=1, dash="dot"))],
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")


# ── 정책 피드백 카드 ──────────────────────────────────────────────────────────
def _render_policy_feedback(row: pd.Series):
    def _v(col, default=""):
        val = row.get(col, None)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return default
        return str(val)

    action   = _v("policy_action_type",         "정보 없음")
    rec      = _v("policy_recommendation",       "")
    reason   = _v("policy_reason",               "")
    desc     = _v("policy_strategy_description", "")
    tags_raw = _v("policy_strategy_tags",        "")
    tags     = [t.strip() for t in tags_raw.split(";") if t.strip()]

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:10px;'>"
            "📋 정책 피드백</div>",
            unsafe_allow_html=True,
        )
        # 조치 유형 배지
        st.markdown(
            f"<div style='margin-bottom:10px;'>"
            f"<span style='background:#EBF2FF;color:#2980B9;padding:3px 12px;"
            f"border-radius:10px;font-size:0.73rem;font-weight:700;'>"
            f"조치 유형 · {action}</span></div>",
            unsafe_allow_html=True,
        )
        # 추천 방향 (세미콜론 구분 → 불릿)
        if rec:
            bullets = [b.strip() for b in rec.split(";") if b.strip()]
            b_html = "".join(
                f"<div style='display:flex;gap:6px;margin-bottom:6px;'>"
                f"<span style='color:#2980B9;flex-shrink:0;margin-top:1px;'>●</span>"
                f"<span style='font-size:0.77rem;color:#2D3748;line-height:1.5;'>{b}</span>"
                f"</div>"
                for b in bullets
            )
            st.markdown(
                "<div style='font-size:0.73rem;font-weight:700;color:#4A5568;"
                "margin-bottom:6px;'>추천 정책 방향</div>" + b_html,
                unsafe_allow_html=True,
            )
        # 판단 근거
        if reason:
            st.markdown(
                f"<div style='background:#F7FAFC;border-left:3px solid #1ABC9C;"
                f"padding:7px 10px;border-radius:4px;margin-top:8px;'>"
                f"<div style='font-size:0.71rem;font-weight:700;color:#4A5568;"
                f"margin-bottom:3px;'>판단 근거</div>"
                f"<div style='font-size:0.75rem;color:#2D3748;line-height:1.55;'>{reason}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        # 전략 설명
        if desc:
            st.markdown(
                f"<div style='background:#F0FFF4;border-left:3px solid #27AE60;"
                f"padding:7px 10px;border-radius:4px;margin-top:8px;'>"
                f"<div style='font-size:0.71rem;font-weight:700;color:#4A5568;"
                f"margin-bottom:3px;'>전략 설명</div>"
                f"<div style='font-size:0.75rem;color:#2D3748;line-height:1.55;'>{desc}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        # 태그
        if tags:
            tag_html = " ".join(
                f"<span style='background:#EBF2FF;color:#2C5282;padding:2px 8px;"
                f"border-radius:8px;font-size:0.67rem;'>{t}</span>"
                for t in tags[:8]
            )
            st.markdown(
                f"<div style='margin-top:10px;line-height:2;'>{tag_html}</div>",
                unsafe_allow_html=True,
            )
        st.markdown(
            "<div style='margin-top:12px;padding-top:6px;border-top:1px solid #F0F4F8;"
            "font-size:0.67rem;color:#A0AEC0;'>"
            "※ 본 피드백은 정책 검토용 분석 결과이며 실제 지원 확정 기준이 아닙니다.</div>",
            unsafe_allow_html=True,
        )


# ── 경남 평균 대비 주요 지표 비교 ────────────────────────────────────────────
def _render_school_avg_comparison(row: pd.Series, df_all: pd.DataFrame):
    metrics  = ["CSI", "CDI", "priority_score"]
    m_labels = ["CSI", "CDI", "우선지원점수"]
    try:
        school_v = [round(float(row.get(m, 0) or 0), 3) for m in metrics]
    except (TypeError, ValueError):
        school_v = [0.0, 0.0, 0.0]
    avg_v = [round(float(df_all[m].mean()), 3) for m in metrics]

    s_colors = ["#1ABC9C", "#E67E22", "#C0392B"]
    a_colors = ["rgba(26,188,156,0.3)", "rgba(230,126,34,0.3)", "rgba(192,57,43,0.3)"]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="해당 학교", x=m_labels, y=school_v,
        marker_color=s_colors,
        text=[f"{v:.3f}" for v in school_v], textposition="outside",
        textfont=dict(size=10),
    ))
    fig.add_trace(go.Bar(
        name="경남 평균", x=m_labels, y=avg_v,
        marker_color=a_colors, marker_line_color="#AAAAAA", marker_line_width=0.5,
        text=[f"{v:.3f}" for v in avg_v], textposition="outside",
        textfont=dict(size=10),
    ))
    y_min = min(min(school_v), min(avg_v)) - 0.1
    y_max = max(max(school_v), max(avg_v)) + 0.15
    fig.update_layout(
        title=dict(text="📈 경남 평균 대비 주요 지표 비교",
                   font=dict(size=12, color="#1E3A5F"), x=0),
        barmode="group", height=300,
        margin=dict(l=10, r=20, t=40, b=40),
        yaxis=dict(title="값", range=[y_min, y_max], tickfont=dict(size=9),
                   zeroline=True, zerolinecolor="#BDC3C7", zerolinewidth=1),
        legend=dict(orientation="h", y=-0.20, x=0, font=dict(size=9)),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")


# ── 세부 지표 테이블 ──────────────────────────────────────────────────────────
def _render_school_detail_table(row: pd.Series):
    score_cols = set(_CSI_SUB + _CDI_SUB + ["CSI", "CDI", "priority_score"])
    rows_html = ""
    for col, kor in _DETAIL_COLS.items():
        if col not in row.index:
            continue
        val = row[col]
        if col in score_cols:
            try:
                val = f"{float(val):.3f}"
            except (TypeError, ValueError):
                val = "확인 필요"
        else:
            val = str(val) if not (isinstance(val, float) and pd.isna(val)) else "확인 필요"
        rows_html += (
            f"<tr>"
            f"<td style='color:#718096;font-size:0.76rem;padding:5px 8px;"
            f"border-bottom:1px solid #F0F4F8;white-space:nowrap;'>{kor}</td>"
            f"<td style='color:#2D3748;font-weight:600;font-size:0.76rem;"
            f"padding:5px 8px;border-bottom:1px solid #F0F4F8;'>{val}</td>"
            f"</tr>"
        )
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📋 세부 지표 테이블</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<table style='width:100%;border-collapse:collapse;'>{rows_html}</table>",
            unsafe_allow_html=True,
        )


# ── 학교 위치 지도 ────────────────────────────────────────────────────────────
def _render_single_school_map(row: pd.Series, height: int = 320):
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📍 학교 위치</div>",
            unsafe_allow_html=True,
        )
        try:
            lat = float(row.get("school_latitude", 0) or 0)
            lon = float(row.get("school_longitude", 0) or 0)
            has_coord = lat != 0.0 and lon != 0.0
        except (TypeError, ValueError):
            has_coord = False

        if not has_coord:
            st.info("해당 학교의 위경도 정보가 없어 지도 표시를 생략합니다.")
            return

        pdisp  = PRIORITY_DISPLAY.get(str(row.get("priority_level", "")), "")
        color  = PRIORITY_COLORS_PYDECK.get(pdisp, DEFAULT_COLOR_PYDECK)
        map_df = pd.DataFrame([{
            "school_name": str(row.get("school_name", "")),
            "sigungu":     str(row.get("sigungu", "")),
            "lat": lat, "lon": lon, "color": color,
        }])
        layer = pdk.Layer(
            "ScatterplotLayer", data=map_df,
            get_position=["lon", "lat"], get_color="color",
            get_radius=600, pickable=True, opacity=0.9, stroked=True, filled=True,
            line_width_min_pixels=2, get_line_color=[255, 255, 255, 200],
            auto_highlight=True,
        )
        view  = pdk.ViewState(latitude=lat, longitude=lon, zoom=11, pitch=0)
        deck  = pdk.Deck(
            layers=[layer], initial_view_state=view,
            map_style="https://basemaps.cartocdn.com/gl/voyager-gl-style/style.json",
            tooltip={
                "html": "<b>{school_name}</b><br/>{sigungu}",
                "style": {"backgroundColor": "#1E3A5F", "color": "white", "fontSize": "12px"},
            },
        )
        st.pydeck_chart(deck, width="stretch", height=height)


# ── 유사 학교 비교 표 ─────────────────────────────────────────────────────────
def _render_similar_schools(row: pd.Series, df: pd.DataFrame):
    group      = row.get("policy_strategy_group", None)
    school_code = row.get("school_code", None)

    with st.container(border=True):
        scolor = STRATEGY_COLORS.get(str(group), "#2980B9")
        st.markdown(
            f"<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            f"padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:6px;'>"
            f"🔗 같은 정책전략 그룹 내 비교 학교</div>"
            f"<div style='font-size:0.74rem;color:#718096;margin-bottom:8px;'>"
            f"그룹: <b style='color:{scolor};'>{group}</b> "
            f"| 우선지원점수 상위 5개교 (★ 선택 학교)</div>",
            unsafe_allow_html=True,
        )
        if group is None:
            st.info("정책전략 그룹 정보가 없습니다.")
            return

        sim = (
            df[df["policy_strategy_group"] == group]
            .sort_values("priority_score", ascending=False)
            .head(6)
            .reset_index(drop=True)
        )
        disp_cols = ["school_code", "school_name", "sigungu", "CSI", "CDI",
                     "priority_score", "priority_level"]
        avail     = [c for c in disp_cols if c in sim.columns]
        sim_disp  = sim[avail].copy()

        # 선택 학교 표시용 마커 추가
        sim_disp["선택"] = sim_disp["school_code"].apply(
            lambda c: "★" if str(c) == str(school_code) else ""
        )
        sim_disp = sim_disp.drop(columns=["school_code"])
        sim_disp = sim_disp.rename(columns={
            "school_name":    "학교명",
            "sigungu":        "시군구",
            "priority_score": "우선지원점수",
            "priority_level": "등급",
        })
        if "등급" in sim_disp.columns:
            sim_disp["등급"] = sim_disp["등급"].map(PRIORITY_DISPLAY).fillna(sim_disp["등급"])
        for c in ["CSI", "CDI", "우선지원점수"]:
            if c in sim_disp.columns:
                sim_disp[c] = sim_disp[c].apply(
                    lambda x: f"{float(x):.3f}" if pd.notna(x) else "-"
                )
        col_order = ["선택"] + [c for c in sim_disp.columns if c != "선택"]
        st.dataframe(sim_disp[col_order].reset_index(drop=True),
                     use_container_width=True, height=252)


# ══════════════════════════════════════════════════════════════════════════════
# ── 9. 유형 분석 탭 ──────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

# 3x3 셀 이름 매핑 (cdi_relative_level, csi_relative_level)
_MATRIX_CELL_NAME = {
    ("수요 상위", "공급 하위"): "핵심 불균형형",
    ("수요 상위", "공급 중위"): "고수요 보완형",
    ("수요 상위", "공급 상위"): "고수요 유지관리형",
    ("수요 중위", "공급 하위"): "잠재 취약형",
    ("수요 중위", "공급 중위"): "평균 관리형",
    ("수요 중위", "공급 상위"): "안정 관리형",
    ("수요 하위", "공급 하위"): "최소 인프라 보완형",
    ("수요 하위", "공급 중위"): "안정 모니터링형",
    ("수요 하위", "공급 상위"): "여유·거점 활용형",
}
_CDI_ROW_ORDER = ["수요 상위", "수요 중위", "수요 하위"]
_CSI_COL_ORDER = ["공급 하위", "공급 중위", "공급 상위"]

# 유형 해석 카드 내용
_TYPE_CARD_INFO = [
    ("접근성 보완형", "#2980B9",
     "Wee센터 접근성 점수가 낮은 학교로, 물리적 거리나 이동 여건으로 인해 상담 기관 이용이 "
     "어려울 수 있는 유형입니다. 이동형 상담, 온라인 상담 연계, 권역별 Wee센터 연결을 검토할 수 있습니다."),
    ("잠재 취약형",   "#9B59B6",
     "수요는 중위 수준이나 공급이 낮은 학교로, 현재 수요가 급격히 높지 않더라도 상담 인프라 기반이 "
     "취약한 유형입니다. 순회상담 또는 기본 상담지원 체계 보완이 필요할 수 있습니다."),
    ("고수요 보완형", "#E67E22",
     "수요가 높고 공급은 중위 수준인 학교로, 기존 인프라를 유지하면서 상담 프로그램 확대나 "
     "인력 보완을 검토할 수 있습니다."),
    ("안정형",        "#27AE60",
     "수요 대비 공급이 비교적 안정적인 학교로, 신규 자원 배치보다는 기존 인프라 유지와 "
     "정기 모니터링이 적절한 유형입니다."),
]

# 정책전략 그룹 → 해석 문구
_STRATEGY_INTERPRET = {
    "최우선 개입형":      "상담수요 대비 공급 부족이 뚜렷하여 우선 지원 검토가 필요한 유형",
    "우선 보완형":        "상담수요가 높거나 기존 인프라 보완 필요성이 있는 유형",
    "인력 취약형":        "전문상담교사 배치 또는 순회상담 연계 검토가 필요한 유형",
    "접근성 보완형":      "이동형·온라인 상담, 권역별 연계가 필요한 유형",
    "고수요 유지관리형":  "수요가 높지만 공급도 확보되어 질 관리·프로그램 고도화가 필요한 유형",
    "최소 인프라 보완형": "수요가 높지 않더라도 기본 상담 인프라 보완이 필요한 유형",
    "안정형":             "현재 지표상 수요 대비 공급이 안정적인 유형",
    "확인 필요형":        "주요 지표 결측으로 추가 확인이 필요한 유형",
}


def show_type_analysis(df: pd.DataFrame):
    # ── 헤더 ─────────────────────────────────────────────────────────────────
    st.markdown(
        "<h1 style='font-size:1.3rem;color:#1E3A5F;margin:0 0 2px 0;font-weight:700;'>"
        "상담수요-상담공급 기반 학교 유형 분석"
        "</h1>"
        "<p style='color:#718096;font-size:0.78rem;margin:0 0 8px 0;'>"
        "상담공급지수(CSI)와 상담수요지수(CDI)의 상대적 수준을 기준으로 학교를 "
        "3×3 수요-공급 매트릭스와 정책전략 그룹으로 분류하여 정책 대응 방향을 확인합니다."
        "</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div style='background:#EBF2FF;border-left:3px solid #2980B9;"
        "padding:6px 12px;border-radius:4px;font-size:0.74rem;color:#2C3E50;"
        "margin-bottom:14px;'>"
        "ℹ️ 본 유형화는 실제 지원 확정 기준이 아니라, 상담지원 인프라 배치와 "
        "정책 검토 우선순위를 파악하기 위한 분석 기준입니다."
        "</div>",
        unsafe_allow_html=True,
    )

    # ── KPI 4개 ──────────────────────────────────────────────────────────────
    _render_type_kpi_cards(df)
    st.markdown("<div style='margin:25px 0;'></div>", unsafe_allow_html=True)

    # ── Row 1: 산점도 (좌) + 정책전략 그룹 바 (우) ──────────────────────────
    sc_col, sg_col = st.columns([1.8, 1.2], gap="small")
    with sc_col:
        _render_csi_cdi_scatter(df)
    with sg_col:
        _render_strategy_group_bar(df)

    # ── Row 2: 3x3 히트맵 (좌) + 3x3 유형 바 (우) ──────────────────────────
    hm_col, mb_col = st.columns([1, 1], gap="small")
    with hm_col:
        _render_3x3_heatmap(df)
    with mb_col:
        _render_matrix_type_bar(df)

    # ── Row 3: 3x3 유형별 주요 특성 표 ─────────────────────────────────────
    _render_matrix_summary_table(df)

    # ── Row 4: 정책전략 그룹별 주요 특성 표 ─────────────────────────────────
    _render_strategy_summary_table(df)

    # ── Row 5: 핵심 유형 해석 카드 4개 ──────────────────────────────────────
    _render_type_cards()

    # ── Row 6: 유형별 우선 검토 학교 (expander) ─────────────────────────────
    _render_type_school_list(df)

    # ── K-means 보조 분석 섹션 ───────────────────────────────────────────────
    _render_kmeans_section()

    st.markdown(
        "<div class='footer-note'>"
        "※ CSI·CDI·우선지원점수 기반 유형화 결과는 정책 검토용 분석 자료이며 "
        "실제 지원 확정 기준이 아닙니다."
        "</div>",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
# ── K-means 보조 분석 섹션 ────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

_KMEANS_PATH = ROOT / "data" / "processed" / "gyeongnam_high_schools_policy_feedback_kmeans.xlsx"

_KMEANS_COLORS = {
    "수요 대비 공급 취약 군집": "#C0392B",
    "기초 인프라 보완 군집":    "#9B59B6",
    "평균 관리 군집":            "#2980B9",
    "상대적 안정 군집":          "#27AE60",
}

_KMEANS_INTERP = {
    "수요 대비 공급 취약 군집": "상담수요 대비 공급 부족 가능성이 높아 인력 배치·Wee클래스·Wee센터 연계 강화 검토가 필요한 군집입니다.",
    "기초 인프라 보완 군집":    "수요와 공급 모두 낮아 기본 상담 인프라 구축 및 보완이 필요한 군집입니다.",
    "평균 관리 군집":            "지표가 평균 수준으로, 현 인프라 유지 및 정기 모니터링이 적절한 군집입니다.",
    "상대적 안정 군집":          "CSI가 높고 우선지원점수가 낮아 현재 지표상 수요 대비 공급이 비교적 안정적인 군집입니다.",
}


@st.cache_data(show_spinner=False)
def _load_kmeans_data() -> pd.DataFrame | None:
    if not _KMEANS_PATH.exists():
        return None
    try:
        return pd.read_excel(_KMEANS_PATH, sheet_name="kmeans_school_table",
                             dtype={"school_code": str})
    except Exception:
        return None


def _render_kmeans_section():
    """K-means 보조 분석 섹션 — 유형 분석 탭 하단."""
    st.markdown(
        "<hr style='border-color:#E2E8F0;margin:28px 0 20px 0;'>"
        "<h2 style='font-size:1.1rem;color:#1E3A5F;margin:0 0 4px 0;font-weight:700;'>"
        "🔬 보조 데이터마이닝 분석: K-means 클러스터링</h2>"
        "<p style='color:#718096;font-size:0.77rem;margin:0 0 8px 0;'>"
        "기존 3×3 수요-공급 매트릭스를 유지하면서, 데이터 기반 유사 집단을 탐색하는 보조 분석입니다.</p>",
        unsafe_allow_html=True,
    )

    km_df = _load_kmeans_data()

    if km_df is None:
        st.info(
            "K-means 분석 결과가 아직 생성되지 않았습니다. "
            "`scripts/16_kmeans_clustering_school_types_2025.py`를 먼저 실행하세요."
        )
        return

    # 유효 데이터만 사용
    km_valid = km_df[km_df["kmeans_cluster_label"] != "확인 필요"].copy()

    # 설명 배너
    st.markdown(
        "<div style='background:#EBF2FF;border-left:3px solid #2980B9;"
        "padding:8px 12px;border-radius:4px;font-size:0.74rem;color:#2C3E50;"
        "margin-bottom:14px;'>"
        "ℹ️ K-means 군집 결과는 실제 지원 확정 기준이 아니라 정책 해석을 보완하는 참고 자료입니다. "
        "k=4, random_state=42, StandardScaler 표준화 적용."
        "</div>",
        unsafe_allow_html=True,
    )

    # ── Row 1: 군집별 학교 수 바 (좌) + CSI-CDI 산점도 (우) ─────────────────
    bar_col, sc_col = st.columns([1, 1.5], gap="small")

    with bar_col:
        cnt = (km_valid["kmeans_cluster_label"]
               .value_counts().reset_index()
               .rename(columns={"count": "학교수", "kmeans_cluster_label": "군집"}))
        cnt = cnt.sort_values("학교수", ascending=True)
        colors = [_KMEANS_COLORS.get(g, "#BDC3C7") for g in cnt["군집"]]
        fig = go.Figure(go.Bar(
            x=cnt["학교수"], y=cnt["군집"], orientation="h",
            marker_color=colors,
            text=cnt["학교수"], textposition="outside", textfont=dict(size=10),
        ))
        fig.update_layout(
            title=dict(text="K-means 군집별 학교 수", font=dict(size=12, color="#1E3A5F"), x=0),
            height=300, margin=dict(l=10, r=40, t=40, b=20),
            xaxis=dict(title="학교 수", tickfont=dict(size=9)),
            yaxis=dict(tickfont=dict(size=10)),
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Malgun Gothic, sans-serif"),
        )
        with st.container(border=True):
            st.plotly_chart(fig, width="stretch")

    with sc_col:
        avg_csi = km_valid["CSI"].mean()
        avg_cdi = km_valid["CDI"].mean()
        color_map = {g: _KMEANS_COLORS.get(g, "#BDC3C7")
                     for g in km_valid["kmeans_cluster_label"].unique()}
        fig2 = px.scatter(
            km_valid, x="CSI", y="CDI",
            color="kmeans_cluster_label",
            color_discrete_map=color_map,
            custom_data=["school_name", "sigungu", "priority_score", "policy_strategy_group"],
            labels={"kmeans_cluster_label": "K-means 군집"},
        )
        fig2.update_traces(
            marker=dict(size=8, opacity=0.85, line=dict(width=0.5, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b> (%{customdata[1]})<br>"
                "CSI: %{x:.3f} | CDI: %{y:.3f}<br>"
                "우선지원점수: %{customdata[2]:.3f}<br>"
                "정책전략 그룹: %{customdata[3]}<extra></extra>"
            ),
        )
        fig2.add_hline(y=avg_cdi, line_dash="dot", line_color="#718096", line_width=1,
                       annotation_text=f"평균 CDI {avg_cdi:.3f}",
                       annotation_position="bottom right",
                       annotation_font=dict(size=9, color="#718096"))
        fig2.add_vline(x=avg_csi, line_dash="dot", line_color="#718096", line_width=1,
                       annotation_text=f"평균 CSI {avg_csi:.3f}",
                       annotation_position="top left",
                       annotation_font=dict(size=9, color="#718096"))
        fig2.update_layout(
            title=dict(text="K-means 기반 CSI-CDI 군집 분포",
                       font=dict(size=12, color="#1E3A5F"), x=0),
            height=300, margin=dict(l=10, r=10, t=40, b=20),
            xaxis=dict(range=[-0.05, 1.05], title="CSI", tickfont=dict(size=9)),
            yaxis=dict(range=[-0.05, 1.05], title="CDI", tickfont=dict(size=9)),
            legend=dict(title="K-means 군집", font=dict(size=9), x=1.01, y=1),
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Malgun Gothic, sans-serif"),
        )
        with st.container(border=True):
            st.plotly_chart(fig2, width="stretch")

    # ── Row 2: 군집별 평균 요약표 ──────────────────────────────────────────
    agg_cols = {
        "CSI": "평균CSI", "CDI": "평균CDI",
        "priority_score": "평균PS",
        "counseling_staff_supply_score": "상담인력공급",
        "wee_class_score": "Wee클래스",
        "wee_center_access_score": "Wee센터접근",
    }
    avail = {k: v for k, v in agg_cols.items() if k in km_valid.columns}
    summary = (km_valid.groupby("kmeans_cluster_label")[list(avail.keys())]
               .mean().round(3).reset_index()
               .rename(columns=avail)
               .rename(columns={"kmeans_cluster_label": "K-means 군집"}))
    cnt_map = km_valid["kmeans_cluster_label"].value_counts().rename("학교수")
    summary = summary.merge(cnt_map, left_on="K-means 군집", right_index=True)
    col_order = ["K-means 군집", "학교수"] + list(avail.values())
    summary = summary[[c for c in col_order if c in summary.columns]]

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📊 K-means 군집별 평균 지표</div>",
            unsafe_allow_html=True,
        )
        st.dataframe(summary, use_container_width=True, height=200)

    # ── Row 3: 기존 유형화 × K-means 비교표 ──────────────────────────────
    with st.expander("▶ 기존 3×3 유형 × K-means 군집 비교표"):
        for cross_col, title in [
            ("supply_demand_matrix_3x3", "3×3 수요-공급 매트릭스"),
            ("policy_strategy_group",    "정책전략 그룹"),
            ("priority_level",           "우선지원등급"),
        ]:
            if cross_col not in km_valid.columns:
                continue
            ct = pd.crosstab(km_valid["kmeans_cluster_label"], km_valid[cross_col])
            st.markdown(
                f"<div style='font-size:0.80rem;font-weight:700;color:#2E5FA3;"
                f"margin:10px 0 4px 0;'>{title}</div>",
                unsafe_allow_html=True,
            )
            st.dataframe(ct, use_container_width=True)

    # ── Row 4: 군집 해석 카드 ────────────────────────────────────────────
    labels_in_data = [lbl for lbl in _KMEANS_INTERP if lbl in km_valid["kmeans_cluster_label"].unique()]
    if labels_in_data:
        cols = st.columns(len(labels_in_data), gap="small")
        for col, lbl in zip(cols, labels_in_data):
            color = _KMEANS_COLORS.get(lbl, "#718096")
            n = int((km_valid["kmeans_cluster_label"] == lbl).sum())
            desc = _KMEANS_INTERP.get(lbl, "")
            with col:
                st.markdown(
                    f"<div style='background:white;border-radius:10px;padding:12px;"
                    f"box-shadow:0 2px 8px rgba(0,0,0,0.07);border-top:3px solid {color};"
                    f"min-height:140px;'>"
                    f"<div style='font-size:0.78rem;font-weight:700;color:{color};"
                    f"margin-bottom:4px;'>{lbl}</div>"
                    f"<div style='font-size:0.68rem;color:#718096;margin-bottom:6px;'>{n}개교</div>"
                    f"<div style='font-size:0.72rem;color:#4A5568;line-height:1.5;'>{desc}</div>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

    st.markdown(
        "<div style='font-size:0.69rem;color:#A0AEC0;margin-top:8px;'>"
        "※ K-means 군집 라벨은 사후 해석이며 실제 지원 확정 기준이 아닙니다. "
        "k=4, 9개 변수(StandardScaler 표준화) 적용."
        "</div>",
        unsafe_allow_html=True,
    )


# ── KPI 카드 4개 ──────────────────────────────────────────────────────────────
def _render_type_kpi_cards(df: pd.DataFrame):
    n_total  = len(df)
    n_access = int((df.get("policy_strategy_group",    pd.Series(dtype=str)) == "접근성 보완형").sum())
    n_top    = int((df.get("policy_strategy_group",    pd.Series(dtype=str)) == "최우선 개입형").sum())
    n_manp   = int((df.get("policy_strategy_group",    pd.Series(dtype=str)) == "인력 취약형").sum())

    c1, c2, c3, c4 = st.columns(4, gap="small")
    specs = [
        (c1, "#2E5FA3", "분석 대상 학교 수", f"{n_total}개교",  "경상남도 일반고 전수"),
        (c2, "#2980B9", "접근성 보완형",      f"{n_access}개교", "Wee센터 접근성 부족"),
        (c3, "#E67E22", "최우선 개입형",      f"{n_top}개교",    "정책전략 최우선 지원"),
        (c4, "#9B59B6", "인력 취약형",        f"{n_manp}개교",   "상담인력 공급 부족"),
    ]
    for col, color, label, value, sub in specs:
        with col:
            st.markdown(
                f"<div style='background:white;border-radius:10px;padding:14px 12px;"
                f"box-shadow:0 2px 8px rgba(0,0,0,0.08);border-top:3px solid {color};"
                f"min-height:95px;'>"
                f"<div style='font-size:0.71rem;color:#718096;margin-bottom:4px;'>{label}</div>"
                f"<div style='font-size:1.4rem;font-weight:700;color:{color};'>{value}</div>"
                f"<div style='font-size:0.67rem;color:#A0AEC0;'>{sub}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )


# ── CSI-CDI 산점도 ────────────────────────────────────────────────────────────
def _render_csi_cdi_scatter(df: pd.DataFrame):
    if "policy_strategy_group" not in df.columns:
        st.warning("policy_strategy_group 컬럼이 없어 산점도를 표시할 수 없습니다.")
        return

    avg_csi = df["CSI"].mean()
    avg_cdi = df["CDI"].mean()

    color_map = {g: STRATEGY_COLORS.get(g, "#BDC3C7") for g in df["policy_strategy_group"].unique()}

    hover_cols = {
        "학교명": "school_name", "시군구": "sigungu",
        "3x3 유형": "supply_demand_matrix_3x3",
        "정책전략 그룹": "policy_strategy_group",
        "우선지원점수": "priority_score",
    }
    custom_data_cols = [c for c in hover_cols.values() if c in df.columns]

    fig = px.scatter(
        df, x="CSI", y="CDI",
        color="policy_strategy_group",
        color_discrete_map=color_map,
        custom_data=custom_data_cols,
        labels={"CSI": "CSI (상담공급지수)", "CDI": "CDI (상담수요지수)",
                "policy_strategy_group": "정책전략 그룹"},
    )
    # hover 템플릿
    cd_labels = [k for k, v in hover_cols.items() if v in custom_data_cols]
    ht = "<b>%{customdata[" + str(custom_data_cols.index("school_name") if "school_name" in custom_data_cols else 0) + "]}</b><br>"
    for i, col in enumerate(custom_data_cols):
        lbl = next((k for k, v in hover_cols.items() if v == col), col)
        ht += f"{lbl}: %{{customdata[{i}]}}<br>"
    ht += f"CSI: %{{x:.3f}}<br>CDI: %{{y:.3f}}<extra></extra>"
    fig.update_traces(hovertemplate=ht, marker=dict(size=8, opacity=0.8, line=dict(width=0.5, color="white")))

    # 평균선
    fig.add_hline(y=avg_cdi, line_dash="dot", line_color="#718096", line_width=1.2,
                  annotation_text=f"평균 CDI {avg_cdi:.2f}",
                  annotation_position="bottom right",
                  annotation_font=dict(size=9, color="#718096"))
    fig.add_vline(x=avg_csi, line_dash="dot", line_color="#718096", line_width=1.2,
                  annotation_text=f"평균 CSI {avg_csi:.2f}",
                  annotation_position="top left",
                  annotation_font=dict(size=9, color="#718096"))

    fig.update_layout(
        title=dict(text="CSI-CDI 기반 학교 유형 분포", font=dict(size=13, color="#1E3A5F"), x=0),
        height=420,
        margin=dict(l=10, r=10, t=45, b=50),
        xaxis=dict(range=[-0.05, 1.05], title="CSI (상담공급지수)", tickfont=dict(size=9)),
        yaxis=dict(range=[-0.05, 1.05], title="CDI (상담수요지수)", tickfont=dict(size=9)),
        legend=dict(title="정책전략 그룹", font=dict(size=9), orientation="v",
                    x=1.01, y=1, xanchor="left"),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div style='font-size:0.72rem;color:#718096;text-align:center;margin-top:-8px;'>"
            "오른쪽으로 갈수록 상담공급 수준이 높고, 위쪽으로 갈수록 상담수요 수준이 높음</div>",
            unsafe_allow_html=True,
        )


# ── 정책전략 그룹별 학교 수 수평 바 ──────────────────────────────────────────
def _render_strategy_group_bar(df: pd.DataFrame):
    if "policy_strategy_group" not in df.columns:
        st.warning("policy_strategy_group 컬럼이 없습니다.")
        return

    grp = (df["policy_strategy_group"].value_counts()
           .reset_index().rename(columns={"count": "학교수", "policy_strategy_group": "그룹"}))
    grp = grp.sort_values("학교수", ascending=True)
    colors = [STRATEGY_COLORS.get(g, "#BDC3C7") for g in grp["그룹"]]

    fig = go.Figure(go.Bar(
        x=grp["학교수"], y=grp["그룹"], orientation="h",
        marker_color=colors,
        text=grp["학교수"], textposition="outside", textfont=dict(size=10),
    ))
    fig.update_layout(
        title=dict(text="정책전략 그룹별 학교 수", font=dict(size=12, color="#1E3A5F"), x=0),
        height=420,
        margin=dict(l=10, r=40, t=45, b=20),
        xaxis=dict(title="학교 수", tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=10)),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")


# ── 3x3 매트릭스 히트맵 ──────────────────────────────────────────────────────
def _render_3x3_heatmap(df: pd.DataFrame):
    if "cdi_relative_level" not in df.columns or "csi_relative_level" not in df.columns:
        st.info("cdi_relative_level 또는 csi_relative_level 컬럼이 없어 3×3 히트맵을 표시할 수 없습니다.")
        return

    pivot = pd.crosstab(df["cdi_relative_level"], df["csi_relative_level"])
    # 행/열 순서 정렬
    row_order = [r for r in _CDI_ROW_ORDER if r in pivot.index]
    col_order = [c for c in _CSI_COL_ORDER if c in pivot.columns]
    pivot = pivot.reindex(index=row_order, columns=col_order, fill_value=0)

    n_total = len(df)
    z_vals, text_vals, hover_vals = [], [], []
    for r in row_order:
        z_row, t_row, h_row = [], [], []
        for c in col_order:
            cnt  = int(pivot.loc[r, c]) if (r in pivot.index and c in pivot.columns) else 0
            pct  = cnt / n_total * 100
            name = _MATRIX_CELL_NAME.get((r, c), f"{r}·{c}")
            z_row.append(cnt)
            t_row.append(f"<b>{name}</b><br>{cnt}개교<br>({pct:.1f}%)")
            h_row.append(f"<b>{name}</b><br>수요: {r} / 공급: {c}<br>{cnt}개교 ({pct:.1f}%)")
        z_vals.append(z_row)
        text_vals.append(t_row)
        hover_vals.append(h_row)

    fig = go.Figure(go.Heatmap(
        z=z_vals,
        x=[f"<b>{c}</b>" for c in col_order],
        y=[f"<b>{r}</b>" for r in row_order],
        text=text_vals,
        texttemplate="%{text}",
        hovertext=hover_vals,
        hovertemplate="%{hovertext}<extra></extra>",
        colorscale=[[0, "#FEF9E7"], [0.4, "#F39C12"], [1, "#C0392B"]],
        showscale=True,
        colorbar=dict(title="학교수", tickfont=dict(size=9), len=0.8),
        xgap=3, ygap=3,
    ))
    fig.update_layout(
        title=dict(text="상담수요-상담공급 3×3 매트릭스 (학교 수)",
                   font=dict(size=12, color="#1E3A5F"), x=0),
        height=370,
        margin=dict(l=10, r=10, t=45, b=20),
        xaxis=dict(title="공급(CSI) 수준", side="bottom", tickfont=dict(size=10)),
        yaxis=dict(title="수요(CDI) 수준", tickfont=dict(size=10)),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif", size=10),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")


# ── 3x3 유형별 학교 수 수평 바 ───────────────────────────────────────────────
def _render_matrix_type_bar(df: pd.DataFrame):
    if "supply_demand_matrix_3x3" not in df.columns:
        st.warning("supply_demand_matrix_3x3 컬럼이 없습니다.")
        return

    mat = (df["supply_demand_matrix_3x3"].value_counts()
           .reset_index().rename(columns={"count": "학교수", "supply_demand_matrix_3x3": "유형"}))
    mat = mat.sort_values("학교수", ascending=True)

    palette = ["#C0392B", "#E67E22", "#F4D03F", "#9B59B6",
               "#2980B9", "#1ABC9C", "#27AE60", "#BDC3C7"]
    colors = [palette[i % len(palette)] for i in range(len(mat))]

    fig = go.Figure(go.Bar(
        x=mat["학교수"], y=mat["유형"], orientation="h",
        marker_color=colors,
        text=mat["학교수"], textposition="outside", textfont=dict(size=10),
    ))
    fig.update_layout(
        title=dict(text="3×3 수요-공급 유형별 학교 수", font=dict(size=12, color="#1E3A5F"), x=0),
        height=370,
        margin=dict(l=10, r=40, t=45, b=20),
        xaxis=dict(title="학교 수", tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=10)),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")


# ── 3x3 유형별 주요 특성 표 ──────────────────────────────────────────────────
def _render_matrix_summary_table(df: pd.DataFrame):
    if "supply_demand_matrix_3x3" not in df.columns:
        return

    sub_score_cols = {
        "상담인력 공급": "counseling_staff_supply_score",
        "Wee클래스":     "wee_class_score",
        "Wee센터 접근": "wee_center_access_score",
        "수요 규모":     "demand_size_score",
        "상담 이용률":   "counseling_use_score",
        "학교폭력 위험": "school_violence_risk_score",
    }
    agg = df.groupby("supply_demand_matrix_3x3").agg(
        학교수=("school_name", "count"),
        평균CSI=("CSI", "mean"),
        평균CDI=("CDI", "mean"),
        평균우선지원점수=("priority_score", "mean"),
        **{k: (v, "mean") for k, v in sub_score_cols.items() if v in df.columns},
    ).round(3).reset_index()
    agg = agg.rename(columns={"supply_demand_matrix_3x3": "3x3 유형"})
    agg = agg.sort_values("평균우선지원점수", ascending=False).reset_index(drop=True)

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📊 3×3 유형별 주요 특성</div>",
            unsafe_allow_html=True,
        )
        st.dataframe(agg, use_container_width=True, height=280)


# ── 정책전략 그룹별 주요 특성 표 ─────────────────────────────────────────────
def _render_strategy_summary_table(df: pd.DataFrame):
    if "policy_strategy_group" not in df.columns:
        return

    # 우선지원 등급별 학교 수 (1등급, 2등급)
    def _cnt(group_df, level):
        return int((group_df["priority_level"] == level).sum()) if "priority_level" in group_df.columns else 0

    rows = []
    for grp, gdf in df.groupby("policy_strategy_group"):
        rows.append({
            "정책전략 그룹":   grp,
            "학교수":          len(gdf),
            "평균CSI":         round(gdf["CSI"].mean(), 3),
            "평균CDI":         round(gdf["CDI"].mean(), 3),
            "평균우선지원점수": round(gdf["priority_score"].mean(), 3),
            "최우선 지원":     _cnt(gdf, "최우선 지원"),
            "우선 지원":       _cnt(gdf, "우선 지원"),
            "주요 정책 해석":  _STRATEGY_INTERPRET.get(str(grp), "-"),
        })
    strat_df = (pd.DataFrame(rows)
                .sort_values("평균우선지원점수", ascending=False)
                .reset_index(drop=True))

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📋 정책전략 그룹별 주요 특성</div>",
            unsafe_allow_html=True,
        )
        st.dataframe(strat_df, use_container_width=True, height=240)


# ── 핵심 유형 해석 카드 4개 ──────────────────────────────────────────────────
def _render_type_cards():
    c1, c2, c3, c4 = st.columns(4, gap="small")
    for col, (title, color, desc) in zip([c1, c2, c3, c4], _TYPE_CARD_INFO):
        with col:
            st.markdown(
                f"<div style='background:white;border-radius:10px;padding:14px 12px;"
                f"box-shadow:0 2px 8px rgba(0,0,0,0.07);border-top:3px solid {color};"
                f"min-height:150px;'>"
                f"<div style='font-size:0.82rem;font-weight:700;color:{color};"
                f"margin-bottom:8px;'>{title}</div>"
                f"<div style='font-size:0.74rem;color:#4A5568;line-height:1.6;'>{desc}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
    st.markdown("<div style='margin:20px 0;'></div>", unsafe_allow_html=True)


# ── 유형별 우선 검토 학교 (expander) ─────────────────────────────────────────
def _render_type_school_list(df: pd.DataFrame):
    if "supply_demand_matrix_3x3" not in df.columns:
        return

    with st.expander("▶ 유형별 우선 검토 학교 예시 (3×3 유형별 우선지원점수 상위 5개교)"):
        disp_cols = ["supply_demand_matrix_3x3", "school_name", "sigungu",
                     "CSI", "CDI", "priority_score", "priority_level",
                     "policy_strategy_group"]
        avail = [c for c in disp_cols if c in df.columns]

        frames = []
        for typ, gdf in df.groupby("supply_demand_matrix_3x3"):
            top5 = gdf.nlargest(5, "priority_score")[avail].copy()
            frames.append(top5)

        if not frames:
            st.info("데이터가 없습니다.")
            return

        result = pd.concat(frames, ignore_index=True)
        result = result.rename(columns={
            "supply_demand_matrix_3x3": "3x3 유형",
            "school_name":              "학교명",
            "sigungu":                  "시군구",
            "priority_score":           "우선지원점수",
            "priority_level":           "우선지원등급",
            "policy_strategy_group":    "정책전략 그룹",
        })
        for c in ["CSI", "CDI", "우선지원점수"]:
            if c in result.columns:
                result[c] = result[c].apply(lambda x: f"{float(x):.3f}" if pd.notna(x) else "-")

        st.dataframe(result.reset_index(drop=True), use_container_width=True, height=400)


# ══════════════════════════════════════════════════════════════════════════════
# ── 10. 시뮬레이션 탭 ────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

# CSI 하위 점수 컬럼
_SIM_CSI_COLS = [
    "counseling_staff_supply_score",
    "wee_class_score",
    "wee_center_access_score",
]

# 보조 등급 분류 (고정 기준)
def _assign_sim_level(score: float, q90: float, q80: float, q20: float) -> str:
    """원본 데이터 분위수 기준으로 시뮬레이션 보조 등급 부여."""
    if pd.isna(score):
        return "확인 필요"
    if score >= q90:
        return "최우선 지원 검토"
    if score >= q80:
        return "우선 지원 검토"
    if score >= q20:
        return "모니터링"
    return "안정"


# 시뮬레이션 계산 핵심 함수
def _run_simulation(
    df: pd.DataFrame,
    target_mask: pd.Series,
    apply_staff: bool,
    apply_wee_class: bool,
    apply_wee_center: bool,
    apply_program: bool,
    q90: float,
    q80: float,
    q20: float,
) -> pd.DataFrame:
    sim = df[target_mask].copy()

    # 하위 점수 복사
    sim["sim_staff"]  = sim["counseling_staff_supply_score"].copy()
    sim["sim_wee"]    = sim["wee_class_score"].copy()
    sim["sim_center"] = sim["wee_center_access_score"].copy()

    # 정책 1: 전문상담교사 배치 강화
    if apply_staff:
        sim.loc[sim["sim_staff"] < 0.4, "sim_staff"] = 0.4

    # 정책 2: Wee클래스 신설
    if apply_wee_class:
        sim.loc[sim["sim_wee"] == 0, "sim_wee"] = 1.0

    # 정책 3: Wee센터 연계 강화
    if apply_wee_center:
        sim.loc[sim["sim_center"] <= 0.4, "sim_center"] = 0.7

    # sim_CSI 계산
    sim["sim_CSI"]    = (sim["sim_staff"] + sim["sim_wee"] + sim["sim_center"]) / 3
    sim["sim_CDI"]    = sim["CDI"]
    sim["sim_PS"]     = sim["sim_CDI"] - sim["sim_CSI"]
    sim["sim_CSI_chg"] = sim["sim_CSI"] - sim["CSI"]
    sim["sim_PS_chg"]  = sim["sim_PS"]  - sim["priority_score"]
    sim["sim_level"]   = sim["sim_PS"].apply(lambda s: _assign_sim_level(s, q90, q80, q20))

    # 적용 정책 문자열
    applied = []
    if apply_staff:     applied.append("전문상담교사 배치")
    if apply_wee_class: applied.append("Wee클래스 신설")
    if apply_wee_center:applied.append("Wee센터 연계")
    if apply_program:   applied.append("프로그램 강화(정성)")
    sim["적용_정책"] = "; ".join(applied) if applied else "(없음)"

    return sim


# 메인 탭 함수
def show_simulation(df: pd.DataFrame):
    # ── 헤더 ─────────────────────────────────────────────────────────────────
    st.markdown(
        "<h1 style='font-size:1.3rem;color:#1E3A5F;margin:0 0 14px 0;font-weight:700;'>"
        "정책 적용 시뮬레이션"
        "</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h2 style='font-size:1.1rem;color:#1E3A5F;margin:0 0 4px 0;font-weight:700;'>"
        "📊 유형별 정책 적용 시뮬레이션"
        "</h2>"
        "<p style='color:#718096;font-size:0.78rem;margin:0 0 8px 0;'>"
        "전문상담교사 배치, Wee클래스 신설, Wee센터 연계 강화 정책 조합을 가정하여 "
        "상담공급지수(CSI)와 우선지원점수의 변화를 비교합니다."
        "</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div style='background:#FEF9E7;border-left:3px solid #F39C12;"
        "padding:6px 12px;border-radius:4px;font-size:0.74rem;color:#7D6608;"
        "margin-bottom:14px;'>"
        "⚠️ 본 시뮬레이션은 실제 정책 효과를 예측하는 모델이 아니라, "
        "지수 산식에 기반한 가상 정책 적용 결과입니다."
        "</div>",
        unsafe_allow_html=True,
    )

    # ── 선택 영역 (2열) ──────────────────────────────────────────────────────
    sel_col, pol_col = st.columns([1, 1], gap="small")

    with sel_col:
        with st.container(border=True):
            st.markdown(
                "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
                "margin-bottom:8px;'>🎯 시뮬레이션 대상 선택</div>",
                unsafe_allow_html=True,
            )
            target_options = [
                "최우선 개입형 학교",
                "우선 보완형 학교",
                "인력 취약형 학교",
                "접근성 보완형 학교",
                "우선지원점수 상위 20개교",
            ]
            target_sel = st.radio(
                "대상 선택", target_options,
                label_visibility="collapsed",
            )
            school_code_sel = None

    with pol_col:
        with st.container(border=True):
            st.markdown(
                "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
                "margin-bottom:8px;'>⚙️ 정책 조건 선택</div>",
                unsafe_allow_html=True,
            )
            apply_staff  = st.checkbox(
                "전문상담교사 배치 또는 순회상담 연계 적용",
                help="counseling_staff_supply_score < 0.4 → 0.4 상향"
            )
            apply_wee    = st.checkbox(
                "Wee클래스 미운영 학교 신설 가정",
                help="wee_class_score == 0 → 1.0 변경"
            )
            apply_center = st.checkbox(
                "Wee센터 연계 강화",
                help="wee_center_access_score ≤ 0.4 → 0.7 상향"
            )
            apply_prog   = st.checkbox(
                "고수요 학교 상담 프로그램 강화 (정성 반영만)",
                help="점수 변화 없음. 정책 피드백 문구에만 반영됩니다."
            )
            st.markdown(
                "<div style='font-size:0.70rem;color:#A0AEC0;margin-top:6px;'>"
                "※ 정책 1~3은 CSI 산식에 반영 | 정책 4는 정성 표시만</div>",
                unsafe_allow_html=True,
            )

    # ── 대상 마스크 생성 ─────────────────────────────────────────────────────
    grp_col = "policy_strategy_group"
    target_map = {
        "최우선 개입형 학교": df[grp_col] == "최우선 개입형",
        "우선 보완형 학교":   df[grp_col] == "우선 보완형",
        "인력 취약형 학교":   df[grp_col] == "인력 취약형",
        "접근성 보완형 학교": df[grp_col] == "접근성 보완형",
        "우선지원점수 상위 20개교": df["priority_score"].rank(ascending=False) <= 20,
    }
    mask = target_map.get(target_sel, pd.Series([False] * len(df), index=df.index))
    n_target = int(mask.sum())

    if n_target == 0:
        st.warning("선택한 조건에 해당하는 학교가 없습니다.")
        return

    # 정책 미선택 안내
    no_policy = not (apply_staff or apply_wee or apply_center or apply_prog)
    if no_policy:
        st.info("정책 조건을 하나 이상 선택하면 시뮬레이션 결과를 확인할 수 있습니다.")
        # 학교별 시뮬레이션은 정책 미선택 시에도 표시
        st.markdown(
            "<hr style='border-color:#E2E8F0;margin:28px 0 20px 0;'>"
            "<h2 style='font-size:1.1rem;color:#1E3A5F;margin:0 0 4px 0;font-weight:700;'>"
            "🏫 학교별 정책 적용 시뮬레이션</h2>"
            "<p style='color:#718096;font-size:0.77rem;margin:0 0 14px 0;'>"
            "개별 학교를 선택하고 정책 조건을 적용하여 6개 하위 지표와 CSI·우선지원점수의 "
            "변화를 레이더 차트로 비교합니다.</p>",
            unsafe_allow_html=True,
        )
        _render_school_sim(df)
        return

    # ── 시뮬레이션 실행 ──────────────────────────────────────────────────────
    # 원본 전체 데이터 분위수 기준값 계산 (priority_level 등급 기준과 동일한 방식)
    ps = df["priority_score"].dropna()
    q90 = float(ps.quantile(0.90))  # 상위 10% 기준
    q80 = float(ps.quantile(0.80))  # 상위 20% 기준
    q20 = float(ps.quantile(0.20))  # 하위 20% 기준

    sim_df = _run_simulation(df, mask, apply_staff, apply_wee, apply_center, apply_prog, q90, q80, q20)

    # ── KPI 카드 4개 ─────────────────────────────────────────────────────────
    _render_sim_kpi_cards(sim_df)
    st.markdown("<div style='margin:25px 0;'></div>", unsafe_allow_html=True)

    # ── Row 2: 전후 그래프 (좌) + 결과 표 (우) ──────────────────────────────
    chart_col, tbl_col = st.columns([1.5, 1], gap="small")
    with chart_col:
        _render_sim_score_chart(sim_df)
    with tbl_col:
        _render_sim_result_table(sim_df)

    # ── Row 3: 정책별 효과 바 (좌) + 해석 박스 (우) ─────────────────────────
    eff_col, int_col = st.columns([1, 1], gap="small")
    with eff_col:
        _render_policy_effect_bar(sim_df, apply_staff, apply_wee, apply_center)
    with int_col:
        _render_sim_interpretation(sim_df, target_sel, apply_staff, apply_wee, apply_center, apply_prog)

    # ── 유의사항 ─────────────────────────────────────────────────────────────
    # ── 학교별 시뮬레이션 섹션 ──────────────────────────────────────────────
    st.markdown(
        "<hr style='border-color:#E2E8F0;margin:28px 0 20px 0;'>"
        "<h2 style='font-size:1.1rem;color:#1E3A5F;margin:0 0 4px 0;font-weight:700;'>"
        "🏫 학교별 정책 적용 시뮬레이션</h2>"
        "<p style='color:#718096;font-size:0.77rem;margin:0 0 14px 0;'>"
        "개별 학교를 선택하고 정책 조건을 적용하여 6개 하위 지표와 CSI·우선지원점수의 "
        "변화를 레이더 차트로 비교합니다.</p>",
        unsafe_allow_html=True,
    )
    _render_school_sim(df)

    st.markdown(
        "<div class='footer-note'>"
        "※ 시뮬레이션 결과는 지수 산식(CSI 평균) 기반 가상 시나리오이며 실제 정책 효과를 보장하지 않습니다. "
        "| CSI = (상담인력공급 + Wee클래스 + Wee센터접근성) / 3 "
        "| 우선지원점수 = CDI − CSI"
        "</div>",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
# ── 학교별 시뮬레이션 함수 ────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

# 레이더 차트용 축 정의
_RADAR_COLS   = _CSI_SUB + _CDI_SUB          # 6개 하위 점수 컬럼
_RADAR_LABELS = [
    "상담인력 공급", "Wee클래스 운영", "Wee센터 접근성",
    "수요 규모",     "상담 이용률",    "학교폭력 위험",
]


def _render_school_sim(df: pd.DataFrame):
    """학교별 정책 적용 시뮬레이션 — 선택 UI + 레이더 차트 + 세부 표."""

    # ── 학교 선택 + 정책 조건 (2열) ──────────────────────────────────────────
    sch_col, pol_col = st.columns([1, 1], gap="small")

    with sch_col:
        with st.container(border=True):
            st.markdown(
                "<div style='font-size:0.83rem;font-weight:700;color:#1E3A5F;"
                "margin-bottom:8px;'>🔍 학교 선택</div>",
                unsafe_allow_html=True,
            )
            df_s = df.sort_values("priority_score", ascending=False).reset_index(drop=True)
            from collections import Counter
            base_labels = df_s.apply(
                lambda r: f"{r['school_name']} ({r['sigungu']})", axis=1
            ).tolist()
            cnt  = Counter(base_labels)
            seen: dict = {}
            sch_opts, sch_codes = [], df_s["school_code"].tolist()
            for lbl, code in zip(base_labels, sch_codes):
                if cnt[lbl] > 1:
                    seen[lbl] = seen.get(lbl, 0) + 1
                    sch_opts.append(f"{lbl} [{code}]")
                else:
                    sch_opts.append(lbl)

            sel_opt  = st.selectbox("학교", sch_opts, index=0,
                                    label_visibility="collapsed")
            sel_code = sch_codes[sch_opts.index(sel_opt)]

    with pol_col:
        with st.container(border=True):
            st.markdown(
                "<div style='font-size:0.83rem;font-weight:700;color:#1E3A5F;"
                "margin-bottom:8px;'>⚙️ 정책 조건</div>",
                unsafe_allow_html=True,
            )
            p1 = st.checkbox("전문상담교사 배치 또는 순회상담 연계 적용",
                             key="sch_p1",
                             help="counseling_staff_supply_score < 0.4 → 0.4 상향")
            p2 = st.checkbox("Wee클래스 신설 가정",
                             key="sch_p2",
                             help="wee_class_score == 0 → 1.0 상향")
            p3 = st.checkbox("Wee센터 연계 강화",
                             key="sch_p3",
                             help="wee_center_access_score ≤ 0.4 → 0.7 상향")
            p4 = st.checkbox("상담 프로그램 강화 (해석 문구만 반영)",
                             key="sch_p4",
                             help="점수 변화 없음. 해석 문구에만 반영됩니다.")

    # ── 선택 학교 row 추출 ────────────────────────────────────────────────────
    row_df = df[df["school_code"] == sel_code]
    if row_df.empty:
        st.warning("선택한 학교 데이터를 찾을 수 없습니다.")
        return
    row = row_df.iloc[0]

    # ── 적용 후 점수 계산 ─────────────────────────────────────────────────────
    def _safe_float(val, default=0.0):
        try:
            return float(val)
        except (TypeError, ValueError):
            return default

    staff_b  = _safe_float(row.get("counseling_staff_supply_score"))
    wee_b    = _safe_float(row.get("wee_class_score"))
    center_b = _safe_float(row.get("wee_center_access_score"))
    dem_b    = _safe_float(row.get("demand_size_score"))
    use_b    = _safe_float(row.get("counseling_use_score"))
    viol_b   = _safe_float(row.get("school_violence_risk_score"))
    csi_b    = _safe_float(row.get("CSI"))
    cdi_b    = _safe_float(row.get("CDI"))
    ps_b     = _safe_float(row.get("priority_score"))

    staff_a  = max(staff_b, 0.4)  if p1 else staff_b
    wee_a    = 1.0                if (p2 and wee_b == 0) else wee_b
    center_a = max(center_b, 0.7) if (p3 and center_b <= 0.4) else center_b
    dem_a    = dem_b
    use_a    = use_b
    viol_a   = viol_b

    csi_a  = (staff_a + wee_a + center_a) / 3
    cdi_a  = cdi_b
    ps_a   = cdi_a - csi_a

    csi_chg = csi_a - csi_b
    ps_chg  = ps_a  - ps_b

    # 결측 체크
    has_missing = any(
        (isinstance(row.get(c), float) and pd.isna(row.get(c)))
        for c in _RADAR_COLS
    )
    if has_missing:
        st.warning("⚠️ 일부 지표 결측으로 레이더 차트 해석에 주의가 필요합니다.")

    # ── KPI 4개 카드 ──────────────────────────────────────────────────────────
    _render_school_sim_kpi(row, csi_b, csi_a, csi_chg, cdi_b, ps_b, ps_a, ps_chg)
    st.markdown("<div style='margin:25px 0;'></div>", unsafe_allow_html=True)

    # ── 레이더 차트 (좌) + 해석 박스·세부 표 (우) ───────────────────────────
    radar_col, interp_col = st.columns([1.2, 1.8], gap="small")

    before_vals = [staff_b, wee_b, center_b, dem_b, use_b, viol_b]
    after_vals  = [staff_a, wee_a, center_a, dem_a, use_a, viol_a]

    with radar_col:
        _render_radar_chart(before_vals, after_vals, _RADAR_LABELS)

    with interp_col:
        _render_school_sim_interpretation(
            row, csi_b, csi_a, csi_chg, ps_b, ps_a, ps_chg,
            p1, p2, p3, p4,
            before_vals, after_vals,
        )
        _render_school_sim_table(
            before_vals, after_vals,
            csi_b, csi_a, cdi_b, ps_b, ps_a,
        )


def _render_school_sim_kpi(
    row: pd.Series,
    csi_b: float, csi_a: float, csi_chg: float,
    cdi_b: float,
    ps_b: float, ps_a: float, ps_chg: float,
):
    """학교별 시뮬레이션 KPI 카드 4개."""
    def _arrow(val, good_neg=False):
        if abs(val) < 1e-6:
            return "<span style='color:#718096;font-size:0.70rem;'>변화 없음</span>"
        improved = (val < 0) if good_neg else (val > 0)
        color    = "#27AE60" if improved else "#C0392B"
        sign     = "+" if val > 0 else ""
        arrow    = "▲" if val > 0 else "▼"
        return f"<span style='color:{color};font-size:0.70rem;'>{arrow} {sign}{val:.3f}</span>"

    sg    = str(row.get("policy_strategy_group", "확인 필요"))
    scolor = STRATEGY_COLORS.get(sg, "#718096")

    c1, c2, c3, c4 = st.columns(4, gap="small")
    specs = [
        (c1, "#1ABC9C", "CSI",
         f"{csi_b:.3f} → <b style='color:#1ABC9C;'>{csi_a:.3f}</b>",
         _arrow(csi_chg, good_neg=False)),
        (c2, "#2980B9", "CDI",
         f"{cdi_b:.3f}",
         "<span style='color:#718096;font-size:0.70rem;'>수요 지표 (불변)</span>"),
        (c3, "#E67E22", "우선지원점수",
         f"{ps_b:.3f} → <b style='color:#E67E22;'>{ps_a:.3f}</b>",
         _arrow(ps_chg, good_neg=True)),
        (c4, scolor, "정책전략 그룹",
         sg,
         "<span style='color:#A0AEC0;font-size:0.67rem;'>기존 그룹 기준 (재분류 없음)</span>"),
    ]
    for col, color, label, value, sub in specs:
        with col:
            st.markdown(
                f"<div style='background:white;border-radius:10px;padding:12px 10px;"
                f"box-shadow:0 2px 8px rgba(0,0,0,0.08);border-top:3px solid {color};"
                f"min-height:95px;'>"
                f"<div style='font-size:0.70rem;color:#718096;margin-bottom:4px;'>{label}</div>"
                f"<div style='font-size:1.0rem;font-weight:700;color:#2D3748;line-height:1.3;"
                f"word-break:keep-all;'>{value}</div>"
                f"<div style='margin-top:4px;'>{sub}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )


def _render_radar_chart(
    before: list, after: list, labels: list
):
    """Plotly Scatterpolar 레이더 차트 — 적용 전(회색) vs 적용 후(청록)."""
    # 닫힌 다각형을 위해 첫 값 반복
    cats   = labels + [labels[0]]
    before_c = before + [before[0]]
    after_c  = after  + [after[0]]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=before_c, theta=cats, name="적용 전",
        fill="toself",
        line=dict(color="#95A5A6", width=2),
        fillcolor="rgba(149,165,166,0.2)",
        hovertemplate="%{theta}: %{r:.3f}<extra>적용 전</extra>",
    ))
    fig.add_trace(go.Scatterpolar(
        r=after_c, theta=cats, name="적용 후",
        fill="toself",
        line=dict(color="#1ABC9C", width=2),
        fillcolor="rgba(26,188,156,0.2)",
        hovertemplate="%{theta}: %{r:.3f}<extra>적용 후</extra>",
    ))
    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 1], tickfont=dict(size=8)),
            angularaxis=dict(tickfont=dict(size=10)),
        ),
        showlegend=True,
        legend=dict(orientation="h", y=-0.12, x=0.3, font=dict(size=10)),
        title=dict(
            text="6개 하위 지표 적용 전후 비교",
            font=dict(size=12, color="#1E3A5F"), x=0.5,
        ),
        height=380,
        margin=dict(l=40, r=40, t=55, b=50),
        paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div style='font-size:0.68rem;color:#A0AEC0;text-align:center;margin-top:-6px;'>"
            "※ 레이더 차트는 보조 시각화이며 정책 확정 근거가 아닙니다.</div>",
            unsafe_allow_html=True,
        )


def _render_school_sim_interpretation(
    row: pd.Series,
    csi_b: float, csi_a: float, csi_chg: float,
    ps_b: float, ps_a: float, ps_chg: float,
    p1: bool, p2: bool, p3: bool, p4: bool,
    before_vals: list, after_vals: list,
):
    """학교별 시뮬레이션 해석 박스."""
    school_name = str(row.get("school_name", "선택 학교"))

    applied = []
    if p1: applied.append("전문상담교사 배치")
    if p2: applied.append("Wee클래스 신설")
    if p3: applied.append("Wee센터 연계 강화")
    if p4: applied.append("상담 프로그램 강화(정성)")
    pol_str = "·".join(applied) if applied else "(정책 미선택)"

    # 개선된 공급 지표 탐지
    improved_axes = []
    labels_csi = ["상담인력 공급", "Wee클래스 운영", "Wee센터 접근성"]
    for lbl, b, a in zip(labels_csi, before_vals[:3], after_vals[:3]):
        if a - b > 1e-6:
            improved_axes.append(lbl)
    improved_str = "·".join(improved_axes) if improved_axes else "없음 (변화 없음)"

    ps_dir = "완화되는 방향" if ps_chg < 0 else ("변화 없음" if abs(ps_chg) < 1e-6 else "악화되는 방향")

    items = [
        ("#2E5FA3", f"대상 학교: <b>{school_name}</b> | 적용 정책: <b>{pol_str}</b>"),
        ("#1ABC9C",
         f"CSI: <b>{csi_b:.3f}</b> → <b>{csi_a:.3f}</b> "
         f"(변화: <b>{'+'if csi_chg>=0 else ''}{csi_chg:.3f}</b>) — "
         f"개선 지표: {improved_str}"),
        ("#E67E22",
         f"우선지원점수: <b>{ps_b:.3f}</b> → <b>{ps_a:.3f}</b> "
         f"(변화: <b>{'+'if ps_chg>=0 else ''}{ps_chg:.3f}</b>) — {ps_dir}"),
        ("#F39C12",
         "본 결과는 지수 산식에 따른 가상 시나리오이며 실제 정책 효과를 보장하지 않습니다. "
         "지역 여건·학교 특성·정책 실행 수준에 따라 실제 효과는 달라질 수 있습니다."),
    ]

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:10px;'>"
            "💡 학교별 시뮬레이션 해석</div>",
            unsafe_allow_html=True,
        )
        for color, text in items:
            st.markdown(
                f"<div style='display:flex;gap:8px;margin-bottom:8px;"
                f"padding:7px 10px;background:#F7FAFC;"
                f"border-radius:6px;border-left:3px solid {color};'>"
                f"<p style='font-size:0.76rem;color:#2D3748;margin:0;line-height:1.6;'>{text}</p>"
                f"</div>",
                unsafe_allow_html=True,
            )


def _render_school_sim_table(
    before_vals: list, after_vals: list,
    csi_b: float, csi_a: float,
    cdi_b: float,
    ps_b: float, ps_a: float,
):
    """적용 전후 세부 점수 표."""
    rows = []
    for lbl, b, a in zip(_RADAR_LABELS, before_vals, after_vals):
        chg = a - b
        rows.append({
            "지표명": lbl,
            "적용 전": f"{b:.3f}",
            "적용 후": f"{a:.3f}",
            "변화량":  f"{'+'if chg>=0 else ''}{chg:.3f}" if abs(chg) > 1e-6 else "-",
        })
    # 지수 행 추가
    for lbl, b, a in [("CSI", csi_b, csi_a), ("CDI (불변)", cdi_b, cdi_b), ("우선지원점수", ps_b, ps_a)]:
        chg = a - b
        rows.append({
            "지표명": f"**{lbl}**",
            "적용 전": f"{b:.3f}",
            "적용 후": f"{a:.3f}",
            "변화량":  f"{'+'if chg>=0 else ''}{chg:.3f}" if abs(chg) > 1e-6 else "-",
        })

    tbl_df = pd.DataFrame(rows)
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📋 적용 전후 세부 점수 비교</div>",
            unsafe_allow_html=True,
        )
        st.dataframe(tbl_df, use_container_width=True, height=340)


# ── KPI 카드: 적용 전후 평균 비교 ────────────────────────────────────────────
def _render_sim_kpi_cards(sim_df: pd.DataFrame):
    n = len(sim_df)
    avg_csi_b = sim_df["CSI"].mean()
    avg_csi_a = sim_df["sim_CSI"].mean()
    avg_cdi   = sim_df["CDI"].mean()
    avg_ps_b  = sim_df["priority_score"].mean()
    avg_ps_a  = sim_df["sim_PS"].mean()
    n_improved = int((sim_df["sim_PS_chg"] < 0).sum())
    pct_imp    = n_improved / n * 100 if n > 0 else 0

    csi_chg = avg_csi_a - avg_csi_b
    ps_chg  = avg_ps_a  - avg_ps_b

    def _arrow_html(val, color_pos="#27AE60", color_neg="#C0392B"):
        if val > 0:
            return f"<span style='color:{color_pos};font-size:0.72rem;'>▲ +{val:.3f}</span>"
        elif val < 0:
            return f"<span style='color:{color_neg};font-size:0.72rem;'>▼ {val:.3f}</span>"
        return "<span style='color:#718096;font-size:0.72rem;'>변화 없음</span>"

    c1, c2, c3, c4 = st.columns(4, gap="small")
    specs = [
        (c1, "#1ABC9C",
         "평균 CSI",
         f"{avg_csi_b:.3f} → <b style='color:#1ABC9C;'>{avg_csi_a:.3f}</b>",
         _arrow_html(csi_chg, "#1ABC9C", "#C0392B")),
        (c2, "#2980B9",
         "평균 CDI",
         f"{avg_cdi:.3f}",
         "<span style='color:#718096;font-size:0.72rem;'>수요 지수 (불변)</span>"),
        (c3, "#E67E22",
         "평균 우선지원점수",
         f"{avg_ps_b:.3f} → <b style='color:#E67E22;'>{avg_ps_a:.3f}</b>",
         _arrow_html(ps_chg, "#C0392B", "#27AE60")),
        (c4, "#9B59B6",
         "점수 개선 학교 수",
         f"{n_improved}개교",
         f"<span style='color:#718096;font-size:0.72rem;'>전체 {n}개교 중 {pct_imp:.0f}%</span>"),
    ]
    for col, color, label, value, sub in specs:
        with col:
            st.markdown(
                f"<div style='background:white;border-radius:10px;padding:14px 12px;"
                f"box-shadow:0 2px 8px rgba(0,0,0,0.08);border-top:3px solid {color};"
                f"min-height:95px;'>"
                f"<div style='font-size:0.70rem;color:#718096;margin-bottom:4px;'>{label}</div>"
                f"<div style='font-size:1.0rem;font-weight:700;color:#2D3748;line-height:1.4;'>{value}</div>"
                f"<div style='margin-top:4px;'>{sub}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )


# ── 전후 우선지원점수 grouped bar ────────────────────────────────────────────
def _render_sim_score_chart(sim_df: pd.DataFrame):
    # 상위 10개교 (기존 priority_score 기준)
    top = sim_df.nlargest(10, "priority_score").copy()
    top["label"] = top.apply(
        lambda r: r.get("school_name", "?") if "school_name" in r.index else "?", axis=1
    )

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="적용 전", x=top["label"], y=top["priority_score"],
        marker_color="rgba(149,165,166,0.7)",
        text=[f"{v:.3f}" for v in top["priority_score"]],
        textposition="outside", textfont=dict(size=9),
    ))
    fig.add_trace(go.Bar(
        name="적용 후", x=top["label"], y=top["sim_PS"],
        marker_color="#2E5FA3",
        text=[f"{v:.3f}" for v in top["sim_PS"]],
        textposition="outside", textfont=dict(size=9),
    ))

    y_min = min(top["sim_PS"].min(), top["priority_score"].min()) - 0.05
    y_max = max(top["sim_PS"].max(), top["priority_score"].max()) + 0.12

    fig.update_layout(
        title=dict(text="정책 적용 전후 우선지원점수 변화 (상위 10개교)",
                   font=dict(size=12, color="#1E3A5F"), x=0),
        barmode="group", height=360,
        margin=dict(l=10, r=10, t=45, b=60),
        yaxis=dict(title="우선지원점수", range=[y_min, y_max],
                   zeroline=True, zerolinecolor="#BDC3C7", tickfont=dict(size=9)),
        xaxis=dict(tickangle=-30, tickfont=dict(size=9)),
        legend=dict(orientation="h", y=-0.22, x=0, font=dict(size=9)),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Malgun Gothic, sans-serif"),
    )
    with st.container(border=True):
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div style='font-size:0.70rem;color:#718096;text-align:center;margin-top:-6px;'>"
            "※ 우선지원점수가 낮아질수록 수요 대비 공급 부족이 완화된 것으로 해석합니다.</div>",
            unsafe_allow_html=True,
        )


# ── 시뮬레이션 결과 표 ────────────────────────────────────────────────────────
def _render_sim_result_table(sim_df: pd.DataFrame):
    disp = sim_df.sort_values("sim_PS_chg", ascending=True).copy()
    cols_map = {
        "school_name":     "학교명",
        "sigungu":         "시군구",
        "적용_정책":        "적용 정책",
        "CSI":             "기존 CSI",
        "sim_CSI":         "시뮬 CSI",
        "sim_CSI_chg":     "CSI 변화",
        "priority_score":  "기존 PS",
        "sim_PS":          "시뮬 PS",
        "sim_PS_chg":      "PS 변화",
        "priority_level":  "기존 등급",
        "sim_level":       "시뮬 등급",
    }
    avail = {k: v for k, v in cols_map.items() if k in disp.columns}
    tbl = disp[list(avail.keys())].rename(columns=avail)

    for c in ["기존 CSI", "시뮬 CSI", "CSI 변화", "기존 PS", "시뮬 PS", "PS 변화"]:
        if c in tbl.columns:
            tbl[c] = tbl[c].apply(lambda x: f"{float(x):.3f}" if pd.notna(x) else "-")

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📋 시뮬레이션 결과 (개선 폭 큰 순)</div>",
            unsafe_allow_html=True,
        )
        st.dataframe(tbl.reset_index(drop=True), use_container_width=True, height=340)


# ── 정책별 평균 공급 점수 개선 효과 바 ───────────────────────────────────────
def _render_policy_effect_bar(
    sim_df: pd.DataFrame,
    apply_staff: bool, apply_wee: bool, apply_center: bool
):
    effects = []
    if apply_staff:
        chg = (sim_df["sim_staff"] - sim_df["counseling_staff_supply_score"]).mean()
        effects.append(("전문상담교사 배치", round(chg, 3), "#1ABC9C"))
    if apply_wee:
        chg = (sim_df["sim_wee"] - sim_df["wee_class_score"]).mean()
        effects.append(("Wee클래스 신설", round(chg, 3), "#2980B9"))
    if apply_center:
        chg = (sim_df["sim_center"] - sim_df["wee_center_access_score"]).mean()
        effects.append(("Wee센터 연계 강화", round(chg, 3), "#E67E22"))

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:8px;'>"
            "📊 정책별 평균 공급 점수 개선 효과</div>",
            unsafe_allow_html=True,
        )
        if not effects:
            st.info("정책 조건을 선택하면 정책별 효과를 확인할 수 있습니다.")
            return

        names  = [e[0] for e in effects]
        values = [e[1] for e in effects]
        colors = [e[2] for e in effects]

        fig = go.Figure(go.Bar(
            x=values, y=names, orientation="h",
            marker_color=colors,
            text=[f"+{v:.3f}" if v >= 0 else f"{v:.3f}" for v in values],
            textposition="outside", textfont=dict(size=11),
        ))
        x_max = max(values) + 0.05 if values else 0.5
        fig.update_layout(
            height=220,
            margin=dict(l=10, r=60, t=10, b=20),
            xaxis=dict(title="평균 점수 변화량", range=[0, x_max], tickfont=dict(size=9)),
            yaxis=dict(tickfont=dict(size=10)),
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Malgun Gothic, sans-serif"),
        )
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            "<div style='font-size:0.69rem;color:#A0AEC0;margin-top:-4px;'>"
            "※ 정책 4(고수요 프로그램 강화)는 정량 지수에 반영되지 않습니다.</div>",
            unsafe_allow_html=True,
        )


# ── 시뮬레이션 해석 박스 ─────────────────────────────────────────────────────
def _render_sim_interpretation(
    sim_df: pd.DataFrame, target_sel: str,
    apply_staff: bool, apply_wee: bool, apply_center: bool, apply_prog: bool
):
    n = len(sim_df)
    avg_csi_b = sim_df["CSI"].mean()
    avg_csi_a = sim_df["sim_CSI"].mean()
    avg_ps_b  = sim_df["priority_score"].mean()
    avg_ps_a  = sim_df["sim_PS"].mean()
    ps_chg    = avg_ps_a - avg_ps_b
    csi_chg   = avg_csi_a - avg_csi_b
    n_imp     = int((sim_df["sim_PS_chg"] < 0).sum())

    applied_names = []
    if apply_staff:  applied_names.append("전문상담교사 배치")
    if apply_wee:    applied_names.append("Wee클래스 신설")
    if apply_center: applied_names.append("Wee센터 연계 강화")
    if apply_prog:   applied_names.append("고수요 프로그램 강화(정성)")
    pol_str = "·".join(applied_names) if applied_names else "(정책 미선택)"

    interp_items = [
        ("📋", "#2E5FA3",
         f"선택 대상: <b>{target_sel}</b> ({n}개교) | 적용 정책: <b>{pol_str}</b>"),
        ("📈", "#1ABC9C",
         f"평균 CSI: <b>{avg_csi_b:.3f}</b> → <b>{avg_csi_a:.3f}</b> "
         f"(변화: <b>{'+'if csi_chg>=0 else ''}{csi_chg:.3f}</b>)"),
        ("🎯", "#E67E22",
         f"평균 우선지원점수: <b>{avg_ps_b:.3f}</b> → <b>{avg_ps_a:.3f}</b> "
         f"(변화: <b>{'+'if ps_chg>=0 else ''}{ps_chg:.3f}</b>) — "
         + ("수요 대비 공급 부족 완화 방향" if ps_chg < 0 else "변화 없거나 증가")),
        ("🏫", "#9B59B6",
         f"우선지원점수 개선 학교: <b>{n_imp}개교</b> / 전체 {n}개교 "
         f"({n_imp/n*100:.0f}%)" if n > 0 else f"개선 학교: {n_imp}개교"),
        ("⚠️", "#F39C12",
         "본 결과는 지수 산식 기반 가상 시나리오이며 실제 정책 효과를 보장하지 않습니다. "
         "지역 여건·학교 특성·정책 실행 수준에 따라 실제 효과는 달라질 수 있습니다."),
    ]

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:10px;'>"
            "💡 시뮬레이션 해석</div>",
            unsafe_allow_html=True,
        )
        for icon, color, text in interp_items:
            st.markdown(
                f"<div style='display:flex;gap:8px;margin-bottom:10px;"
                f"padding:8px 10px;background:#F7FAFC;"
                f"border-radius:6px;border-left:3px solid {color};'>"
                f"<span style='font-size:1rem;flex-shrink:0;'>{icon}</span>"
                f"<p style='font-size:0.76rem;color:#2D3748;margin:0;line-height:1.6;'>{text}</p>"
                f"</div>",
                unsafe_allow_html=True,
            )


# ── 7. 메인 ────────────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════
# ── 11. 데이터 설명 탭 ───────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def show_data_description(df: pd.DataFrame):

    # ── 헤더 ─────────────────────────────────────────────────────────────────
    st.markdown(
        "<h1 style='font-size:1.3rem;color:#1E3A5F;margin:0 0 2px 0;font-weight:700;'>"
        "데이터 및 지표 산출 기준 설명"
        "</h1>"
        "<p style='color:#718096;font-size:0.78rem;margin:0 0 8px 0;'>"
        "본 대시보드에서 사용한 데이터 출처, 핵심 변수, 지수 산출식, 정규화 방식, "
        "유형화 기준, 정책 피드백 로직을 설명합니다."
        "</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div style='background:#EBF2FF;border-left:3px solid #2980B9;"
        "padding:6px 12px;border-radius:4px;font-size:0.74rem;color:#2C3E50;"
        "margin-bottom:18px;'>"
        "ℹ️ 본 대시보드는 경상남도 일반고등학교 상담지원 인프라의 상대적 수급 불균형을 "
        "파악하기 위한 의사결정 지원 도구이며, 실제 지원 확정 기준은 아닙니다."
        "</div>",
        unsafe_allow_html=True,
    )

    # ── Row 1: 데이터 출처 / 핵심 지표 정의 / 정규화 방식 ────────────────────
    _MH = "min-height:290px;"   # 3개 카드 공통 최소 높이

    c1, c2, c3 = st.columns(3, gap="small")

    with c1:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#2E5FA3;"
                "margin-bottom:10px;'>📦 1. 데이터 출처</div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.76rem;font-weight:700;color:#2D3748;'>· 교육통계 (KESS·나이스)</div>"
                "<div style='font-size:0.72rem;color:#718096;padding-left:10px;line-height:1.5;'>학교코드, 학교명, 학생 수, 교원 수, 전문상담교사 수 등 기본 정보</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.76rem;font-weight:700;color:#2D3748;'>· 학교알리미</div>"
                "<div style='font-size:0.72rem;color:#718096;padding-left:10px;line-height:1.5;'>Wee클래스 운영 여부, 개인·집단 상담 건수, 학교폭력 실태조사 자료</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.76rem;font-weight:700;color:#2D3748;'>· 경남교육청 Wee센터 현황</div>"
                "<div style='font-size:0.72rem;color:#718096;padding-left:10px;line-height:1.5;'>Wee센터명, 주소 — 직선거리 기반 접근성 산출</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.76rem;font-weight:700;color:#2D3748;'>· 카카오 Local API</div>"
                "<div style='font-size:0.72rem;color:#718096;padding-left:10px;line-height:1.5;'>학교·Wee센터 주소 지오코딩 (위경도 변환)</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.76rem;font-weight:700;color:#2D3748;'>· 수작업 입력</div>"
                "<div style='font-size:0.72rem;color:#718096;padding-left:10px;line-height:1.5;'>2023~2025년 학교폭력 피해 응답 학생 수, 실태조사 참여 학생 수</div></div>"
                "<div style='font-size:0.68rem;color:#A0AEC0;margin-top:4px;'>"
                "※ 모든 데이터는 2025년 기준이며, 상담·학교폭력 자료는 2023~2025년 활용</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    with c2:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#2E5FA3;"
                "margin-bottom:10px;'>📐 2. 핵심 지표 정의</div>"
                "<div style='margin-bottom:10px;padding:7px 10px;border-left:3px solid #1ABC9C;background:#F7FAFC;border-radius:4px;'>"
                "<div style='font-size:0.77rem;font-weight:700;color:#1ABC9C;margin-bottom:2px;'>상담공급지수 (CSI)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;line-height:1.5;'>학교가 보유한 상담지원 자원의 수준을 나타내는 종합 지수 (0~1)</div></div>"
                "<div style='margin-bottom:10px;padding:7px 10px;border-left:3px solid #E67E22;background:#F7FAFC;border-radius:4px;'>"
                "<div style='font-size:0.77rem;font-weight:700;color:#E67E22;margin-bottom:2px;'>상담수요지수 (CDI)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;line-height:1.5;'>해당 학교에서 상담지원이 필요할 가능성을 나타내는 종합 지수 (0~1)</div></div>"
                "<div style='margin-bottom:10px;padding:7px 10px;border-left:3px solid #C0392B;background:#F7FAFC;border-radius:4px;'>"
                "<div style='font-size:0.77rem;font-weight:700;color:#C0392B;margin-bottom:2px;'>우선지원점수 (PS)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;line-height:1.5;'>CDI − CSI. 값이 클수록 수요 대비 공급 부족 가능성이 높음</div></div>"
                "<div style='margin-bottom:10px;padding:7px 10px;border-left:3px solid #9B59B6;background:#F7FAFC;border-radius:4px;'>"
                "<div style='font-size:0.77rem;font-weight:700;color:#9B59B6;margin-bottom:2px;'>우선지원등급</div>"
                "<div style='font-size:0.72rem;color:#4A5568;line-height:1.5;'>PS 분위수 기준 4단계 분류 (최우선 지원 / 우선 지원 / 모니터링 / 안정)</div></div>"
                "</div>",
                unsafe_allow_html=True,
            )

    with c3:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#2E5FA3;"
                "margin-bottom:10px;'>⚖️ 3. 정규화 방식</div>"
                "<div style='font-size:0.76rem;color:#2D3748;margin-bottom:8px;'>"
                "모든 하위 지표는 학교 간 비교 가능성을 위해 <b>0~1 범위</b>로 정규화합니다.</div>"
                "<div style='background:#EBF2FF;border-radius:6px;padding:10px 12px;"
                "font-size:0.78rem;color:#1E3A5F;font-family:monospace;"
                "text-align:center;margin-bottom:12px;'>"
                "정규화 점수 = (값 − 최솟값) / (최댓값 − 최솟값)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;'>· 정방향 지표: 값이 클수록 점수 높음</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;'>· 일부 지표는 구간·이진 기준 직접 적용</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;'>· Wee클래스: 운영(1.0) / 미운영(0.0)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;'>· 실제 0과 결측(미공시)은 구분하여 처리</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    # ── Row 2: CSI 산출 / CDI 산출 / 정책전략 그룹 ──────────────────────────
    _MH2 = "min-height:255px;"

    d1, d2, d3 = st.columns(3, gap="small")

    with d1:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH2}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#1ABC9C;"
                "margin-bottom:8px;'>📊 4. CSI 산출 기준</div>"
                "<div style='background:#E8F8F5;border-radius:6px;padding:8px 12px;"
                "font-size:0.78rem;color:#1E3A5F;font-family:monospace;"
                "text-align:center;margin-bottom:10px;'>"
                "CSI = (상담인력 공급 + Wee클래스 + Wee센터 접근성) / 3</div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.74rem;font-weight:700;color:#1ABC9C;'>· 상담인력 공급</div>"
                "<div style='font-size:0.70rem;color:#718096;padding-left:8px;line-height:1.5;'>전문상담교사 0명 → 0.0 | 1명 이상이고 1인당 학생 500명 이상 → 0.4 | 1인당 250~500명 → 0.7 | 1인당 250명 미만 → 1.0</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.74rem;font-weight:700;color:#1ABC9C;'>· Wee클래스 운영</div>"
                "<div style='font-size:0.70rem;color:#718096;padding-left:8px;line-height:1.5;'>운영→1.0 | 미운영→0.0 | 결측→확인 필요</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.74rem;font-weight:700;color:#1ABC9C;'>· Wee센터 접근성</div>"
                "<div style='font-size:0.70rem;color:#718096;padding-left:8px;line-height:1.5;'>직선거리: 5km↓→1.0 | 5~10km→0.7 | 10~15km→0.4 | 15km↑→0.1</div></div>"
                "<div style='font-size:0.68rem;color:#A0AEC0;margin-top:4px;'>"
                "※ Wee센터 접근성은 직선거리 기준이며 실제 이동시간을 의미하지 않습니다.</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    with d2:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH2}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#E67E22;"
                "margin-bottom:8px;'>📊 5. CDI 산출 기준</div>"
                "<div style='background:#FEF9E7;border-radius:6px;padding:8px 12px;"
                "font-size:0.78rem;color:#1E3A5F;font-family:monospace;"
                "text-align:center;margin-bottom:10px;'>"
                "CDI = (수요 규모 + 실제 상담 이용 + 학교폭력 위험) / 3</div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.74rem;font-weight:700;color:#E67E22;'>· 수요 규모</div>"
                "<div style='font-size:0.70rem;color:#718096;padding-left:8px;line-height:1.5;'>학생 수 250↓→0.3 | 250~500→0.6 | 500↑→1.0</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.74rem;font-weight:700;color:#E67E22;'>· 실제 상담 이용</div>"
                "<div style='font-size:0.70rem;color:#718096;padding-left:8px;line-height:1.5;'>최근 3개년 학생·학부모 통합 상담 건수 평균 및 학생 수 대비 비율 → Min-Max 정규화 후 평균</div></div>"
                "<div style='margin-bottom:8px;'><div style='font-size:0.74rem;font-weight:700;color:#E67E22;'>· 학교폭력 위험</div>"
                "<div style='font-size:0.70rem;color:#718096;padding-left:8px;line-height:1.5;'>2023~2025년 피해 응답률 평균 → Min-Max 정규화</div></div>"
                "<div style='font-size:0.68rem;color:#A0AEC0;margin-top:4px;'>"
                "※ 학교폭력 위험 점수는 피해 응답 자료 기반 상대적 지표입니다.</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    with d3:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH2}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#9B59B6;"
                "margin-bottom:8px;'>👥 6. 정책전략 그룹 설명</div>"
                "<div style='display:flex;align-items:flex-start;gap:7px;margin-bottom:7px;'>"
                "<span style='background:#C0392B;color:white;padding:1px 7px;border-radius:8px;font-size:0.67rem;font-weight:700;white-space:nowrap;flex-shrink:0;'>최우선 개입형</span>"
                "<span style='font-size:0.71rem;color:#4A5568;line-height:1.5;'>PS 상위·핵심 불균형 → 최우선 정책 검토</span></div>"
                "<div style='display:flex;align-items:flex-start;gap:7px;margin-bottom:7px;'>"
                "<span style='background:#E67E22;color:white;padding:1px 7px;border-radius:8px;font-size:0.67rem;font-weight:700;white-space:nowrap;flex-shrink:0;'>우선 보완형</span>"
                "<span style='font-size:0.71rem;color:#4A5568;line-height:1.5;'>수요 높거나 공급 보완 필요</span></div>"
                "<div style='display:flex;align-items:flex-start;gap:7px;margin-bottom:7px;'>"
                "<span style='background:#9B59B6;color:white;padding:1px 7px;border-radius:8px;font-size:0.67rem;font-weight:700;white-space:nowrap;flex-shrink:0;'>인력 취약형</span>"
                "<span style='font-size:0.71rem;color:#4A5568;line-height:1.5;'>상담인력 공급 점수 낮음</span></div>"
                "<div style='display:flex;align-items:flex-start;gap:7px;margin-bottom:7px;'>"
                "<span style='background:#2980B9;color:white;padding:1px 7px;border-radius:8px;font-size:0.67rem;font-weight:700;white-space:nowrap;flex-shrink:0;'>접근성 보완형</span>"
                "<span style='font-size:0.71rem;color:#4A5568;line-height:1.5;'>Wee센터 접근성 점수 낮음</span></div>"
                "<div style='display:flex;align-items:flex-start;gap:7px;margin-bottom:7px;'>"
                "<span style='background:#27AE60;color:white;padding:1px 7px;border-radius:8px;font-size:0.67rem;font-weight:700;white-space:nowrap;flex-shrink:0;'>안정형</span>"
                "<span style='font-size:0.71rem;color:#4A5568;line-height:1.5;'>수요 대비 공급 비교적 안정</span></div>"
                "</div>",
                unsafe_allow_html=True,
            )

    # ── Row 3: 분석 흐름 다이어그램 ─────────────────────────────────────────
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "margin-bottom:12px;'>🔄 7. 전체 분석 흐름</div>",
            unsafe_allow_html=True,
        )
        steps = [
            ("📥", "데이터 수집", "공공기관 원천 데이터"),
            ("🔧", "데이터 정제", "표준화·단위 통일·결측 처리"),
            ("📐", "지표 정규화", "0~1 범위 변환"),
            ("📊", "CSI·CDI 산출", "하위 지표 평균"),
            ("🎯", "PS 계산", "CDI − CSI"),
            ("🗂️", "3×3 유형화", "수요·공급 분위 조합"),
            ("📋", "정책 피드백", "그룹별 권고 생성"),
            ("🖥️", "대시보드", "시각화 및 시뮬레이션"),
        ]
        cols = st.columns(len(steps))
        for i, (icon, title, desc) in enumerate(steps):
            with cols[i]:
                arrow = "→" if i < len(steps) - 1 else ""
                st.markdown(
                    f"<div style='background:white;border:1px solid #E2E8F0;"
                    f"border-top:3px solid #2E5FA3;border-radius:8px;"
                    f"padding:10px 8px;text-align:center;position:relative;'>"
                    f"<div style='font-size:1.2rem;'>{icon}</div>"
                    f"<div style='font-size:0.74rem;font-weight:700;color:#1E3A5F;"
                    f"margin:4px 0 2px;'>{title}</div>"
                    f"<div style='font-size:0.65rem;color:#718096;line-height:1.4;'>{desc}</div>"
                    f"</div>",
                    unsafe_allow_html=True,
                )
        st.markdown("<div style='margin-bottom:20px;'></div>", unsafe_allow_html=True)

    # ── Row 4: 우선지원점수·등급 (좌) + 3x3 매트릭스 (우) ───────────────────
    ps_col, mx_col = st.columns([1, 1], gap="small")

    with ps_col:
        with st.container(border=True):
            st.markdown(
                "<div style='font-size:0.88rem;font-weight:700;color:#C0392B;"
                "margin-bottom:8px;'>🎯 8. 우선지원점수 및 등급 기준</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div style='background:#FDEDEC;border-radius:6px;padding:8px 12px;"
                "font-size:0.80rem;color:#C0392B;font-family:monospace;"
                "text-align:center;margin-bottom:10px;font-weight:700;'>"
                "Priority Score (PS) = CDI − CSI"
                "</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div style='font-size:0.74rem;color:#2D3748;margin-bottom:10px;'>"
                "값이 <b>클수록</b> 수요 대비 공급 부족 가능성이 높습니다. "
                "값이 <b>0 미만</b>이면 공급이 수요보다 높은 상태로 해석할 수 있습니다. "
                "단, 현재 모델은 CSI·CDI 구성 지표에 균일 가중치를 적용하고 있어, "
                "지표 설계 구조상 공급 수준이 과대평가될 수 있는 한계가 존재합니다. "
                "따라서 우선지원점수가 음수인 경우에도 실제 지원 필요성을 단정하기 어려우며, "
                "현장 맥락과 함께 해석하는 것이 적절합니다.</div>",
                unsafe_allow_html=True,
            )
            grade_data = [
                ("1등급(최우선)", "#C0392B", "PS 상위 10%",      "최우선 지원 검토"),
                ("2등급(우선)",   "#E67E22", "PS 상위 10~20%",   "우선 지원 검토"),
                ("3등급(모니터링)","#F4D03F", "PS 중위 60%",      "정기 모니터링"),
                ("4등급(안정)",   "#27AE60", "PS 하위 20%",      "현 상태 유지"),
            ]
            rows_html = ""
            for grade, color, crit, action in grade_data:
                rows_html += (
                    f"<tr>"
                    f"<td style='padding:5px 8px;border-bottom:1px solid #F0F4F8;'>"
                    f"<span style='background:{color};color:white;padding:1px 8px;"
                    f"border-radius:8px;font-size:0.68rem;font-weight:700;'>{grade}</span></td>"
                    f"<td style='padding:5px 8px;border-bottom:1px solid #F0F4F8;"
                    f"font-size:0.72rem;color:#718096;'>{crit}</td>"
                    f"<td style='padding:5px 8px;border-bottom:1px solid #F0F4F8;"
                    f"font-size:0.72rem;color:#2D3748;'>{action}</td>"
                    f"</tr>"
                )
            st.markdown(
                f"<table style='width:100%;border-collapse:collapse;'>"
                f"<tr style='background:#EBF2FF;'>"
                f"<th style='padding:5px 8px;font-size:0.73rem;color:#1E3A5F;'>등급</th>"
                f"<th style='padding:5px 8px;font-size:0.73rem;color:#1E3A5F;'>기준</th>"
                f"<th style='padding:5px 8px;font-size:0.73rem;color:#1E3A5F;'>해석</th>"
                f"</tr>{rows_html}</table>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div style='font-size:0.68rem;color:#A0AEC0;margin-top:8px;'>"
                "※ 분위수 기반 정책 검토용 등급으로 법정 기준이 아닙니다.</div>"
                "<div style='min-height:5px;'></div>",
                unsafe_allow_html=True,
            )

    with mx_col:
        with st.container(border=True):
            st.markdown(
                "<div style='font-size:0.88rem;font-weight:700;color:#2980B9;"
                "margin-bottom:8px;'>🗂️ 9. 3×3 수요-공급 매트릭스</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div style='font-size:0.73rem;color:#4A5568;margin-bottom:10px;'>"
                "CDI·CSI 분위수(상위/중위/하위)를 조합하여 9개 유형으로 분류합니다. "
                "고정 기준만 사용 시 CDI가 특정 구간에 집중되는 한계를 분위 기준으로 보완합니다.</div>",
                unsafe_allow_html=True,
            )
            # 3x3 그리드 HTML 테이블 (셀 설명 포함)
            cell_map = {
                ("상위", "하위"): ("핵심 불균형형", "#C0392B", "수요 높음·공급 낮음 → 즉각 지원 필요"),
                ("상위", "중위"): ("고수요 보완형", "#E67E22", "수요 높음·공급 보통 → 공급 보완 검토"),
                ("상위", "상위"): ("고수요 유지관리형", "#F4D03F", "수요·공급 모두 높음 → 질 관리 중점"),
                ("중위", "하위"): ("잠재 취약형", "#9B59B6", "수요 보통·공급 낮음 → 기반 보완 필요"),
                ("중위", "중위"): ("평균 관리형", "#2980B9", "수요·공급 모두 보통 → 현 수준 유지"),
                ("중위", "상위"): ("안정 관리형", "#1ABC9C", "수요 보통·공급 높음 → 안정 운영"),
                ("하위", "하위"): ("최소 인프라 보완형", "#95A5A6", "수요·공급 모두 낮음 → 기본 인프라 구축"),
                ("하위", "중위"): ("안정 모니터링형", "#27AE60", "수요 낮음·공급 보통 → 정기 모니터링"),
                ("하위", "상위"): ("여유·거점 활용형", "#27AE60", "수요 낮음·공급 높음 → 거점 활용 검토"),
            }
            row_orders = ["상위", "중위", "하위"]
            col_orders = ["하위", "중위", "상위"]
            header = (
                "<table style='width:100%;border-collapse:collapse;font-size:0.70rem;'>"
                "<tr><th style='padding:12px 5px;text-align:center;background:#F0F4F8;color:#718096;'>"
                "수요↓ / 공급→</th>"
            )
            for c in col_orders:
                header += (f"<th style='padding:12px 5px;text-align:center;background:#EBF2FF;"
                           f"color:#1E3A5F;font-weight:700;'>공급 {c}</th>")
            header += "</tr>"
            body = ""
            for r in row_orders:
                body += (f"<tr><td style='padding:18px 5px;text-align:center;background:#EBF2FF;"
                         f"color:#1E3A5F;font-weight:700;vertical-align:middle;'>수요 {r}</td>")
                for c in col_orders:
                    name, color, desc = cell_map.get((r, c), ("?", "#BDC3C7", ""))
                    body += (
                        f"<td style='padding:14px 6px;text-align:center;"
                        f"background:{color}22;border:1px solid #E2E8F0;vertical-align:middle;'>"
                        f"<div style='font-size:0.68rem;color:{color};font-weight:700;"
                        f"margin-bottom:4px;'>{name}</div>"
                        f"<div style='font-size:0.60rem;color:#718096;line-height:1.4;'>{desc}</div>"
                        f"</td>"
                    )
                body += "</tr>"
            st.markdown(
                header + body + "</table>",
                unsafe_allow_html=True,
            )
            st.markdown("<div style='min-height:52px;'></div>", unsafe_allow_html=True)

    # ── Row 5: 변수 설명표 ───────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown(
            "<div style='font-size:0.88rem;font-weight:700;color:#1E3A5F;"
            "padding-bottom:6px;border-bottom:1px solid #E8EEF6;margin-bottom:10px;'>"
            "📋 10. 주요 변수 설명표</div>",
            unsafe_allow_html=True,
        )
        var_data = [
            ("counseling_staff_supply_score", "상담인력 공급 점수", "전문상담교사 수·학생 수 기반 공급 수준", "CSI 구성", "높을수록 공급 양호"),
            ("wee_class_score",               "Wee클래스 운영 점수", "Wee클래스 운영 여부 (0/1)",              "CSI 구성", "1.0=운영"),
            ("wee_center_access_score",        "Wee센터 접근성 점수", "학교↔Wee센터 직선거리 기반",             "CSI 구성", "높을수록 접근 용이"),
            ("CSI",                            "상담공급지수",         "위 3개 점수 평균",                        "최종 지수", "높을수록 공급 양호"),
            ("demand_size_score",              "수요 규모 점수",       "학생 수 구간 기반 잠재 수요",             "CDI 구성", "높을수록 잠재 수요 큼"),
            ("counseling_use_score",           "상담 이용률 점수",     "3개년 상담 건수 및 학생 대비 비율",       "CDI 구성", "높을수록 실제 이용 활발"),
            ("school_violence_risk_score",     "학교폭력 위험 점수",   "2023~2025 피해 응답률 평균 정규화",       "CDI 구성", "높을수록 위험 수준 높음"),
            ("CDI",                            "상담수요지수",         "위 3개 점수 평균",                        "최종 지수", "높을수록 수요 높음"),
            ("priority_score",                 "우선지원점수",         "CDI − CSI",                               "최종 지수", "클수록 공급 부족 가능성↑"),
            ("priority_level",                 "우선지원등급",         "PS 분위수 기반 4단계 등급",               "분류 결과", "최우선~안정"),
            ("csi_relative_level",             "CSI 상대 수준",        "CSI 분위수 기반 공급 하·중·상위",        "유형화 기준", "-"),
            ("cdi_relative_level",             "CDI 상대 수준",        "CDI 분위수 기반 수요 하·중·상위",        "유형화 기준", "-"),
            ("supply_demand_matrix_3x3",       "3×3 매트릭스 유형",    "cdi_relative × csi_relative 조합",       "유형화 결과", "9개 유형"),
            ("policy_strategy_group",          "정책전략 그룹",        "PS·지표 패턴 기반 5개 그룹 분류",        "정책 분류", "5개 그룹"),
            ("policy_recommendation",          "정책 추천 방향",       "그룹별 맞춤 정책 방향 텍스트",            "정책 피드백", "-"),
            ("policy_reason",                  "판단 근거",            "해당 학교 지표 패턴 기반 설명",           "정책 피드백", "-"),
        ]
        var_df = pd.DataFrame(var_data, columns=["변수명","한글명","설명","지수 구분","값의 방향"])
        st.dataframe(var_df, use_container_width=True, height=460)

    # ── Row 6: 결측 처리 / 시뮬레이션 한계 / 전체 한계 ─────────────────────
    _MH6 = "min-height:175px;"   # 13번 카드(6항목) 기준 높이 통일

    l1, l2, l3 = st.columns(3, gap="small")

    with l1:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH6}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#2E5FA3;"
                "margin-bottom:8px;'>🔍 11. 정규화·결측 처리</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #2E5FA3;line-height:1.5;'>전문상담교사 0명·Wee클래스 미운영·피해 응답 0명 → 실제 0으로 처리</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #2E5FA3;line-height:1.5;'>자료 없음·미공시·확인 불가 → 결측으로 유지</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #2E5FA3;line-height:1.5;'>CSI·CDI는 구성요소 모두 존재 시 계산 (결측 제외 평균은 참고용)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #2E5FA3;line-height:1.5;'>결측 여부는 분석 결과표에 '확인 필요'로 표기</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    with l2:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH6}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#E67E22;"
                "margin-bottom:8px;'>⚙️ 12. 시뮬레이션 해석 한계</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #E67E22;line-height:1.5;'>실제 정책 효과 예측 모델이 아니라 지수 산식 기반 가상 시나리오</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #E67E22;line-height:1.5;'>CDI(수요 지표)는 시뮬레이션에서 변경하지 않음</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #E67E22;line-height:1.5;'>하위 점수 일부가 이진·구간화되어 미세 변화가 충분히 드러나지 않을 수 있음</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:6px;padding-left:8px;border-left:2px solid #E67E22;line-height:1.5;'>향후 하위 지표를 연속형으로 전환하면 시뮬레이션 정밀도 향상 가능</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    with l3:
        with st.container(border=True):
            st.markdown(
                f"<div style='{_MH6}'>"
                "<div style='font-size:0.88rem;font-weight:700;color:#C0392B;"
                "margin-bottom:8px;'>⚠️ 13. 전체 한계 및 유의사항</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:5px;padding-left:8px;border-left:2px solid #C0392B;line-height:1.5;'>동일가중치 적용 — 지표 간 상대적 중요도 미반영</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:5px;padding-left:8px;border-left:2px solid #C0392B;line-height:1.5;'>Wee센터 접근성은 직선거리 기준 (대중교통·이동시간 미반영)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:5px;padding-left:8px;border-left:2px solid #C0392B;line-height:1.5;'>학교폭력 위험 점수는 피해 응답 자료 중심 (가해·목격 자료 미포함)</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:5px;padding-left:8px;border-left:2px solid #C0392B;line-height:1.5;'>상담 질·교사 업무 부담·대기시간 등 정성 지표 미반영</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:5px;padding-left:8px;border-left:2px solid #C0392B;line-height:1.5;'>수작업 입력 자료의 오류 가능성 존재</div>"
                "<div style='font-size:0.72rem;color:#4A5568;margin-bottom:5px;padding-left:8px;border-left:2px solid #C0392B;line-height:1.5;'>정책전략 그룹은 실제 지원 확정이 아닌 정책 검토용 분류</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    # ── 하단 유의사항 배너 ────────────────────────────────────────────────────
    st.markdown(
        "<div style='background:#FEF9E7;border:1px solid #F39C12;border-radius:8px;"
        "padding:14px 18px;margin-top:8px;'>"
        "<div style='font-size:0.82rem;font-weight:700;color:#7D6608;margin-bottom:8px;'>"
        "⚠️ 유의사항 (Disclaimer)</div>"
        "<ul style='margin:0;padding-left:18px;'>"
        "<li style='font-size:0.74rem;color:#7D6608;margin-bottom:5px;line-height:1.6;'>"
        "본 대시보드는 파일럿 프로젝트 단계의 분석 결과로, 실제 정책 의사결정을 대체하지 않습니다.</li>"
        "<li style='font-size:0.74rem;color:#7D6608;margin-bottom:5px;line-height:1.6;'>"
        "실제 활용 시에는 최신 공식 데이터 확인, 현장 검증, 전문가 자문 등 추가 검토가 필요합니다.</li>"
        "<li style='font-size:0.74rem;color:#7D6608;line-height:1.6;'>"
        "본 분석 결과의 해석 및 활용에 따른 최종 책임은 사용자에게 있습니다.</li>"
        "</ul></div>",
        unsafe_allow_html=True,
    )


def main():
    if not DATA_PATH.exists():
        st.error(
            f"입력 파일을 찾을 수 없습니다.\n\n경로: `{DATA_PATH}`")
        st.stop()

    df = load_data(DATA_PATH, SHEET).copy()
    df["priority_display"] = (
        df["priority_level"].map(PRIORITY_DISPLAY).fillna(df["priority_level"])
    )
    if df.empty:
        st.error("데이터가 비어 있습니다.")
        st.stop()

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"필수 변수 누락: `{missing}`")
        st.stop()

    selected, sigungu_f, priority_f = render_sidebar(df)
    tab_key = selected.split(" ", 1)[-1]

    if tab_key == "현황 개요":
        show_overview(df, sigungu_f, priority_f)
    elif tab_key == "지역별 분석":
        show_regional(df)
    elif tab_key == "학교 검색":
        show_school_search(df)
    elif tab_key == "유형 분석":
        show_type_analysis(df)
    elif tab_key == "시뮬레이션":
        show_simulation(df)
    elif tab_key == "데이터 설명":
        show_data_description(df)
    else:
        st.markdown("<br>", unsafe_allow_html=True)
        st.info(
            f"**{tab_key}** 탭은 추후 구현 예정입니다.\n\n"
            "현재는 **📋 현황 개요** 탭에서 주요 분석 결과를 확인할 수 있습니다.")


if __name__ == "__main__":
    main()
