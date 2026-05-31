"""
최종 분석 결과 기반 EDA 및 시각화 생성 (전체 한글화)

입력 파일:
  data/processed/gyeongnam_high_schools_policy_feedback.xlsx (policy_feedback_table 시트)

출력 파일:
  outputs/tables/전체_요약통계_2025.csv
  outputs/tables/지수별_기술통계_2025.csv
  outputs/tables/등급별_빈도표_2025.csv
  outputs/tables/수요공급유형별_빈도표_2025.csv
  outputs/tables/정책조치유형별_빈도표_2025.csv
  outputs/tables/시군구별_지수_요약_2025.csv
  outputs/tables/우선지원상위20개교_2025.csv
  outputs/tables/취약학교목록_2025.csv
  outputs/figures/상담공급지수_분포_2025.png
  outputs/figures/상담수요지수_분포_2025.png
  outputs/figures/우선지원점수_분포_2025.png
  outputs/figures/우선지원등급별_학교수_2025.png
  outputs/figures/수요공급유형별_학교수_2025.png
  outputs/figures/공급지수_수요지수_산점도_2025.png
  outputs/figures/시군구별_우선지원점수_2025.png
  outputs/figures/정책조치유형별_학교수_2025.png
  outputs/figures/공급지수_하위변수_평균비교_2025.png
  outputs/figures/수요지수_하위변수_평균비교_2025.png
  docs/eda_findings_2025.md
"""

import pathlib
import sys
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

# ── 한글 폰트 설정 ─────────────────────────────────────────────────────────
plt.rcParams["font.family"]        = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT        = pathlib.Path(__file__).resolve().parent.parent
PROCESSED   = ROOT / "data" / "processed"
TABLES_DIR  = ROOT / "outputs" / "tables"
FIGURES_DIR = ROOT / "outputs" / "figures"
DOCS_DIR    = ROOT / "docs"

for d in [TABLES_DIR, FIGURES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

INPUT_PATH = PROCESSED / "gyeongnam_high_schools_policy_feedback.xlsx"

# ── 영문 → 한글 변수명 매핑 ──────────────────────────────────────────────────
COL_MAP = {
    "school_code":                   "학교코드",
    "school_name":                   "학교명",
    "sido":                          "시도",
    "sigungu":                       "시군구",
    "counseling_staff_supply_score": "상담인력 공급 점수",
    "wee_class_score":               "Wee클래스 운영 점수",
    "wee_center_access_score":       "Wee센터 접근성 점수",
    "demand_size_score":             "상담수요 규모 점수",
    "counseling_use_score":          "실제 상담 이용 점수",
    "school_violence_risk_score":    "학교폭력 위험 점수",
    "CSI":                           "상담공급지수",
    "CDI":                           "상담수요지수",
    "priority_score":                "우선지원점수",
    "priority_level":                "우선지원등급",
    "csi_level":                     "공급수준등급",
    "cdi_level":                     "수요수준등급",
    "supply_demand_type":            "수요공급유형",
    "policy_action_type":            "정책조치유형",
    "policy_recommendation":         "정책권고사항",
    "policy_reason":                 "정책근거",
}

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 입력 파일 로드
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 입력 파일 로드")

if not INPUT_PATH.exists():
    print(f"[ERROR] 파일 없음: {INPUT_PATH}")
    sys.exit(1)

xl = pd.ExcelFile(INPUT_PATH)
sheet_used = "policy_feedback_table" if "policy_feedback_table" in xl.sheet_names else xl.sheet_names[0]
if sheet_used != "policy_feedback_table":
    print(f"  [WARN] 'policy_feedback_table' 없음 - '{sheet_used}' 시트 사용")

raw = pd.read_excel(INPUT_PATH, sheet_name=sheet_used, dtype={"school_code": str})
n_rows = len(raw)
print(f"  로드 완료: {n_rows}행 x {len(raw.columns)}열")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. 필수 변수 확인 및 한글 변수명으로 변환
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 2] 필수 변수 확인 및 한글 변수명 변환")

REQUIRED_ENG = list(COL_MAP.keys())
missing = [c for c in REQUIRED_ENG if c not in raw.columns]
if missing:
    print(f"  [ERROR] 누락 변수: {missing}")
    sys.exit(1)

# 한글로 변수명 일괄 변환
df = raw.rename(columns=COL_MAP)
print(f"  변환 완료: {len(df.columns)}개 컬럼 -> 한글화")
print(f"  컬럼 목록: {list(df.columns)}")

# 분석용 상수
CDI_MEAN = df["상담수요지수"].mean()
CSI_MEAN = df["상담공급지수"].mean()
USE_P80  = df["실제 상담 이용 점수"].quantile(0.80)
VIOL_P80 = df["학교폭력 위험 점수"].quantile(0.80)

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. 전체 요약통계
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 3] 전체 요약통계")

rows = []
def add(항목, 값): rows.append({"항목": 항목, "값": 값})

add("전체 학교 수",            n_rows)
add("시군구 수",               df["시군구"].nunique())
for col in ["상담공급지수", "상담수요지수", "우선지원점수"]:
    s = df[col].dropna()
    add(f"{col} 평균",   round(s.mean(),   4))
    add(f"{col} 최솟값", round(s.min(),    4))
    add(f"{col} 최댓값", round(s.max(),    4))
    add(f"{col} 중앙값", round(s.median(), 4))
add("우선지원점수 양수 학교 수",         (df["우선지원점수"] > 0).sum())
add("최우선 지원 학교 수",               (df["우선지원등급"] == "최우선 지원").sum())
add("우선 지원 학교 수",                 (df["우선지원등급"] == "우선 지원").sum())
add("A유형(수요높음-공급낮음) 학교 수",  (df["수요공급유형"] == "A. 수요높음-공급낮음").sum())

overall_df = pd.DataFrame(rows)
overall_df.to_csv(TABLES_DIR / "전체_요약통계_2025.csv", index=False, encoding="utf-8-sig")
print("  저장: 전체_요약통계_2025.csv")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. 지수별 기술통계
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 4] 지수별 기술통계")

INDEX_COLS = [
    "상담공급지수", "상담수요지수", "우선지원점수",
    "상담인력 공급 점수", "Wee클래스 운영 점수", "Wee센터 접근성 점수",
    "상담수요 규모 점수", "실제 상담 이용 점수", "학교폭력 위험 점수",
]
desc = (
    df[INDEX_COLS]
    .describe(percentiles=[0.25, 0.5, 0.75])
    .T
    .rename(columns={"count": "학교수", "mean": "평균", "std": "표준편차",
                     "min": "최솟값", "25%": "Q1", "50%": "중앙값",
                     "75%": "Q3", "max": "최댓값"})
    .round(4)
    .reset_index()
    .rename(columns={"index": "변수명"})
)
desc.to_csv(TABLES_DIR / "지수별_기술통계_2025.csv", index=False, encoding="utf-8-sig")
print("  저장: 지수별_기술통계_2025.csv")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. 등급 및 유형별 빈도표
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 5] 등급 및 유형별 빈도표")

def freq_table(series, order=None, col_label="구분"):
    cnt = series.value_counts()
    if order:
        cnt = cnt.reindex([o for o in order if o in cnt.index]).fillna(0).astype(int)
    pct = (cnt / n_rows * 100).round(1)
    return pd.DataFrame({col_label: cnt.index, "학교수": cnt.values, "비율(%)": pct.values})

# 등급별 빈도표 (우선지원등급 + 공급수준등급 + 수요수준등급)
level_rows = []
for col, lbl, order in [
    ("우선지원등급", "우선지원등급", ["최우선 지원", "우선 지원", "모니터링", "안정", "확인 필요"]),
    ("공급수준등급", "공급수준등급", ["공급 낮음", "공급 보통", "공급 양호", "공급 높음", "확인 필요"]),
    ("수요수준등급", "수요수준등급", ["수요 낮음", "수요 보통", "수요 높음", "수요 매우 높음", "확인 필요"]),
]:
    ft = freq_table(df[col], order, "등급")
    ft.insert(0, "변수", lbl)
    level_rows.append(ft)
pd.concat(level_rows, ignore_index=True).to_csv(
    TABLES_DIR / "등급별_빈도표_2025.csv", index=False, encoding="utf-8-sig")
print("  저장: 등급별_빈도표_2025.csv")

type_order = ["A. 수요높음-공급낮음", "B. 수요높음-공급높음",
              "C. 수요낮음-공급낮음", "D. 수요낮음-공급높음", "확인 필요"]
freq_table(df["수요공급유형"], type_order, "유형").to_csv(
    TABLES_DIR / "수요공급유형별_빈도표_2025.csv", index=False, encoding="utf-8-sig")
print("  저장: 수요공급유형별_빈도표_2025.csv")

action_order = [
    "최우선 지원 검토", "우선 지원 검토", "수요-공급 불균형 우선 개선",
    "고수요 학교 모니터링 및 프로그램 강화", "권역별 순회상담 또는 최소 인프라 보완",
    "현 수준 유지 또는 거점학교 활용", "확인 필요",
]
freq_table(df["정책조치유형"], action_order, "조치유형").to_csv(
    TABLES_DIR / "정책조치유형별_빈도표_2025.csv", index=False, encoding="utf-8-sig")
print("  저장: 정책조치유형별_빈도표_2025.csv")

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. 시군구별 요약표
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 6] 시군구별 요약표")

sigungu_df = (
    df.groupby("시군구")
    .agg(
        학교수              = ("학교코드",              "count"),
        상담공급지수_평균    = ("상담공급지수",           "mean"),
        상담수요지수_평균    = ("상담수요지수",           "mean"),
        우선지원점수_평균    = ("우선지원점수",           "mean"),
        최우선지원_수        = ("우선지원등급", lambda x: (x == "최우선 지원").sum()),
        우선지원_수          = ("우선지원등급", lambda x: (x == "우선 지원").sum()),
        A유형_수             = ("수요공급유형", lambda x: (x == "A. 수요높음-공급낮음").sum()),
        상담인력공급점수_평균 = ("상담인력 공급 점수",    "mean"),
        Wee클래스점수_평균   = ("Wee클래스 운영 점수",   "mean"),
        Wee센터접근성점수_평균= ("Wee센터 접근성 점수",  "mean"),
        상담수요규모점수_평균 = ("상담수요 규모 점수",    "mean"),
        실제상담이용점수_평균 = ("실제 상담 이용 점수",   "mean"),
        학교폭력위험점수_평균 = ("학교폭력 위험 점수",    "mean"),
    )
    .round(4)
    .reset_index()
    .sort_values("우선지원점수_평균", ascending=False)
)
sigungu_df.to_csv(TABLES_DIR / "시군구별_지수_요약_2025.csv", index=False, encoding="utf-8-sig")
print("  저장: 시군구별_지수_요약_2025.csv")

# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 우선지원 상위 20개교
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 7] 우선지원 상위 20개교")

TOP_COLS = ["학교코드", "학교명", "시군구", "상담공급지수", "상담수요지수",
            "우선지원점수", "우선지원등급", "수요공급유형", "정책조치유형", "정책권고사항"]
top20 = (
    df[df["우선지원점수"].notna()]
    .nlargest(20, "우선지원점수")[TOP_COLS]
    .reset_index(drop=True)
)
top20.index += 1
top20.index.name = "순위"
top20.to_csv(TABLES_DIR / "우선지원상위20개교_2025.csv", encoding="utf-8-sig")
print(f"  저장: 우선지원상위20개교_2025.csv ({len(top20)}개교)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. 하위 점수 취약 학교 목록
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 8] 하위 점수 취약 학교 목록")

LOW_COLS = ["학교코드", "학교명", "시군구",
            "상담인력 공급 점수", "Wee클래스 운영 점수", "Wee센터 접근성 점수",
            "상담수요 규모 점수", "실제 상담 이용 점수", "학교폭력 위험 점수",
            "상담공급지수", "상담수요지수", "우선지원점수", "우선지원등급", "수요공급유형"]

cond_staff  = df["상담인력 공급 점수"] < 0.4
cond_wee    = df["Wee클래스 운영 점수"] == 0
cond_center = df["Wee센터 접근성 점수"] <= 0.4
cond_viol   = df["학교폭력 위험 점수"] >= VIOL_P80
cond_use    = df["실제 상담 이용 점수"] >= USE_P80
any_cond    = cond_staff | cond_wee | cond_center | cond_viol | cond_use

low_df = df[any_cond][LOW_COLS].copy()

# 취약 조건 문자열 생성
def build_weakness(row):
    parts = []
    if row["상담인력 공급 점수"] < 0.4:          parts.append("상담인력 공급 낮음")
    if row["Wee클래스 운영 점수"] == 0:           parts.append("Wee클래스 미운영")
    if row["Wee센터 접근성 점수"] <= 0.4:         parts.append("Wee센터 접근성 낮음")
    if row["학교폭력 위험 점수"] >= VIOL_P80:     parts.append("학교폭력 위험 상위20%")
    if row["실제 상담 이용 점수"] >= USE_P80:     parts.append("상담이용 상위20%")
    return "; ".join(parts)

low_df = low_df.copy()
low_df["취약조건"] = low_df.apply(build_weakness, axis=1)
low_df = low_df.sort_values("우선지원점수", ascending=False).reset_index(drop=True)
low_df.to_csv(TABLES_DIR / "취약학교목록_2025.csv", index=False, encoding="utf-8-sig")
print(f"  저장: 취약학교목록_2025.csv ({len(low_df)}개교)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. 그래프 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 9] 그래프 생성")

PALETTE = {
    "A. 수요높음-공급낮음": "#E74C3C",
    "B. 수요높음-공급높음": "#E67E22",
    "C. 수요낮음-공급낮음": "#3498DB",
    "D. 수요낮음-공급높음": "#2ECC71",
    "확인 필요":           "#95A5A6",
}
LEVEL_PALETTE = {
    "최우선 지원": "#C0392B",
    "우선 지원":   "#E67E22",
    "모니터링":    "#3498DB",
    "안정":        "#27AE60",
    "확인 필요":   "#95A5A6",
}

def save_fig(fname, dpi=300):
    path = FIGURES_DIR / fname
    plt.savefig(path, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"  저장: {fname}")

# ── 그래프 1: 상담공급지수(CSI) 히스토그램 ─────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
ax.hist(df["상담공급지수"].dropna(), bins=20, color="#3498DB", edgecolor="white", linewidth=0.6)
ax.axvline(CSI_MEAN, color="#E74C3C", linestyle="--", linewidth=1.5,
           label=f"평균 {CSI_MEAN:.3f}")
ax.set_title("상담공급지수(CSI) 분포", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("상담공급지수", fontsize=11)
ax.set_ylabel("학교 수", fontsize=11)
ax.legend(fontsize=10)
ax.set_xlim(0, 1)
sns.despine()
save_fig("상담공급지수_분포_2025.png")

# ── 그래프 2: 상담수요지수(CDI) 히스토그램 ─────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
ax.hist(df["상담수요지수"].dropna(), bins=20, color="#E67E22", edgecolor="white", linewidth=0.6)
ax.axvline(CDI_MEAN, color="#2C3E50", linestyle="--", linewidth=1.5,
           label=f"평균 {CDI_MEAN:.3f}")
ax.set_title("상담수요지수(CDI) 분포", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("상담수요지수", fontsize=11)
ax.set_ylabel("학교 수", fontsize=11)
ax.legend(fontsize=10)
ax.set_xlim(0, 1)
sns.despine()
save_fig("상담수요지수_분포_2025.png")

# ── 그래프 3: 우선지원점수 히스토그램 ─────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
ax.hist(df["우선지원점수"].dropna(), bins=25, color="#8E44AD", edgecolor="white", linewidth=0.6)
ax.axvline(0, color="#E74C3C", linestyle="-", linewidth=2, label="기준선 (0)")
ax.axvline(df["우선지원점수"].mean(), color="#F39C12", linestyle="--", linewidth=1.5,
           label=f"평균 {df['우선지원점수'].mean():.3f}")
ax.set_title("우선지원점수(Priority Score) 분포", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("우선지원점수 (상담수요지수 - 상담공급지수)", fontsize=11)
ax.set_ylabel("학교 수", fontsize=11)
ax.legend(fontsize=10)
sns.despine()
save_fig("우선지원점수_분포_2025.png")

# ── 그래프 4: 우선지원등급별 학교 수 ───────────────────────────────────────
order4  = ["최우선 지원", "우선 지원", "모니터링", "안정"]
cnt4    = df["우선지원등급"].value_counts().reindex(order4).fillna(0).astype(int)
colors4 = [LEVEL_PALETTE[l] for l in order4]
fig, ax = plt.subplots(figsize=(8, 5))
bars = ax.bar(order4, cnt4.values, color=colors4, edgecolor="white", linewidth=0.6)
for bar, val in zip(bars, cnt4.values):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
            f"{val}개교\n({val/n_rows*100:.1f}%)", ha="center", va="bottom", fontsize=10)
ax.set_title("우선지원등급별 학교 수", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("우선지원등급", fontsize=11)
ax.set_ylabel("학교 수", fontsize=11)
ax.set_ylim(0, cnt4.max() * 1.25)
sns.despine()
save_fig("우선지원등급별_학교수_2025.png")

# ── 그래프 5: 수요공급유형별 학교 수 ───────────────────────────────────────
order5  = ["A. 수요높음-공급낮음", "B. 수요높음-공급높음",
           "C. 수요낮음-공급낮음", "D. 수요낮음-공급높음"]
cnt5    = df["수요공급유형"].value_counts().reindex(order5).fillna(0).astype(int)
colors5 = [PALETTE[t] for t in order5]
labels5 = ["A유형\n수요높음\n공급낮음", "B유형\n수요높음\n공급높음",
           "C유형\n수요낮음\n공급낮음", "D유형\n수요낮음\n공급높음"]
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar(labels5, cnt5.values, color=colors5, edgecolor="white", linewidth=0.6)
for bar, val in zip(bars, cnt5.values):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
            f"{val}개교\n({val/n_rows*100:.1f}%)", ha="center", va="bottom", fontsize=10)
ax.set_title("상담수요-공급 유형별 학교 수", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("수요공급 유형", fontsize=11)
ax.set_ylabel("학교 수", fontsize=11)
ax.set_ylim(0, cnt5.max() * 1.28)
sns.despine()
save_fig("수요공급유형별_학교수_2025.png")

# ── 그래프 6: 상담공급지수-상담수요지수 산점도 ─────────────────────────────
fig, ax = plt.subplots(figsize=(9, 7))
for sdt, grp in df.groupby("수요공급유형"):
    color = PALETTE.get(sdt, "#95A5A6")
    ax.scatter(grp["상담공급지수"], grp["상담수요지수"],
               c=color, label=sdt, alpha=0.75, s=60,
               edgecolors="white", linewidths=0.4)
ax.axvline(CSI_MEAN, color="#2C3E50", linestyle="--", linewidth=1.2,
           label=f"상담공급지수 평균 ({CSI_MEAN:.3f})")
ax.axhline(CDI_MEAN, color="#7F8C8D", linestyle="--", linewidth=1.2,
           label=f"상담수요지수 평균 ({CDI_MEAN:.3f})")
offset = 0.02
ax.text(0.02,             CDI_MEAN + offset, "C유형\n수요낮음-공급낮음", fontsize=8, color="#3498DB", va="bottom")
ax.text(CSI_MEAN + offset, CDI_MEAN + offset, "B유형\n수요높음-공급높음", fontsize=8, color="#E67E22", va="bottom")
ax.text(0.02,             0.02,              "C유형 영역",               fontsize=7, color="#aaa")
ax.text(CSI_MEAN + offset, 0.02,             "D유형\n수요낮음-공급높음", fontsize=8, color="#2ECC71", va="bottom")
ax.set_title("상담공급지수와 상담수요지수 산점도", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("상담공급지수(CSI)", fontsize=11)
ax.set_ylabel("상담수요지수(CDI)", fontsize=11)
ax.set_xlim(0, 1)
ax.set_ylim(0, 0.75)
ax.legend(fontsize=9, loc="upper left", framealpha=0.8)
sns.despine()
save_fig("공급지수_수요지수_산점도_2025.png")

# ── 그래프 7: 시군구별 평균 우선지원점수 가로 막대 ─────────────────────────
sg      = sigungu_df.sort_values("우선지원점수_평균", ascending=True)
colors7 = ["#E74C3C" if v > -0.15 else "#3498DB" for v in sg["우선지원점수_평균"]]
fig, ax = plt.subplots(figsize=(9, 8))
bars = ax.barh(sg["시군구"], sg["우선지원점수_평균"],
               color=colors7, edgecolor="white", linewidth=0.5)
ax.axvline(0, color="#2C3E50", linewidth=1.2)
for bar, val in zip(bars, sg["우선지원점수_평균"]):
    ax.text(val + (0.005 if val >= 0 else -0.005),
            bar.get_y() + bar.get_height() / 2,
            f"{val:.3f}", va="center",
            ha="left" if val >= 0 else "right", fontsize=8.5)
ax.set_title("시군구별 평균 우선지원점수", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("평균 우선지원점수 (상담수요지수 - 상담공급지수)", fontsize=11)
ax.set_ylabel("시군구", fontsize=11)
sns.despine()
plt.tight_layout()
save_fig("시군구별_우선지원점수_2025.png")

# ── 그래프 8: 정책조치유형별 학교 수 ──────────────────────────────────────
order8  = [a for a in action_order if (df["정책조치유형"] == a).sum() > 0]
cnt8    = df["정책조치유형"].value_counts().reindex(order8).fillna(0).astype(int)
labels8 = ["최우선\n지원 검토", "우선\n지원 검토", "수요-공급\n불균형\n우선 개선",
           "고수요\n모니터링", "순회상담\n인프라 보완", "현 수준 유지\n거점 활용"][:len(order8)]
colors8 = ["#C0392B", "#E67E22", "#E74C3C", "#F39C12", "#3498DB", "#27AE60"][:len(order8)]
fig, ax = plt.subplots(figsize=(11, 5))
bars = ax.bar(labels8, cnt8.values, color=colors8, edgecolor="white", linewidth=0.6)
for bar, val in zip(bars, cnt8.values):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.4,
            f"{val}개교", ha="center", va="bottom", fontsize=10)
ax.set_title("정책 조치 유형별 학교 수", fontsize=14, fontweight="bold", pad=12)
ax.set_xlabel("정책 조치 유형", fontsize=11)
ax.set_ylabel("학교 수", fontsize=11)
ax.set_ylim(0, cnt8.max() * 1.22)
sns.despine()
save_fig("정책조치유형별_학교수_2025.png")

# ── 그래프 9: 상담공급지수 하위변수 평균 비교 ─────────────────────────────
csi_labels = ["상담인력\n공급 점수", "Wee클래스\n운영 점수", "Wee센터\n접근성 점수"]
csi_cols   = ["상담인력 공급 점수", "Wee클래스 운영 점수", "Wee센터 접근성 점수"]
csi_means  = [df[c].mean() for c in csi_cols]
fig, ax = plt.subplots(figsize=(7, 5))
bars = ax.bar(csi_labels, csi_means, color=["#E74C3C", "#3498DB", "#2ECC71"],
              edgecolor="white", linewidth=0.6)
for bar, val in zip(bars, csi_means):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.01,
            f"{val:.3f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
ax.axhline(CSI_MEAN, color="#7F8C8D", linestyle="--", linewidth=1,
           label=f"상담공급지수 전체 평균 ({CSI_MEAN:.3f})")
ax.set_title("상담공급지수(CSI) 하위 변수 평균 비교", fontsize=14, fontweight="bold", pad=12)
ax.set_ylabel("평균 점수", fontsize=11)
ax.set_ylim(0, 1.1)
ax.legend(fontsize=9)
sns.despine()
save_fig("공급지수_하위변수_평균비교_2025.png")

# ── 그래프 10: 상담수요지수 하위변수 평균 비교 ────────────────────────────
cdi_labels = ["상담수요\n규모 점수", "실제 상담\n이용 점수", "학교폭력\n위험 점수"]
cdi_cols   = ["상담수요 규모 점수", "실제 상담 이용 점수", "학교폭력 위험 점수"]
cdi_means  = [df[c].mean() for c in cdi_cols]
fig, ax = plt.subplots(figsize=(7, 5))
bars = ax.bar(cdi_labels, cdi_means, color=["#E67E22", "#9B59B6", "#E74C3C"],
              edgecolor="white", linewidth=0.6)
for bar, val in zip(bars, cdi_means):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.01,
            f"{val:.3f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
ax.axhline(CDI_MEAN, color="#7F8C8D", linestyle="--", linewidth=1,
           label=f"상담수요지수 전체 평균 ({CDI_MEAN:.3f})")
ax.set_title("상담수요지수(CDI) 하위 변수 평균 비교", fontsize=14, fontweight="bold", pad=12)
ax.set_ylabel("평균 점수", fontsize=11)
ax.set_ylim(0, 1.1)
ax.legend(fontsize=9)
sns.despine()
save_fig("수요지수_하위변수_평균비교_2025.png")

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. EDA 해석 문서 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 10] EDA 해석 문서 생성")

top5_sg = sigungu_df.head(5)[["시군구", "우선지원점수_평균", "최우선지원_수", "우선지원_수"]].to_string(index=False)
a_cnt   = (df["수요공급유형"] == "A. 수요높음-공급낮음").sum()
b_cnt   = (df["수요공급유형"] == "B. 수요높음-공급높음").sum()
c_cnt   = (df["수요공급유형"] == "C. 수요낮음-공급낮음").sum()
d_cnt   = (df["수요공급유형"] == "D. 수요낮음-공급높음").sum()
top1    = df.nlargest(1, "우선지원점수")[["학교명", "시군구", "우선지원점수"]].iloc[0]

findings_md = f"""# EDA 분석 결과 해석

## 1. 전체 분석 대상
- 분석 대상: 경상남도 일반고등학교 {n_rows}개교 (18개 시군구)
- 분석 기준연도: 2023~2025학년도

## 2. 상담공급지수(CSI) 분포 해석
- 평균: {CSI_MEAN:.4f}, 범위: {df['상담공급지수'].min():.4f} ~ {df['상담공급지수'].max():.4f}
- CSI 평균(0.65)은 표면적으로 양호한 수준이나, 하위 구성요소 간 격차가 뚜렷하다.
  - Wee클래스 운영 점수 평균 {df['Wee클래스 운영 점수'].mean():.3f}: 대부분 학교 운영 중
  - Wee센터 접근성 점수 평균 {df['Wee센터 접근성 점수'].mean():.3f}: 접근성 양호
  - **상담인력 공급 점수 평균 {df['상담인력 공급 점수'].mean():.3f}: 심각한 취약 수준**
- 상담인력 부족이 Wee클래스·센터 인프라 수치에 가려져 CSI가 과대 추정될 가능성이 있다.

## 3. 상담수요지수(CDI) 분포 해석
- 평균: {CDI_MEAN:.4f}, 범위: {df['상담수요지수'].min():.4f} ~ {df['상담수요지수'].max():.4f}
- 상담수요 규모 점수 평균 {df['상담수요 규모 점수'].mean():.3f}: 학생 수 기반 수요 잠재성 높음
- **실제 상담 이용 점수 평균 {df['실제 상담 이용 점수'].mean():.3f}: 낮음** — 미충족 수요 가능성
- 학교폭력 위험 점수 평균 {df['학교폭력 위험 점수'].mean():.3f}: 대체로 낮음

## 4. 우선지원점수(Priority Score) 분포 해석
- 평균: {df['우선지원점수'].mean():.4f}, 범위: {df['우선지원점수'].min():.4f} ~ {df['우선지원점수'].max():.4f}
- {(df['우선지원점수'] > 0).sum()}개교만 양수(수요 > 공급), {(df['우선지원점수'] <= 0).sum()}개교는 음수(공급 >= 수요)
- 분위수 기반 상대비교로 상위 20% 약 29개교를 우선 지원 검토 대상으로 도출

## 5. 우선지원등급별 해석
- 최우선 지원 {(df['우선지원등급']=='최우선 지원').sum()}개교 + 우선 지원 {(df['우선지원등급']=='우선 지원').sum()}개교 = 총 {(df['우선지원등급'].isin(['최우선 지원','우선 지원'])).sum()}개교(약 20%)
- 우선지원점수 최고 학교: {top1['학교명']}({top1['시군구']}, 우선지원점수={top1['우선지원점수']:.4f})
- 모니터링 {(df['우선지원등급']=='모니터링').sum()}개교(59.6%): 현 수준 유지, 정기 모니터링 필요
- 안정 {(df['우선지원등급']=='안정').sum()}개교(20.5%): 상대적으로 공급 여유 있는 학교

## 6. CSI-CDI 유형화 결과 해석
- A유형(수요높음-공급낮음): {a_cnt}개교({a_cnt/n_rows*100:.1f}%) — 핵심 우선지원 대상
- B유형(수요높음-공급높음): {b_cnt}개교({b_cnt/n_rows*100:.1f}%) — 가장 많음, 고수요 유지관리 필요
- C유형(수요낮음-공급낮음): {c_cnt}개교({c_cnt/n_rows*100:.1f}%) — 인프라 최소 보완 필요
- D유형(수요낮음-공급높음): {d_cnt}개교({d_cnt/n_rows*100:.1f}%) — 상대적 안정, 거점 활용 검토 가능

## 7. 시군구별 우선지원 경향
우선지원점수 평균 상위 시군구 (내림차순):
{top5_sg}

## 8. 정책 피드백 유형별 해석
- 상담인력 공급 점수 < 0.4: {(df['상담인력 공급 점수'] < 0.4).sum()}개교({(df['상담인력 공급 점수'] < 0.4).sum()/n_rows*100:.1f}%) — 전문상담교사 배치 우선 검토
- Wee클래스 미운영: {(df['Wee클래스 운영 점수'] == 0).sum()}개교 — Wee클래스 설치 또는 공동 운영 검토
- Wee센터 접근성 낮음: {(df['Wee센터 접근성 점수'] <= 0.4).sum()}개교 — 이동형·온라인 상담 연계 검토
- 학교폭력 위험 상위 20%: {(df['학교폭력 위험 점수'] >= VIOL_P80).sum()}개교 — 학교폭력 상담지원 집중 검토

## 9. 보고서·PPT 핵심 해석 문장
1. "경남 일반고 146개교 중 약 20%인 29개교가 상담 수요 대비 공급이 상대적으로 부족하여 우선 지원 검토 대상으로 도출되었다."
2. "상담인력 공급 점수 평균(0.27)이 Wee클래스(0.89)·Wee센터 접근성(0.80) 점수에 비해 현저히 낮아, 인프라는 갖춰져 있으나 전문 인력 배치가 취약한 구조임을 확인하였다."
3. "실제 상담 이용 점수 평균(0.28)이 낮아, 학생 수 규모 대비 상담 이용률이 낮거나 미충족 수요가 존재할 가능성이 있다."
4. "A유형(수요높음-공급낮음) 20개교는 상담 수요와 공급 간 불균형이 가장 두드러진 핵심 우선지원 유형이다."

## 10. 주의사항 및 한계
- 본 분석의 모든 등급·유형·피드백은 실제 지원 확정이 아니라 정책 검토 우선순위이다.
- CSI와 CDI는 동일가중치 기반으로 산출되어 구성요소의 실제 중요도를 반드시 반영한다고 할 수 없다.
- 상담 이용 건수 기반의 실제 상담 이용 점수는 학교별 기록 방식 차이로 인한 편차가 있을 수 있다.
- 학교별 특수 상황, 지역 교통 여건 등 정성적 요인은 정량 분석에서 완전히 포착되지 않는다.
"""

MD_PATH = DOCS_DIR / "eda_findings_2025.md"
MD_PATH.write_text(findings_md, encoding="utf-8")
print(f"  저장: eda_findings_2025.md")

# ════════════════════════════════════════════════════════════════════════════
# 최종 요약 출력
# ════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("EDA 및 시각화 생성 완료 (전체 한글화)")
print("=" * 60)
print(f"\n요약표: {len(list(TABLES_DIR.glob('*.csv')))}개 파일 -> {TABLES_DIR}")
print(f"그래프: {len(list(FIGURES_DIR.glob('*.png')))}개 파일 -> {FIGURES_DIR}")
print(f"해석 문서: {MD_PATH}")
