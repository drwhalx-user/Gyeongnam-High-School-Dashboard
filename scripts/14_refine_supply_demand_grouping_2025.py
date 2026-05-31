"""
CDI-CSI 상대 분위수 기반 정교한 학교 유형화 및 정책전략 그룹 생성

기존 절대 기준 분류(csi_level, cdi_level, supply_demand_type)를 유지하면서,
CDI가 0.3~0.6 구간에 집중되는 한계를 보완하기 위해
상대 분위수 기반 3x3 수요-공급 매트릭스와 정책전략 그룹을 추가한다.

입력 파일:
  data/processed/gyeongnam_high_schools_policy_feedback.xlsx

출력 파일:
  data/processed/gyeongnam_high_schools_policy_feedback_refined.xlsx
  outputs/tables/refined_supply_demand_matrix_summary_2025.csv
  outputs/tables/refined_policy_strategy_summary_2025.csv
  outputs/tables/refined_sigungu_strategy_summary_2025.csv
  outputs/figures/ (그래프 영문명 + 국문명 각 1쌍)
  docs/refined_grouping_methodology_2025.md
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

for d in [PROCESSED, TABLES_DIR, FIGURES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

INPUT_PATH  = PROCESSED / "gyeongnam_high_schools_policy_feedback.xlsx"
OUTPUT_PATH = PROCESSED / "gyeongnam_high_schools_policy_feedback_refined.xlsx"

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

df = pd.read_excel(INPUT_PATH, sheet_name=sheet_used, dtype={"school_code": str})
n_rows = len(df)
print(f"  로드 완료: {n_rows}행 x {len(df.columns)}열 (시트: {sheet_used})")
print(f"  컬럼: {list(df.columns)}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. 필수 변수 확인
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 2] 필수 변수 확인")

REQUIRED = [
    "school_code", "school_name", "sido", "sigungu",
    "counseling_staff_supply_score", "wee_class_score", "wee_center_access_score",
    "demand_size_score", "counseling_use_score", "school_violence_risk_score",
    "CSI", "CDI", "priority_score", "priority_level",
    "csi_level", "cdi_level", "supply_demand_type",
    "policy_action_type", "policy_recommendation", "policy_reason",
]
missing = [c for c in REQUIRED if c not in df.columns]
if missing:
    print(f"  [ERROR] 누락 변수: {missing}")
    sys.exit(1)
print("  필수 변수 전체 확인 완료")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. CDI 현재 분포 확인 (기존 분류의 한계 출력)
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 3] 기존 분류 분포 확인")

print("  [기존 cdi_level 분포]")
for lv, cnt in df["cdi_level"].value_counts().items():
    print(f"    {lv}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

print("  [기존 csi_level 분포]")
for lv, cnt in df["csi_level"].value_counts().items():
    print(f"    {lv}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

print(f"\n  CDI 범위: {df['CDI'].min():.4f} ~ {df['CDI'].max():.4f}")
print(f"  CDI 0.3~0.6 구간 학교 수: {((df['CDI']>=0.3) & (df['CDI']<0.6)).sum()}개교 "
      f"({((df['CDI']>=0.3) & (df['CDI']<0.6)).sum()/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. CDI 상대 분위수 기준 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 4] CDI 상대 분위수 기준 생성")

cdi_valid = df["CDI"].dropna()
CDI_P20 = cdi_valid.quantile(0.20)
CDI_P80 = cdi_valid.quantile(0.80)
print(f"  CDI P20 = {CDI_P20:.4f}, P80 = {CDI_P80:.4f}")

def assign_cdi_relative(v):
    if pd.isna(v):    return "확인 필요"
    elif v > CDI_P80: return "수요 상위"
    elif v >= CDI_P20: return "수요 중위"
    else:             return "수요 하위"

df["cdi_relative_level"] = df["CDI"].apply(assign_cdi_relative)
print("  cdi_relative_level 분포:")
for lv in ["수요 상위", "수요 중위", "수요 하위", "확인 필요"]:
    cnt = (df["cdi_relative_level"] == lv).sum()
    print(f"    {lv}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. CSI 상대 분위수 기준 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 5] CSI 상대 분위수 기준 생성")

csi_valid = df["CSI"].dropna()
CSI_P20 = csi_valid.quantile(0.20)
CSI_P80 = csi_valid.quantile(0.80)
print(f"  CSI P20 = {CSI_P20:.4f}, P80 = {CSI_P80:.4f}")

def assign_csi_relative(v):
    if pd.isna(v):    return "확인 필요"
    elif v > CSI_P80: return "공급 상위"
    elif v >= CSI_P20: return "공급 중위"
    else:             return "공급 하위"

df["csi_relative_level"] = df["CSI"].apply(assign_csi_relative)
print("  csi_relative_level 분포:")
for lv in ["공급 상위", "공급 중위", "공급 하위", "확인 필요"]:
    cnt = (df["csi_relative_level"] == lv).sum()
    print(f"    {lv}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. 3x3 수요-공급 매트릭스 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 6] 3x3 수요-공급 매트릭스 생성")

MATRIX_MAP = {
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

def assign_matrix(row):
    d = row["cdi_relative_level"]
    s = row["csi_relative_level"]
    if d == "확인 필요" or s == "확인 필요":
        return "확인 필요"
    return MATRIX_MAP.get((d, s), "확인 필요")

df["supply_demand_matrix_3x3"] = df[["cdi_relative_level", "csi_relative_level"]].apply(
    assign_matrix, axis=1
)

MATRIX_ORDER = [
    "핵심 불균형형", "고수요 보완형", "고수요 유지관리형",
    "잠재 취약형", "평균 관리형", "안정 관리형",
    "최소 인프라 보완형", "안정 모니터링형", "여유·거점 활용형", "확인 필요",
]
print("  supply_demand_matrix_3x3 분포:")
for t in MATRIX_ORDER:
    cnt = (df["supply_demand_matrix_3x3"] == t).sum()
    if cnt > 0:
        print(f"    {t}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 정책전략 그룹 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 7] 정책전략 그룹 생성")

def assign_strategy_group(row):
    pl  = row["priority_level"]
    m3  = row["supply_demand_matrix_3x3"]
    css = row["counseling_staff_supply_score"]
    wca = row["wee_center_access_score"]

    # 결측 확인
    if pd.isna(row["CDI"]) or pd.isna(row["CSI"]) or pd.isna(row["priority_score"]):
        return "확인 필요형"
    # 최우선 개입형 (priority_level 우선 또는 핵심 불균형형)
    if pl == "최우선 지원" or m3 == "핵심 불균형형":
        return "최우선 개입형"
    # 우선 보완형
    if pl == "우선 지원" or m3 == "고수요 보완형":
        return "우선 보완형"
    # 고수요 유지관리형
    if m3 == "고수요 유지관리형":
        return "고수요 유지관리형"
    # 인력 취약형
    if css < 0.4:
        return "인력 취약형"
    # 접근성 보완형
    if wca <= 0.4:
        return "접근성 보완형"
    # 최소 인프라 보완형
    if m3 in ("최소 인프라 보완형", "잠재 취약형"):
        return "최소 인프라 보완형"
    # 안정형
    return "안정형"

df["policy_strategy_group"] = df.apply(assign_strategy_group, axis=1)

# 정책전략 태그 (해당되는 모든 조건)
def build_strategy_tags(row):
    tags = []
    if pd.isna(row["CDI"]) or pd.isna(row["CSI"]):
        return "지표 결측"
    if row["priority_level"] == "최우선 지원":    tags.append("priority:최우선지원")
    if row["priority_level"] == "우선 지원":      tags.append("priority:우선지원")
    if row["supply_demand_matrix_3x3"] == "핵심 불균형형": tags.append("matrix:핵심불균형")
    if row["supply_demand_matrix_3x3"] == "고수요 보완형": tags.append("matrix:고수요보완")
    if row["supply_demand_matrix_3x3"] == "고수요 유지관리형": tags.append("matrix:고수요유지")
    if row["counseling_staff_supply_score"] < 0.4: tags.append("weak:상담인력부족")
    if row["wee_center_access_score"] <= 0.4:      tags.append("weak:Wee센터접근성낮음")
    if row["supply_demand_matrix_3x3"] in ("최소 인프라 보완형", "잠재 취약형"):
        tags.append("matrix:인프라취약")
    return "; ".join(tags) if tags else "해당없음"

df["policy_strategy_tags"] = df.apply(build_strategy_tags, axis=1)

# 정책전략 설명
STRATEGY_DESC = {
    "최우선 개입형":       "상담수요가 상대적으로 높고 상담공급이 부족하거나 우선지원점수가 높아 교육청 차원의 우선 지원 검토가 필요한 유형",
    "우선 보완형":         "상담수요가 높은 편으로 기존 인프라 보강 또는 상담 프로그램 확대 검토가 필요한 유형",
    "고수요 유지관리형":   "상담수요가 높지만 공급도 비교적 확보되어 있어 기존 인프라 유지와 프로그램 질 관리가 필요한 유형",
    "인력 취약형":         "전문상담교사 공급 점수가 낮아 전문상담교사 배치 또는 순회상담 연계 검토가 필요한 유형",
    "접근성 보완형":       "Wee센터 접근성이 낮아 온라인 상담, 이동형 상담 또는 권역별 Wee센터 연계 강화가 필요한 유형",
    "최소 인프라 보완형":  "상담수요는 높지 않더라도 공급 기반이 낮아 권역별 순회상담 등 최소 인프라 보완이 필요한 유형",
    "안정형":              "현재 지표상 수요 대비 공급이 상대적으로 안정적인 유형",
    "확인 필요형":         "주요 지표 결측으로 추가 확인이 필요한 유형",
}
df["policy_strategy_description"] = df["policy_strategy_group"].map(STRATEGY_DESC)

STRATEGY_ORDER = [
    "최우선 개입형", "우선 보완형", "고수요 유지관리형",
    "인력 취약형", "접근성 보완형", "최소 인프라 보완형",
    "안정형", "확인 필요형",
]
print("  policy_strategy_group 분포:")
for g in STRATEGY_ORDER:
    cnt = (df["policy_strategy_group"] == g).sum()
    if cnt > 0:
        print(f"    {g}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. 요약표 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 8] 요약표 생성")

SUB_COLS = [
    "counseling_staff_supply_score", "wee_class_score", "wee_center_access_score",
    "demand_size_score", "counseling_use_score", "school_violence_risk_score",
]

# 8-1. 3x3 매트릭스 요약
matrix_agg = (
    df.groupby("supply_demand_matrix_3x3")
    .agg(
        학교수              = ("school_code",    "count"),
        CSI_평균            = ("CSI",            "mean"),
        CDI_평균            = ("CDI",            "mean"),
        priority_score_평균 = ("priority_score", "mean"),
        상담인력공급점수_평균 = ("counseling_staff_supply_score", "mean"),
        Wee클래스점수_평균   = ("wee_class_score",               "mean"),
        Wee센터접근성점수_평균= ("wee_center_access_score",      "mean"),
        상담수요규모점수_평균 = ("demand_size_score",             "mean"),
        실제상담이용점수_평균 = ("counseling_use_score",          "mean"),
        학교폭력위험점수_평균 = ("school_violence_risk_score",    "mean"),
    )
    .round(4)
    .reset_index()
)
matrix_agg["비율(%)"] = (matrix_agg["학교수"] / n_rows * 100).round(1)
matrix_agg["_sort"] = matrix_agg["supply_demand_matrix_3x3"].map(
    {t: i for i, t in enumerate(MATRIX_ORDER)}
).fillna(99)
matrix_agg = matrix_agg.sort_values("_sort").drop(columns=["_sort"])

matrix_agg.to_csv(
    TABLES_DIR / "refined_supply_demand_matrix_summary_2025.csv",
    index=False, encoding="utf-8-sig")
print("  저장: refined_supply_demand_matrix_summary_2025.csv")
print(matrix_agg[["supply_demand_matrix_3x3","학교수","비율(%)","CSI_평균","CDI_평균","priority_score_평균"]].to_string(index=False))

# 8-2. 정책전략 그룹 요약
strategy_agg = (
    df.groupby("policy_strategy_group")
    .agg(
        학교수              = ("school_code",    "count"),
        CSI_평균            = ("CSI",            "mean"),
        CDI_평균            = ("CDI",            "mean"),
        priority_score_평균 = ("priority_score", "mean"),
        최우선지원_수       = ("priority_level", lambda x: (x == "최우선 지원").sum()),
        우선지원_수         = ("priority_level", lambda x: (x == "우선 지원").sum()),
    )
    .round(4)
    .reset_index()
)
strategy_agg["비율(%)"] = (strategy_agg["학교수"] / n_rows * 100).round(1)
strategy_agg["_sort"] = strategy_agg["policy_strategy_group"].map(
    {g: i for i, g in enumerate(STRATEGY_ORDER)}
).fillna(99)
strategy_agg = strategy_agg.sort_values("_sort").drop(columns=["_sort"])

strategy_agg.to_csv(
    TABLES_DIR / "refined_policy_strategy_summary_2025.csv",
    index=False, encoding="utf-8-sig")
print("\n  저장: refined_policy_strategy_summary_2025.csv")
print(strategy_agg[["policy_strategy_group","학교수","비율(%)","CSI_평균","CDI_평균"]].to_string(index=False))

# 8-3. 시군구별 정책전략 요약
sigungu_agg = (
    df.groupby("sigungu")
    .agg(
        학교수              = ("school_code",    "count"),
        CSI_평균            = ("CSI",            "mean"),
        CDI_평균            = ("CDI",            "mean"),
        priority_score_평균 = ("priority_score", "mean"),
        핵심불균형형_수      = ("supply_demand_matrix_3x3", lambda x: (x == "핵심 불균형형").sum()),
        최우선개입형_수      = ("policy_strategy_group",   lambda x: (x == "최우선 개입형").sum()),
        우선보완형_수        = ("policy_strategy_group",   lambda x: (x == "우선 보완형").sum()),
        인력취약형_수        = ("policy_strategy_group",   lambda x: (x == "인력 취약형").sum()),
        접근성보완형_수      = ("policy_strategy_group",   lambda x: (x == "접근성 보완형").sum()),
    )
    .round(4)
    .reset_index()
    .sort_values("priority_score_평균", ascending=False)
)
sigungu_agg.to_csv(
    TABLES_DIR / "refined_sigungu_strategy_summary_2025.csv",
    index=False, encoding="utf-8-sig")
print("\n  저장: refined_sigungu_strategy_summary_2025.csv")

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. 그래프 생성 (영문명 + 국문명 각 1쌍)
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 9] 그래프 생성 (영문 + 국문 각 1쌍)")

STRATEGY_PALETTE = {
    "최우선 개입형":      "#C0392B",
    "우선 보완형":        "#E67E22",
    "고수요 유지관리형":  "#F1C40F",
    "인력 취약형":        "#9B59B6",
    "접근성 보완형":      "#3498DB",
    "최소 인프라 보완형": "#1ABC9C",
    "안정형":             "#27AE60",
    "확인 필요형":        "#95A5A6",
}
MATRIX_PALETTE = {
    "핵심 불균형형":    "#C0392B",
    "고수요 보완형":    "#E67E22",
    "고수요 유지관리형":"#F1C40F",
    "잠재 취약형":      "#9B59B6",
    "평균 관리형":      "#3498DB",
    "안정 관리형":      "#2ECC71",
    "최소 인프라 보완형":"#1ABC9C",
    "안정 모니터링형":  "#27AE60",
    "여유·거점 활용형": "#BDC3C7",
    "확인 필요":        "#95A5A6",
}

def save_fig_both(eng_name, kor_name, dpi=300):
    """영문명과 국문명 두 파일로 저장"""
    for fname in [eng_name, kor_name]:
        plt.savefig(FIGURES_DIR / fname, dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"  저장: {eng_name}  /  {kor_name}")

# ── 그래프 1: 3x3 매트릭스 학교 수 막대그래프 ─────────────────────────────
present_types = [t for t in MATRIX_ORDER if (df["supply_demand_matrix_3x3"] == t).sum() > 0]
cnt_matrix = df["supply_demand_matrix_3x3"].value_counts().reindex(present_types).fillna(0).astype(int)
colors_m = [MATRIX_PALETTE.get(t, "#aaa") for t in present_types]

fig, ax = plt.subplots(figsize=(12, 5))
bars = ax.bar(range(len(present_types)), cnt_matrix.values, color=colors_m,
              edgecolor="white", linewidth=0.6)
ax.set_xticks(range(len(present_types)))
ax.set_xticklabels(present_types, fontsize=9, rotation=20, ha="right")
for bar, val in zip(bars, cnt_matrix.values):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
            f"{val}개교\n({val/n_rows*100:.1f}%)", ha="center", va="bottom", fontsize=8.5)
ax.set_title("3x3 수요-공급 매트릭스 유형별 학교 수", fontsize=13, fontweight="bold", pad=12)
ax.set_ylabel("학교 수", fontsize=11)
ax.set_ylim(0, cnt_matrix.max() * 1.28)
sns.despine()
plt.tight_layout()
save_fig_both(
    "refined_matrix_3x3_counts_2025.png",
    "정교화_3x3매트릭스_유형별학교수_2025.png"
)

# ── 그래프 2: 정책전략 그룹 학교 수 막대그래프 ────────────────────────────
present_groups = [g for g in STRATEGY_ORDER if (df["policy_strategy_group"] == g).sum() > 0]
cnt_strategy = df["policy_strategy_group"].value_counts().reindex(present_groups).fillna(0).astype(int)
colors_s = [STRATEGY_PALETTE.get(g, "#aaa") for g in present_groups]

fig, ax = plt.subplots(figsize=(11, 5))
bars = ax.bar(range(len(present_groups)), cnt_strategy.values, color=colors_s,
              edgecolor="white", linewidth=0.6)
ax.set_xticks(range(len(present_groups)))
ax.set_xticklabels(present_groups, fontsize=9, rotation=15, ha="right")
for bar, val in zip(bars, cnt_strategy.values):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
            f"{val}개교\n({val/n_rows*100:.1f}%)", ha="center", va="bottom", fontsize=9)
ax.set_title("정책전략 그룹별 학교 수", fontsize=13, fontweight="bold", pad=12)
ax.set_ylabel("학교 수", fontsize=11)
ax.set_ylim(0, cnt_strategy.max() * 1.25)
sns.despine()
plt.tight_layout()
save_fig_both(
    "refined_strategy_group_counts_2025.png",
    "정교화_정책전략그룹별학교수_2025.png"
)

# ── 그래프 3: 3x3 매트릭스 산점도 (CSI x CDI, 매트릭스 유형 색상) ──────────
fig, ax = plt.subplots(figsize=(10, 7))
for mtype, grp in df.groupby("supply_demand_matrix_3x3"):
    color = MATRIX_PALETTE.get(mtype, "#aaa")
    ax.scatter(grp["CSI"], grp["CDI"], c=color, label=mtype,
               alpha=0.80, s=65, edgecolors="white", linewidths=0.4)

# 분위수 기준선
ax.axvline(CSI_P20, color="#7F8C8D", linestyle=":", linewidth=1.0, alpha=0.7)
ax.axvline(CSI_P80, color="#7F8C8D", linestyle=":", linewidth=1.0, alpha=0.7)
ax.axhline(CDI_P20, color="#BDC3C7", linestyle=":", linewidth=1.0, alpha=0.7)
ax.axhline(CDI_P80, color="#BDC3C7", linestyle=":", linewidth=1.0, alpha=0.7)

# 구간 레이블
ax.text(CSI_P20 + 0.01, 0.70, f"CSI P20\n({CSI_P20:.2f})", fontsize=7, color="#7F8C8D")
ax.text(CSI_P80 + 0.01, 0.70, f"CSI P80\n({CSI_P80:.2f})", fontsize=7, color="#7F8C8D")
ax.text(0.02, CDI_P20 + 0.01, f"CDI P20 ({CDI_P20:.2f})", fontsize=7, color="#BDC3C7")
ax.text(0.02, CDI_P80 + 0.01, f"CDI P80 ({CDI_P80:.2f})", fontsize=7, color="#BDC3C7")

ax.set_title("CSI-CDI 산점도: 3x3 수요-공급 매트릭스 유형", fontsize=13, fontweight="bold", pad=12)
ax.set_xlabel("상담공급지수(CSI)", fontsize=11)
ax.set_ylabel("상담수요지수(CDI)", fontsize=11)
ax.set_xlim(0, 1.05)
ax.set_ylim(0, 0.80)
ax.legend(fontsize=8, loc="upper left", framealpha=0.85, ncol=2)
sns.despine()
save_fig_both(
    "refined_matrix_3x3_scatter_2025.png",
    "정교화_3x3매트릭스_산점도_2025.png"
)

# ── 그래프 4: 시군구별 정책전략 그룹 누적 막대그래프 ──────────────────────
sg_pivot = (
    df.groupby(["sigungu", "policy_strategy_group"])
    .size()
    .unstack(fill_value=0)
)
# 우선지원점수 평균 내림차순 정렬
sg_order = sigungu_agg.sort_values("priority_score_평균", ascending=False)["sigungu"].tolist()
sg_pivot = sg_pivot.reindex(sg_order)

# 존재하는 그룹만 사용, 순서 고정
cols_to_plot = [g for g in STRATEGY_ORDER if g in sg_pivot.columns]
colors_stk   = [STRATEGY_PALETTE.get(g, "#aaa") for g in cols_to_plot]

fig, ax = plt.subplots(figsize=(12, 7))
bottom = np.zeros(len(sg_pivot))
for col, color in zip(cols_to_plot, colors_stk):
    vals = sg_pivot[col].values
    ax.barh(sg_pivot.index, vals, left=bottom, color=color,
            label=col, edgecolor="white", linewidth=0.4)
    bottom += vals
ax.set_title("시군구별 정책전략 그룹 분포", fontsize=13, fontweight="bold", pad=12)
ax.set_xlabel("학교 수", fontsize=11)
ax.set_ylabel("시군구", fontsize=11)
ax.legend(fontsize=8, loc="lower right", framealpha=0.85)
ax.set_xlim(0, sg_pivot.sum(axis=1).max() * 1.05)
sns.despine()
plt.tight_layout()
save_fig_both(
    "refined_sigungu_strategy_stacked_2025.png",
    "정교화_시군구별정책전략분포_2025.png"
)

# ── 그래프 5: 정책전략 그룹별 평균 CSI·CDI 비교 (점 그래프) ───────────────
present_groups_full = [g for g in STRATEGY_ORDER if g in strategy_agg["policy_strategy_group"].values]
sa = strategy_agg.set_index("policy_strategy_group").reindex(present_groups_full)

x = np.arange(len(present_groups_full))
width = 0.35

fig, ax = plt.subplots(figsize=(12, 5))
ax.bar(x - width / 2, sa["CSI_평균"], width, label="상담공급지수(CSI)",
       color="#3498DB", edgecolor="white")
ax.bar(x + width / 2, sa["CDI_평균"], width, label="상담수요지수(CDI)",
       color="#E67E22", edgecolor="white")
ax.set_xticks(x)
ax.set_xticklabels(present_groups_full, fontsize=9, rotation=15, ha="right")
ax.set_title("정책전략 그룹별 평균 상담공급지수 · 상담수요지수", fontsize=13, fontweight="bold", pad=12)
ax.set_ylabel("평균 지수", fontsize=11)
ax.set_ylim(0, 1.0)
ax.legend(fontsize=10)
ax.axhline(df["CSI"].mean(), color="#3498DB", linestyle="--", linewidth=0.8, alpha=0.5)
ax.axhline(df["CDI"].mean(), color="#E67E22", linestyle="--", linewidth=0.8, alpha=0.5)
sns.despine()
plt.tight_layout()
save_fig_both(
    "refined_strategy_csi_cdi_comparison_2025.png",
    "정교화_정책전략그룹_공급수요지수비교_2025.png"
)

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. 최종 엑셀 파일 저장
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 10] 최종 엑셀 파일 저장")

# 정렬: 정책전략 그룹 우선 -> 우선지원점수 내림차순
GROUP_SORT = {g: i for i, g in enumerate(STRATEGY_ORDER)}
df["_gsort"] = df["policy_strategy_group"].map(GROUP_SORT).fillna(99)
out_df = (
    df.sort_values(["_gsort", "priority_score"], ascending=[True, False])
    .drop(columns=["_gsort"])
    .reset_index(drop=True)
)

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    out_df.to_excel(writer,       sheet_name="refined_policy_feedback_table",      index=False)
    matrix_agg.to_excel(writer,   sheet_name="refined_supply_demand_matrix_summary", index=False)
    strategy_agg.to_excel(writer, sheet_name="refined_policy_strategy_summary",    index=False)
    sigungu_agg.to_excel(writer,  sheet_name="refined_sigungu_strategy_summary",   index=False)

print(f"  저장 완료: {OUTPUT_PATH}")
print(f"  시트: refined_policy_feedback_table ({len(out_df)}행 x {len(out_df.columns)}열)")
print(f"  신규 변수 6개: cdi_relative_level, csi_relative_level, supply_demand_matrix_3x3,")
print(f"                policy_strategy_group, policy_strategy_tags, policy_strategy_description")

# ════════════════════════════════════════════════════════════════════════════
# STEP 11. 방법론 문서 저장
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 11] 방법론 문서 저장")

matrix_dist_str = "\n".join(
    f"| {t} | {(df['supply_demand_matrix_3x3']==t).sum()}개교 "
    f"| {(df['supply_demand_matrix_3x3']==t).sum()/n_rows*100:.1f}% |"
    for t in MATRIX_ORDER if (df["supply_demand_matrix_3x3"]==t).sum() > 0
)
strategy_dist_str = "\n".join(
    f"| {g} | {(df['policy_strategy_group']==g).sum()}개교 "
    f"| {(df['policy_strategy_group']==g).sum()/n_rows*100:.1f}% |"
    for g in STRATEGY_ORDER if (df["policy_strategy_group"]==g).sum() > 0
)

methodology_md = f"""# 상대 분위수 기반 정교화 유형화 방법론

## 1. 기존 절대 기준 분류의 한계
기존 cdi_level은 절대 구간(0.3/0.6/0.8)을 기준으로 분류하였다.
그러나 경남 일반고의 CDI가 {((df['CDI']>=0.3) & (df['CDI']<0.6)).sum()}개교({((df['CDI']>=0.3) & (df['CDI']<0.6)).sum()/n_rows*100:.1f}%)가
0.3~0.6 구간에 집중되어 대부분 '수요 보통'으로 분류되었다.
이는 학교 간 차별화된 정책 우선순위 도출에 한계가 있었다.

## 2. 상대 분위수 기준을 추가한 이유
경남 일반고 146개교의 실제 분포를 반영하여 학교 간 상대적 위치를 파악하기 위해
CDI와 CSI 각각의 분위수 기준(P20, P80)을 이용한 보조 분류를 추가하였다.
이 분류는 기존 절대 기준 분류를 대체하는 것이 아니라, 이를 보완하는 목적으로 사용한다.

## 3. CDI 상대 수준 분류 기준 (cdi_relative_level)
| 구간 | 등급 |
|------|------|
| CDI > P80({CDI_P80:.4f}) | 수요 상위 |
| P20({CDI_P20:.4f}) <= CDI <= P80 | 수요 중위 |
| CDI < P20 | 수요 하위 |
| 결측 | 확인 필요 |

## 4. CSI 상대 수준 분류 기준 (csi_relative_level)
| 구간 | 등급 |
|------|------|
| CSI > P80({CSI_P80:.4f}) | 공급 상위 |
| P20({CSI_P20:.4f}) <= CSI <= P80 | 공급 중위 |
| CSI < P20 | 공급 하위 |
| 결측 | 확인 필요 |

## 5. 3x3 수요-공급 매트릭스 구성 (supply_demand_matrix_3x3)
| CDI 수준 \\ CSI 수준 | 공급 하위 | 공급 중위 | 공급 상위 |
|-----------------|---------|---------|---------|
| **수요 상위** | 핵심 불균형형 | 고수요 보완형 | 고수요 유지관리형 |
| **수요 중위** | 잠재 취약형 | 평균 관리형 | 안정 관리형 |
| **수요 하위** | 최소 인프라 보완형 | 안정 모니터링형 | 여유·거점 활용형 |

### 산출 분포
| 유형 | 학교 수 | 비율 |
|------|---------|------|
{matrix_dist_str}

## 6. policy_strategy_group 생성 규칙 (우선순위 순서)
| 우선순위 | 조건 | 부여 그룹 |
|---------|------|---------|
| 1 | CDI·CSI·priority_score 결측 | 확인 필요형 |
| 2 | priority_level=최우선 지원 또는 3x3=핵심 불균형형 | 최우선 개입형 |
| 3 | priority_level=우선 지원 또는 3x3=고수요 보완형 | 우선 보완형 |
| 4 | 3x3=고수요 유지관리형 | 고수요 유지관리형 |
| 5 | counseling_staff_supply_score < 0.4 | 인력 취약형 |
| 6 | wee_center_access_score <= 0.4 | 접근성 보완형 |
| 7 | 3x3=최소 인프라 보완형 또는 잠재 취약형 | 최소 인프라 보완형 |
| 8 | 그 외 | 안정형 |

### 산출 분포
| 그룹 | 학교 수 | 비율 |
|------|---------|------|
{strategy_dist_str}

## 7. policy_strategy_tags를 별도로 생성한 이유
하나의 학교가 여러 취약 조건에 동시에 해당할 수 있다.
policy_strategy_group은 가장 우선도가 높은 단일 그룹만 부여하므로,
나머지 조건 정보가 소실될 수 있다.
policy_strategy_tags는 해당되는 모든 조건 태그를 세미콜론으로 연결하여
정책 피드백 세분화 및 대시보드 필터링에 활용할 수 있도록 보완 제공한다.

## 8. 분석의 성격
이 분류는 실제 지원 확정이 아니라 **정책 검토를 위한 우선순위 분석 기준**이다.
최종 지원 여부는 현장 방문, 담당 교사 의견, 전문가 자문, 예산 등을 종합하여 결정해야 한다.

## 9. 한계
1. **상대 기준의 가변성**: 분위수 기준(P20, P80)은 현재 146개교의 분포에 기반하며,
   매년 데이터 분포 변화에 따라 기준값이 달라질 수 있다.
2. **동일값 처리**: 동일값이 P20 또는 P80 경계에 걸릴 경우 구간별 학교 수가
   정확히 20/60/20%로 나뉘지 않을 수 있다.
3. **추가 검토 필요**: 실제 정책 적용 시 예산, 인력, 학교 현장 의견, 전문가 자문이
   추가로 필요하며, 정량 지표만으로 정책 결정을 내려서는 안 된다.
"""

MD_PATH = DOCS_DIR / "refined_grouping_methodology_2025.md"
MD_PATH.write_text(methodology_md, encoding="utf-8")
print(f"  저장 완료: {MD_PATH}")

# ════════════════════════════════════════════════════════════════════════════
# 최종 요약 출력
# ════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("정교화 유형화 및 정책전략 그룹 생성 완료")
print("=" * 60)
print(f"\n대상 학교: {n_rows}개교")
print(f"신규 변수: cdi_relative_level, csi_relative_level, supply_demand_matrix_3x3,")
print(f"          policy_strategy_group, policy_strategy_tags, policy_strategy_description")
print(f"\n분위수 기준값:")
print(f"  CDI P20={CDI_P20:.4f}, P80={CDI_P80:.4f}")
print(f"  CSI P20={CSI_P20:.4f}, P80={CSI_P80:.4f}")
print(f"\n생성 파일:")
print(f"  {OUTPUT_PATH}")
print(f"  {TABLES_DIR / 'refined_supply_demand_matrix_summary_2025.csv'}")
print(f"  {TABLES_DIR / 'refined_policy_strategy_summary_2025.csv'}")
print(f"  {TABLES_DIR / 'refined_sigungu_strategy_summary_2025.csv'}")
print(f"  그래프 5종 x 2(영문+국문) = 10개 PNG")
print(f"  {MD_PATH}")
