"""
K-means 클러스터링 기반 보조 유형화 분석 (2025)

목적:
  기존 3×3 수요-공급 매트릭스 및 정책전략 그룹은 유지하되,
  K-means 클러스터링을 보조 분석으로 추가하여 데이터상 유사 집단을 탐색한다.

입력:
  data/processed/gyeongnam_high_schools_policy_feedback_refined.xlsx
  시트: refined_policy_feedback_table

출력:
  data/processed/gyeongnam_high_schools_policy_feedback_kmeans.xlsx
  outputs/tables/kmeans_cluster_summary_2025.csv
  outputs/tables/kmeans_cluster_cross_tab_2025.csv
  outputs/tables/kmeans_cluster_school_list_2025.csv
  outputs/figures/kmeans_elbow_method_2025.png
  outputs/figures/kmeans_silhouette_score_2025.png
  outputs/figures/kmeans_csi_cdi_scatter_2025.png
  docs/kmeans_clustering_methodology_2025.md
"""

import pathlib
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT        = pathlib.Path(__file__).resolve().parent.parent
PROCESSED   = ROOT / "data" / "processed"
TABLES_DIR  = ROOT / "outputs" / "tables"
FIGURES_DIR = ROOT / "outputs" / "figures"
DOCS_DIR    = ROOT / "docs"

for d in [TABLES_DIR, FIGURES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

INPUT_PATH  = PROCESSED / "gyeongnam_high_schools_policy_feedback_refined.xlsx"
OUTPUT_PATH = PROCESSED / "gyeongnam_high_schools_policy_feedback_kmeans.xlsx"
SHEET_NAME  = "refined_policy_feedback_table"

# 한글 폰트 설정 (Windows)
def _set_korean_font():
    candidates = ["Malgun Gothic", "NanumGothic", "AppleGothic", "sans-serif"]
    for name in candidates:
        if any(name.lower() in f.name.lower() for f in fm.fontManager.ttflist):
            plt.rcParams["font.family"] = name
            plt.rcParams["axes.unicode_minus"] = False
            return name
    plt.rcParams["axes.unicode_minus"] = False
    return "default"

FONT = _set_korean_font()

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 입력 파일 불러오기
# ════════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("STEP 1. 입력 파일 불러오기")
print("=" * 60)

if not INPUT_PATH.exists():
    print(f"[ERROR] 입력 파일 없음: {INPUT_PATH}")
    raise SystemExit(1)

xl = pd.ExcelFile(INPUT_PATH)
if SHEET_NAME in xl.sheet_names:
    df = pd.read_excel(INPUT_PATH, sheet_name=SHEET_NAME, dtype={"school_code": str})
    print(f"  시트 사용: {SHEET_NAME}")
else:
    first = xl.sheet_names[0]
    df = pd.read_excel(INPUT_PATH, sheet_name=first, dtype={"school_code": str})
    print(f"  [WARN] '{SHEET_NAME}' 없음 → 첫 번째 시트 사용: {first}")

print(f"  로드 완료: {len(df)}행 × {len(df.columns)}열")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. 필수 변수 확인
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 2. 필수 변수 확인")

REQUIRED = [
    "school_code", "school_name", "sido", "sigungu",
    "counseling_staff_supply_score", "wee_class_score", "wee_center_access_score",
    "demand_size_score", "counseling_use_score", "school_violence_risk_score",
    "CSI", "CDI", "priority_score",
    "cdi_relative_level", "csi_relative_level",
    "supply_demand_matrix_3x3", "policy_strategy_group", "priority_level",
]
missing_cols = [c for c in REQUIRED if c not in df.columns]
if missing_cols:
    print(f"[ERROR] 누락 변수: {missing_cols}")
    raise SystemExit(1)
print(f"  필수 변수 {len(REQUIRED)}개 모두 확인")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. 클러스터링 변수 및 결측 처리
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 3. 클러스터링 변수 선정 및 결측 처리")

CLUSTER_VARS = [
    "counseling_staff_supply_score", "wee_class_score", "wee_center_access_score",
    "demand_size_score", "counseling_use_score", "school_violence_risk_score",
]

df_valid = df.dropna(subset=CLUSTER_VARS).copy()
df_miss  = df[df[CLUSTER_VARS].isna().any(axis=1)].copy()

print(f"  클러스터링 사용 학교: {len(df_valid)}개교")
print(f"  결측으로 제외 학교:   {len(df_miss)}개교")
if len(df_miss) > 0:
    print(f"  제외 학교 목록: {df_miss['school_name'].tolist()}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. 표준화
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 4. StandardScaler 표준화")

scaler  = StandardScaler()
X_scaled = scaler.fit_transform(df_valid[CLUSTER_VARS])
print(f"  표준화 완료: {X_scaled.shape}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. 적정 군집 수 탐색 (k=2~6)
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 5. 적정 군집 수 탐색 (k=2~6)")

k_range    = range(2, 7)
inertias   = []
silhouettes = []

for k in k_range:
    km = KMeans(n_clusters=k, random_state=42, n_init=20)
    labels = km.fit_predict(X_scaled)
    inertias.append(km.inertia_)
    sil = silhouette_score(X_scaled, labels)
    silhouettes.append(sil)
    print(f"  k={k}: inertia={km.inertia_:.2f}, silhouette={sil:.4f}")

# Elbow 그래프
fig, ax = plt.subplots(figsize=(7, 4))
ax.plot(list(k_range), inertias, marker="o", color="#2E5FA3", linewidth=2)
ax.set_xlabel("군집 수 (k)"); ax.set_ylabel("Inertia")
ax.set_title("K-means Elbow Method"); ax.grid(True, alpha=0.3)
plt.tight_layout()
elbow_path = FIGURES_DIR / "kmeans_elbow_method_2025.png"
fig.savefig(elbow_path, dpi=150, bbox_inches="tight")
plt.close()
print(f"  저장: {elbow_path.name}")

# Silhouette 그래프
fig, ax = plt.subplots(figsize=(7, 4))
ax.plot(list(k_range), silhouettes, marker="s", color="#C0392B", linewidth=2)
ax.set_xlabel("군집 수 (k)"); ax.set_ylabel("Silhouette Score")
ax.set_title("K-means Silhouette Score"); ax.grid(True, alpha=0.3)
plt.tight_layout()
sil_path = FIGURES_DIR / "kmeans_silhouette_score_2025.png"
fig.savefig(sil_path, dpi=150, bbox_inches="tight")
plt.close()
print(f"  저장: {sil_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. K-means 실행 (selected_k=4)
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 6. K-means 실행 (k=4)")

selected_k = 4
km_final   = KMeans(n_clusters=selected_k, random_state=42, n_init=20)
df_valid["kmeans_cluster"] = km_final.fit_predict(X_scaled)
print(f"  군집 분포: {df_valid['kmeans_cluster'].value_counts().sort_index().to_dict()}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 군집 라벨 자동 부여
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 7. 군집 라벨 자동 부여")

cluster_means = df_valid.groupby("kmeans_cluster")[["CSI","CDI","priority_score"]].mean()
print(cluster_means.round(3))

assigned = {}
remaining = set(cluster_means.index)

# 라벨 1: PS 최고 + CSI 최저 → 수요 대비 공급 취약
c_vuln = cluster_means["priority_score"].idxmax()
if cluster_means.loc[c_vuln, "CSI"] < cluster_means["CSI"].median():
    assigned[c_vuln] = "수요 대비 공급 취약 군집"
    remaining.discard(c_vuln)

# 라벨 2: PS 최저 + CSI 최고 → 상대적 안정
rem_df = cluster_means.loc[list(remaining)]
if not rem_df.empty:
    c_stable = rem_df["priority_score"].idxmin()
    if cluster_means.loc[c_stable, "CSI"] >= cluster_means["CSI"].median():
        assigned[c_stable] = "상대적 안정 군집"
        remaining.discard(c_stable)

# 라벨 3: CDI 높고 CSI도 높음 → 고수요 유지관리
rem_df = cluster_means.loc[list(remaining)]
if len(rem_df) >= 2:
    c_high = rem_df["CDI"].idxmax()
    if cluster_means.loc[c_high, "CSI"] >= rem_df["CSI"].median():
        assigned[c_high] = "고수요 유지관리 군집"
        remaining.discard(c_high)

# 라벨 4: 나머지 → 기초 인프라 보완
for c in remaining:
    assigned[c] = "기초 인프라 보완 군집"

# 미배정 처리
for c in cluster_means.index:
    if c not in assigned:
        assigned[c] = f"Cluster {c}"

label_map = assigned
df_valid["kmeans_cluster_label"] = df_valid["kmeans_cluster"].map(label_map)

for c, lbl in sorted(label_map.items()):
    n = (df_valid["kmeans_cluster"] == c).sum()
    print(f"  Cluster {c} → {lbl} ({n}개교)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. 결측 학교에 "확인 필요" 표시 후 전체 병합
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 8. 전체 데이터 병합")

df_miss = df_miss.copy()
df_miss["kmeans_cluster"]       = "확인 필요"
df_miss["kmeans_cluster_label"] = "확인 필요"

df_all = pd.concat([df_valid, df_miss], ignore_index=True)
df_all = df_all.sort_values("school_code").reset_index(drop=True)
print(f"  전체 {len(df_all)}개교 병합 완료")

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. 군집별 요약표
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 9. 군집별 요약표 생성")

POLICY_INTERPRET = {
    "수요 대비 공급 취약 군집": "상담수요 대비 공급 부족 가능성이 높아 인력 배치, Wee클래스·Wee센터 연계 강화 검토가 필요한 군집",
    "상대적 안정 군집":         "현재 지표상 수요 대비 공급이 비교적 안정적으로, 유지·모니터링 중심의 관리가 적절한 군집",
    "고수요 유지관리 군집":     "수요와 공급이 모두 높아 기존 인프라 유지와 프로그램 질 관리가 필요한 군집",
    "기초 인프라 보완 군집":    "수요와 공급 모두 낮아 기본 상담 인프라 구축 및 보완이 필요한 군집",
}

df_valid_only = df_all[df_all["kmeans_cluster"] != "확인 필요"].copy()
df_valid_only["kmeans_cluster"] = df_valid_only["kmeans_cluster"].astype(int)

agg_cols = ["CSI","CDI","priority_score",
            "counseling_staff_supply_score","wee_class_score","wee_center_access_score",
            "demand_size_score","counseling_use_score","school_violence_risk_score"]

summary_rows = []
for c in sorted(df_valid_only["kmeans_cluster"].unique()):
    sub  = df_valid_only[df_valid_only["kmeans_cluster"] == c]
    lbl  = label_map[c]
    row  = {"군집번호": c, "군집 라벨": lbl, "학교수": len(sub)}
    for col in agg_cols:
        row[f"평균_{col}"] = round(sub[col].mean(), 3)
    row["최우선_지원_학교수"] = int((sub["priority_level"] == "최우선 지원").sum())
    row["우선_지원_학교수"]   = int((sub["priority_level"] == "우선 지원").sum())
    row["대표_정책_해석"]     = POLICY_INTERPRET.get(lbl, "-")
    summary_rows.append(row)

summary_df = pd.DataFrame(summary_rows)
summary_path = TABLES_DIR / "kmeans_cluster_summary_2025.csv"
summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")
print(f"  저장: {summary_path.name}")
print(summary_df[["군집번호","군집 라벨","학교수","평균_CSI","평균_CDI","평균_priority_score"]].to_string(index=False))

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. 기존 유형화 × K-means 교차표
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 10. 기존 유형화 × K-means 비교표 생성")

cross_sections = []

for cross_col in ["supply_demand_matrix_3x3", "policy_strategy_group", "priority_level"]:
    ct = pd.crosstab(
        df_valid_only["kmeans_cluster_label"],
        df_valid_only[cross_col],
        margins=True, margins_name="합계"
    )
    # 구분 헤더 추가
    header = pd.DataFrame(
        [[f"=== {cross_col} 교차표 ==="] + [""] * (len(ct.columns)-1)],
        columns=ct.columns
    )
    cross_sections.append(header)
    cross_sections.append(ct.reset_index())
    cross_sections.append(pd.DataFrame([[""] * len(ct.columns)], columns=ct.columns))
    print(f"\n  × {cross_col}")
    print(ct)

cross_path = TABLES_DIR / "kmeans_cluster_cross_tab_2025.csv"
pd.concat(cross_sections, ignore_index=True).to_csv(
    cross_path, index=False, encoding="utf-8-sig"
)
print(f"\n  저장: {cross_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 11. 학교별 결과 목록
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 11. 학교별 K-means 결과 목록 생성")

list_cols = [
    "school_code","school_name","sido","sigungu",
    "CSI","CDI","priority_score","priority_level",
    "supply_demand_matrix_3x3","policy_strategy_group",
    "kmeans_cluster","kmeans_cluster_label",
]
school_list = df_all[[c for c in list_cols if c in df_all.columns]].copy()
school_list = school_list.sort_values(
    ["kmeans_cluster_label","priority_score"],
    ascending=[True, False]
).reset_index(drop=True)

list_path = TABLES_DIR / "kmeans_cluster_school_list_2025.csv"
school_list.to_csv(list_path, index=False, encoding="utf-8-sig")
print(f"  저장: {list_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 12. CSI-CDI 산점도 (군집 색상)
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 12. CSI-CDI 산점도 생성")

CLUSTER_COLORS = {
    "수요 대비 공급 취약 군집": "#C0392B",
    "고수요 유지관리 군집":     "#E67E22",
    "기초 인프라 보완 군집":    "#9B59B6",
    "상대적 안정 군집":         "#27AE60",
}

fig, ax = plt.subplots(figsize=(9, 7))

for lbl, grp in df_valid_only.groupby("kmeans_cluster_label"):
    color = CLUSTER_COLORS.get(lbl, "#BDC3C7")
    ax.scatter(grp["CSI"], grp["CDI"], c=color, label=lbl,
               alpha=0.8, s=60, edgecolors="white", linewidths=0.5)

avg_csi = df_valid_only["CSI"].mean()
avg_cdi = df_valid_only["CDI"].mean()
ax.axvline(avg_csi, color="gray", linestyle="--", linewidth=1, alpha=0.7,
           label=f"평균 CSI ({avg_csi:.3f})")
ax.axhline(avg_cdi, color="gray", linestyle=":",  linewidth=1, alpha=0.7,
           label=f"평균 CDI ({avg_cdi:.3f})")

ax.set_xlabel("CSI (상담공급지수)", fontsize=11)
ax.set_ylabel("CDI (상담수요지수)", fontsize=11)
ax.set_title("K-means 기반 CSI-CDI 학교 군집 분포", fontsize=13, fontweight="bold")
ax.set_xlim(-0.05, 1.05); ax.set_ylim(-0.05, 1.05)
ax.legend(loc="upper left", fontsize=9, framealpha=0.9)
ax.grid(True, alpha=0.2)
plt.tight_layout()

scatter_path = FIGURES_DIR / "kmeans_csi_cdi_scatter_2025.png"
fig.savefig(scatter_path, dpi=150, bbox_inches="tight")
plt.close()
print(f"  저장: {scatter_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 13. 결과 엑셀 저장
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 13. 결과 엑셀 저장")

# 결측 학교 점검표
missing_check = df_miss[["school_code","school_name","sigungu"]].copy() if len(df_miss) > 0 else pd.DataFrame(columns=["school_code","school_name","sigungu"])
missing_check["비고"] = "K-means 학습 제외 (결측)"

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    df_all.to_excel(writer, sheet_name="kmeans_school_table", index=False)
    summary_df.to_excel(writer, sheet_name="kmeans_cluster_summary", index=False)
    school_list.to_excel(writer, sheet_name="kmeans_cluster_school_list", index=False)
    missing_check.to_excel(writer, sheet_name="kmeans_missing_check", index=False)

print(f"  저장: {OUTPUT_PATH.name}")
print(f"  시트: kmeans_school_table, kmeans_cluster_summary, kmeans_cluster_school_list, kmeans_missing_check")

# ════════════════════════════════════════════════════════════════════════════
# STEP 14. 방법론 문서 생성
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 14. 방법론 문서 생성")

# k별 지표 정리
k_table = "\n".join([
    f"| {k} | {inertia:.2f} | {sil:.4f} |"
    for k, inertia, sil in zip(k_range, inertias, silhouettes)
])

label_table = "\n".join([
    f"| Cluster {c} | {lbl} |"
    for c, lbl in sorted(label_map.items())
])

cluster_desc = "\n".join([
    f"- **{lbl}** ({(df_valid_only['kmeans_cluster_label']==lbl).sum()}개교): "
    f"평균 CSI={cluster_means.loc[[k for k,v in label_map.items() if v==lbl][0],'CSI']:.3f}, "
    f"CDI={cluster_means.loc[[k for k,v in label_map.items() if v==lbl][0],'CDI']:.3f}, "
    f"PS={cluster_means.loc[[k for k,v in label_map.items() if v==lbl][0],'priority_score']:.3f}"
    for lbl in [label_map[c] for c in sorted(label_map)]
    if lbl != "확인 필요"
])

md_content = f"""# K-means 클러스터링 기반 보조 유형화 분석 방법론

## 1. 분석 목적

본 분석은 기존 3×3 수요-공급 매트릭스 및 정책전략 그룹을 **유지**하면서,
K-means 클러스터링을 보조적으로 적용하여 학교 간 지표 구조의 유사성을 탐색한다.
기존 기준 기반 유형화가 데이터상 유사 집단과 어느 정도 일치하는지 확인하고,
정책전략 그룹의 해석력을 보완하는 데 목적이 있다.

> **핵심 문장**: 본 연구는 기존 3×3 수요-공급 매트릭스를 주 분석 기준으로 유지하되,
> K-means 클러스터링을 보조적으로 적용하여 학교 간 지표 구조의 유사성을 탐색하였다.
> 이를 통해 기준 기반 유형화가 데이터상 유사 집단과 어느 정도 일치하는지 확인하고,
> 정책전략 그룹의 해석력을 보완하였다.

## 2. 기존 3×3 유형화와 K-means의 관계

| 구분 | 기존 3×3 유형화 | K-means 클러스터링 |
|---|---|---|
| 분류 기준 | 사전 정의된 분위수 기준 | 데이터 거리 기반 자동 탐색 |
| 역할 | 주 분석 기준 (유지) | 보조 분석 (타당성 점검) |
| 결과 | supply_demand_matrix_3x3 | kmeans_cluster_label |

K-means 결과는 기존 유형화를 **대체하지 않는다.**

## 3. 사용 변수 (9개)

CSI 구성:
- counseling_staff_supply_score
- wee_class_score
- wee_center_access_score

CDI 구성:
- demand_size_score
- counseling_use_score
- school_violence_risk_score

종합 지수:
- CSI, CDI, priority_score

## 4. 결측 처리

- 9개 변수 중 결측 있는 학교: K-means 학습에서 제외
- 제외 학교: {len(df_miss)}개교
- 결과 파일에 kmeans_cluster = "확인 필요"로 표기

## 5. 표준화

StandardScaler 적용 (평균 0, 표준편차 1)
해석은 원래 점수 기준 평균값으로 실시

## 6. k 탐색 결과 (k=2~6)

| k | Inertia | Silhouette Score |
|---|---|---|
{k_table}

## 7. 최종 k=4 선택 이유

- 수요 높음/낮음 × 공급 높음/낮음 구조와 해석적으로 연결 용이
- 기존 정책전략 그룹(5개)과 비교 가능한 적정 수준
- Silhouette Score 결과와 해석 가능성을 종합 판단

## 8. 군집 라벨 해석

| 군집 | 라벨 |
|---|---|
{label_table}

{cluster_desc}

## 9. 한계

- 표본 수가 146개로 크지 않아 군집 안정성이 낮을 수 있음
- 일부 하위 점수(wee_class_score 등)가 이진화·구간화되어 군집 경계가 불명확할 수 있음
- CSI·CDI·priority_score는 하위 점수로부터 계산된 파생변수이므로 변수 중복성 존재
- K-means는 거리 기반 알고리즘이므로 변수 선택·표준화 방식에 결과가 민감하게 반응
- 군집 라벨은 사후 해석이며 실제 정책 지원 확정 기준이 아님

## 10. 실행 정보

- 분석 일시: {pd.Timestamp.now().strftime('%Y-%m-%d')}
- 사용 데이터: gyeongnam_high_schools_policy_feedback_refined.xlsx
- random_state=42, n_init=20, selected_k=4
"""

md_path = DOCS_DIR / "kmeans_clustering_methodology_2025.md"
md_path.write_text(md_content, encoding="utf-8")
print(f"  저장: {md_path.name}")

# ════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("K-means 클러스터링 분석 완료")
print("=" * 60)
print(f"  결과 엑셀:  {OUTPUT_PATH.name}")
print(f"  요약표:     kmeans_cluster_summary_2025.csv")
print(f"  비교표:     kmeans_cluster_cross_tab_2025.csv")
print(f"  학교 목록:  kmeans_cluster_school_list_2025.csv")
print(f"  그래프:     elbow / silhouette / scatter")
print(f"  문서:       kmeans_clustering_methodology_2025.md")
print()
print("다음 단계:")
print("  streamlit run app.py")
