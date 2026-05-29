"""
학교별 정책 추천 적합도 점수 산출 (2025)

목적:
  각 학교별 상담지원 관련 6개 정책 대안에 대한 적합도 점수를 산출하고
  1~3순위 추천 정책과 근거 문장을 자동 생성한다.

입력 우선순위:
  1. data/processed/gyeongnam_high_schools_policy_feedback_kmeans.xlsx
  2. data/processed/gyeongnam_high_schools_policy_feedback_refined.xlsx

출력:
  data/processed/gyeongnam_high_schools_policy_recommendation_scores.xlsx
  outputs/tables/policy_fit_score_summary_2025.csv
  outputs/tables/policy_top_recommendations_2025.csv
  docs/policy_fit_score_methodology_2025.md
"""

import pathlib
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT       = pathlib.Path(__file__).resolve().parent.parent
PROCESSED  = ROOT / "data" / "processed"
TABLES_DIR = ROOT / "outputs" / "tables"
DOCS_DIR   = ROOT / "docs"

for d in [TABLES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

KMEANS_PATH  = PROCESSED / "gyeongnam_high_schools_policy_feedback_kmeans.xlsx"
REFINED_PATH = PROCESSED / "gyeongnam_high_schools_policy_feedback_refined.xlsx"
OUTPUT_PATH  = PROCESSED / "gyeongnam_high_schools_policy_recommendation_scores.xlsx"

# 정책 정의
POLICIES = {
    "fit_counselor_assignment":    "전문상담교사 배치 또는 순회상담 연계",
    "fit_wee_class":               "Wee클래스 신설 또는 운영 보완",
    "fit_wee_center_linkage":      "Wee센터 연계 강화",
    "fit_school_violence_support": "학교폭력 피해 관련 상담지원 강화",
    "fit_high_demand_program":     "고수요 상담 프로그램 확대",
    "fit_monitoring":              "현 수준 유지 및 정기 모니터링",
}
POLICY_COLS = list(POLICIES.keys())
POLICY_NAMES = list(POLICIES.values())

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 입력 파일 로드
# ════════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("STEP 1. 입력 파일 불러오기")
print("=" * 60)

def _load_sheet(path: pathlib.Path, preferred_sheet: str) -> tuple[pd.DataFrame, str, str]:
    xl = pd.ExcelFile(path)
    if preferred_sheet in xl.sheet_names:
        sheet = preferred_sheet
    else:
        sheet = xl.sheet_names[0]
        print(f"  [WARN] '{preferred_sheet}' 없음 → '{sheet}' 사용")
    df = pd.read_excel(path, sheet_name=sheet, dtype={"school_code": str})
    return df, path.name, sheet

if KMEANS_PATH.exists():
    df, src_file, src_sheet = _load_sheet(KMEANS_PATH, "kmeans_school_table")
else:
    print(f"  [INFO] kmeans 파일 없음 → refined 파일 사용")
    df, src_file, src_sheet = _load_sheet(REFINED_PATH, "refined_policy_feedback_table")

print(f"  파일: {src_file} / 시트: {src_sheet}")
print(f"  로드 완료: {len(df)}행 × {len(df.columns)}열")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. 필수 변수 확인
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 2. 필수 변수 확인")

REQUIRED = [
    "school_code", "school_name", "sido", "sigungu",
    "counseling_staff_supply_score", "wee_class_score", "wee_center_access_score",
    "demand_size_score", "counseling_use_score", "school_violence_risk_score",
    "CSI", "CDI", "priority_score", "priority_level",
    "supply_demand_matrix_3x3", "policy_strategy_group",
]
OPTIONAL = ["kmeans_cluster", "kmeans_cluster_label", "policy_recommendation", "policy_reason"]

missing = [c for c in REQUIRED if c not in df.columns]
if missing:
    print(f"[ERROR] 필수 변수 누락: {missing}")
    raise SystemExit(1)

optional_present = [c for c in OPTIONAL if c in df.columns]
print(f"  필수 변수 {len(REQUIRED)}개 확인 완료")
print(f"  선택 변수 존재: {optional_present}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. priority_need_score 생성 (Min-Max 정규화)
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 3. priority_need_score 생성")

ps = df["priority_score"]
ps_min, ps_max = ps.min(), ps.max()
df["priority_need_score"] = (ps - ps_min) / (ps_max - ps_min) if ps_max != ps_min else 0.5
print(f"  priority_score 범위: {ps_min:.3f} ~ {ps_max:.3f}")
print(f"  priority_need_score 범위: {df['priority_need_score'].min():.3f} ~ {df['priority_need_score'].max():.3f}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. 정책별 적합도 점수 산출
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 4. 정책별 적합도 점수 산출")

def _safe(df: pd.DataFrame, col: str) -> pd.Series:
    """컬럼이 있으면 반환, 없으면 NaN 시리즈 반환."""
    return df[col] if col in df.columns else pd.Series(np.nan, index=df.index)

# 하위 변수 추출
staff  = _safe(df, "counseling_staff_supply_score")
wee_c  = _safe(df, "wee_class_score")
wee_ct = _safe(df, "wee_center_access_score")
dem    = _safe(df, "demand_size_score")
use    = _safe(df, "counseling_use_score")
viol   = _safe(df, "school_violence_risk_score")
csi    = _safe(df, "CSI")
need   = df["priority_need_score"]

missing_notes = []

# 정책 A: 전문상담교사 배치
df["fit_counselor_assignment"] = (
    0.50 * (1 - staff) + 0.25 * dem + 0.25 * need
).clip(0, 1).round(3)

# 정책 B: Wee클래스 신설
df["fit_wee_class"] = (
    0.60 * (1 - wee_c) + 0.20 * need + 0.20 * use
).clip(0, 1).round(3)

# 정책 C: Wee센터 연계
df["fit_wee_center_linkage"] = (
    0.60 * (1 - wee_ct) + 0.20 * need + 0.20 * use
).clip(0, 1).round(3)

# 정책 D: 학교폭력 상담 강화
df["fit_school_violence_support"] = (
    0.60 * viol + 0.20 * use + 0.20 * need
).clip(0, 1).round(3)

# 정책 E: 고수요 프로그램
df["fit_high_demand_program"] = (
    0.40 * dem + 0.40 * use + 0.20 * need
).clip(0, 1).round(3)

# 정책 F: 현 수준 모니터링
df["fit_monitoring"] = (
    0.50 * csi + 0.30 * (1 - need) + 0.20 * (1 - viol)
).clip(0, 1).round(3)

for col in POLICY_COLS:
    n_null = df[col].isna().sum()
    print(f"  {col}: 평균={df[col].mean():.3f}, 결측={n_null}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. 추천 정책 1~3순위 생성
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 5. 추천 정책 1~3순위 생성")

def _rank_policies(row: pd.Series) -> list[tuple[str, float]]:
    """6개 점수 내림차순 정렬 → 동점 시 정책 정의 순서 우선."""
    scored = [(POLICIES[col], row[col]) for col in POLICY_COLS if pd.notna(row[col])]
    scored.sort(key=lambda x: (-x[1], POLICY_NAMES.index(x[0])))
    return scored

def _make_reason(row: pd.Series, top_policy: str, top_score: float) -> str:
    """추천 근거 문장 생성."""
    school = row.get("school_name", "해당 학교")
    base = f"해당 학교는 현재 지표 체계상 '{top_policy}'의 적합도가 가장 높게 산출되었다."

    detail = ""
    if top_policy == "전문상담교사 배치 또는 순회상담 연계":
        detail = (f"상담인력 공급 점수({row.get('counseling_staff_supply_score', 'N/A'):.3f})가 낮고 "
                  f"상담수요 규모가 상대적으로 높아 전문상담인력 확충 검토가 필요하다.")
    elif top_policy == "Wee클래스 신설 또는 운영 보완":
        detail = (f"Wee클래스 운영 점수({row.get('wee_class_score', 'N/A'):.3f})가 낮아 "
                  f"상담공간 및 기본 상담체계 보완이 필요한 것으로 나타났다.")
    elif top_policy == "Wee센터 연계 강화":
        detail = (f"Wee센터 접근성 점수({row.get('wee_center_access_score', 'N/A'):.3f})가 낮아 "
                  f"이동형·온라인 상담 또는 권역별 연계 강화 검토가 필요하다.")
    elif top_policy == "학교폭력 피해 관련 상담지원 강화":
        detail = (f"학교폭력 위험 점수({row.get('school_violence_risk_score', 'N/A'):.3f})가 높아 "
                  f"피해 학생 대상 상담지원 강화 검토가 우선적으로 필요하다.")
    elif top_policy == "고수요 상담 프로그램 확대":
        detail = (f"상담수요 규모와 실제 상담 이용 수준이 모두 높아 "
                  f"상담 프로그램 다양화 및 확대 검토가 적절한 것으로 나타났다.")
    elif top_policy == "현 수준 유지 및 정기 모니터링":
        detail = (f"CSI({row.get('CSI', 'N/A'):.3f})가 높고 우선지원 필요성이 상대적으로 낮아 "
                  f"현재 인프라 유지와 정기 모니터링이 적절한 것으로 나타났다.")

    suffix = "이는 실제 지원 확정이 아니라 정책 검토 우선순위 판단을 위한 추천 결과이다."
    return f"{base} {detail} {suffix}"

recs_1, scores_1 = [], []
recs_2, scores_2 = [], []
recs_3, scores_3 = [], []
reasons          = []
missing_flags    = []

for _, row in df.iterrows():
    ranked = _rank_policies(row)
    # 1순위
    p1, s1 = (ranked[0][0], ranked[0][1]) if len(ranked) >= 1 else ("확인 필요", np.nan)
    p2, s2 = (ranked[1][0], ranked[1][1]) if len(ranked) >= 2 else ("확인 필요", np.nan)
    p3, s3 = (ranked[2][0], ranked[2][1]) if len(ranked) >= 3 else ("확인 필요", np.nan)
    recs_1.append(p1); scores_1.append(round(s1, 3) if not np.isnan(s1) else np.nan)
    recs_2.append(p2); scores_2.append(round(s2, 3) if not np.isnan(s2) else np.nan)
    recs_3.append(p3); scores_3.append(round(s3, 3) if not np.isnan(s3) else np.nan)

    # 추천 근거
    reasons.append(_make_reason(row, p1, s1) if p1 != "확인 필요" else "지표 결측으로 추천 근거 생성 불가")

    # 확인 필요 플래그
    flag = ""
    if p1 == "현 수준 유지 및 정기 모니터링" and row.get("priority_level") in ["최우선 지원", "우선 지원"]:
        flag = "확인 필요: 우선지원등급 학교에 모니터링 1순위"
    missing_flags.append(flag)

df["recommended_policy_1"]       = recs_1
df["recommended_policy_1_score"] = scores_1
df["recommended_policy_2"]       = recs_2
df["recommended_policy_2_score"] = scores_2
df["recommended_policy_3"]       = recs_3
df["recommended_policy_3_score"] = scores_3
df["recommended_policy_reason"]  = reasons
df["policy_fit_missing_note"]    = missing_flags

print(f"  1순위 분포:\n{pd.Series(recs_1).value_counts().to_string()}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. 정책별 요약표
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 6. 정책별 요약표 생성")

summary_rows = []
for col, name in POLICIES.items():
    scores = df[col].dropna()
    n_rank1 = (df["recommended_policy_1"] == name).sum()
    n_rank2 = ((df["recommended_policy_1"] == name) | (df["recommended_policy_2"] == name)).sum()
    n_rank3 = ((df["recommended_policy_1"] == name) | (df["recommended_policy_2"] == name) |
               (df["recommended_policy_3"] == name)).sum()
    summary_rows.append({
        "정책명":           name,
        "평균_적합도":      round(scores.mean(), 3),
        "최고_적합도":      round(scores.max(), 3),
        "최저_적합도":      round(scores.min(), 3),
        "1순위_학교수":     int(n_rank1),
        "2순위이내_학교수": int(n_rank2),
        "3순위이내_학교수": int(n_rank3),
    })

summary_df = pd.DataFrame(summary_rows)
summary_path = TABLES_DIR / "policy_fit_score_summary_2025.csv"
summary_df.to_csv(summary_path, index=False, encoding="utf-8-sig")
print(f"  저장: {summary_path.name}")
print(summary_df[["정책명","평균_적합도","1순위_학교수"]].to_string(index=False))

# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 학교별 추천 결과표
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 7. 학교별 추천 결과표 생성")

list_cols = [
    "school_code","school_name","sido","sigungu",
    "CSI","CDI","priority_score","priority_level",
    "policy_strategy_group","supply_demand_matrix_3x3",
    "kmeans_cluster_label",
    "recommended_policy_1","recommended_policy_1_score",
    "recommended_policy_2","recommended_policy_2_score",
    "recommended_policy_3","recommended_policy_3_score",
    "recommended_policy_reason",
]
avail = [c for c in list_cols if c in df.columns]
top_recs = (df[avail]
            .sort_values(["priority_score","recommended_policy_1_score"],
                         ascending=[False, False])
            .reset_index(drop=True))

rec_path = TABLES_DIR / "policy_top_recommendations_2025.csv"
top_recs.to_csv(rec_path, index=False, encoding="utf-8-sig")
print(f"  저장: {rec_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. 정책별 상위 10개교 (엑셀 시트용)
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 8. 정책별 상위 10개교 생성")

top_by_policy_rows = []
for col, name in POLICIES.items():
    top10 = df.nlargest(10, col)[
        ["school_name","sigungu", col,
         "CSI","CDI","priority_score","priority_level","policy_strategy_group"]
    ].copy()
    top10.insert(0, "정책명", name)
    top10.insert(1, "순위", range(1, len(top10)+1))
    top10 = top10.rename(columns={col: "적합도_점수"})
    top_by_policy_rows.append(top10)

top_by_policy_df = pd.concat(top_by_policy_rows, ignore_index=True)

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. 결과 엑셀 저장
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 9. 결과 엑셀 저장")

# 결측 점검표
missing_check = df[df["policy_fit_missing_note"] != ""][
    ["school_code","school_name","sigungu","policy_fit_missing_note"]
].copy() if (df["policy_fit_missing_note"] != "").any() else pd.DataFrame(
    columns=["school_code","school_name","sigungu","policy_fit_missing_note"]
)

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="policy_recommendation_table", index=False)
    summary_df.to_excel(writer, sheet_name="policy_fit_score_summary", index=False)
    top_by_policy_df.to_excel(writer, sheet_name="top_recommendations_by_policy", index=False)
    missing_check.to_excel(writer, sheet_name="policy_fit_missing_check", index=False)

print(f"  저장: {OUTPUT_PATH.name}")
print(f"  시트: policy_recommendation_table, policy_fit_score_summary, "
      f"top_recommendations_by_policy, policy_fit_missing_check")

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. 방법론 문서
# ════════════════════════════════════════════════════════════════════════════
print("\nSTEP 10. 방법론 문서 생성")

formula_table = "\n".join([
    f"| {name} | {col} | — |"
    for col, name in POLICIES.items()
])

md = f"""# 학교별 정책 추천 적합도 점수 산출 방법론

## 1. 목적

기존 규칙 기반 정책 피드백 문장을 **정량적으로 보완**하기 위해, 각 학교의 하위 지표와
우선지원점수를 활용하여 6개 정책 대안별 적합도 점수(0~1)를 산출한다.

> **보고서용 핵심 문장**: 본 연구는 학교별 상담지원 정책 대안을 보다 구체적으로
> 제시하기 위해 정책별 적합도 점수를 산출하였다. 전문상담교사 배치, Wee클래스 신설,
> Wee센터 연계 강화, 학교폭력 상담지원 강화 등 주요 정책 대안에 대해 학교별 하위
> 지표와 우선지원점수를 결합하여 추천 우선순위를 계산하였다.

## 2. 기존 정책 피드백과의 관계

| 구분 | 기존 정책 피드백 | 정책 적합도 점수 |
|---|---|---|
| 생성 방식 | 규칙 기반 텍스트 생성 | 하위 지표 가중합 산출 |
| 출력 | 문장형 피드백 | 0~1 점수 + 순위 |
| 역할 | 주 정책 피드백 (유지) | 보완 분석 |

정책 적합도 점수는 기존 피드백을 **대체하지 않으며 보완**한다.

## 3. 정책 대안 및 변수

| 정책명 | 변수명 | 핵심 지표 |
|---|---|---|
{formula_table}

## 4. priority_need_score 산출

priority_score를 Min-Max 정규화:
- priority_need_score = (PS - PS_min) / (PS_max - PS_min)
- PS 범위: {ps_min:.3f} ~ {ps_max:.3f}

## 5. 정책별 산식

**A. 전문상담교사 배치**
fit_counselor_assignment = 0.50×(1-상담인력) + 0.25×수요규모 + 0.25×need

**B. Wee클래스 신설**
fit_wee_class = 0.60×(1-Wee클래스) + 0.20×need + 0.20×이용률

**C. Wee센터 연계**
fit_wee_center_linkage = 0.60×(1-Wee접근성) + 0.20×need + 0.20×이용률

**D. 학교폭력 상담 강화**
fit_school_violence_support = 0.60×학폭위험 + 0.20×이용률 + 0.20×need

**E. 고수요 프로그램**
fit_high_demand_program = 0.40×수요규모 + 0.40×이용률 + 0.20×need

**F. 현 수준 모니터링**
fit_monitoring = 0.50×CSI + 0.30×(1-need) + 0.20×(1-학폭위험)

## 6. 추천 순위 생성

- 6개 점수 내림차순 정렬 → 1~3순위 배정
- 동점: A→B→C→D→E→F 순서 우선
- F가 1순위 + 최우선/우선 지원 등급 → "확인 필요" 플래그

## 7. 한계

- 가중치는 분석 목적에 맞춘 운영 기준으로 전문가 자문을 거치지 않음
- 일부 하위 점수가 이진화·구간화되어 점수 변화가 제한적으로 표현될 수 있음
- 실제 예산, 인력 수급, 현장 의견은 반영되지 않음
- 추천 결과는 실제 지원 확정이 아니라 정책 검토 우선순위임

## 8. 실행 정보

- 분석 일시: {pd.Timestamp.now().strftime('%Y-%m-%d')}
- 입력 파일: {src_file} / {src_sheet}
- 대상 학교: {len(df)}개교
"""

md_path = DOCS_DIR / "policy_fit_score_methodology_2025.md"
md_path.write_text(md, encoding="utf-8")
print(f"  저장: {md_path.name}")

print("\n" + "=" * 60)
print("정책 적합도 점수 산출 완료")
print("=" * 60)
print(f"  결과 엑셀: {OUTPUT_PATH.name}")
print(f"  요약표:   policy_fit_score_summary_2025.csv")
print(f"  추천표:   policy_top_recommendations_2025.csv")
print(f"  문서:     policy_fit_score_methodology_2025.md")
print()
print("다음 단계: streamlit run app.py")
