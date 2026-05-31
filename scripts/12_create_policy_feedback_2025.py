"""
CSI-CDI 기반 학교 유형화 및 정책 피드백 로직 생성

입력 파일:
  data/processed/gyeongnam_high_schools_priority_summary.xlsx (priority_summary_table 시트)

출력 파일:
  data/processed/gyeongnam_high_schools_policy_feedback.xlsx
  outputs/tables/supply_demand_type_summary_2025.csv
  outputs/tables/policy_feedback_summary_2025.csv
  outputs/tables/sigungu_priority_summary_2025.csv
  docs/policy_feedback_logic_2025.md
"""

import pathlib
import sys

import numpy as np
import pandas as pd

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT          = pathlib.Path(__file__).resolve().parent.parent
PROCESSED_DIR = ROOT / "data" / "processed"
TABLES_DIR    = ROOT / "outputs" / "tables"
DOCS_DIR      = ROOT / "docs"

for d in [PROCESSED_DIR, TABLES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

INPUT_PATH   = PROCESSED_DIR / "gyeongnam_high_schools_priority_summary.xlsx"
OUTPUT_PATH  = PROCESSED_DIR / "gyeongnam_high_schools_policy_feedback.xlsx"

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 입력 파일 로드
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 입력 파일 로드")

if not INPUT_PATH.exists():
    print(f"[ERROR] 파일 없음: {INPUT_PATH}")
    sys.exit(1)

xl = pd.ExcelFile(INPUT_PATH)
if "priority_summary_table" in xl.sheet_names:
    sheet_used = "priority_summary_table"
else:
    sheet_used = xl.sheet_names[0]
    print(f"  [WARN] 'priority_summary_table' 없음 - '{sheet_used}' 시트 사용")

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
]
missing_vars = [c for c in REQUIRED if c not in df.columns]
if missing_vars:
    print(f"  [ERROR] 누락 변수: {missing_vars}")
    sys.exit(1)

print("  필수 변수 전체 확인 완료")
for col in ["CSI", "CDI", "priority_score"]:
    print(f"    {col}: 비결측={df[col].notna().sum()}, 평균={df[col].mean():.4f}")

# 분위수 기준값 산출 (policy_recommendation 조건용)
use_p80 = df["counseling_use_score"].quantile(0.80)
viol_p80 = df["school_violence_risk_score"].quantile(0.80)
cdi_mean = df["CDI"].mean()
csi_mean = df["CSI"].mean()
print(f"\n  2x2 유형화 기준 - CDI 평균: {cdi_mean:.4f}, CSI 평균: {csi_mean:.4f}")
print(f"  counseling_use_score P80: {use_p80:.4f}")
print(f"  school_violence_risk_score P80: {viol_p80:.4f}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. CSI 수준 분류
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 3] CSI 수준 분류")

def assign_csi_level(v):
    if pd.isna(v):   return "확인 필요"
    elif v < 0.3:    return "공급 낮음"
    elif v < 0.6:    return "공급 보통"
    elif v < 0.8:    return "공급 양호"
    else:            return "공급 높음"

df["csi_level"] = df["CSI"].apply(assign_csi_level)
csi_level_order = ["공급 낮음", "공급 보통", "공급 양호", "공급 높음", "확인 필요"]
print("  csi_level 분포:")
for lv in csi_level_order:
    cnt = (df["csi_level"] == lv).sum()
    print(f"    {lv}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. CDI 수준 분류
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 4] CDI 수준 분류")

def assign_cdi_level(v):
    if pd.isna(v):   return "확인 필요"
    elif v < 0.3:    return "수요 낮음"
    elif v < 0.6:    return "수요 보통"
    elif v < 0.8:    return "수요 높음"
    else:            return "수요 매우 높음"

df["cdi_level"] = df["CDI"].apply(assign_cdi_level)
cdi_level_order = ["수요 낮음", "수요 보통", "수요 높음", "수요 매우 높음", "확인 필요"]
print("  cdi_level 분포:")
for lv in cdi_level_order:
    cnt = (df["cdi_level"] == lv).sum()
    print(f"    {lv}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. CSI-CDI 2x2 유형화
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 5] CSI-CDI 2x2 유형화")

def assign_supply_demand_type(row):
    cdi, csi = row["CDI"], row["CSI"]
    if pd.isna(cdi) or pd.isna(csi):
        return "확인 필요"
    demand_high = cdi >= cdi_mean
    supply_high = csi >= csi_mean
    if     demand_high and not supply_high: return "A. 수요높음-공급낮음"
    elif   demand_high and     supply_high: return "B. 수요높음-공급높음"
    elif not demand_high and not supply_high: return "C. 수요낮음-공급낮음"
    else:                                   return "D. 수요낮음-공급높음"

df["supply_demand_type"] = df[["CDI", "CSI"]].apply(assign_supply_demand_type, axis=1)

type_order = [
    "A. 수요높음-공급낮음", "B. 수요높음-공급높음",
    "C. 수요낮음-공급낮음", "D. 수요낮음-공급높음", "확인 필요",
]
print(f"  기준 - CDI 평균: {cdi_mean:.4f}, CSI 평균: {csi_mean:.4f}")
print("  supply_demand_type 분포:")
for t in type_order:
    cnt = (df["supply_demand_type"] == t).sum()
    print(f"    {t}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. 정책 조치 유형 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 6] 정책 조치 유형 생성")

def assign_policy_action_type(row):
    pl  = row["priority_level"]
    sdt = row["supply_demand_type"]
    if pl == "최우선 지원":              return "최우선 지원 검토"
    elif pl == "우선 지원":              return "우선 지원 검토"
    elif sdt == "A. 수요높음-공급낮음":  return "수요-공급 불균형 우선 개선"
    elif sdt == "B. 수요높음-공급높음":  return "고수요 학교 모니터링 및 프로그램 강화"
    elif sdt == "C. 수요낮음-공급낮음":  return "권역별 순회상담 또는 최소 인프라 보완"
    elif sdt == "D. 수요낮음-공급높음":  return "현 수준 유지 또는 거점학교 활용"
    else:                                return "확인 필요"

df["policy_action_type"] = df.apply(assign_policy_action_type, axis=1)
action_order = [
    "최우선 지원 검토", "우선 지원 검토",
    "수요-공급 불균형 우선 개선", "고수요 학교 모니터링 및 프로그램 강화",
    "권역별 순회상담 또는 최소 인프라 보완", "현 수준 유지 또는 거점학교 활용", "확인 필요",
]
print("  policy_action_type 분포:")
for a in action_order:
    cnt = (df["policy_action_type"] == a).sum()
    if cnt > 0:
        print(f"    {a}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 세부 정책 피드백 문장 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 7] 세부 정책 피드백 문장 생성")

def build_recommendation(row):
    parts = []
    if row["counseling_staff_supply_score"] < 0.4:
        parts.append("전문상담교사 배치 또는 순회상담교사 연계 검토")
    if row["wee_class_score"] == 0:
        parts.append("Wee클래스 신설 또는 운영 현황 재확인 검토")
    if row["wee_center_access_score"] <= 0.4:
        parts.append("Wee센터 접근성이 낮아 이동형 상담, 온라인 상담, 권역별 연계 강화 검토")
    if row["demand_size_score"] == 1.0:
        parts.append("대규모 학교로 상담수요 규모가 커 상담 인력 보강 검토")
    if row["counseling_use_score"] >= use_p80:
        parts.append("실제 상담 이용 수준이 높아 상담 프로그램 확대 또는 상담 인력 보완 검토")
    if row["school_violence_risk_score"] >= viol_p80:
        parts.append("학교폭력 피해 응답률 기반 위험 점수가 높아 학교폭력 관련 상담지원 강화 검토")
    if pd.notna(row["priority_score"]) and row["priority_score"] > 0:
        parts.append("상담수요지수가 상담공급지수보다 높아 우선 지원 검토 필요")
    return "; ".join(parts) if parts else "현 수준 유지 및 정기 모니터링"

df["policy_recommendation"] = df.apply(build_recommendation, axis=1)
print(f"  policy_recommendation 생성 완료")
print(f"    '현 수준 유지 및 정기 모니터링': {(df['policy_recommendation'] == '현 수준 유지 및 정기 모니터링').sum()}개교")
print(f"    복수 조건 해당 (세미콜론 포함): {df['policy_recommendation'].str.contains(';').sum()}개교")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. 정책 근거 문장 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 8] 정책 근거 문장 생성")

def build_reason(row):
    cdi = row["CDI"]
    csi = row["CSI"]
    ps  = row["priority_score"]

    # CDI vs CSI 관계
    if pd.isna(cdi) or pd.isna(csi):
        rel = "CSI 또는 CDI를 산출할 수 없어 관계를 판단하기 어렵다"
    elif cdi > csi:
        rel = (f"해당 학교는 상담수요지수(CDI={cdi:.3f})가 상담공급지수(CSI={csi:.3f})보다 높아 "
               "수요 대비 공급이 상대적으로 부족한 학교로 분류된다")
    elif cdi < csi:
        rel = (f"해당 학교는 상담공급지수(CSI={csi:.3f})가 상담수요지수(CDI={cdi:.3f})보다 높아 "
               "현재 공급이 수요를 상대적으로 충족하는 것으로 분류된다")
    else:
        rel = (f"해당 학교는 상담공급지수(CSI={csi:.3f})와 상담수요지수(CDI={cdi:.3f})가 동일하여 "
               "수요와 공급이 균형 상태로 분류된다")

    # 우선지원점수
    if pd.notna(ps):
        ps_desc = f"우선지원점수는 {ps:.3f}로, " + (
            "수요가 공급보다 높은 상태이다." if ps > 0
            else "공급이 수요보다 높거나 비슷한 상태이다."
        )
    else:
        ps_desc = "우선지원점수를 산출할 수 없다."

    # 취약 하위 점수 파악
    weak = []
    if row["counseling_staff_supply_score"] < 0.4:
        weak.append("상담인력 공급 점수")
    if row["wee_class_score"] == 0:
        weak.append("Wee클래스 미운영")
    if row["wee_center_access_score"] <= 0.4:
        weak.append("Wee센터 접근성 점수")
    if row["school_violence_risk_score"] >= viol_p80:
        weak.append("학교폭력 위험 점수")

    if weak:
        weak_desc = f"특히 {', '.join(weak)}이 취약하여 관련 지원 검토가 필요하다."
    else:
        weak_desc = "주요 하위 점수에서 두드러진 취약 항목은 확인되지 않는다."

    tail = "이 결과는 실제 지원 확정이 아니라 정책 검토 우선순위를 나타내며, 현장 의견 및 추가 검토를 거쳐 최종 판단해야 한다."

    return f"{rel}. {ps_desc} {weak_desc} {tail}"

df["policy_reason"] = df.apply(build_reason, axis=1)
print(f"  policy_reason 생성 완료")

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. 요약표 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 9] 요약표 생성")

# 9-1. supply_demand_type 요약
type_summary = (
    df.groupby("supply_demand_type")
    .agg(
        학교수        = ("school_code", "count"),
        CSI_평균      = ("CSI",           "mean"),
        CDI_평균      = ("CDI",           "mean"),
        priority_score_평균 = ("priority_score", "mean"),
    )
    .round(4)
    .reset_index()
)
# 유형 해석 추가
type_desc = {
    "A. 수요높음-공급낮음": "핵심 우선지원 유형",
    "B. 수요높음-공급높음": "고수요 유지관리 유형",
    "C. 수요낮음-공급낮음": "최소 인프라 보완 유형",
    "D. 수요낮음-공급높음": "안정 또는 거점 활용 유형",
    "확인 필요":           "판단 불가",
}
type_summary["해석"] = type_summary["supply_demand_type"].map(type_desc)
# 순서 정렬
type_summary["_sort"] = type_summary["supply_demand_type"].map(
    {t: i for i, t in enumerate(type_order)}
).fillna(99)
type_summary = type_summary.sort_values("_sort").drop(columns=["_sort"])

TYPE_CSV = TABLES_DIR / "supply_demand_type_summary_2025.csv"
type_summary.to_csv(TYPE_CSV, index=False, encoding="utf-8-sig")
print(f"  저장: {TYPE_CSV}")
print(type_summary[["supply_demand_type","학교수","CSI_평균","CDI_평균","priority_score_평균"]].to_string(index=False))

# 9-2. policy_feedback_summary
print()
action_cnt  = df["policy_action_type"].value_counts().rename_axis("구분").reset_index(name="학교수")
action_cnt.insert(0, "분류", "policy_action_type")

level_cnt   = df["priority_level"].value_counts().rename_axis("구분").reset_index(name="학교수")
level_cnt.insert(0, "분류", "priority_level")

# 주요 피드백 키워드별 학교 수
keywords = {
    "전문상담교사 배치 또는 순회상담교사 연계 검토":             "상담인력 배치 검토",
    "Wee클래스 신설 또는 운영 현황 재확인 검토":               "Wee클래스 검토",
    "이동형 상담, 온라인 상담, 권역별 연계 강화 검토":           "Wee센터 접근성 보완 검토",
    "대규모 학교로 상담수요 규모가 커 상담 인력 보강 검토":       "대규모 인력 보강 검토",
    "상담 프로그램 확대 또는 상담 인력 보완 검토":               "고이용 상담 확대 검토",
    "학교폭력 관련 상담지원 강화 검토":                         "학교폭력 상담 강화 검토",
    "상담수요지수가 상담공급지수보다 높아 우선 지원 검토 필요":    "수요>공급 우선 지원",
    "현 수준 유지 및 정기 모니터링":                            "현 수준 유지",
}
kw_rows = []
for kw, label in keywords.items():
    cnt = df["policy_recommendation"].str.contains(kw, regex=False).sum()
    kw_rows.append({"분류": "policy_recommendation_keyword", "구분": label, "학교수": cnt})
kw_df = pd.DataFrame(kw_rows)

feedback_summary = pd.concat([action_cnt, level_cnt, kw_df], ignore_index=True)
FEEDBACK_CSV = TABLES_DIR / "policy_feedback_summary_2025.csv"
feedback_summary.to_csv(FEEDBACK_CSV, index=False, encoding="utf-8-sig")
print(f"  저장: {FEEDBACK_CSV}")

# 9-3. sigungu_priority_summary
sigungu_summary = (
    df.groupby("sigungu")
    .agg(
        학교수              = ("school_code", "count"),
        CSI_평균            = ("CSI",           "mean"),
        CDI_평균            = ("CDI",           "mean"),
        priority_score_평균 = ("priority_score", "mean"),
        최우선지원_수       = ("priority_level", lambda x: (x == "최우선 지원").sum()),
        우선지원_수         = ("priority_level", lambda x: (x == "우선 지원").sum()),
        A유형_수            = ("supply_demand_type", lambda x: (x == "A. 수요높음-공급낮음").sum()),
    )
    .round(4)
    .reset_index()
    .sort_values("priority_score_평균", ascending=False)
)
SIGUNGU_CSV = TABLES_DIR / "sigungu_priority_summary_2025.csv"
sigungu_summary.to_csv(SIGUNGU_CSV, index=False, encoding="utf-8-sig")
print(f"  저장: {SIGUNGU_CSV}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. 최종 정책 피드백 엑셀 파일 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 10] 정책 피드백 엑셀 파일 생성")

FEEDBACK_COLS = [
    "school_code", "school_name", "sido", "sigungu",
    "counseling_staff_supply_score", "wee_class_score", "wee_center_access_score",
    "demand_size_score", "counseling_use_score", "school_violence_risk_score",
    "CSI", "CDI", "priority_score", "priority_level",
    "csi_level", "cdi_level", "supply_demand_type",
    "policy_action_type", "policy_recommendation", "policy_reason",
]
feedback_df = df[[c for c in FEEDBACK_COLS if c in df.columns]].copy()

# priority_level 기준 정렬 (최우선->우선->모니터링->안정)
LEVEL_SORT_MAP = {"최우선 지원": 0, "우선 지원": 1, "모니터링": 2, "안정": 3, "확인 필요": 4}
feedback_df["_sort"] = feedback_df["priority_level"].map(LEVEL_SORT_MAP).fillna(9)
feedback_df = (
    feedback_df
    .sort_values(["_sort", "priority_score"], ascending=[True, False])
    .drop(columns=["_sort"])
    .reset_index(drop=True)
)

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    feedback_df.to_excel(writer,       sheet_name="policy_feedback_table",      index=False)
    type_summary.to_excel(writer,      sheet_name="supply_demand_type_summary",  index=False)
    feedback_summary.to_excel(writer,  sheet_name="policy_feedback_summary",     index=False)
    sigungu_summary.to_excel(writer,   sheet_name="sigungu_priority_summary",    index=False)

print(f"  저장 완료: {OUTPUT_PATH}")
print(f"  시트: policy_feedback_table ({len(feedback_df)}행 x {len(feedback_df.columns)}열)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 11. 방법론 문서 저장
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 11] 방법론 문서 저장")

type_cnt = df["supply_demand_type"].value_counts()

methodology_md = f"""# 정책 피드백 로직 방법론

## 1. 정책 피드백 로직의 목적
경남 일반고 146개교의 CSI(상담공급지수)와 CDI(상담수요지수) 분석 결과를
실제 정책 검토에 활용 가능한 형태로 가공한다.
각 학교를 수요-공급 유형으로 분류하고, 하위 점수에 기반한 세부 피드백 문장을 생성하여
대시보드 및 보고서에서 학교별 맞춤형 정책 근거로 활용할 수 있도록 한다.

## 2. CSI-CDI 2x2 유형화 기준
CDI와 CSI 각각의 전체 평균값을 기준으로 수요 높음/낮음, 공급 높음/낮음을 구분한다.

| 기준 변수 | 기준값 |
|----------|--------|
| CDI 평균 (demand_high 기준) | {cdi_mean:.4f} |
| CSI 평균 (supply_high 기준) | {csi_mean:.4f} |

| 유형 | 조건 | 해석 | 학교 수 |
|------|------|------|---------|
| A. 수요높음-공급낮음 | CDI >= {cdi_mean:.4f} & CSI < {csi_mean:.4f} | 핵심 우선지원 유형 | {type_cnt.get('A. 수요높음-공급낮음', 0)}개교 |
| B. 수요높음-공급높음 | CDI >= {cdi_mean:.4f} & CSI >= {csi_mean:.4f} | 고수요 유지관리 유형 | {type_cnt.get('B. 수요높음-공급높음', 0)}개교 |
| C. 수요낮음-공급낮음 | CDI < {cdi_mean:.4f} & CSI < {csi_mean:.4f} | 최소 인프라 보완 유형 | {type_cnt.get('C. 수요낮음-공급낮음', 0)}개교 |
| D. 수요낮음-공급높음 | CDI < {cdi_mean:.4f} & CSI >= {csi_mean:.4f} | 안정 또는 거점 활용 유형 | {type_cnt.get('D. 수요낮음-공급높음', 0)}개교 |

## 3. priority_level 활용 방식
분위수 기반 우선지원등급(priority_level)을 policy_action_type 결정에 1순위로 활용한다.
priority_level이 "최우선 지원" 또는 "우선 지원"인 학교는
supply_demand_type과 관계없이 우선 지원 검토 대상으로 분류된다.

## 4. policy_action_type 생성 규칙
다음 우선순위로 하나의 policy_action_type을 부여한다.

| 우선순위 | 조건 | 부여값 |
|---------|------|--------|
| 1 | priority_level = 최우선 지원 | 최우선 지원 검토 |
| 2 | priority_level = 우선 지원 | 우선 지원 검토 |
| 3 | A유형 | 수요-공급 불균형 우선 개선 |
| 4 | B유형 | 고수요 학교 모니터링 및 프로그램 강화 |
| 5 | C유형 | 권역별 순회상담 또는 최소 인프라 보완 |
| 6 | D유형 | 현 수준 유지 또는 거점학교 활용 |
| 7 | 기타 | 확인 필요 |

## 5. policy_recommendation 생성 규칙
각 학교의 하위 점수를 조건으로 검토하여 해당하는 문장을 세미콜론으로 연결한다.
조건에 해당하지 않으면 "현 수준 유지 및 정기 모니터링"을 부여한다.

| 조건 | 추가 문장 |
|------|----------|
| counseling_staff_supply_score < 0.4 | 전문상담교사 배치 또는 순회상담교사 연계 검토 |
| wee_class_score == 0 | Wee클래스 신설 또는 운영 현황 재확인 검토 |
| wee_center_access_score <= 0.4 | 이동형 상담, 온라인 상담, 권역별 연계 강화 검토 |
| demand_size_score == 1.0 | 대규모 학교로 상담 인력 보강 검토 |
| counseling_use_score >= P80({use_p80:.4f}) | 상담 프로그램 확대 또는 상담 인력 보완 검토 |
| school_violence_risk_score >= P80({viol_p80:.4f}) | 학교폭력 관련 상담지원 강화 검토 |
| priority_score > 0 | 수요지수가 공급지수보다 높아 우선 지원 검토 필요 |

## 6. 시군구별 요약 방식
18개 시군구별로 다음을 집계한다.
- 학교 수, 평균 CSI, 평균 CDI, 평균 priority_score
- 최우선 지원 / 우선 지원 / A유형 학교 수
- priority_score 평균 기준 내림차순 정렬

## 7. 정책 피드백의 성격
이 파일에서 생성된 모든 정책 피드백(policy_action_type, policy_recommendation, policy_reason)은
실제 지원 확정이 아니라 **정책 검토 우선순위**를 나타낸다.
최종 지원 여부는 현장 방문, 담당 교사 의견, 지역 여건 등을 종합하여 별도로 결정해야 한다.

## 8. 산출 결과 요약
- 대상 학교: {n_rows}개교 / 시군구: {df['sigungu'].nunique()}개
- CDI 평균: {cdi_mean:.4f}, CSI 평균: {csi_mean:.4f}
- supply_demand_type A유형 (핵심 우선지원): {type_cnt.get('A. 수요높음-공급낮음', 0)}개교
- 최우선 지원 + 우선 지원: {(df['priority_level'].isin(['최우선 지원','우선 지원'])).sum()}개교

## 9. 한계
1. **운영 기준의 임의성**: 점수 기준(평균값, P80 등)은 분석 목적에 맞춘 운영 기준이며,
   절대적 기준으로 해석해서는 안 된다.
2. **추가 검토 필요 요소**: 실제 배치 가능 인력, 예산, 학교 현장 의견은 정량 지표에 반영되지 않는다.
3. **개별 특수 상황**: 학교별 특수 상황(학생 특성, 지역 여건, 관리자 리더십 등)은
   정량 지표만으로 완전히 설명하기 어렵다.
"""

METHODOLOGY_MD = DOCS_DIR / "policy_feedback_logic_2025.md"
METHODOLOGY_MD.write_text(methodology_md, encoding="utf-8")
print(f"  저장 완료: {METHODOLOGY_MD}")

# ════════════════════════════════════════════════════════════════════════════
# 최종 요약 출력
# ════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("CSI-CDI 기반 학교 유형화 및 정책 피드백 생성 완료")
print("=" * 60)
print(f"\n대상 학교: {n_rows}개교 / 시군구: {df['sigungu'].nunique()}개")

print("\nsupply_demand_type 분포:")
for t in type_order:
    cnt = (df["supply_demand_type"] == t).sum()
    print(f"  {t}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

print("\npolicy_action_type 분포:")
for a in action_order:
    cnt = (df["policy_action_type"] == a).sum()
    if cnt > 0:
        print(f"  {a}: {cnt}개교 ({cnt/n_rows*100:.1f}%)")

print("\n생성 파일:")
for p in [OUTPUT_PATH, TYPE_CSV, FEEDBACK_CSV, SIGUNGU_CSV, METHODOLOGY_MD]:
    print(f"  {p}")
