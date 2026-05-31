"""
상담공급지수(CSI) 산출 및 CSI 분석용 엑셀 파일 생성.

CSI = (counseling_staff_supply_score + wee_class_score + wee_center_access_score) / 3

- 세 구성요소가 모두 존재할 때만 공식 CSI를 산출한다.
- 하나라도 결측이면 CSI = NaN, 단 결측 제외 평균(csi_available_mean)을 별도 산출한다.
- wee_class(0/1) → wee_class_score(0.0/1.0) 변환을 이번 단계에서 수행한다.
- 원본 입력 파일은 덮어쓰지 않는다.

입력 파일:
  data/processed/gyeongnam_general_high_schools_with_wee_access_score_2025.csv

출력 파일:
  data/processed/gyeongnam_high_schools_master_table.xlsx
  data/processed/gyeongnam_high_schools_CSI.xlsx
  outputs/tables/csi_summary_2025.csv
  outputs/tables/csi_missing_check_2025.csv
  docs/csi_methodology_2025.md
"""

import sys
import pathlib

import numpy as np
import pandas as pd

# ── 경로 설정 ────────────────────────────────────────────────────────────────
ROOT          = pathlib.Path(__file__).resolve().parent.parent
PROCESSED_DIR = ROOT / "data" / "processed"
TABLES_DIR    = ROOT / "outputs" / "tables"
DOCS_DIR      = ROOT / "docs"

for d in [PROCESSED_DIR, TABLES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 입력 파일 로드
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 입력 파일 로드")

INPUT_PATH = PROCESSED_DIR / "gyeongnam_general_high_schools_with_wee_access_score_2025.csv"

if not INPUT_PATH.exists():
    print(f"[ERROR] 입력 파일 없음: {INPUT_PATH}")
    sys.exit(1)

df = pd.read_csv(INPUT_PATH, encoding="utf-8-sig",
                 dtype={"school_code": str, "postcode": str})

print(f"  → {len(df)}행 × {len(df.columns)}열 로드 완료")

# 학교 단위 데이터 확인 (school_code 중복 없어야 함)
if df["school_code"].duplicated().any():
    print("[WARNING] school_code에 중복이 있습니다. 데이터를 확인하세요.")
else:
    print(f"  → 학교 단위 데이터 확인 완료 (school_code 중복 없음)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. 필수 변수 존재 여부 확인
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 2] 필수 변수 확인")

# 식별 변수 후보 탐색
def find_col(df, primary, candidates):
    if primary in df.columns:
        return primary
    for c in candidates:
        if c in df.columns:
            print(f"  [INFO] '{primary}' 없음 → '{c}' 대체 사용")
            return c
    return None

id_cols = {
    "school_code": find_col(df, "school_code", ["학교코드", "SCHUL_CODE", "school_id"]),
    "school_name": find_col(df, "school_name", ["학교명", "SCHUL_NM"]),
    "sido"       : find_col(df, "sido",        ["시도", "SIDO"]),
    "sigungu"    : find_col(df, "sigungu",     ["시군구", "지역", "district"]),
}

# CSI 구성요소 변수
csi_cols = {
    "counseling_staff_supply_score": find_col(
        df, "counseling_staff_supply_score",
        ["상담인력공급점수", "staff_supply_score"]),
    "wee_class": find_col(
        df, "wee_class",
        ["위클래스", "wee_class_yn", "wee_class_operation"]),
    "wee_center_access_score": find_col(
        df, "wee_center_access_score",
        ["위센터접근성", "wee_access_score"]),
}

# 핵심 변수 없으면 중단
missing_required = [k for k, v in csi_cols.items() if v is None]
if missing_required:
    print(f"[ERROR] 다음 변수를 찾지 못했습니다: {missing_required}")
    print("        입력 파일 또는 변수명을 확인하세요.")
    sys.exit(1)

for col, found in {**id_cols, **csi_cols}.items():
    status = "[OK]" if found else "[MISSING]"
    n_miss = df[found].isna().sum() if found else "-"
    print(f"  {status} {col:<40s} → '{found}'  결측={n_miss}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. wee_class_score 생성
# wee_class 원본(0/1/NaN)을 float 점수(0.0/1.0/NaN)로 변환한다.
# wee_class == 0: 실제 미운영 → 0.0으로 확정 처리
# wee_class 결측: 미운영으로 단정하지 않음 → NaN 유지
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 3] wee_class_score 생성")

wee_class_col = csi_cols["wee_class"]
wc = pd.to_numeric(df[wee_class_col], errors="coerce")   # 숫자형 변환

# 0, 1 외 값 확인
unexpected = wc.dropna()[~wc.dropna().isin([0, 1])]
if not unexpected.empty:
    print(f"  [WARNING] wee_class에 0/1 외의 값 감지: {unexpected.unique().tolist()}")
    print(f"           해당 학교 수: {len(unexpected)}개 — csi_missing_check에 기록됩니다.")
else:
    print(f"  [OK] wee_class 고유값: {sorted(wc.dropna().unique().tolist())} (0과 1만 존재)")

# wee_class_score 매핑
df["wee_class_score"] = wc.map({1: 1.0, 0: 0.0})
# 0/1 외 값은 NaN으로 남겨 결측으로 처리

n1  = int((df["wee_class_score"] == 1.0).sum())
n0  = int((df["wee_class_score"] == 0.0).sum())
nan = int(df["wee_class_score"].isna().sum())
print(f"  wee_class_score = 1.0 (운영)  : {n1}개교")
print(f"  wee_class_score = 0.0 (미운영): {n0}개교")
print(f"  wee_class_score = NaN (결측)  : {nan}개교")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. CSI 및 파생 변수 산출
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 4] CSI 산출")

comp_cols = [
    csi_cols["counseling_staff_supply_score"],
    "wee_class_score",
    csi_cols["wee_center_access_score"],
]

# csi_components_available: 결측이 아닌 구성요소 수
df["csi_components_available"] = df[comp_cols].notna().sum(axis=1).astype(int)

# csi_missing_components: 결측 구성요소명 목록 (세미콜론 구분)
def list_missing(row):
    missing = [c for c in comp_cols if pd.isna(row[c])]
    return "; ".join(missing) if missing else ""

df["csi_missing_components"] = df.apply(list_missing, axis=1)

# CSI: 세 구성요소 모두 존재할 때만 계산 (하나라도 결측이면 NaN)
df["CSI"] = df[comp_cols].apply(
    lambda row: row.mean() if row.notna().all() else float("nan"),
    axis=1
)

# csi_available_mean: 결측 제외 평균 (민감도 분석·보조 지표용)
df["csi_available_mean"] = df[comp_cols].mean(axis=1, skipna=True)
# 세 구성요소 모두 결측인 경우 csi_available_mean도 NaN
df.loc[df["csi_components_available"] == 0, "csi_available_mean"] = float("nan")

# 결과 요약
n_csi_ok  = int(df["CSI"].notna().sum())
n_csi_nan = int(df["CSI"].isna().sum())
print(f"  CSI 산출 완료: {n_csi_ok}개교")
print(f"  CSI 결측     : {n_csi_nan}개교")
if n_csi_ok > 0:
    print(f"  CSI 평균     : {df['CSI'].mean():.4f}")
    print(f"  CSI 최솟값   : {df['CSI'].min():.4f}")
    print(f"  CSI 최댓값   : {df['CSI'].max():.4f}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. CSI 점수 구간 분류 (임시)
# 0.0~0.3: 낮음 / 0.3~0.6: 보통 / 0.6~0.8: 양호 / 0.8~1.0: 높음
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 5] CSI 점수 구간 분류")

def csi_grade(csi_val):
    if pd.isna(csi_val):
        return "확인 필요"
    if csi_val < 0.3:
        return "공급 낮음"
    elif csi_val < 0.6:
        return "공급 보통"
    elif csi_val < 0.8:
        return "공급 양호"
    else:
        return "공급 높음"

df["csi_grade"] = df["CSI"].apply(csi_grade)

grade_order = ["공급 낮음", "공급 보통", "공급 양호", "공급 높음", "확인 필요"]
for g in grade_order:
    cnt = int((df["csi_grade"] == g).sum())
    print(f"  {g}: {cnt}개교  ({cnt/len(df)*100:.1f}%)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. CSI 요약표 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 6] CSI 요약표 생성")

summary_records = [
    {"항목": "전체 학교 수",                    "값": len(df)},
    {"항목": "CSI 산출 완료 학교 수",            "값": n_csi_ok},
    {"항목": "CSI 결측 학교 수",                "값": n_csi_nan},
    {"항목": "CSI 평균",                        "값": round(df["CSI"].mean(), 4)},
    {"항목": "CSI 중앙값",                      "값": round(df["CSI"].median(), 4)},
    {"항목": "CSI 최솟값",                      "값": round(df["CSI"].min(), 4)},
    {"항목": "CSI 최댓값",                      "값": round(df["CSI"].max(), 4)},
    {"항목": "--- 구성요소별 ---",              "값": ""},
    {"항목": "counseling_staff_supply_score 평균",
     "값": round(df[csi_cols["counseling_staff_supply_score"]].mean(), 4)},
    {"항목": "counseling_staff_supply_score 결측",
     "값": int(df[csi_cols["counseling_staff_supply_score"]].isna().sum())},
    {"항목": "wee_class_score 평균",
     "값": round(df["wee_class_score"].mean(), 4)},
    {"항목": "wee_class_score 결측",
     "값": int(df["wee_class_score"].isna().sum())},
    {"항목": "wee_center_access_score 평균",
     "값": round(df[csi_cols["wee_center_access_score"]].mean(), 4)},
    {"항목": "wee_center_access_score 결측",
     "값": int(df[csi_cols["wee_center_access_score"]].isna().sum())},
    {"항목": "--- CSI 구간별 학교 수 ---",      "값": ""},
]
for g in grade_order:
    summary_records.append({
        "항목": f"CSI 구간: {g}",
        "값": int((df["csi_grade"] == g).sum()),
    })

df_summary = pd.DataFrame(summary_records)
summary_path = TABLES_DIR / "csi_summary_2025.csv"
df_summary.to_csv(summary_path, index=False, encoding="utf-8-sig")
print(f"  → 저장: {summary_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 결측 점검표 생성
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 7] 결측 점검표 생성")

check_display_cols = [
    "school_code", "school_name", "sido", "sigungu",
    csi_cols["counseling_staff_supply_score"],
    wee_class_col, "wee_class_score",
    csi_cols["wee_center_access_score"],
    "CSI", "csi_available_mean",
    "csi_components_available", "csi_missing_components",
]
# 표준 변수명으로 통일 (별칭 대체 시 원본 열명이 다를 수 있으므로)
check_display_cols = list(dict.fromkeys(check_display_cols))  # 중복 제거

# 점검 대상 행 수집 (중복 포함 허용 — 여러 조건 해당 가능)
check_masks = {
    "counseling_staff_supply_score 결측":
        df[csi_cols["counseling_staff_supply_score"]].isna(),
    "wee_class 결측":
        df[wee_class_col].isna(),
    "wee_class_score 결측":
        df["wee_class_score"].isna(),
    "wee_center_access_score 결측":
        df[csi_cols["wee_center_access_score"]].isna(),
    "CSI 결측":
        df["CSI"].isna(),
    "csi_components_available < 3":
        df["csi_components_available"] < 3,
    "wee_class 값 이상 (0/1 외)":
        pd.to_numeric(df[wee_class_col], errors="coerce").notna() &
        ~pd.to_numeric(df[wee_class_col], errors="coerce").isin([0, 1]),
}

check_rows = []
for reason, mask in check_masks.items():
    subset = df[mask][check_display_cols].copy()
    subset.insert(0, "check_reason", reason)
    check_rows.append(subset)

if check_rows and any(len(r) > 0 for r in check_rows):
    df_check = pd.concat([r for r in check_rows if len(r) > 0], ignore_index=True)
else:
    # 모든 항목 이상 없으면 대표 행 1개만 기록
    df_check = pd.DataFrame([{"check_reason": "이상 없음 — 전 146개교 CSI 정상 산출"}])

check_path = TABLES_DIR / "csi_missing_check_2025.csv"
df_check.to_csv(check_path, index=False, encoding="utf-8-sig")
print(f"  → 저장: {check_path.name}  ({len(df_check)}건)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. Master Table 엑셀 저장
# 전체 변수에 wee_class_score, CSI 관련 변수 추가한 완전 통합 데이터셋
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 8] Master Table 엑셀 저장")

master_path = PROCESSED_DIR / "gyeongnam_high_schools_master_table.xlsx"

with pd.ExcelWriter(master_path, engine="openpyxl") as writer:
    # 시트 1: master_table (전체 변수)
    df.to_excel(writer, sheet_name="master_table", index=False)

    # 시트 2: csi_summary
    df_summary.to_excel(writer, sheet_name="csi_summary", index=False)

    # 시트 3: csi_missing_check
    df_check.to_excel(writer, sheet_name="csi_missing_check", index=False)

print(f"  → 저장: {master_path.name}")
print(f"    시트: master_table ({len(df)}행 × {len(df.columns)}열) / csi_summary / csi_missing_check")

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. CSI 전용 엑셀 저장
# CSI 분석에 필요한 핵심 12개 변수만 추출
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 9] CSI 전용 엑셀 저장")

csi_only_cols = [
    "school_code", "school_name", "sido", "sigungu",
    csi_cols["counseling_staff_supply_score"],
    wee_class_col,
    "wee_class_score",
    csi_cols["wee_center_access_score"],
    "CSI",
    "csi_available_mean",
    "csi_components_available",
    "csi_missing_components",
    "csi_grade",
]
# 표준 변수명 확보 (별칭 사용 시 이름 통일)
csi_only_cols = [c for c in csi_only_cols if c in df.columns]
df_csi = df[csi_only_cols].copy()

# 열 이름을 표준 변수명으로 정규화
rename_map = {
    csi_cols["counseling_staff_supply_score"]: "counseling_staff_supply_score",
    wee_class_col:                            "wee_class",
    csi_cols["wee_center_access_score"]:      "wee_center_access_score",
}
df_csi = df_csi.rename(columns=rename_map)

csi_path = PROCESSED_DIR / "gyeongnam_high_schools_CSI.xlsx"

with pd.ExcelWriter(csi_path, engine="openpyxl") as writer:
    # 시트 1: CSI_table
    df_csi.to_excel(writer, sheet_name="CSI_table", index=False)

    # 시트 2: csi_summary
    df_summary.to_excel(writer, sheet_name="csi_summary", index=False)

    # 시트 3: csi_missing_check
    df_check.to_excel(writer, sheet_name="csi_missing_check", index=False)

print(f"  → 저장: {csi_path.name}")
print(f"    시트: CSI_table ({len(df_csi)}행 × {len(df_csi.columns)}열) / csi_summary / csi_missing_check")
print(f"    포함 변수: {list(df_csi.columns)}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. 방법론 문서 작성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 10] 방법론 문서 작성")

doc_lines = [
    "# 상담공급지수(CSI) 방법론 문서 (2025)",
    "",
    "## 목적",
    "",
    "경상남도 일반고등학교 146개교의 상담 서비스 공급 수준을",
    "다차원적으로 측정하기 위해 상담공급지수(Counseling Supply Index, CSI)를 산출한다.",
    "CSI는 학교 내부 상담 인프라(전문상담교사 배치 수준, Wee클래스 운영 여부)와",
    "외부 상담 자원 접근성(Wee센터 직선거리)을 동등 가중치로 통합한 복합 지수이다.",
    "",
    "## CSI 산출식",
    "",
    "```",
    "CSI = (counseling_staff_supply_score + wee_class_score + wee_center_access_score) / 3",
    "```",
    "",
    "세 구성요소를 동등 가중치(각 1/3)로 산술 평균하여 산출한다.",
    "CSI의 이론적 범위는 0.0 ~ 1.0이며, 값이 클수록 상담 공급 수준이 높음을 의미한다.",
    "",
    "## 구성요소별 산출 기준",
    "",
    "### 1. 상담인력 공급 점수 (counseling_staff_supply_score)",
    "",
    "전문상담교사 1인당 학생 수(students_per_counselor)를 기준으로 구간화한다.",
    "",
    "| 기준 | 점수 |",
    "|------|------|",
    "| 전문상담교사 미배치 (counselor_count = 0) | 0.0 |",
    "| 1인당 학생 500명 이상 | 0.4 |",
    "| 1인당 학생 250명 이상 ~ 500명 미만 | 0.7 |",
    "| 1인당 학생 250명 미만 | 1.0 |",
    "| counselor_count 결측 | NaN |",
    "",
    "- 학생 수는 students_per_counselor 계산 시 간접 반영되며, CSI에 직접 포함하지 않는다.",
    "- 전문상담교사 0명(미배치)은 실제 0으로 처리하며, 결측과 구분한다.",
    "",
    "### 2. Wee클래스 운영 점수 (wee_class_score)",
    "",
    "학교알리미 공시 데이터 기반 Wee클래스 운영 여부(wee_class)를 점수화한다.",
    "",
    "| wee_class 값 | wee_class_score | 의미 |",
    "|-------------|-----------------|------|",
    "| 1 | 1.0 | Wee클래스 운영 |",
    "| 0 | 0.0 | 실제 미운영 (확정 0값) |",
    "| NaN | NaN | 정보 불명 (미운영으로 단정하지 않음) |",
    "",
    "- wee_class = 0은 실제 미운영으로 보고 0.0으로 처리한다.",
    "- wee_class 결측은 미운영으로 단정하지 않고 NaN으로 유지한다.",
    "",
    "### 3. Wee센터 접근성 점수 (wee_center_access_score)",
    "",
    "카카오 Local API 지오코딩 결과와 Haversine 직선거리 계산 기반 구간화.",
    "각 학교에서 가장 가까운 Wee센터(경남 19개)까지의 직선거리를 사용한다.",
    "",
    "| 직선거리(km) | 점수 |",
    "|------------|------|",
    "| 5km 미만 | 1.0 |",
    "| 5km 이상 ~ 10km 미만 | 0.7 |",
    "| 10km 이상 ~ 15km 미만 | 0.4 |",
    "| 15km 이상 | 0.1 |",
    "| 거리 산출 불가 | NaN |",
    "",
    "- 구간 기준은 실측 최대 직선거리(20.08 km) 기반으로 변별력을 고려하여 설정.",
    "- 직선거리는 실제 이동거리·이동시간이 아님.",
    "",
    "## 결측 처리 기준",
    "",
    "| 상황 | CSI | csi_available_mean |",
    "|------|-----|--------------------|",
    "| 3개 구성요소 모두 존재 | 공식 계산 | 공식 계산 (= CSI) |",
    "| 1~2개 결측 | NaN | 비결측 구성요소 평균 |",
    "| 3개 모두 결측 | NaN | NaN |",
    "",
    "## csi_available_mean을 별도 생성한 이유",
    "",
    "`csi_available_mean`은 결측이 포함된 학교에 대해 민감도 분석이나",
    "보조 지표로 활용하기 위해 별도로 계산한 값이다.",
    "공식 CSI는 세 구성요소가 모두 존재할 때만 산출하여 엄밀성을 유지하되,",
    "`csi_available_mean`을 통해 결측 1~2개인 학교의 상대적 수준을 참고할 수 있다.",
    "",
    "## CSI 점수 구간 해석 (임시 기준)",
    "",
    "| CSI 범위 | 등급 | 해석 |",
    "|---------|------|------|",
    "| 0.0 이상 ~ 0.3 미만 | 공급 낮음 | 상담 인프라 전반적으로 취약 |",
    "| 0.3 이상 ~ 0.6 미만 | 공급 보통 | 일부 인프라 갖춰졌으나 개선 필요 |",
    "| 0.6 이상 ~ 0.8 미만 | 공급 양호 | 상담 공급 수준 양호 |",
    "| 0.8 이상 ~ 1.0 이하 | 공급 높음 | 상담 인프라 전반적으로 충실 |",
    "",
    "- 위 구간 기준은 탐색적 분석을 위한 임시 분류이며,",
    "  정책적 활용 시 전문가 검토를 통한 재조정이 필요하다.",
    "",
    "## 한계",
    "",
    "- **동일 가중치 적용의 한계**: 세 구성요소에 동등 가중치(1/3)를 부여하였으나,",
    "  구성요소별 중요도가 실제로 동일하다는 근거가 충분하지 않다.",
    "  향후 전문가 조사(AHP 등)를 통한 가중치 산출이 필요하다.",
    "- **Wee센터 접근성은 직선거리 기준**: 실제 도로 거리 및 대중교통 접근성을",
    "  반영하지 못하며, 특히 산간·도서 지역에서 과대 추정될 수 있다.",
    "- **상담 인프라의 질적 수준 미반영**: 전문상담교사 자격 수준, 상담 공간의",
    "  적절성, Wee클래스 프로그램 품질 등 질적 지표는 포함되지 않았다.",
    "- **데이터 최신성 확인 필요**: 전문상담교사 수(KESS 2025) 및 Wee클래스",
    "  운영 여부(학교알리미 2025 공시)는 조사 기준일 시점 기준이며,",
    "  학기 중 변동이 반영되지 않을 수 있다.",
]

doc_path = DOCS_DIR / "csi_methodology_2025.md"
doc_path.write_text("\n".join(doc_lines), encoding="utf-8")
print(f"  → 저장: {doc_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# 최종 요약
# ════════════════════════════════════════════════════════════════════════════
print()
print("=" * 60)
print("최종 요약")
print("=" * 60)
print(f"  전체 학교 수           : {len(df)}개교")
print(f"  CSI 산출 완료          : {n_csi_ok}개교")
print(f"  CSI 결측               : {n_csi_nan}개교")
print(f"  CSI 평균               : {df['CSI'].mean():.4f}")
print(f"  CSI 중앙값             : {df['CSI'].median():.4f}")
print(f"  CSI 최솟값             : {df['CSI'].min():.4f}")
print(f"  CSI 최댓값             : {df['CSI'].max():.4f}")
print()
print("  [구성요소 평균]")
print(f"  counseling_staff_supply_score : {df[csi_cols['counseling_staff_supply_score']].mean():.4f}")
print(f"  wee_class_score               : {df['wee_class_score'].mean():.4f}")
print(f"  wee_center_access_score       : {df[csi_cols['wee_center_access_score']].mean():.4f}")
print()
print("  [CSI 구간별 학교 수]")
for g in grade_order:
    cnt = int((df["csi_grade"] == g).sum())
    print(f"  {g}: {cnt}개교  ({cnt/len(df)*100:.1f}%)")
print("=" * 60)
print()
print("[생성 파일]")
print("  data/processed/gyeongnam_high_schools_master_table.xlsx")
print("  data/processed/gyeongnam_high_schools_CSI.xlsx")
print("  outputs/tables/csi_summary_2025.csv")
print("  outputs/tables/csi_missing_check_2025.csv")
print("  docs/csi_methodology_2025.md")
