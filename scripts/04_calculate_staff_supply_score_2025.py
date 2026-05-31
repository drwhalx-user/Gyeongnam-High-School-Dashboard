"""
상담인력 공급 점수(counseling_staff_supply_score) 산출
── CSI(상담공급지수)의 구성요소 중 상담인력 공급 점수를 구간화 방식으로 생성한다.

점수 구간화 기준:
  counselor_count == 0                                  → 0.0 (미배치 확정)
  counselor_count >= 1 AND students_per_counselor >= 500 → 0.4
  counselor_count >= 1 AND 250 <= s_p_c < 500           → 0.7
  counselor_count >= 1 AND students_per_counselor < 250  → 1.0
  counselor_count 결측                                  → NaN
  counselor_count >= 1 AND students_per_counselor 결측  → NaN

입력: data/processed/gyeongnam_general_high_schools_with_counseling_use_2025.csv
출력:
  data/processed/gyeongnam_general_high_schools_with_staff_supply_score_2025.csv
  outputs/tables/staff_supply_score_summary_2025.csv
  outputs/tables/staff_supply_score_missing_check_2025.csv
  docs/staff_supply_score_methodology_2025.md
"""

import sys
import pathlib
import numpy as np
import pandas as pd

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT = pathlib.Path(__file__).resolve().parent.parent

INPUT_FILE   = ROOT / "data" / "processed" / "gyeongnam_general_high_schools_with_counseling_use_2025.csv"
OUTPUT_FILE  = ROOT / "data" / "processed" / "gyeongnam_general_high_schools_with_staff_supply_score_2025.csv"
SUMMARY_CSV  = ROOT / "outputs" / "tables" / "staff_supply_score_summary_2025.csv"
MISSING_CSV  = ROOT / "outputs" / "tables" / "staff_supply_score_missing_check_2025.csv"
METHOD_MD    = ROOT / "docs"               / "staff_supply_score_methodology_2025.md"

# 필요한 폴더 자동 생성
for _d in [OUTPUT_FILE.parent, SUMMARY_CSV.parent, METHOD_MD.parent]:
    _d.mkdir(parents=True, exist_ok=True)

# 실제 열 이름 (사전 확인 완료)
COL_COUNSELOR = "counselor_count"
COL_STUDENT   = "student_count"
COL_SPC       = "students_per_counselor"  # 상담교사 1인당 학생 수


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 : 입력 파일 로드
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 입력 파일 로드")

if not INPUT_FILE.exists():
    print(f"[ERROR] 입력 파일이 없습니다: {INPUT_FILE}")
    sys.exit(1)

df = pd.read_csv(INPUT_FILE, encoding="utf-8-sig")
print(f"  → {len(df)}개교 × {len(df.columns)}열 로드 완료")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 : 필수 변수 존재 확인 및 유사 변수 탐색
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 2] 필수 변수 확인")

required = {COL_COUNSELOR: "전문상담교사 수",
            COL_STUDENT  : "총 학생 수",
            COL_SPC      : "상담교사 1인당 학생 수"}
missing_cols = []

for col, desc in required.items():
    if col in df.columns:
        print(f"  [OK] '{col}' ({desc}) 확인")
    else:
        # 유사 변수명 탐색
        keywords = col.lower().replace("_", " ").split()
        candidates = [c for c in df.columns
                      if any(kw in c.lower() for kw in keywords)]
        print(f"  [경고] '{col}' ({desc}) 없음. 후보 변수: {candidates}")
        missing_cols.append(col)

if missing_cols:
    print("[ERROR] 필수 변수 미존재 — 열 이름을 확인하고 코드를 수정하세요.")
    sys.exit(1)

print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 : 숫자형 변환 및 inf → NaN 처리
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 3] 숫자형 변환 및 inf 처리")

df[COL_COUNSELOR] = pd.to_numeric(df[COL_COUNSELOR], errors="coerce")
df[COL_SPC]       = pd.to_numeric(df[COL_SPC],       errors="coerce")

# inf / -inf → NaN (0 나누기 결과가 남아 있는 경우 대비)
n_inf = np.isinf(df[COL_SPC]).sum()
if n_inf > 0:
    print(f"  [주의] students_per_counselor inf 값 {n_inf}건 → NaN 변환")
df[COL_SPC] = df[COL_SPC].replace([np.inf, -np.inf], np.nan)

print(f"  counselor_count  : 결측 {df[COL_COUNSELOR].isna().sum()}개 / "
      f"0값 {(df[COL_COUNSELOR]==0).sum()}개 / "
      f"1이상 {(df[COL_COUNSELOR]>=1).sum()}개")
print(f"  students_per_counselor: 결측 {df[COL_SPC].isna().sum()}개 / "
      f"범위 {df[COL_SPC].min():.1f}~{df[COL_SPC].max():.1f}")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 : counseling_staff_supply_score 산출 (구간화)
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 4] counseling_staff_supply_score 산출")

def assign_score(row) -> float:
    """
    상담인력 공급 점수 구간화 함수.

    규칙 우선순위:
    1. counselor_count 결측 → NaN
    2. counselor_count == 0 → 0.0 (students_per_counselor 무관)
    3. counselor_count >= 1 AND students_per_counselor 결측/inf → NaN
    4. counselor_count >= 1 AND students_per_counselor >= 500 → 0.4
    5. counselor_count >= 1 AND 250 <= students_per_counselor < 500 → 0.7
    6. counselor_count >= 1 AND students_per_counselor < 250 → 1.0
    """
    cc  = row[COL_COUNSELOR]
    spc = row[COL_SPC]

    # 규칙 1: counselor_count 자체 결측 → 미배치 단정 불가
    if pd.isna(cc):
        return float("nan")

    # 규칙 2: 전문상담교사 미배치 확정
    if cc == 0:
        return 0.0

    # 이하 counselor_count >= 1 케이스
    # 규칙 3: students_per_counselor 결측 또는 inf
    if pd.isna(spc) or np.isinf(spc):
        return float("nan")

    # 규칙 4~6: 구간화
    if spc >= 500:
        return 0.4
    elif spc >= 250:
        return 0.7
    else:
        return 1.0

df["counseling_staff_supply_score"] = df.apply(assign_score, axis=1)

# 결과 분포 출력
score_counts = df["counseling_staff_supply_score"].value_counts(dropna=False).sort_index()
print("  점수 분포:")
for score, cnt in score_counts.items():
    label = "NaN (결측)" if pd.isna(score) else str(score)
    print(f"    {label}: {cnt}개교")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 : 점수 요약표 저장
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 5] 점수 요약표 저장")

summary_rows = []
score_labels = {
    0.0: "0.0 — 전문상담교사 미배치",
    0.4: "0.4 — 1인당 학생 500명 이상",
    0.7: "0.7 — 1인당 학생 250~499명",
    1.0: "1.0 — 1인당 학생 250명 미만",
}
total = len(df)

for score_val, label in score_labels.items():
    cnt = int((df["counseling_staff_supply_score"] == score_val).sum())
    summary_rows.append({
        "score"      : score_val,
        "label"      : label,
        "count"      : cnt,
        "ratio(%)"   : round(cnt / total * 100, 2),
    })

n_missing_score = int(df["counseling_staff_supply_score"].isna().sum())
summary_rows.append({
    "score"   : "NaN",
    "label"   : "NaN — 결측 (counselor_count 또는 students_per_counselor 미확인)",
    "count"   : n_missing_score,
    "ratio(%)": round(n_missing_score / total * 100, 2),
})

df_summary = pd.DataFrame(summary_rows)
df_summary.to_csv(SUMMARY_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 점수 요약표 저장: {SUMMARY_CSV.name}")
print(df_summary.to_string(index=False))
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 : 결측 및 확인 필요 학교 목록 저장
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 6] 결측·확인 필요 학교 점검")

check_cols = ["school_name", "sigungu", COL_COUNSELOR, COL_STUDENT,
              COL_SPC, "counseling_staff_supply_score"]

missing_rows = []

# counselor_count 결측
for _, r in df[df[COL_COUNSELOR].isna()].iterrows():
    missing_rows.append({**{c: r[c] for c in check_cols},
                         "check_reason": "counselor_count 결측"})

# counselor_count >= 1 이지만 students_per_counselor 결측
mask_spc_miss = (df[COL_COUNSELOR] >= 1) & df[COL_SPC].isna()
for _, r in df[mask_spc_miss].iterrows():
    missing_rows.append({**{c: r[c] for c in check_cols},
                         "check_reason": "counselor_count>=1 이나 students_per_counselor 결측"})

# counseling_staff_supply_score 최종 결측
for _, r in df[df["counseling_staff_supply_score"].isna()].iterrows():
    if not any(row["school_name"] == r["school_name"] for row in missing_rows):
        missing_rows.append({**{c: r[c] for c in check_cols},
                             "check_reason": "score 결측 (기타)"})

if missing_rows:
    df_missing = pd.DataFrame(missing_rows)
    df_missing.to_csv(MISSING_CSV, index=False, encoding="utf-8-sig")
    print(f"  [주의] 확인 필요 학교 {len(missing_rows)}개교 → {MISSING_CSV.name}")
    print(df_missing[["school_name", "sigungu", COL_COUNSELOR,
                       COL_SPC, "check_reason"]].to_string(index=False))
else:
    # 결측 없음: 빈 파일 저장
    pd.DataFrame(columns=check_cols + ["check_reason"]).to_csv(
        MISSING_CSV, index=False, encoding="utf-8-sig")
    print(f"  → 확인 필요 학교 없음 (전 {total}개교 점수 확정)")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 7 : 최종 CSV 저장
# ═══════════════════════════════════════════════════════════════════════════════
df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
print(f"[INFO] 최종 파일 저장: {OUTPUT_FILE.name}")
print(f"       열 수: {len(df.columns)}  행 수: {len(df)}")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 8 : 방법론 문서 작성
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 8] 방법론 문서 작성")

# 실제 점수별 학교 수 집계 (문서 삽입용)
score_dist_lines = []
for score_val, label in score_labels.items():
    cnt = int((df["counseling_staff_supply_score"] == score_val).sum())
    pct = round(cnt / total * 100, 1)
    score_dist_lines.append(f"| {score_val} | {label.split(' — ')[1]} | {cnt}개교 ({pct}%) |")
if n_missing_score:
    score_dist_lines.append(f"| NaN | 결측 | {n_missing_score}개교 |")

doc_lines = [
    "# 상담인력 공급 점수 산출 방법론 (2025)",
    "",
    "## 목적",
    "",
    "상담공급지수(CSI)의 구성요소 중 **상담인력 공급 점수**를 산출한다.",
    "전문상담교사 배치 여부와 1인당 담당 학생 수를 기준으로",
    "학교별 상담인력 공급 여건을 0.0~1.0 사이의 점수로 수치화한다.",
    "",
    "## 사용 변수",
    "",
    "| 변수명 | 설명 | 출처 |",
    "|---|---|---|",
    "| counselor_count | 전문상담교사 수 (명) | KESS 교육통계 2025 |",
    "| student_count | 총 학생 수 (명) | KESS 교육통계 2025 |",
    "| students_per_counselor | 상담교사 1인당 학생 수 (= student_count / counselor_count) | 파생 변수 |",
    "| counseling_staff_supply_score | 상담인력 공급 점수 (0.0 ~ 1.0) | 본 단계 산출 |",
    "",
    "## 점수 구간화 기준",
    "",
    "| 점수 | 조건 | 해석 |",
    "|---|---|---|",
    "| 0.0 | counselor_count = 0 | 전문상담교사 미배치 |",
    "| 0.4 | counselor_count ≥ 1 AND students_per_counselor ≥ 500 | 배치되어 있으나 1인 담당 학생 과다 |",
    "| 0.7 | counselor_count ≥ 1 AND 250 ≤ students_per_counselor < 500 | 적정 수준 미만 |",
    "| 1.0 | counselor_count ≥ 1 AND students_per_counselor < 250 | 공급 여건 양호 |",
    "| NaN | counselor_count 결측, 또는 counselor_count ≥ 1이나 students_per_counselor 결측 | 확인 불가 |",
    "",
    "## 실제 산출 결과 분포",
    "",
    "| 점수 | 조건 요약 | 학교 수 |",
    "|---|---|---|",
] + score_dist_lines + [
    "",
    "## counselor_count 0값과 결측값 처리 기준",
    "",
    "- `counselor_count = 0` : KESS 원본에서 실제 0으로 확인 → **미배치 확정, 0.0점 부여**",
    "  - `students_per_counselor`가 NaN이어도 counselor_count=0이 명확하므로 0.0 처리",
    "- `counselor_count` 결측 : 미배치로 단정할 수 없으므로 → **NaN 유지**",
    "- `students_per_counselor` inf/-inf : 0나누기 결과로 발생 가능 → **NaN으로 변환 후 처리**",
    "",
    "## students_per_counselor 해석 방식",
    "",
    "- 값이 **낮을수록** 상담교사 1인이 담당하는 학생 수가 적음 → 공급 여건이 좋음",
    "- 값이 **높을수록** 1인당 부담이 과중 → 공급 여건이 나쁨",
    "- 따라서 점수는 students_per_counselor가 낮을수록 높게 부여한다 (역방향 구간화)",
    "",
    "## CSI 구성요소로서의 위치",
    "",
    "상담인력 공급 점수(`counseling_staff_supply_score`)는",
    "**상담공급지수(CSI)** 의 구성요소 중 하나로 사용된다.",
    "",
    "```",
    "CSI = (상담인력 공급 점수 + 실제 상담 이용 점수 + ...) / 구성요소 수",
    "```",
    "",
    "- 실제 상담 이용 점수(`counseling_use_score`)는 scripts/03에서 산출 완료",
    "- 추가 구성요소는 추후 학교폭력 관련 공시자료 등을 활용하여 확장 예정",
    "",
    "## 한계",
    "",
    "- 상담교사 **수**와 **학생 수**만 반영하며, 다음 항목은 직접 반영하지 못함:",
    "  - 상담의 질 (상담교사 역량, 수련 이력 등)",
    "  - 상담교사의 실질적 업무 부담 (행정 업무, 담임 병행 등)",
    "  - 비정규 상담 인력(계약직 상담사, 외부 연계 인력) 여부",
    "  - Wee클래스 운영 여부와의 상호작용",
    "- 구간화 기준(250명, 500명)은 선행 연구 및 정책 기준을 참고한 설정이며,",
    "  추후 근거 문헌 인용 또는 민감도 분석이 필요하다.",
    "- 분교장 및 폐교·휴교 학교는 이미 필터링되어 있으나,",
    "  특수 학급 통합 운영 학교의 학생 수 왜곡 가능성은 별도 검토가 필요하다.",
]

METHOD_MD.write_text("\n".join(doc_lines), encoding="utf-8")
print(f"[INFO] 방법론 문서 저장: {METHOD_MD.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# 최종 요약
# ═══════════════════════════════════════════════════════════════════════════════
print("=" * 55)
print("최종 요약")
print("=" * 55)
print(f"  입력 학교 수           : {total}개교")
print(f"  score = 0.0 (미배치)   : {int((df['counseling_staff_supply_score']==0.0).sum())}개교")
print(f"  score = 0.4 (500명 이상): {int((df['counseling_staff_supply_score']==0.4).sum())}개교")
print(f"  score = 0.7 (250~499명) : {int((df['counseling_staff_supply_score']==0.7).sum())}개교")
print(f"  score = 1.0 (250명 미만): {int((df['counseling_staff_supply_score']==1.0).sum())}개교")
print(f"  score = NaN (결측)     : {n_missing_score}개교")
print(f"  score 평균             : {df['counseling_staff_supply_score'].mean():.4f}")
print("=" * 55)
