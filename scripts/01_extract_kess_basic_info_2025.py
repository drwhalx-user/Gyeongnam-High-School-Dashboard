"""
KESS 2025 학교별 주요 통계에서 경남 일반고 기본 정보를 추출한다.

원본 파일 : data/raw/kess_school_statistics_2025.xlsx
  - 시트   : 학교별 주요 통계
  - 실제 헤더 행 : 12번째 행(0-indexed) → skiprows=12, header=0
  - 조사기준일   : 2025-10-01 / 추출일 : 2026-02-06

출력 파일:
  - data/processed/gyeongnam_general_high_schools_basic_2025.csv
  - outputs/tables/kess_basic_info_missing_summary_2025.csv
  - docs/kess_basic_info_variable_mapping_2025.md
"""

import sys
import pathlib
import pandas as pd

# ── 경로 설정 ────────────────────────────────────────────────────────────────
ROOT = pathlib.Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw"
PROCESSED_DIR = ROOT / "data" / "processed"
TABLES_DIR = ROOT / "outputs" / "tables"
DOCS_DIR = ROOT / "docs"

# 필요한 출력 폴더를 없으면 생성
for d in [PROCESSED_DIR, TABLES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── 원본 파일 탐색 ────────────────────────────────────────────────────────────
# raw 폴더 안의 .xlsx/.xls 파일 중 'kess_school_statistics_2025' 패턴 검색
candidates = list(RAW_DIR.glob("*kess*school*statistics*2025*.xlsx")) + \
             list(RAW_DIR.glob("*kess*school*statistics*2025*.xls"))

if not candidates:
    print("[ERROR] data/raw/ 폴더에서 kess_school_statistics_2025 파일을 찾지 못했습니다.")
    print("        확인 필요 경로:", RAW_DIR)
    sys.exit(1)

RAW_FILE = candidates[0]
print(f"[INFO] 원본 파일 사용: {RAW_FILE.name}")

# ── 원본 데이터 로드 ──────────────────────────────────────────────────────────
# 실제 열 이름은 13번째 행(0-indexed row 12)에 위치 → skiprows=12
SHEET_NAME = "학교별 주요 통계"
print(f"[INFO] 시트 '{SHEET_NAME}' 로드 중...")
df_raw = pd.read_excel(RAW_FILE, sheet_name=SHEET_NAME, skiprows=12, header=0)
print(f"[INFO] 원본 데이터 로드 완료 — 전체 {len(df_raw):,}행 × {len(df_raw.columns)}열")

# ── 필수 열 존재 여부 검증 ────────────────────────────────────────────────────
# 실제 원본 열 이름 (개행 문자 포함)
REQUIRED_COLS = {
    "KEDI학교코드"        : "school_code",
    "학교명"              : "school_name",
    "시도"                : "sido",
    "행정구"              : "sigungu",
    "교육\n(지원)청"      : "education_office",
    "고등학교\n유형"      : "school_type",
    "우편번호"            : "postcode",
    "주소"                : "address",
    "학생수_총계_계"      : "student_count",
    "교원수_총계_계"      : "teacher_count",
    "교원수_정규_상담_계" : "counselor_count",
}

# 필터링에 필요한 열
FILTER_COLS = ["학교급", "상태", "본분교", "학교\n세부유형"]

missing_cols = [c for c in list(REQUIRED_COLS.keys()) + FILTER_COLS
                if c not in df_raw.columns]
if missing_cols:
    print("[ERROR] 아래 열을 원본 파일에서 찾지 못했습니다. 열 이름을 직접 확인하세요.")
    for c in missing_cols:
        print(f"  - {repr(c)}")
    print("\n[INFO] 원본 파일의 전체 열 목록:")
    for i, col in enumerate(df_raw.columns):
        print(f"  [{i:3d}] {repr(col)}")
    sys.exit(1)

print("[INFO] 필수 열 검증 통과")

# ── 필터링 ────────────────────────────────────────────────────────────────────
# 1) 경상남도 소재 학교 (원본에서 '경남'으로 표기)
mask_sido = df_raw["시도"] == "경남"

# 2) 고등학교 학교급
mask_schoollevel = df_raw["학교급"] == "고등학교"

# 3) 일반고 유형만 유지 (특성화고·마이스터고·특목고·자율고 제외)
mask_type = df_raw["고등학교\n유형"] == "일반고"

# 4) 폐교·휴교 제외 (상태 열 기준)
mask_active = ~df_raw["상태"].isin(["폐(원)교", "휴(원)교"])

# 5) 분교장 제외 (본교만 유지)
mask_main = df_raw["본분교"] == "본교"

# 6) 각종학교 제외 (세부유형 기준 이중 확인)
mask_not_misc = df_raw["학교\n세부유형"] != "각종학교(고교)"

df_filtered = df_raw[mask_sido & mask_schoollevel & mask_type &
                     mask_active & mask_main & mask_not_misc].copy()

print(f"[INFO] 필터링 후 경남 일반고(활성·본교): {len(df_filtered)}개교")

if len(df_filtered) == 0:
    print("[ERROR] 필터링 결과가 0건입니다. 필터 조건을 재확인하세요.")
    sys.exit(1)

# ── 열 선택 및 표준 변수명으로 변경 ──────────────────────────────────────────
df_selected = df_filtered[list(REQUIRED_COLS.keys())].rename(columns=REQUIRED_COLS)

# ── 전문상담교사 처리 ─────────────────────────────────────────────────────────
# 원본 데이터에서 0값과 NaN을 구분하여 플래그 생성
# (현재 원본 기준: NaN 없음, 0값 다수 존재)
df_selected["counselor_is_zero"] = df_selected["counselor_count"] == 0
df_selected["counselor_is_missing"] = df_selected["counselor_count"].isna()

# students_per_counselor: 상담교사 수가 0이거나 결측이면 NaN 처리
# (무한대·0나누기 방지)
df_selected["students_per_counselor"] = df_selected.apply(
    lambda row: (
        row["student_count"] / row["counselor_count"]
        if pd.notna(row["counselor_count"])
           and row["counselor_count"] > 0
           and pd.notna(row["student_count"])
           and row["student_count"] > 0
        else float("nan")
    ),
    axis=1,
)

# ── 결측값 요약표 생성 ────────────────────────────────────────────────────────
# 표준 변수명 열 대상으로 결측 현황 집계
standard_cols = list(REQUIRED_COLS.values()) + ["students_per_counselor"]
missing_summary = pd.DataFrame({
    "variable"      : standard_cols,
    "total"         : len(df_selected),
    "missing_count" : [df_selected[c].isna().sum() for c in standard_cols],
    "zero_count"    : [
        int((df_selected[c] == 0).sum()) if pd.api.types.is_numeric_dtype(df_selected[c]) else 0
        for c in standard_cols
    ],
})
missing_summary["missing_rate(%)"] = (
    missing_summary["missing_count"] / missing_summary["total"] * 100
).round(2)

missing_summary_path = TABLES_DIR / "kess_basic_info_missing_summary_2025.csv"
missing_summary.to_csv(missing_summary_path, index=False, encoding="utf-8-sig")
print(f"[INFO] 결측값 요약표 저장 완료: {missing_summary_path}")

# ── 변수 매핑표 저장 (Markdown) ───────────────────────────────────────────────
mapping_lines = [
    "# KESS 기본 정보 변수 매핑표 (2025)",
    "",
    f"- 원본 파일: `{RAW_FILE.name}`",
    f"- 시트: `{SHEET_NAME}`",
    "- 조사기준일: 2025-10-01",
    "",
    "| 표준 변수명 | 원본 열 이름 | 설명 |",
    "|---|---|---|",
    "| school_code | KEDI학교코드 | KEDI 부여 고유 학교 식별 코드 |",
    "| school_name | 학교명 | 학교 공식 명칭 |",
    "| sido | 시도 | 광역 시·도 (경남) |",
    "| sigungu | 행정구 | 시·군·구 단위 행정구역 |",
    "| education_office | 교육(지원)청 | 소속 교육(지원)청 |",
    "| school_type | 고등학교 유형 | 일반고/특성화고/특목고/자율고 구분 |",
    "| postcode | 우편번호 | 5자리 우편번호 |",
    "| address | 주소 | 도로명 주소 |",
    "| student_count | 학생수_총계_계 | 전체 재학생 수(명) |",
    "| teacher_count | 교원수_총계_계 | 전체 교원 수(정규+기간제, 명) |",
    "| counselor_count | 교원수_정규_상담_계 | 정규직 전문상담교사 수(명) |",
    "| students_per_counselor | (파생) | student_count / counselor_count; 상담교사 0명 또는 결측이면 NaN |",
    "",
    "## 필터링 조건",
    "",
    "| 조건 | 원본 열 | 값 |",
    "|---|---|---|",
    "| 경상남도 | 시도 | 경남 |",
    "| 고등학교 | 학교급 | 고등학교 |",
    "| 일반고 | 고등학교 유형 | 일반고 |",
    "| 활성 학교 | 상태 | 폐(원)교·휴(원)교 제외 |",
    "| 본교 | 본분교 | 본교 |",
    "| 각종학교 제외 | 학교 세부유형 | 각종학교(고교) 제외 |",
    "",
    "## 상담교사 0값 처리 방침",
    "",
    "- `counselor_count == 0`: 조사기준일 기준 전문상담교사 미배치 (실제 0)",
    "- `counselor_count` 결측: 미응답 또는 해당 없음",
    "- `students_per_counselor`: 위 두 경우 모두 `NaN`으로 저장 (무한대·0나누기 방지)",
]

mapping_path = DOCS_DIR / "kess_basic_info_variable_mapping_2025.md"
mapping_path.write_text("\n".join(mapping_lines), encoding="utf-8")
print(f"[INFO] 변수 매핑표 저장 완료: {mapping_path}")

# ── 최종 데이터셋 저장 ────────────────────────────────────────────────────────
output_cols = [
    "school_code", "school_name", "sido", "sigungu", "education_office",
    "school_type", "postcode", "address", "student_count", "teacher_count",
    "counselor_count", "students_per_counselor",
]
df_output = df_selected[output_cols].sort_values("sigungu").reset_index(drop=True)

output_path = PROCESSED_DIR / "gyeongnam_general_high_schools_basic_2025.csv"
df_output.to_csv(output_path, index=False, encoding="utf-8-sig")
print(f"[INFO] 최종 데이터셋 저장 완료: {output_path}")

# ── 요약 출력 ─────────────────────────────────────────────────────────────────
print()
print("=" * 55)
print("추출 결과 요약")
print("=" * 55)
print(f"  학교 수          : {len(df_output)}개교")
print(f"  학생 수 합계     : {df_output['student_count'].sum():,.0f}명")
print(f"  교원 수 합계     : {df_output['teacher_count'].sum():,.0f}명")
print(f"  상담교사 0명 학교: {(df_output['counselor_count'] == 0).sum()}개교")
print(f"  상담교사 결측    : {df_output['counselor_count'].isna().sum()}개교")
print(f"  상담교사 배치 학교: {(df_output['counselor_count'] > 0).sum()}개교")
print("=" * 55)
