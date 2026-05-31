"""
학교 위경도 병합 스크립트 (대시보드 입력 파일 업데이트)

1. 기존 지오코딩 결과 파일에서 school_latitude, school_longitude 추출
2. refined_policy_feedback_table 시트에 위경도 컬럼 추가
3. 병합 점검표 생성

입력:
  data/processed/gyeongnam_general_high_schools_geocoded_2025.csv   (우선)
  data/processed/gyeongnam_general_high_schools_with_wee_access_score_2025.csv  (보조)

출력:
  data/processed/gyeongnam_high_schools_policy_feedback_refined.xlsx  (위경도 컬럼 추가)
  outputs/tables/dashboard_geocoding_merge_check_2025.csv
"""

import pathlib
import sys
import warnings
warnings.filterwarnings("ignore")

import pandas as pd

# ── 경로 ─────────────────────────────────────────────────────────────────────
ROOT         = pathlib.Path(__file__).resolve().parent.parent
PROCESSED    = ROOT / "data" / "processed"
TABLES_DIR   = ROOT / "outputs" / "tables"
TABLES_DIR.mkdir(parents=True, exist_ok=True)

REFINED_PATH = PROCESSED / "gyeongnam_high_schools_policy_feedback_refined.xlsx"
REFINED_SHEET = "refined_policy_feedback_table"

GEO_PRIMARY  = PROCESSED / "gyeongnam_general_high_schools_geocoded_2025.csv"
GEO_FALLBACK = PROCESSED / "gyeongnam_general_high_schools_with_wee_access_score_2025.csv"

CHECK_PATH   = TABLES_DIR / "dashboard_geocoding_merge_check_2025.csv"

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 지오코딩 소스 파일 탐색
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 지오코딩 소스 파일 탐색")

geo_path = None
for candidate in [GEO_PRIMARY, GEO_FALLBACK]:
    if candidate.exists():
        geo_path = candidate
        print(f"  사용 파일: {candidate.name}")
        break

if geo_path is None:
    print("[ERROR] 지오코딩 파일을 찾을 수 없습니다.")
    sys.exit(1)

geo_raw = pd.read_csv(geo_path, encoding="utf-8-sig", dtype={"school_code": str})
print(f"  로드 완료: {len(geo_raw)}행 x {len(geo_raw.columns)}열")

# ── 위경도 컬럼 탐색 ─────────────────────────────────────────────────────────
def find_col(df, candidates):
    """후보 컬럼명 리스트에서 첫 번째로 존재하는 컬럼 반환"""
    for c in candidates:
        if c in df.columns:
            return c
    return None

lat_col = find_col(geo_raw, ["school_latitude",  "latitude",  "lat",  "위도"])
lon_col = find_col(geo_raw, ["school_longitude", "longitude", "lon", "lng", "경도"])
code_col = find_col(geo_raw, ["school_code"])
name_col = find_col(geo_raw, ["school_name"])

if lat_col is None or lon_col is None:
    print(f"[ERROR] 위경도 컬럼을 찾을 수 없습니다. 컬럼 목록: {list(geo_raw.columns)}")
    sys.exit(1)

print(f"  위도 컬럼: {lat_col}")
print(f"  경도 컬럼: {lon_col}")

# 좌표 유효성 확인
geo_valid = geo_raw[[c for c in [code_col, name_col, lat_col, lon_col] if c]].copy()
geo_valid = geo_valid.rename(columns={lat_col: "school_latitude", lon_col: "school_longitude"})
if code_col and code_col != "school_code":
    geo_valid = geo_valid.rename(columns={code_col: "school_code"})
if name_col and name_col != "school_name":
    geo_valid = geo_valid.rename(columns={name_col: "school_name"})

n_valid = geo_valid[["school_latitude","school_longitude"]].notna().all(axis=1).sum()
print(f"  유효 좌표 보유: {n_valid}개교 / 전체 {len(geo_valid)}개교")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. refined 파일 로드 (전체 시트 보존용)
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 2] refined 파일 로드")

if not REFINED_PATH.exists():
    print(f"[ERROR] 파일 없음: {REFINED_PATH}")
    sys.exit(1)

xl = pd.ExcelFile(REFINED_PATH)
print(f"  시트 목록: {xl.sheet_names}")

# 전체 시트 읽기
all_sheets = {}
for sheet in xl.sheet_names:
    dtype_arg = {"school_code": str} if sheet == REFINED_SHEET else {}
    all_sheets[sheet] = pd.read_excel(REFINED_PATH, sheet_name=sheet, dtype=dtype_arg)
    print(f"  [{sheet}] {len(all_sheets[sheet])}행 x {len(all_sheets[sheet].columns)}열")

df = all_sheets[REFINED_SHEET]
n_original = len(df)
print(f"\n  기준 시트({REFINED_SHEET}) 원본 행 수: {n_original}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. 위경도 병합
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 3] 위경도 병합")

# 병합 전 기존 위경도 컬럼 제거 (재실행 시 중복 방지)
for col in ["school_latitude", "school_longitude"]:
    if col in df.columns:
        df = df.drop(columns=[col])
        print(f"  기존 {col} 컬럼 제거 (재실행 중복 방지)")

# ── 3-1. school_code 기준 병합 ───────────────────────────────────────────────
n_code_matched = 0
n_name_matched = 0
failed_list    = []

if "school_code" in df.columns and "school_code" in geo_valid.columns:
    geo_by_code = geo_valid[["school_code","school_latitude","school_longitude"]].drop_duplicates("school_code")
    df = df.merge(geo_by_code, on="school_code", how="left")
    n_code_matched = df[["school_latitude","school_longitude"]].notna().all(axis=1).sum()
    print(f"  school_code 기준 병합 성공: {n_code_matched}개교")
else:
    print("  [WARN] school_code 컬럼 없음 → school_name 단독 병합으로 진행")
    df["school_latitude"]  = None
    df["school_longitude"] = None

# ── 3-2. school_name 보조 병합 (code 미매칭 행) ──────────────────────────────
if "school_name" in df.columns and "school_name" in geo_valid.columns:
    unmatched_mask = df[["school_latitude","school_longitude"]].isna().any(axis=1)
    n_unmatched = unmatched_mask.sum()

    if n_unmatched > 0:
        print(f"  school_code 미매칭 {n_unmatched}개교 → school_name 보조 병합 시도")
        geo_by_name = geo_valid[["school_name","school_latitude","school_longitude"]].drop_duplicates("school_name")
        geo_by_name = geo_by_name.rename(columns={
            "school_latitude":  "_lat_sup",
            "school_longitude": "_lon_sup",
        })
        df = df.merge(geo_by_name, on="school_name", how="left")

        # 미매칭 행에만 보조값 채우기
        lat_fill = df["school_latitude"].isna() & df["_lat_sup"].notna()
        lon_fill = df["school_longitude"].isna() & df["_lon_sup"].notna()
        df.loc[lat_fill, "school_latitude"]  = df.loc[lat_fill, "_lat_sup"]
        df.loc[lon_fill, "school_longitude"] = df.loc[lon_fill, "_lon_sup"]

        n_name_matched = int(lat_fill.sum())
        print(f"  school_name 보조 병합 성공: {n_name_matched}개교")

        if n_name_matched > 0:
            name_matched_schools = df.loc[lat_fill, "school_name"].tolist()
            print(f"    해당 학교: {name_matched_schools}")

        df = df.drop(columns=["_lat_sup", "_lon_sup"])

# ── 3-3. 병합 최종 확인 ──────────────────────────────────────────────────────
n_matched = df[["school_latitude","school_longitude"]].notna().all(axis=1).sum()
n_missing = n_original - n_matched
failed_list = df.loc[
    df[["school_latitude","school_longitude"]].isna().any(axis=1),
    "school_name"
].tolist() if "school_name" in df.columns else []

assert len(df) == n_original, f"행 수 변동! 원본: {n_original}, 병합 후: {len(df)}"
print(f"\n  병합 결과:")
print(f"    위경도 병합 성공: {n_matched}개교 ({n_matched/n_original*100:.1f}%)")
print(f"    위경도 결측:     {n_missing}개교")
if failed_list:
    print(f"    결측 학교: {failed_list}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. 병합 점검표 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 4] 병합 점검표 생성")

check_rows = [
    {"항목": "전체 학교 수",              "값": n_original},
    {"항목": "위경도 병합 성공 학교 수",  "값": n_matched},
    {"항목": "위경도 결측 학교 수",       "값": n_missing},
    {"항목": "school_code 기준 병합 수",  "값": n_code_matched},
    {"항목": "school_name 보조 병합 수",  "값": n_name_matched},
    {"항목": "병합 실패 학교 목록",        "값": "; ".join(failed_list) if failed_list else "(없음)"},
    {"항목": "사용 지오코딩 소스 파일",   "값": geo_path.name},
]
check_df = pd.DataFrame(check_rows)
check_df.to_csv(CHECK_PATH, index=False, encoding="utf-8-sig")
print(f"  저장: {CHECK_PATH.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. refined 엑셀 저장 (전체 시트 보존)
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 5] refined 엑셀 저장")

# 업데이트된 시트 교체
all_sheets[REFINED_SHEET] = df

with pd.ExcelWriter(REFINED_PATH, engine="openpyxl") as writer:
    for sheet_name, sheet_df in all_sheets.items():
        idx = (sheet_name == REFINED_SHEET)   # 기준 시트만 school_code str 유지
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"  저장 완료: {REFINED_PATH.name}")
print(f"  [{REFINED_SHEET}] {len(df)}행 x {len(df.columns)}열")
print(f"  추가 컬럼: school_latitude, school_longitude")
print(f"  전체 시트 보존: {list(all_sheets.keys())}")

# ════════════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("위경도 병합 완료")
print("=" * 60)
print(f"  위경도 병합 성공: {n_matched}/{n_original}개교")
print(f"  결과 파일: {REFINED_PATH.name}")
print(f"  점검표:   {CHECK_PATH.name}")
