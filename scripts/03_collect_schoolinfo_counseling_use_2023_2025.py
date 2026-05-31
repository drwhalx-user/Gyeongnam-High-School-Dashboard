"""
학교알리미 API (apiType=61) 2024·2025·2026 공시연도 순환 호출
→ 경남 일반고별 상담 건수 수집 및 실제 상담 이용 점수(counseling_use_score) 산출

공시연도(pbanYr) → 실제 데이터 시점(data_year) 매핑:
  pbanYr=2024  →  data_year=2023
  pbanYr=2025  →  data_year=2024
  pbanYr=2026  →  data_year=2025

파일명·변수명은 모두 실제 데이터 시점 기준(2023~2025)으로 저장한다.
"""

import sys, time, json, re, pathlib
import requests
import pandas as pd

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT = pathlib.Path(__file__).resolve().parent.parent

SGG_FILE   = ROOT / "시도시군구코드.xlsx"
KESS_FILE  = ROOT / "data" / "processed" / "gyeongnam_general_high_schools_with_wee_class_2025.csv"
RAW_JSON   = ROOT / "data" / "raw"       / "schoolinfo_counseling_2023_2025_all_sigungu.json"
CLEAN_CSV  = ROOT / "data" / "processed" / "schoolinfo_counseling_2023_2025_clean.csv"
MERGED_CSV = ROOT / "data" / "processed" / "gyeongnam_general_high_schools_with_counseling_use_2025.csv"
CALL_LOG   = ROOT / "outputs" / "tables" / "counseling_api_call_log_2023_2025.csv"
MISS_CSV   = ROOT / "outputs" / "tables" / "counseling_missing_summary_2023_2025.csv"
SCORE_CSV  = ROOT / "outputs" / "tables" / "counseling_use_score_summary_2023_2025.csv"
MAP_MD     = ROOT / "docs"               / "counseling_use_variable_mapping_2023_2025.md"

for _d in [RAW_JSON.parent, CLEAN_CSV.parent, CALL_LOG.parent, MAP_MD.parent]:
    _d.mkdir(parents=True, exist_ok=True)

# ── API 설정 ──────────────────────────────────────────────────────────────────
API_URL = "http://www.schoolinfo.go.kr/openApi.do"
API_KEY = "39bf3d7a293447baba1702f71307e5ad"

# 공시연도(API) → 실제 데이터 시점 매핑
PBAN_TO_DATA = {2024: 2023, 2025: 2024, 2026: 2025}

# 실제 상담 건수 필드 (API 검증 완료)
TEACHER_FIELD  = "COSE_CNSL_TLGM_TCR_FGR"    # 전담교사 상담 건수
EXTERNAL_FIELD = "COSE_CNSL_EXTRL_SPLST_FGR"  # 외부전문가 상담 건수


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 : 시군구 코드 로드
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 시도시군구코드.xlsx 로드")

df_sgg = pd.read_excel(SGG_FILE, skiprows=2, header=0)
df_sgg.columns = ["시도명", "시도코드", "시군구명", "시군구코드"]

df_gn = df_sgg[df_sgg["시도명"] == "경상남도"].copy()
df_gn["시군구코드_str"] = df_gn["시군구코드"].astype(str).str.zfill(5)

# 창원시 시 단위(48120)는 API 데이터 없음 → 제외 (구 단위 5개 코드 사용)
df_gn = df_gn[df_gn["시군구코드_str"] != "48120"]
sgg_list = sorted(zip(df_gn["시군구코드_str"], df_gn["시군구명"]))

print(f"  → 호출 대상: {len(sgg_list)}개 시군구\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 : API 순환 호출 (공시연도 2024·2025·2026)
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 2] API 순환 호출 시작")

all_raw     = {}   # {str(pban_yr): {sgg_code: 원본 응답}} → JSON 저장용
all_records = []   # 수집된 전체 레코드
call_log    = []   # 시군구×연도 호출 로그

for pban_yr, data_yr in PBAN_TO_DATA.items():
    print(f"\n  ── 공시연도 {pban_yr}  (데이터 시점 {data_yr}) ──")
    all_raw[str(pban_yr)] = {}

    for sgg_code, sgg_name in sgg_list:
        params = {
            "apiKey": API_KEY, "apiType": "61",
            "pbanYr": str(pban_yr), "sidoCode": "48",
            "schulKndCode": "04", "sggCode": sgg_code,
        }
        try:
            resp = requests.get(API_URL, params=params, timeout=20)
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            err = str(exc)[:120]
            print(f"    [ERR]  {sgg_name}({sgg_code}): {err}")
            all_raw[str(pban_yr)][sgg_code] = {"error": err}
            call_log.append({
                "pban_year": pban_yr, "data_year": data_yr,
                "시군구명": sgg_name, "시군구코드": sgg_code,
                "resultCode": "error", "resultMsg": err, "응답건수": 0,
            })
            time.sleep(0.5)
            continue

        rc    = data.get("resultCode", "")
        msg   = data.get("resultMsg", "")
        items = data.get("list", [])
        n     = len(items)
        all_raw[str(pban_yr)][sgg_code] = data

        if rc == "success":
            for item in items:
                # 출처 메타데이터 추가
                item["_pban_year"] = pban_yr
                item["_data_year"] = data_yr
                item["_sgg_code"]  = sgg_code
                item["_sgg_name"]  = sgg_name
            all_records.extend(items)
            print(f"    [OK]   {sgg_name}({sgg_code}): {n}건")
        else:
            print(f"    [FAIL] {sgg_name}({sgg_code}): {msg[:60]}")

        call_log.append({
            "pban_year": pban_yr, "data_year": data_yr,
            "시군구명": sgg_name, "시군구코드": sgg_code,
            "resultCode": rc, "resultMsg": msg,
            "응답건수": n if rc == "success" else 0,
        })
        time.sleep(0.3)  # 서버 부하 방지

print(f"\n  → 총 수집 레코드: {len(all_records)}건")

# 원본 JSON 저장
with open(RAW_JSON, "w", encoding="utf-8") as f:
    json.dump(all_raw, f, ensure_ascii=False, indent=2)
print(f"[INFO] 원본 JSON 저장: {RAW_JSON.name}")

# 호출 로그 저장
pd.DataFrame(call_log).to_csv(CALL_LOG, index=False, encoding="utf-8-sig")
print(f"[INFO] 호출 로그 저장: {CALL_LOG.name}\n")

if not all_records:
    print("[ERROR] 수집된 레코드가 없습니다. 호출 로그를 확인하세요.")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 : 상담 건수 필드 탐색 및 검증
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 3] 상담 건수 필드 탐색")

sample = next((r for r in all_records if r.get("HS_KND_SC_NM") == "일반고등학교"), all_records[0])

cnsl_keywords = {"cnsl", "tlgm", "extrl", "splst", "상담", "건수", "실시", "횟수", "stud", "parn"}
candidate_fields = [k for k in sample if not k.startswith("_") and
                    any(kw in k.lower() for kw in cnsl_keywords)]
print(f"  후보 필드: {candidate_fields}")

for fld in [TEACHER_FIELD, EXTERNAL_FIELD]:
    if fld in sample:
        print(f"  → '{fld}' 확인: 예시값={sample[fld]!r}")
    else:
        print(f"  [ERROR] '{fld}' 필드를 찾을 수 없습니다. 전체 필드:")
        for k, v in sample.items():
            if not k.startswith("_"):
                print(f"    {k!r}: {v!r}")
        sys.exit(1)
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 : DataFrame 변환 / 학교 단위 분리 / 일반고 필터링
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 4] 데이터 변환 및 일반고 필터링")

df_api = pd.DataFrame(all_records)

# 학교명 있는 행 = 학교 단위
mask_school = df_api["SCHUL_NM"].notna() & df_api["SCHUL_NM"].str.strip().ne("")
df_non_school = df_api[~mask_school]
if len(df_non_school):
    print(f"  [주의] 학교명 없는 행(시군 단위 추정) {len(df_non_school)}건 → 병합 제외")

df_school = df_api[mask_school].copy()

# 일반고등학교만 유지
df_general = df_school[df_school["HS_KND_SC_NM"] == "일반고등학교"].copy()
print(f"  → 일반고 레코드: {len(df_general)}건  "
      f"({df_general['_data_year'].unique().tolist()}년도)\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 : 상담 건수 변수 생성 및 숫자 변환
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 5] 상담 건수 변수 생성")

def safe_numeric(series: pd.Series) -> pd.Series:
    """쉼표·공백·하이픈 제거 후 숫자 변환; 변환 불가 → NaN"""
    cleaned = (series.astype(str)
               .str.replace(r"[,\s]", "", regex=True)
               .str.strip()
               .replace({"": None, "-": None, "nan": None, "None": None}))
    return pd.to_numeric(cleaned, errors="coerce")

df_general["teacher_counseling_count"]  = safe_numeric(df_general[TEACHER_FIELD])
df_general["external_counseling_count"] = safe_numeric(df_general[EXTERNAL_FIELD])

# 통합 상담 건수: 두 값 합산 (한쪽이 결측이면 나머지만, 둘 다 결측이면 NaN)
df_general["total_counseling_count"] = (
    df_general["teacher_counseling_count"].fillna(0) +
    df_general["external_counseling_count"].fillna(0)
)
both_missing = (df_general["teacher_counseling_count"].isna() &
                df_general["external_counseling_count"].isna())
df_general.loc[both_missing, "total_counseling_count"] = float("nan")

print("  원본 필드 분포:")
for yr in sorted(df_general["_data_year"].unique()):
    sub = df_general[df_general["_data_year"] == yr]
    print(f"    data_year={yr}: "
          f"전담 평균={sub['teacher_counseling_count'].mean():.1f}  "
          f"외부 평균={sub['external_counseling_count'].mean():.1f}")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 : 정제 데이터 저장 (학교×연도 long format)
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 6] 정제 데이터 저장 (clean CSV)")

rename_map = {
    "SCHUL_CODE": "school_code_api",
    "SCHUL_NM"  : "school_name",
    "_pban_year": "pban_year",
    "_data_year": "data_year",
    "_sgg_code" : "sgg_code",
    "_sgg_name" : "sigungu",
}
df_clean = df_general.rename(columns={k: v for k, v in rename_map.items()
                                       if k in df_general.columns}).copy()
df_clean["sido_code"] = "48"

clean_cols = ["school_code_api", "school_name", "sido_code", "sgg_code", "sigungu",
              "pban_year", "data_year",
              "teacher_counseling_count", "external_counseling_count", "total_counseling_count"]
df_clean[clean_cols].to_csv(CLEAN_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 정제 데이터 저장: {CLEAN_CSV.name}  ({len(df_clean)}행)\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 7 : 3개년 평균 산출 (data_year 기준: 2023·2024·2025)
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 7] 3개년 평균 산출")

df_agg = (
    df_clean
    .groupby("school_name", as_index=False)
    .agg(
        avg_teacher_counseling_count_3yr  = ("teacher_counseling_count",  "mean"),
        avg_external_counseling_count_3yr = ("external_counseling_count", "mean"),
        avg_total_counseling_count_3yr    = ("total_counseling_count",    "mean"),
    )
)

# counseling_years_available: total_counseling_count 비결측 연도 수
yr_valid = (
    df_clean[df_clean["total_counseling_count"].notna()]
    .groupby("school_name")["data_year"]
    .nunique()
    .reset_index()
    .rename(columns={"data_year": "counseling_years_available"})
)
df_agg = df_agg.merge(yr_valid, on="school_name", how="left")
df_agg["counseling_years_available"] = df_agg["counseling_years_available"].fillna(0).astype(int)

print(f"  집계 학교 수: {len(df_agg)}개교")
print(f"  연도 보유 분포: {df_agg['counseling_years_available'].value_counts().sort_index().to_dict()}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 8 : KESS 기본정보와 병합
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 8] KESS+Wee클래스 기본정보 파일 병합")

df_kess = pd.read_csv(KESS_FILE, encoding="utf-8-sig")
n_kess  = len(df_kess)
print(f"  KESS 학교 수: {n_kess}개교")

def normalize_name(name: str) -> str:
    """학교명 정규화: 공백·괄호 내용·특수문자 제거 후 소문자 변환 (원본 보존)"""
    s = str(name).strip()
    s = re.sub(r"[\s ]+", "", s)
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^\w가-힣]", "", s)
    return s.lower()

df_kess["_key"] = df_kess["school_name"].apply(normalize_name)
df_agg["_key"]  = df_agg["school_name"].apply(normalize_name)

# 중복 키 경고 처리
dup = df_agg[df_agg.duplicated("_key", keep=False)]
if len(dup):
    print(f"  [경고] API 집계 중복 키 {len(dup)}건: {dup['school_name'].tolist()}")

df_agg_dedup = df_agg.drop_duplicates("_key", keep="first")

merge_cols = ["_key",
              "avg_teacher_counseling_count_3yr",
              "avg_external_counseling_count_3yr",
              "avg_total_counseling_count_3yr",
              "counseling_years_available"]
df_merged = df_kess.merge(df_agg_dedup[merge_cols], on="_key", how="left").drop(columns=["_key"])

# 행 수 유지 확인
assert len(df_merged) == n_kess, f"[ERROR] 병합 후 행 수 불일치: {len(df_merged)} ≠ {n_kess}"

matched   = df_merged["avg_total_counseling_count_3yr"].notna().sum()
unmatched = df_merged["avg_total_counseling_count_3yr"].isna().sum()
print(f"  → 매칭 성공: {matched}개교 / 미매칭: {unmatched}개교")

if unmatched:
    print("  [검증] 미매칭 학교:")
    for _, r in df_merged[df_merged["avg_total_counseling_count_3yr"].isna()].iterrows():
        print(f"    {r['sigungu']} / {r['school_name']}")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 9 : 학생 수 대비 통합 상담 건수
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 9] 학생 수 대비 통합 상담 건수 산출")

if "student_count" not in df_merged.columns:
    candidates = [c for c in df_merged.columns if "학생" in c or "student" in c.lower()]
    print(f"  [ERROR] 'student_count' 없음. 후보 변수: {candidates} — 확인 필요")
    sys.exit(1)

df_merged["counseling_count_per_student"] = df_merged.apply(
    lambda r: (
        r["avg_total_counseling_count_3yr"] / r["student_count"]
        if pd.notna(r["avg_total_counseling_count_3yr"])
           and pd.notna(r["student_count"])
           and r["student_count"] > 0
        else float("nan")
    ),
    axis=1,
)
print(f"  → 산출 완료 (결측: {df_merged['counseling_count_per_student'].isna().sum()}개교)\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 10 : Min-Max 정규화 및 counseling_use_score 산출
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 10] Min-Max 정규화 및 counseling_use_score 산출")

def minmax_normalize(series: pd.Series) -> pd.Series:
    """결측 제외 후 Min-Max 정규화; max==min이면 비결측값을 0으로 처리"""
    mn, mx = series.min(), series.max()
    if pd.isna(mn) or pd.isna(mx) or mn == mx:
        return series.apply(lambda v: 0.0 if pd.notna(v) else float("nan"))
    return (series - mn) / (mx - mn)

df_merged["norm_avg_total_counseling_count_3yr"] = minmax_normalize(
    df_merged["avg_total_counseling_count_3yr"])
df_merged["norm_counseling_count_per_student"]   = minmax_normalize(
    df_merged["counseling_count_per_student"])

# counseling_use_score: 두 norm 값의 평균 (결측 제외 후 산출)
def calc_score(row):
    vals = [row["norm_avg_total_counseling_count_3yr"],
            row["norm_counseling_count_per_student"]]
    valid = [v for v in vals if pd.notna(v)]
    return (sum(valid) / len(valid), len(valid)) if valid else (float("nan"), 0)

results = df_merged.apply(calc_score, axis=1)
df_merged["counseling_use_score"]                = results.apply(lambda x: x[0])
df_merged["counseling_use_components_available"] = results.apply(lambda x: x[1])

print("  counseling_use_score 기술통계:")
print(df_merged["counseling_use_score"].describe().round(4).to_string())
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 11 : 이상값 점검
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 11] 이상값 점검")

# 전담·외부 상담 건수 평균이 모두 0인 학교
zero_both = df_merged[
    (df_merged["avg_teacher_counseling_count_3yr"].fillna(-1) == 0) &
    (df_merged["avg_external_counseling_count_3yr"].fillna(-1) == 0)
]
if len(zero_both):
    print(f"  [주의] 전담·외부 상담 건수 3개년 평균 모두 0인 학교: {len(zero_both)}개교")
    for nm in zero_both["school_name"].tolist():
        print(f"    - {nm}")

# 통합 상담 건수 상위 5개교
top5 = df_merged.nlargest(5, "avg_total_counseling_count_3yr")[
    ["school_name", "sigungu", "avg_total_counseling_count_3yr",
     "counseling_count_per_student", "counseling_use_score"]]
print(f"\n  통합 상담 건수 상위 5개교:")
print(top5.to_string(index=False))

# counseling_count_per_student 상위 5개교
top5_per = df_merged.nlargest(5, "counseling_count_per_student")[
    ["school_name", "sigungu", "student_count",
     "avg_total_counseling_count_3yr", "counseling_count_per_student"]]
print(f"\n  학생 수 대비 상담 건수 상위 5개교:")
print(top5_per.to_string(index=False))
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 12 : 최종 CSV 저장
# ═══════════════════════════════════════════════════════════════════════════════
df_merged.to_csv(MERGED_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 병합 최종 파일 저장: {MERGED_CSV.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 13 : 결측 요약표
# ═══════════════════════════════════════════════════════════════════════════════
new_cols = [
    "avg_teacher_counseling_count_3yr", "avg_external_counseling_count_3yr",
    "avg_total_counseling_count_3yr", "counseling_years_available",
    "counseling_count_per_student",
    "norm_avg_total_counseling_count_3yr", "norm_counseling_count_per_student",
    "counseling_use_score", "counseling_use_components_available",
]
miss_rows = []
for col in new_cols:
    if col not in df_merged.columns:
        continue
    tot   = len(df_merged)
    nmiss = int(df_merged[col].isna().sum())
    nzero = int((df_merged[col] == 0).sum()) if pd.api.types.is_numeric_dtype(df_merged[col]) else 0
    miss_rows.append({
        "variable"       : col,
        "total"          : tot,
        "missing_count"  : nmiss,
        "zero_count"     : nzero,
        "missing_rate(%)": round(nmiss / tot * 100, 2),
    })
pd.DataFrame(miss_rows).to_csv(MISS_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 결측 요약표 저장: {MISS_CSV.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 14 : 점수 분포 요약표
# ═══════════════════════════════════════════════════════════════════════════════
score_cols = [
    "school_name", "sigungu",
    "avg_teacher_counseling_count_3yr", "avg_external_counseling_count_3yr",
    "avg_total_counseling_count_3yr", "counseling_years_available",
    "counseling_count_per_student",
    "norm_avg_total_counseling_count_3yr", "norm_counseling_count_per_student",
    "counseling_use_score", "counseling_use_components_available",
]
df_merged[score_cols].to_csv(SCORE_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 점수 요약표 저장: {SCORE_CSV.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 15 : 변수 매핑 문서 작성
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 15] 변수 매핑 문서 작성")

all_api_keys = sorted({k for r in all_records for k in r if not k.startswith("_")})

def _ex(field):
    for r in all_records:
        v = r.get(field)
        if v is not None:
            return str(v)[:40]
    return ""

doc_lines = [
    "# 상담 이용 변수 매핑표 (데이터 시점 2023~2025)",
    "",
    "## API 호출 정보",
    "",
    "| 파라미터 | 값 | 설명 |",
    "|---|---|---|",
    "| apiType | 61 | 학생·학부모 상담계획 및 실시 현황 |",
    "| pbanYr | 2024, 2025, 2026 | 공시연도 (실제 데이터 시점: 2023, 2024, 2025) |",
    "| sidoCode | 48 | 경상남도 |",
    "| sggCode | 시군구별 코드 | 시도시군구코드.xlsx 참조 |",
    "| schulKndCode | 04 | 고등학교 |",
    f"| 요청 URL | `{API_URL}?apiKey={{KEY}}&apiType=61&pbanYr={{PBAN}}&sidoCode=48&sggCode={{SGG}}&schulKndCode=04` | |",
    "",
    "## 공시연도 → 실제 데이터 시점 매핑",
    "",
    "| pbanYr (API 파라미터) | data_year (실제 데이터 시점) |",
    "|---|---|",
    "| 2024 | 2023 |",
    "| 2025 | 2024 |",
    "| 2026 | 2025 |",
    "",
    "> 파일명·변수명은 모두 실제 데이터 시점(2023~2025) 기준으로 표기한다.",
    "",
    "## 상담 건수 원본 필드",
    "",
    "| 표준 변수명 | 원본 API 필드 | 설명 |",
    "|---|---|---|",
    f"| teacher_counseling_count | {TEACHER_FIELD} | 전담교사 상담 건수 (연도별) |",
    f"| external_counseling_count | {EXTERNAL_FIELD} | 외부전문가 상담 건수 (연도별) |",
    "| total_counseling_count | (파생) | 전담+외부 합계 (연도별) |",
    "",
    "> apiType=61에는 학생·학부모 상담 건수 별도 구분 필드 없음.",
    "> 전담교사 상담 건수와 외부전문가 상담 건수를 합산하여 통합 상담 건수로 사용한다.",
    "",
    "## 3개년 평균 변수",
    "",
    "| 표준 변수명 | 설명 |",
    "|---|---|",
    "| avg_teacher_counseling_count_3yr | 2023·2024·2025 전담교사 상담 건수 평균 |",
    "| avg_external_counseling_count_3yr | 2023·2024·2025 외부전문가 상담 건수 평균 |",
    "| avg_total_counseling_count_3yr | 2023·2024·2025 통합 상담 건수 평균 |",
    "| counseling_years_available | 평균 산출에 사용된 연도 수 (최대 3) |",
    "",
    "- 일부 연도만 존재하는 경우 존재하는 연도의 평균 산출",
    "- 3개년 모두 결측이면 평균값 NaN 유지",
    "",
    "## 학생 수 대비 통합 상담 건수",
    "",
    "- `counseling_count_per_student = avg_total_counseling_count_3yr / student_count`",
    "- student_count가 0 또는 결측이면 NaN (0나누기 방지)",
    "",
    "## Min-Max 정규화",
    "",
    "- 공식: `normalized = (x - min) / (max - min)`",
    "- 대상: `avg_total_counseling_count_3yr`, `counseling_count_per_student`",
    "- max == min인 경우: 비결측값을 0으로 처리 (정규화 불능 명시)",
    "- 정규화 변수명: `norm_avg_total_counseling_count_3yr`, `norm_counseling_count_per_student`",
    "",
    "## counseling_use_score 산출 (실제 상담 이용 점수)",
    "",
    "- `counseling_use_score = mean(norm_avg_total_counseling_count_3yr, norm_counseling_count_per_student)`",
    "- 한 지표만 존재하면 해당 지표 단독 사용",
    "- 두 지표 모두 결측이면 NaN",
    "- `counseling_use_components_available`: 사용된 구성 지표 수 (0~2)",
    "",
    "## CDI 구조 (향후 참고)",
    "",
    "- CDI = (상담 수요 규모 점수 + 실제 상담 이용 점수 + 학교폭력 위험 점수) / 3",
    "- 현 단계 산출 항목: **실제 상담 이용 점수 (`counseling_use_score`)**",
    "- 상담 수요 규모 점수: 추후 학생 수 기준 구간화 또는 Min-Max 정규화 방식 결정",
    "- 학교폭력 위험 점수: 추후 학교폭력 관련 공시자료 확보 후 산출",
    "",
    "## 결측 처리 기준",
    "",
    "- 상담 건수 0: 실제 0 (미운영) 유지",
    "- 공란·변환 불가 값: NaN 처리",
    "- 3개년 모두 결측: 평균값 NaN 유지",
    "",
    "## 병합 기준",
    "",
    "- KEDI 학교코드(KESS, 480XXXXX 형식)와 학교알리미 코드(API, SXXXXXXXXX 형식)는",
    "  코드 체계가 상이하여 직접 병합 불가",
    "- **school_name 정규화 병합**: 공백·괄호·특수문자 제거 후 소문자 매칭",
    "- 미매칭 학교는 상담 건수 관련 변수 NaN 처리, 실행 시 검증 로그 출력",
    "",
    "## API 전체 응답 필드 목록",
    "",
    "| 필드명 | 예시값 |",
    "|---|---|",
] + [f"| {k} | {_ex(k)} |" for k in all_api_keys] + [
    "",
    "## 한계 및 확인 필요 사항",
    "",
    "- 2023년 이전 공시연도 데이터는 API에서 제공하지 않음 (최근 3년 제한, 현재 기준 2024~2026)",
    "- school_name 정규화 병합으로 동명이교 오매핑 가능 → 결과 수동 검토 권장",
    "- 창원시 시 단위(sggCode=48120)는 API 데이터 없음 → 구 단위 코드로 대체 호출",
]

MAP_MD.write_text("\n".join(doc_lines), encoding="utf-8")
print(f"[INFO] 변수 매핑 문서 저장: {MAP_MD.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# 최종 요약
# ═══════════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("최종 요약")
print("=" * 60)
print(f"  API 호출 총 횟수   : {len(PBAN_TO_DATA)}개연도 × {len(sgg_list)}개시군구 = {len(PBAN_TO_DATA)*len(sgg_list)}회")
print(f"  수집 일반고 레코드 : {len(df_general)}건 (학교×연도 단위)")
print(f"  KESS 기본정보 학교 : {n_kess}개교")
print(f"  병합 매칭 성공     : {matched}개교")
print(f"  병합 미매칭(결측)  : {unmatched}개교")
print(f"  counseling_use_score 평균 : {df_merged['counseling_use_score'].mean():.4f}")
print(f"  counseling_use_score 범위 : "
      f"{df_merged['counseling_use_score'].min():.4f} ~ "
      f"{df_merged['counseling_use_score'].max():.4f}")
print("=" * 60)
