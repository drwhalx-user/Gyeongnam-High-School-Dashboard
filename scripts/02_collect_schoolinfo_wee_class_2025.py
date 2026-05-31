"""
학교알리미 API (apiType=61) 순환 호출 → 경남 고등학교별 Wee클래스 운영 여부 추출
기존 KESS 기본정보 테이블에 wee_class 이진 변수를 추가하여 병합한다.

입력:
  시도시군구코드.xlsx                                            → 경남 시군구 코드
  data/processed/gyeongnam_general_high_schools_basic_2025.csv  → KESS 기본정보

출력:
  data/raw/schoolinfo_wee_class_2025_all_sigungu.json
  data/processed/schoolinfo_wee_class_2025_clean.csv
  data/processed/gyeongnam_general_high_schools_with_wee_class_2025.csv
  outputs/tables/wee_class_missing_summary_2025.csv
  outputs/tables/wee_class_api_call_log_2025.csv
  docs/wee_class_variable_mapping_2025.md
"""

import sys
import time
import json
import re
import pathlib

import requests
import pandas as pd

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
ROOT = pathlib.Path(__file__).resolve().parent.parent

SGG_CODE_FILE = ROOT / "시도시군구코드.xlsx"
KESS_FILE     = ROOT / "data" / "processed" / "gyeongnam_general_high_schools_basic_2025.csv"
RAW_JSON      = ROOT / "data" / "raw"        / "schoolinfo_wee_class_2025_all_sigungu.json"
CLEAN_CSV     = ROOT / "data" / "processed"  / "schoolinfo_wee_class_2025_clean.csv"
MERGED_CSV    = ROOT / "data" / "processed"  / "gyeongnam_general_high_schools_with_wee_class_2025.csv"
MISSING_CSV   = ROOT / "outputs" / "tables"  / "wee_class_missing_summary_2025.csv"
CALL_LOG_CSV  = ROOT / "outputs" / "tables"  / "wee_class_api_call_log_2025.csv"
MAPPING_MD    = ROOT / "docs"                / "wee_class_variable_mapping_2025.md"

# 필요한 폴더 자동 생성
for _dir in [RAW_JSON.parent, CLEAN_CSV.parent, MISSING_CSV.parent, MAPPING_MD.parent]:
    _dir.mkdir(parents=True, exist_ok=True)

# ── API 설정 ──────────────────────────────────────────────────────────────────
API_URL = "http://www.schoolinfo.go.kr/openApi.do"
API_KEY = "39bf3d7a293447baba1702f71307e5ad"
API_BASE_PARAMS = {
    "apiKey"      : API_KEY,
    "apiType"     : "61",
    "pbanYr"      : "2025",
    "sidoCode"    : "48",
    "schulKndCode": "04",
}

# ── Wee클래스 이진화 매핑 ─────────────────────────────────────────────────────
WEE_POSITIVE = {"y", "예", "설치", "운영", "있음", "유", "o", "1"}
WEE_NEGATIVE = {"n", "아니오", "미설치", "미운영", "없음", "무", "x", "0"}

def binarize_wee(raw_value):
    """WEE_CINSTL_YN 원본값을 1 / 0 / NaN으로 변환"""
    if pd.isna(raw_value) or str(raw_value).strip() == "":
        return float("nan")
    v = str(raw_value).strip().lower()
    if v in WEE_POSITIVE:
        return 1
    if v in WEE_NEGATIVE:
        return 0
    return float("nan")  # 확인 불가 → 결측


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 : 시군구 코드 파일 로드 및 경남 코드 추출
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 시도시군구코드.xlsx 로드")

# 파일 구조: row 0=파일명(불필요), row 1=공백, row 2=헤더 → skiprows=2
df_sgg = pd.read_excel(SGG_CODE_FILE, skiprows=2, header=0)
df_sgg.columns = ["시도명", "시도코드", "시군구명", "시군구코드"]

# 경상남도 행 추출 후 시군구코드를 5자리 문자열로 보존
df_gyeongnam = df_sgg[df_sgg["시도명"] == "경상남도"].copy()
df_gyeongnam["시군구코드_str"] = df_gyeongnam["시군구코드"].astype(str).str.zfill(5)

# 창원시 시 단위(48120)는 API에서 데이터 미제공 → 제외
# (구 단위 코드 48121·48123·48125·48127·48129가 파일 내 이미 포함)
SKIP_CODES = {"48120"}
df_gyeongnam = df_gyeongnam[~df_gyeongnam["시군구코드_str"].isin(SKIP_CODES)]

# 코드 → 시군구명 매핑 딕셔너리 및 정렬된 목록 생성
sgg_map  = dict(zip(df_gyeongnam["시군구코드_str"], df_gyeongnam["시군구명"]))
sgg_list = sorted(sgg_map.items())   # [(코드, 시군구명), ...]

print(f"  → 경상남도 시군구 코드 {len(sgg_list)}개 확인:")
for code, name in sgg_list:
    print(f"     {name}({code})")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 : 학교알리미 API 순환 호출
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 2] 학교알리미 API 순환 호출 시작")

all_raw_responses = {}   # {sgg_code: 원본 응답 dict} → JSON으로 저장
all_records       = []   # 성공한 시군구의 list 항목 누적
call_log          = []   # 시군구별 호출 로그

for sgg_code, sgg_name in sgg_list:
    params = {**API_BASE_PARAMS, "sggCode": sgg_code}

    # 요청 실패 시 전체 중단 없이 로그만 남기고 계속 진행
    try:
        resp = requests.get(API_URL, params=params, timeout=20)
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        err_msg = str(exc)[:120]
        print(f"  [ERR]  {sgg_name}({sgg_code}): {err_msg}")
        all_raw_responses[sgg_code] = {"error": err_msg}
        call_log.append({
            "시군구명": sgg_name, "시군구코드": sgg_code,
            "resultCode": "error", "resultMsg": err_msg,
            "응답건수": 0, "비고": "요청 예외",
        })
        time.sleep(0.5)
        continue

    rc    = data.get("resultCode", "")
    msg   = data.get("resultMsg", "")
    items = data.get("list", [])
    n     = len(items)

    all_raw_responses[sgg_code] = data

    if rc == "success":
        # 각 레코드에 출처 시군구 정보 추가
        for item in items:
            item["_sgg_code"] = sgg_code
            item["_sgg_name"] = sgg_name
        all_records.extend(items)
        print(f"  [OK]   {sgg_name}({sgg_code}): {n}건")
        call_log.append({
            "시군구명": sgg_name, "시군구코드": sgg_code,
            "resultCode": rc, "resultMsg": msg,
            "응답건수": n, "비고": "",
        })
    else:
        print(f"  [FAIL] {sgg_name}({sgg_code}): {msg[:60]}")
        call_log.append({
            "시군구명": sgg_name, "시군구코드": sgg_code,
            "resultCode": rc, "resultMsg": msg,
            "응답건수": 0, "비고": "API 실패",
        })

    time.sleep(0.3)  # 서버 부하 방지

print(f"\n  → 총 수집 레코드: {len(all_records)}건")

# API 원본 응답 전체 저장
with open(RAW_JSON, "w", encoding="utf-8") as f:
    json.dump(all_raw_responses, f, ensure_ascii=False, indent=2)
print(f"[INFO] 원본 JSON 저장: {RAW_JSON.name}")

# 호출 로그 저장
df_log = pd.DataFrame(call_log)
df_log.to_csv(CALL_LOG_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] API 호출 로그 저장: {CALL_LOG_CSV.name}\n")

if not all_records:
    print("[ERROR] 수집된 레코드가 없습니다. API 호출 로그를 확인하세요.")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 : Wee클래스 필드 탐색 및 확인
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 3] Wee클래스 관련 필드 탐색")

sample = all_records[0]

# 키 이름에 Wee 관련 키워드가 포함된 후보 필드 탐색
wee_keywords = {"wee", "위클래스", "클래스", "설치", "운영"}
candidate_fields = [
    k for k in sample.keys()
    if not k.startswith("_") and any(kw in k.lower() for kw in wee_keywords)
]

# 실제 API 응답에서 확인된 Wee클래스 필드명
WEE_FIELD = "WEE_CINSTL_YN"

if WEE_FIELD in sample:
    print(f"  → Wee클래스 필드 확인: '{WEE_FIELD}'")
    print(f"     첫 번째 레코드 예시: {sample[WEE_FIELD]!r}")
else:
    print(f"  [경고] '{WEE_FIELD}' 필드 없음. 후보 필드 탐색:")
    if candidate_fields:
        for f in candidate_fields:
            print(f"    - {f!r}: {sample.get(f)!r}")
        WEE_FIELD = candidate_fields[0]
        print(f"  → 후보 첫 번째 '{WEE_FIELD}'를 사용합니다.")
    else:
        print("  [ERROR] Wee클래스 관련 필드를 찾지 못했습니다. 전체 필드 목록:")
        for k, v in sample.items():
            if not k.startswith("_"):
                print(f"    {k!r}: {v!r}")
        sys.exit(1)
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 : DataFrame 변환 및 학교 단위 / 비학교 단위 분리
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 4] 데이터 변환 및 학교 단위 분리")

df_api = pd.DataFrame(all_records)

# 학교명(SCHUL_NM)이 있는 행만 학교 단위로 판별
mask_school = df_api["SCHUL_NM"].notna() & df_api["SCHUL_NM"].str.strip().ne("")
df_school     = df_api[mask_school].copy()
df_non_school = df_api[~mask_school].copy()

if len(df_non_school) > 0:
    print(f"  [주의] 학교명 없는 행(시군 단위 추정) {len(df_non_school)}건 → 병합 제외:")
    for _, row in df_non_school.iterrows():
        print(f"    {row.get('_sgg_name')} | SCHUL_CODE={row.get('SCHUL_CODE')}")

print(f"  → 학교 단위 레코드: {len(df_school)}건\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 : Wee클래스 이진화
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 5] Wee클래스 이진화")

df_school["wee_class_raw"] = df_school[WEE_FIELD]
df_school["wee_class"]     = df_school["wee_class_raw"].apply(binarize_wee)

print("  원본값 분포 (WEE_CINSTL_YN → wee_class):")
for val, cnt in df_school["wee_class_raw"].value_counts(dropna=False).items():
    print(f"    {val!r}: {cnt}건  →  wee_class = {binarize_wee(val)}")
print()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 : 정제 데이터 저장 (schoolinfo_wee_class_2025_clean.csv)
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 6] 정제 데이터 저장")

# 표준 변수명으로 열 이름 변경
rename_map = {
    "SCHUL_CODE": "school_code_api",
    "SCHUL_NM"  : "school_name",
    "_sgg_code" : "sgg_code",
    "_sgg_name" : "sigungu",
}
df_clean = df_school.rename(columns={k: v for k, v in rename_map.items() if k in df_school.columns}).copy()
df_clean["sido_code"] = "48"
df_clean["pban_year"] = 2025

clean_out_cols = [
    "school_code_api", "school_name", "sido_code", "sgg_code", "sigungu",
    "pban_year", "wee_class_raw", "wee_class",
]
clean_out_cols = [c for c in clean_out_cols if c in df_clean.columns]
df_clean[clean_out_cols].to_csv(CLEAN_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 정제 데이터 저장: {CLEAN_CSV.name}  ({len(df_clean)}행)\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 7 : KESS 기본정보 파일과 병합
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 7] KESS 기본정보 파일 병합")

df_kess = pd.read_csv(KESS_FILE, encoding="utf-8-sig")
n_kess  = len(df_kess)
print(f"  KESS 기본정보 학교 수: {n_kess}개교")

def normalize_name(name: str) -> str:
    """학교명 정규화: 공백·괄호 내용·특수문자 제거 후 소문자 변환"""
    s = str(name).strip()
    s = re.sub(r"[\s ]+", "", s)   # 모든 공백 제거
    s = re.sub(r"\(.*?\)", "", s)       # (괄호 내용) 제거
    s = re.sub(r"[^\w가-힣]", "", s)    # 특수문자 제거
    return s.lower()

# 양쪽에 정규화 키 추가 (원본 학교명은 그대로 보존)
df_kess["_name_key"]  = df_kess["school_name"].apply(normalize_name)

# API 정제 데이터에서 학교명별 wee_class 추출
wee_lookup = df_clean[["school_name", "wee_class_raw", "wee_class"]].copy()
wee_lookup["_name_key"] = wee_lookup["school_name"].apply(normalize_name)

# 동명이교 중복 경고 처리 → 첫 번째 항목 사용
dup_keys = wee_lookup[wee_lookup.duplicated("_name_key", keep=False)]["_name_key"].unique()
if len(dup_keys) > 0:
    print(f"  [경고] 정규화 후 중복 학교명 {len(dup_keys)}건 → 첫 번째 항목 사용:")
    for k in dup_keys:
        dnames = wee_lookup[wee_lookup["_name_key"] == k]["school_name"].tolist()
        print(f"    key={k!r} → {dnames}")

wee_lookup_dedup = wee_lookup.drop_duplicates("_name_key", keep="first")

# KESS 기본정보에 wee_class 병합 (left join → KESS 행 수 유지)
df_merged = df_kess.merge(
    wee_lookup_dedup[["_name_key", "wee_class_raw", "wee_class"]],
    on="_name_key",
    how="left",
)
df_merged = df_merged.drop(columns=["_name_key"])

# 병합 행 수 검증 (행 증가 시 오류)
if len(df_merged) != n_kess:
    print(f"[ERROR] 병합 후 행 수 불일치: {len(df_merged)} ≠ {n_kess}")
    print("        중복 키 발생 가능 — wee_lookup_dedup 확인 필요")
    sys.exit(1)

matched   = df_merged["wee_class"].notna().sum()
unmatched = df_merged["wee_class"].isna().sum()
print(f"  → 매칭 성공: {matched}개교 / 미매칭(결측): {unmatched}개교")

# 미매칭 학교 검증 로그 출력
if unmatched > 0:
    print("  [검증] 미매칭 학교 목록 (API 데이터 없음 또는 학교명 불일치):")
    for _, row in df_merged[df_merged["wee_class"].isna()].iterrows():
        print(f"    {row['sigungu']} / {row['school_name']}  (school_code={row['school_code']})")

print(f"  → 병합 후 행 수: {len(df_merged)}개교 (KESS 원본 행 수 유지 확인)\n")

df_merged.to_csv(MERGED_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 병합 최종 파일 저장: {MERGED_CSV.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 8 : 결측 요약표 생성
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 8] 결측 요약표 생성")

check_cols = [c for c in list(df_kess.columns) + ["wee_class_raw", "wee_class"]
              if c in df_merged.columns]
miss_rows = []
for col in check_cols:
    total  = len(df_merged)
    n_miss = int(df_merged[col].isna().sum())
    n_zero = int((df_merged[col] == 0).sum()) if pd.api.types.is_numeric_dtype(df_merged[col]) else 0
    miss_rows.append({
        "variable"       : col,
        "total"          : total,
        "missing_count"  : n_miss,
        "zero_count"     : n_zero,
        "missing_rate(%)": round(n_miss / total * 100, 2),
    })

pd.DataFrame(miss_rows).to_csv(MISSING_CSV, index=False, encoding="utf-8-sig")
print(f"[INFO] 결측 요약표 저장: {MISSING_CSV.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 9 : 변수 매핑 문서 작성
# ═══════════════════════════════════════════════════════════════════════════════
print("[STEP 9] 변수 매핑 문서 작성")

all_api_keys = sorted({
    k for rec in all_records
    for k in rec.keys()
    if not k.startswith("_")
})

def _example_val(field: str) -> str:
    for rec in all_records:
        v = rec.get(field)
        if v is not None:
            return str(v)[:40]
    return ""

lines = [
    "# Wee클래스 변수 매핑표 (2025)",
    "",
    "## API 호출 정보",
    "",
    "| 파라미터 | 값 | 설명 |",
    "|---|---|---|",
    "| apiType | 61 | 학생·학부모 상담계획 및 실시 현황 |",
    "| pbanYr | 2025 | 공시 연도 |",
    "| sidoCode | 48 | 경상남도 |",
    "| sggCode | 시군구별 코드 | 시도시군구코드.xlsx 참조 |",
    "| schulKndCode | 04 | 고등학교 |",
    f"| 요청 URL 형식 | `{API_URL}?apiKey={{KEY}}&apiType=61&pbanYr=2025&sidoCode=48&sggCode={{SGG}}&schulKndCode=04` | |",
    "",
    "## Wee클래스 변수",
    "",
    "| 표준 변수명 | 원본 API 필드 | 설명 |",
    "|---|---|---|",
    f"| wee_class_raw | {WEE_FIELD} | API 원본 응답값 그대로 보존 |",
    "| wee_class | (파생) | 이진화 결과 (1=운영, 0=미운영, NaN=확인불가) |",
    "",
    "## wee_class 이진화 기준",
    "",
    "| 원본값 | wee_class | 비고 |",
    "|---|---|---|",
    "| Y, y, 예, 설치, 운영, 있음, 유, O, o, 1 | 1 | Wee클래스 운영 |",
    "| N, n, 아니오, 미설치, 미운영, 없음, 무, X, x, 0 | 0 | Wee클래스 미운영 |",
    "| 공란, 기타 | NaN | 결측 처리 — 실제 0과 구분 |",
    "",
    "## 결측 처리 기준",
    "",
    "- `wee_class == 0` : API 응답값 N 확인 → 실제 미운영",
    "- `wee_class == NaN` : API 미매칭 또는 원본값 불명확 → 0과 구분하여 보존",
    "",
    "## 병합 기준",
    "",
    "- KEDI 학교코드(KESS)와 학교알리미 학교코드(API)는 코드 체계가 상이하여 직접 병합 불가",
    "  - KESS: 480041002 형식 (KEDI 부여 8자리 숫자)",
    "  - API : S160000464 형식 (학교알리미 자체 코드)",
    "- **1차: school_name 정규화 병합** (공백·괄호·특수문자 제거 후 소문자 매칭)",
    "- 미매칭 학교는 wee_class = NaN 보존, 실행 시 검증 로그 출력",
    "",
    "## API 전체 응답 필드 목록",
    "",
    "| 필드명 | 예시값 |",
    "|---|---|",
]
for k in all_api_keys:
    lines.append(f"| {k} | {_example_val(k)} |")

lines += [
    "",
    "## 한계 및 확인 필요 사항",
    "",
    "- 창원시 시 단위(sggCode=48120)는 API 데이터 없음 → 구 단위 코드(48121·48123·48125·48127·48129)로 대체 호출",
    "- 공시 예외 학교(`PBAN_EXCP_YN=Y`)는 응답에 포함되나 실제 운영 여부 별도 확인 필요",
    "- school_name 기준 병합으로 동명이교 존재 시 오매핑 가능 → 결과 수동 검토 권장",
    f"- API 시군구별 호출 로그: `outputs/tables/wee_class_api_call_log_2025.csv`",
]

MAPPING_MD.write_text("\n".join(lines), encoding="utf-8")
print(f"[INFO] 변수 매핑 문서 저장: {MAPPING_MD.name}\n")


# ═══════════════════════════════════════════════════════════════════════════════
# 최종 요약
# ═══════════════════════════════════════════════════════════════════════════════
wee_1 = int((df_merged["wee_class"] == 1).sum())
wee_0 = int((df_merged["wee_class"] == 0).sum())

print("=" * 55)
print("최종 요약")
print("=" * 55)
print(f"  API 호출 시군구 수  : {len(sgg_list)}개")
print(f"  API 수집 학교 수    : {len(df_clean)}개교")
print(f"  KESS 기본정보 학교  : {n_kess}개교")
print(f"  병합 매칭 성공      : {matched}개교")
print(f"  병합 미매칭(결측)   : {unmatched}개교")
print(f"  Wee클래스 운영  (1) : {wee_1}개교")
print(f"  Wee클래스 미운영(0) : {wee_0}개교")
print(f"  Wee클래스 확인불가  : {unmatched}개교")
print("=" * 55)
