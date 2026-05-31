"""
경남교육청 교육지원청 Wee센터 현황 PDF에서 Wee센터 기준 데이터셋을 생성한다.

입력 파일:
  data/raw/2025. 경남교육청 교육지원청 Wee센터 현황.pdf

출력 파일:
  data/processed/gyeongnam_wee_centers_2025.csv
  outputs/tables/wee_center_extraction_check_2025.csv
  outputs/tables/wee_center_missing_summary_2025.csv
  docs/wee_center_data_source_2025.md
"""

import re
import sys
import pathlib
import pandas as pd
import pdfplumber

# ── 경로 설정 ────────────────────────────────────────────────────────────────
ROOT         = pathlib.Path(__file__).resolve().parent.parent
RAW_DIR      = ROOT / "data" / "raw"
PROCESSED_DIR = ROOT / "data" / "processed"
TABLES_DIR   = ROOT / "outputs" / "tables"
DOCS_DIR     = ROOT / "docs"

# 출력 폴더 없으면 자동 생성
for d in [PROCESSED_DIR, TABLES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── 입력 파일 경로 ────────────────────────────────────────────────────────────
# 실제 파일은 data/raw/ 바로 아래에 위치 (하위 폴더 없음)
PDF_FILENAME = "2025. 경남교육청 교육지원청 Wee센터 현황.pdf"
PDF_PATH     = RAW_DIR / PDF_FILENAME

# ── 필수 포함 확인 지역 목록 (요구사항 명세 기준) ─────────────────────────────
REQUIRED_REGIONS = [
    "창원", "마산분원", "진주", "통영", "사천", "김해", "밀양",
    "거제", "양산", "의령", "함안", "창녕", "고성", "남해",
    "하동", "산청", "함양", "거창", "합천",
]

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. PDF 파일 존재 확인
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 1] PDF 파일 확인")

if not PDF_PATH.exists():
    print(f"[ERROR] PDF 파일이 존재하지 않습니다: {PDF_PATH}")
    print("        data/raw/ 폴더에 PDF 파일을 배치한 뒤 다시 실행하세요.")
    sys.exit(1)

print(f"  → 파일 확인: {PDF_PATH.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. PDF 페이지 수 확인 및 텍스트 추출 가능 여부 점검
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 2] PDF 페이지 수 및 텍스트 추출 가능 여부 확인")

with pdfplumber.open(PDF_PATH) as pdf:
    n_pages = len(pdf.pages)
    print(f"  → 총 페이지 수: {n_pages}")

    # 첫 페이지 텍스트 샘플로 추출 가능 여부 확인
    sample_text = pdf.pages[0].extract_text()

if not sample_text or len(sample_text.strip()) < 10:
    print("[ERROR] PDF에서 텍스트를 추출할 수 없습니다.")
    print("        스캔 이미지 기반 PDF일 가능성이 있습니다 → OCR 처리 필요.")
    print("        임의 데이터를 생성하지 않으므로 작업을 중단합니다.")
    sys.exit(1)

print(f"  → 텍스트 추출 가능 확인 (샘플 길이: {len(sample_text.strip())}자)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 3. PDF 표 추출
# pdfplumber extract_tables()는 셀 단위로 파싱하므로
# 단순 텍스트 추출보다 표 구조를 정확히 복원한다.
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 3] PDF 표(지역·전화번호·주소) 추출")

with pdfplumber.open(PDF_PATH) as pdf:
    page = pdf.pages[0]
    tables = page.extract_tables()

if not tables:
    print("[ERROR] PDF에서 표를 감지하지 못했습니다. 작업을 중단합니다.")
    sys.exit(1)

raw_table = tables[0]   # 표가 1개임을 확인했으므로 첫 번째 표 사용
print(f"  → 감지된 표: {len(tables)}개 / 사용할 표: {len(raw_table)}행 (헤더 포함)")

# 헤더 행 제거 (첫 행: ['지역', '전화번호', '주소'])
header = raw_table[0]
data_rows = raw_table[1:]
print(f"  → 헤더: {header}")
print(f"  → 데이터 행 수: {len(data_rows)}개 (Wee센터 수)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. 지역명 정규화 함수
# PDF 셀에서 줄바꿈이 포함된 지역명(예: '마산\n분원')을 정리한다.
# ════════════════════════════════════════════════════════════════════════════
def normalize_region(raw_region: str) -> str:
    """PDF 지역 열 값을 정규화한다."""
    if raw_region is None:
        return ""
    # 줄바꿈 제거 후 공백 합치기 (예: '마산\n분원' → '마산분원')
    return re.sub(r"\s+", "", raw_region.strip())


# ════════════════════════════════════════════════════════════════════════════
# STEP 5. 주소 정제 함수
# 주소 셀 내 줄바꿈을 공백으로 합치고,
# 끝에 붙어 있는 Wee센터 명칭 표현을 제거한다.
# 건물 층수·기관 위치 정보(예: '4층', '마산교육지원센터 2층')는 유지한다.
# ════════════════════════════════════════════════════════════════════════════
def clean_address(raw_addr: str) -> str:
    """주소 셀을 정제하여 순수 주소 문자열을 반환한다."""
    if raw_addr is None:
        return ""

    # 줄바꿈을 공백으로 합침
    addr = " ".join(raw_addr.split("\n")).strip()

    # 창원 특수 케이스: '창원교육지원청Wee센터' (공백 없이 붙어 있음) 제거
    addr = re.sub(r"\s*창원교육지원청Wee센터\s*$", "", addr).strip()

    # 마산 분원 특수 케이스: '창원Wee센터 마산분원' 제거 (2층은 앞에서 처리됨)
    addr = re.sub(r"\s*창원Wee센터\s*마산분원\s*$", "", addr).strip()

    # 일반 케이스: 끝의 'Wee센터' 제거
    # 단, '산청교육지원청 Wee센터' 형태도 여기서 처리됨 → '산청교육지원청' 유지
    addr = re.sub(r"\s*Wee센터\s*$", "", addr).strip()

    return addr


# ════════════════════════════════════════════════════════════════════════════
# STEP 6. 센터명(wee_center_name) 생성 함수
# 지역명 기준으로 공식 센터명을 부여한다.
# ════════════════════════════════════════════════════════════════════════════
def make_center_name(region_key: str) -> str:
    """정규화된 지역 키로 공식 Wee센터명을 생성한다."""
    if region_key == "창원":
        return "창원교육지원청 Wee센터"
    elif region_key == "마산분원":
        return "창원Wee센터 마산분원"
    else:
        return f"{region_key}교육지원청 Wee센터"


# ════════════════════════════════════════════════════════════════════════════
# STEP 7. 행별 파싱 및 표준 변수 생성
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 7] 행별 파싱 및 표준 변수 생성")

records = []
for idx, row in enumerate(data_rows, start=1):
    # 셀이 3개 미만이면 형식 오류 → 빈 값으로 채워 계속 진행
    region_raw = row[0] if len(row) > 0 else None
    phone_raw  = row[1] if len(row) > 1 else None
    addr_raw   = row[2] if len(row) > 2 else None

    region_key = normalize_region(region_raw)     # 지역명 정규화
    center_name = make_center_name(region_key)    # 센터명 생성
    address     = clean_address(addr_raw)         # 주소 정제
    phone       = phone_raw.strip() if phone_raw else ""

    record = {
        "wee_center_id"    : f"WEE_GN_{idx:03d}",
        "sigungu"          : region_key,
        "wee_center_name"  : center_name,
        "phone"            : phone if phone else None,
        "address"          : address if address else None,
        "source_file"      : PDF_FILENAME,
        "base_year"        : 2025,
        "extraction_check" : "needs_review",      # 자동 추출이므로 검수 필요 표시
    }
    records.append(record)
    print(f"  [{idx:02d}] {region_key:8s}  |  {center_name}  |  {address[:40]}...")

df = pd.DataFrame(records)
print(f"\n  → 파싱 완료: {len(df)}개 Wee센터")

# ════════════════════════════════════════════════════════════════════════════
# STEP 8. 데이터 검증
# 필수 포함 지역 누락 여부, 결측값 체크
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 8] 데이터 검증")

extracted_regions = set(df["sigungu"].dropna().tolist())

# 8-1. 필수 지역 누락 확인
missing_regions = [r for r in REQUIRED_REGIONS if r not in extracted_regions]
extra_regions   = [r for r in extracted_regions if r not in REQUIRED_REGIONS]

if missing_regions:
    print(f"  [WARNING] 누락된 지역: {missing_regions}")
else:
    print(f"  [OK] 필수 지역 19개 모두 포함 확인")

if extra_regions:
    print(f"  [INFO] 명세 외 추가 지역 감지: {extra_regions}")

# 8-2. 변수별 결측 확인
check_cols = ["sigungu", "phone", "address"]
for col in check_cols:
    n_miss = df[col].isna().sum()
    flag   = "[WARNING]" if n_miss > 0 else "[OK]"
    print(f"  {flag} {col} 결측: {n_miss}개")

# ════════════════════════════════════════════════════════════════════════════
# STEP 9. 검수표 생성 (누락 지역 + 결측 행 목록)
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 9] 검수표 생성")

check_records = []

# 9-1. 필수 지역 누락 기록
for region in missing_regions:
    check_records.append({
        "check_type"   : "missing_region",
        "wee_center_id": "",
        "sigungu"      : region,
        "issue"        : f"필수 포함 지역 '{region}'이 PDF에서 추출되지 않음",
        "action"       : "PDF 원본 재확인 필요",
    })

# 9-2. 전화번호 결측 행 기록
for _, row_data in df[df["phone"].isna()].iterrows():
    check_records.append({
        "check_type"   : "missing_phone",
        "wee_center_id": row_data["wee_center_id"],
        "sigungu"      : row_data["sigungu"],
        "issue"        : "전화번호 결측",
        "action"       : "PDF 원본 재확인 필요",
    })

# 9-3. 주소 결측 행 기록
for _, row_data in df[df["address"].isna()].iterrows():
    check_records.append({
        "check_type"   : "missing_address",
        "wee_center_id": row_data["wee_center_id"],
        "sigungu"      : row_data["sigungu"],
        "issue"        : "주소 결측",
        "action"       : "임의 보완하지 않음 — PDF 원본 재확인 필요",
    })

# 9-4. 문제 없으면 "이상 없음" 1행 기록
if not check_records:
    check_records.append({
        "check_type"   : "all_clear",
        "wee_center_id": "",
        "sigungu"      : "",
        "issue"        : "이상 없음",
        "action"       : "전 항목 정상 추출 확인됨",
    })

df_check = pd.DataFrame(check_records)

check_path = TABLES_DIR / "wee_center_extraction_check_2025.csv"
df_check.to_csv(check_path, index=False, encoding="utf-8-sig")
print(f"  → 검수표 저장: {check_path.name}  ({len(df_check)}건)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 10. 결측 요약표 생성
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 10] 결측 요약표 생성")

target_cols = ["wee_center_id", "sigungu", "wee_center_name", "phone", "address",
               "source_file", "base_year", "extraction_check"]

missing_summary_records = []
for col in target_cols:
    n_miss  = int(df[col].isna().sum())
    n_empty = int((df[col].astype(str).str.strip() == "").sum()) if col in df.columns else 0
    missing_summary_records.append({
        "variable"     : col,
        "total"        : len(df),
        "missing_count": n_miss,
        "missing_rate(%)": round(n_miss / len(df) * 100, 2),
    })

df_missing = pd.DataFrame(missing_summary_records)

missing_path = TABLES_DIR / "wee_center_missing_summary_2025.csv"
df_missing.to_csv(missing_path, index=False, encoding="utf-8-sig")
print(f"  → 결측 요약표 저장: {missing_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# STEP 11. 최종 데이터셋 저장
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 11] 최종 데이터셋 저장")

output_cols = [
    "wee_center_id", "sigungu", "wee_center_name",
    "phone", "address",
    "source_file", "base_year", "extraction_check",
]
df_output = df[output_cols].copy()

output_path = PROCESSED_DIR / "gyeongnam_wee_centers_2025.csv"
df_output.to_csv(output_path, index=False, encoding="utf-8-sig")
print(f"  → 저장 완료: {output_path.name}  ({len(df_output)}행 × {len(output_cols)}열)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 12. 방법론 문서 작성
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 12] 방법론 문서 작성")

doc_lines = [
    "# Wee센터 기준 데이터셋 방법론 문서 (2025)",
    "",
    "## 자료 개요",
    "",
    f"- **자료명**: 2025. 경남교육청 교육지원청 Wee센터 현황",
    f"- **원본 파일명**: `{PDF_FILENAME}`",
    f"- **기준연도**: 2025",
    f"- **출처**: 경상남도교육청 내부 자료 (PDF)",
    "",
    "## 추출 방식",
    "",
    "- **라이브러리**: `pdfplumber` (`page.extract_tables()`)",
    "- **추출 단위**: 표 셀 단위 파싱 (단순 텍스트 추출 대비 구조 보존)",
    "- **추출 페이지**: 1페이지 (전체 1페이지)",
    "- **추출 행 수**: 19개 (헤더 제외)",
    "- **텍스트 추출 가능 여부**: 확인됨 (스캔 이미지 아님)",
    "",
    "## 표준 변수명",
    "",
    "| 변수명 | 설명 |",
    "|--------|------|",
    "| `wee_center_id` | Wee센터 고유 ID (WEE_GN_001 ~ WEE_GN_019, PDF 행 순서) |",
    "| `sigungu` | PDF 지역 열 정규화값 (줄바꿈 제거, 예: '마산분원') |",
    "| `wee_center_name` | 공식 센터명 (아래 규칙 참조) |",
    "| `phone` | 전화번호 (원본값 그대로) |",
    "| `address` | 정제된 주소 (아래 규칙 참조) |",
    "| `source_file` | 원본 PDF 파일명 |",
    "| `base_year` | 기준연도 (2025) |",
    "| `extraction_check` | 자동 추출 검수 상태 (`needs_review` 고정) |",
    "",
    "## 센터명 생성 규칙",
    "",
    "| 지역 키 | wee_center_name |",
    "|---------|----------------|",
    "| 창원 | 창원교육지원청 Wee센터 |",
    "| 마산분원 | 창원Wee센터 마산분원 |",
    "| 그 외 | {지역}교육지원청 Wee센터 |",
    "",
    "## 주소 정리 기준",
    "",
    "1. 셀 내 줄바꿈(`\\n`)을 공백으로 합쳐 단일 주소 문자열로 변환",
    "2. 끝에 붙은 Wee센터 명칭 표현을 제거 (순서대로 적용)",
    "   - `창원교육지원청Wee센터` (창원 특수 케이스)",
    "   - `창원Wee센터 마산분원` (마산 분원 특수 케이스)",
    "   - `Wee센터` (일반 케이스)",
    "3. 앞뒤 공백 제거",
    "4. 건물 층수·기관 위치 정보는 주소에 유지",
    "   - 예: `진주교육지원청 4층`, `밀양영재교육원 1층`, `산청교육지원청`",
    "5. 주소가 결측인 경우 임의 보완하지 않음",
    "",
    "## 검수 안내",
    "",
    "- **자동 추출 한계**: PDF 표의 셀 경계가 모호하거나 글자가 겹치는 경우 오추출 가능",
    "- **모든 행의 `extraction_check = 'needs_review'`** 상태이므로,",
    "  데이터셋 활용 전 원본 PDF와 대조 검수가 필요하다.",
    "- 검수 완료 후 `extraction_check` 값을 `'verified'`로 업데이트 권장",
    "",
    "## 이후 활용 계획",
    "",
    "- **지오코딩**: `address` 열을 카카오·네이버 지오코딩 API에 입력하여",
    "  위도(latitude)·경도(longitude) 좌표를 추가한다.",
    "- **직선거리 계산**: 경남 일반고 146개교 좌표와 Wee센터 좌표 간",
    "  Haversine 공식으로 직선거리(km)를 산출한다.",
    "- **접근성 점수 산출**: 직선거리 기반 구간화 점수를 학교별 CSI 구성 요소에 반영한다.",
]

doc_path = DOCS_DIR / "wee_center_data_source_2025.md"
doc_path.write_text("\n".join(doc_lines), encoding="utf-8")
print(f"  → 방법론 문서 저장: {doc_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# 최종 요약
# ════════════════════════════════════════════════════════════════════════════
print()
print("=" * 60)
print("최종 요약")
print("=" * 60)
print(f"  추출 Wee센터 수     : {len(df_output)}개")
print(f"  필수 지역 누락      : {len(missing_regions)}개  {missing_regions if missing_regions else ''}")
print(f"  전화번호 결측       : {df_output['phone'].isna().sum()}개")
print(f"  주소 결측           : {df_output['address'].isna().sum()}개")
print(f"  extraction_check    : needs_review ({len(df_output)}개 전체 — 검수 필요)")
print("=" * 60)
print()
print("[생성 파일]")
print(f"  data/processed/gyeongnam_wee_centers_2025.csv")
print(f"  outputs/tables/wee_center_extraction_check_2025.csv")
print(f"  outputs/tables/wee_center_missing_summary_2025.csv")
print(f"  docs/wee_center_data_source_2025.md")
