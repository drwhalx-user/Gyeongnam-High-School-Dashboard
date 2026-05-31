"""
Wee센터 접근성 점수(wee_center_access_score) 재산출 스크립트.

지오코딩과 직선거리 계산은 수행하지 않는다.
기존에 산출된 wee_center_distance_km 변수를 그대로 사용하고,
새로운 거리 구간 기준에 따라 점수만 다시 계산한다.

점수 구간 기준 (수정 후):
  wee_center_distance_km <  5      → 1.0
  5  <= distance_km < 10           → 0.7
  10 <= distance_km < 15           → 0.4
  distance_km >= 15                → 0.1
  distance_km 결측                 → NaN

수정 사유:
  실제 지오코딩 결과 최대 직선거리가 약 20km 수준으로
  기존 30km 이상 구간 해당 학교가 없어 변별력이 낮았음.
  접근성 점수 변별력 제고를 위해 15km 이상을 최저 접근성 구간으로 조정.

입력 파일:
  data/processed/gyeongnam_general_high_schools_with_wee_access_score_2025.csv

출력 파일:
  data/processed/gyeongnam_general_high_schools_with_wee_access_score_2025.csv  (덮어쓰기)
  outputs/tables/wee_center_access_score_summary_2025.csv                       (갱신)
  outputs/tables/wee_center_access_score_revised_summary_2025.csv               (신규)
  docs/wee_center_access_score_methodology_2025.md                              (갱신)
"""

import sys
import math
import pathlib

import pandas as pd

# ── 경로 설정 ────────────────────────────────────────────────────────────────
ROOT          = pathlib.Path(__file__).resolve().parent.parent
PROCESSED_DIR = ROOT / "data" / "processed"
TABLES_DIR    = ROOT / "outputs" / "tables"
DOCS_DIR      = ROOT / "docs"

for d in [PROCESSED_DIR, TABLES_DIR, DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ════════════════════════════════════════════════════════════════════════════
# STEP 1. 입력 파일 로드 및 필수 변수 확인
# ════════════════════════════════════════════════════════════════════════════
print("[STEP 1] 입력 파일 로드 및 필수 변수 확인")

INPUT_PATH = PROCESSED_DIR / "gyeongnam_general_high_schools_with_wee_access_score_2025.csv"

if not INPUT_PATH.exists():
    print(f"[ERROR] 입력 파일 없음: {INPUT_PATH}")
    sys.exit(1)

df = pd.read_csv(INPUT_PATH, encoding="utf-8-sig",
                 dtype={"school_code": str, "postcode": str})

print(f"  → {len(df)}행 × {len(df.columns)}열 로드 완료")

# wee_center_distance_km 존재 여부 확인
if "wee_center_distance_km" not in df.columns:
    print("[ERROR] 'wee_center_distance_km' 변수가 없습니다.")
    print("        지오코딩·거리 계산 스크립트를 먼저 실행하세요.")
    sys.exit(1)

dist = df["wee_center_distance_km"]
print(f"  wee_center_distance_km : 결측={dist.isna().sum()}개 / 범위={dist.min():.2f}~{dist.max():.2f} km")

# wee_center_access_score 기존 값 확인 (수정 전 분포 기록용)
if "wee_center_access_score" in df.columns:
    old_score = df["wee_center_access_score"].copy()
    print(f"  기존 wee_center_access_score 분포: {old_score.value_counts().sort_index().to_dict()}")
else:
    old_score = pd.Series([None] * len(df))
    print("  [INFO] 기존 wee_center_access_score 없음 — 신규 산출")

# ════════════════════════════════════════════════════════════════════════════
# STEP 2. 새 점수 구간 기준 정의
# 기존: <5→1.0 / 5~15→0.7 / 15~30→0.4 / ≥30→0.1
# 수정: <5→1.0 / 5~10→0.7 / 10~15→0.4 / ≥15→0.1
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 2] 새 점수 구간 기준 정의")
print("  < 5.0 km        → 1.0")
print("  5.0 ~ 10.0 km   → 0.7")
print("  10.0 ~ 15.0 km  → 0.4")
print("  >= 15.0 km      → 0.1")
print("  결측             → NaN")


def assign_access_score(distance_km) -> float:
    """
    직선거리(km) 기준으로 Wee센터 접근성 점수를 반환한다.
    결측(NaN)이면 NaN을 반환한다.
    결측 거리와 15km 이상 거리를 혼동하지 않도록 NaN 체크를 먼저 수행한다.
    """
    if pd.isna(distance_km) or math.isnan(float(distance_km)):
        return float("nan")
    d = float(distance_km)
    if d < 5.0:
        return 1.0
    elif d < 10.0:
        return 0.7
    elif d < 15.0:
        return 0.4
    else:                   # 15.0 이상 (결측은 이미 위에서 처리)
        return 0.1


# ════════════════════════════════════════════════════════════════════════════
# STEP 3. wee_center_access_score 재산출
# wee_center_distance_km 값만 사용하며 좌표·API는 건드리지 않는다.
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 3] wee_center_access_score 재산출")

df["wee_center_access_score"] = df["wee_center_distance_km"].apply(assign_access_score)

new_score = df["wee_center_access_score"]
print(f"  새 점수 분포: {new_score.value_counts().sort_index().to_dict()}")
print(f"  NaN: {new_score.isna().sum()}개")

# ════════════════════════════════════════════════════════════════════════════
# STEP 4. 최종 결과 파일 저장 (기존 파일 덮어쓰기 — 백업 생성 안 함)
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 4] 최종 결과 파일 저장")

df.to_csv(INPUT_PATH, index=False, encoding="utf-8-sig")
print(f"  → 저장 완료: {INPUT_PATH.name}  ({len(df)}행 × {len(df.columns)}열)")

# ════════════════════════════════════════════════════════════════════════════
# STEP 5. 점수 요약표 생성
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 5] 점수 요약표 생성")

# 새 기준 레이블
new_labels = {
    1.0: "1.0 — 5km 미만",
    0.7: "0.7 — 5km 이상 10km 미만",
    0.4: "0.4 — 10km 이상 15km 미만",
    0.1: "0.1 — 15km 이상",
}
# 기존 기준 레이블 (비교용)
old_labels = {
    1.0: "1.0 — 5km 미만",
    0.7: "0.7 — 5km 이상 15km 미만",
    0.4: "0.4 — 15km 이상 30km 미만",
    0.1: "0.1 — 30km 이상",
}

def make_summary(score_series, labels) -> pd.DataFrame:
    rows = []
    for sv, lbl in labels.items():
        cnt = int((score_series == sv).sum())
        rows.append({"score": sv, "label": lbl,
                     "count": cnt,
                     "ratio(%)": round(cnt / len(score_series) * 100, 2)})
    n_nan = int(score_series.isna().sum())
    rows.append({"score": float("nan"), "label": "NaN — 거리 산출 불가",
                 "count": n_nan,
                 "ratio(%)": round(n_nan / len(score_series) * 100, 2)})
    return pd.DataFrame(rows)


# 5-1. 기존 summary 파일 갱신 (새 기준)
df_new_summary = make_summary(new_score, new_labels)
summary_path = TABLES_DIR / "wee_center_access_score_summary_2025.csv"
df_new_summary.to_csv(summary_path, index=False, encoding="utf-8-sig")
print(f"  요약표(갱신) 저장: {summary_path.name}")
print(df_new_summary.to_string(index=False))

# 5-2. 기존↔수정 비교표 저장 (revised)
df_old_cnt = make_summary(old_score, old_labels).rename(
    columns={"count": "count_old", "ratio(%)": "ratio_old(%)", "label": "label_old"})
df_new_cnt = make_summary(new_score, new_labels).rename(
    columns={"count": "count_new", "ratio(%)": "ratio_new(%)", "label": "label_new"})

df_revised = pd.merge(
    df_old_cnt[["score", "label_old", "count_old", "ratio_old(%)"]],
    df_new_cnt[["score", "label_new", "count_new", "ratio_new(%)"]],
    on="score", how="outer"
).sort_values("score", na_position="last").reset_index(drop=True)

revised_path = TABLES_DIR / "wee_center_access_score_revised_summary_2025.csv"
df_revised.to_csv(revised_path, index=False, encoding="utf-8-sig")
print(f"\n  비교표(신규) 저장: {revised_path.name}")
print(df_revised.to_string(index=False))

# ════════════════════════════════════════════════════════════════════════════
# STEP 6. 방법론 문서 갱신
# ════════════════════════════════════════════════════════════════════════════
print("\n[STEP 6] 방법론 문서 갱신")

doc_lines = [
    "# Wee센터 접근성 점수 방법론 문서 (2025)",
    "",
    "## 목적",
    "",
    "경상남도 일반고등학교 146개교 각각에 대해 가장 가까운 Wee센터까지의",
    "직선거리를 산출하고, 거리 구간에 따라 Wee센터 접근성 점수를 부여한다.",
    "이 점수는 상담공급지수(CSI)의 구성 요소로 활용된다.",
    "",
    "## 사용 입력 자료",
    "",
    "| 자료 | 파일명 |",
    "|------|--------|",
    "| 경남 일반고 기본 정보 + 상담인력 공급 점수 | `gyeongnam_general_high_schools_with_staff_supply_score_2025.csv` |",
    "| 경남 Wee센터 기준 데이터셋 | `gyeongnam_wee_centers_2025.csv` |",
    "",
    "## 카카오 Local API 지오코딩 방식",
    "",
    "- **API**: 카카오 주소 검색 API (`/v2/local/search/address.json`)",
    "- **반환값**: `documents[0].x` (경도), `documents[0].y` (위도)",
    "- **API 키 관리**: 코드에 하드코딩하지 않고 프로젝트 루트 `.env` 파일의",
    "  `KAKAO_REST_API_KEY` 환경변수로 관리함. `.env`는 `.gitignore`에 포함.",
    "- **재시도 전략**:",
    "  1. 1차: 괄호 제거 + 중복 공백 제거한 주소로 시도",
    "  2. 2차: 기관명·층수·부가정보를 추가 제거한 주소로 재시도",
    "  3. 3차(수동 보정): 3개 주소(창원문성고등학교·창원 Wee센터·의령 Wee센터)는",
    "     주소 오류 수정 후 수동 재시도하여 좌표를 직접 입력.",
    "- **캐시**: 동일 주소 중복 API 호출 방지용 딕셔너리 캐시 적용",
    "- **호출 간격**: 0.05초 sleep (초당 최대 20회 이하 유지)",
    "",
    "## Haversine 공식 기반 직선거리 계산",
    "",
    "```",
    "R = 6371 km (지구 평균 반지름)",
    "a = sin²(Δlat/2) + cos(lat1)·cos(lat2)·sin²(Δlon/2)",
    "d = 2R · arcsin(√a)",
    "```",
    "",
    "- 각 학교와 지오코딩 성공한 모든 Wee센터 간 거리를 계산한 후 최솟값 선택",
    "- 학교 또는 Wee센터 좌표 결측이면 거리 = NaN",
    "",
    "## 점수 구간화 기준 (수정 후)",
    "",
    "| 거리(km) | wee_center_access_score |",
    "|---------|------------------------|",
    "| 5 미만 | 1.0 |",
    "| 5 이상 ~ 10 미만 | 0.7 |",
    "| 10 이상 ~ 15 미만 | 0.4 |",
    "| 15 이상 | 0.1 |",
    "| 거리 산출 불가 | NaN |",
    "",
    "### 수정 사유",
    "",
    "- 실제 지오코딩 결과 경남 일반고 146개교의 최근접 Wee센터까지의",
    "  최대 직선거리가 약 20.08 km 수준으로 나타남.",
    "- 기존 기준(15~30 km = 0.4, 30 km 이상 = 0.1)의 30 km 이상 구간에",
    "  해당하는 학교가 없어 점수 변별력이 낮았음.",
    "- 이에 따라 15 km 이상을 가장 낮은 접근성 구간(0.1)으로 조정하여",
    "  4개 구간 모두 실제 데이터 범위 내에서 유효하게 작동하도록 수정.",
    "- 직선거리 기준이므로 실제 이동거리 또는 이동시간과는 차이가 있을 수 있음.",
    "",
    "## 결측 처리 기준",
    "",
    "- 지오코딩 실패 → 좌표 = NaN → 거리 = NaN → 접근성 점수 = NaN",
    "- 주소 결측인 경우 임의 보완하지 않음",
    "- 지오코딩 성공이더라도 경상남도 좌표 범위(위도 34.5~35.9, 경도 127.5~129.3)",
    "  벗어난 경우 `out_of_gyeongnam_check_needed`로 표시하고 점검표에 기록",
    "",
    "## 한계 및 주의사항",
    "",
    "- **직선거리 ≠ 실제 이동거리**: Haversine 공식은 지구 곡률을 반영한 최단 경로이며,",
    "  실제 도로거리 또는 대중교통 이동시간과 다를 수 있음.",
    "- **지오코딩 오류 가능성**: 도로명주소 오타, 신설 도로 미반영 등으로",
    "  일부 주소가 잘못된 좌표로 변환될 수 있음.",
    "- **Wee센터 마산분원**: 창원시 마산회원구에 위치하는 창원교육지원청 산하 분원으로,",
    "  별도 Wee센터(WEE_GN_002)로 취급하여 독립적인 거리 계산 대상으로 포함.",
    "",
    "## 향후 개선 방향",
    "",
    "- 카카오 길찾기 API 또는 OSRM을 활용한 실제 도로거리 기반 접근성 산출",
    "- 대중교통 노선 데이터를 활용한 이동시간 기반 접근성 산출",
    "- 지오코딩 실패 학교는 수기 좌표 입력 후 재계산 권장",
]

doc_path = DOCS_DIR / "wee_center_access_score_methodology_2025.md"
doc_path.write_text("\n".join(doc_lines), encoding="utf-8")
print(f"  → 방법론 문서 저장: {doc_path.name}")

# ════════════════════════════════════════════════════════════════════════════
# 최종 요약
# ════════════════════════════════════════════════════════════════════════════
print()
print("=" * 58)
print("최종 요약")
print("=" * 58)
print(f"  전체 학교 수         : {len(df)}개교")
print(f"  score = 1.0 (5km 미만)     : {int((new_score == 1.0).sum())}개교")
print(f"  score = 0.7 (5~10km)       : {int((new_score == 0.7).sum())}개교")
print(f"  score = 0.4 (10~15km)      : {int((new_score == 0.4).sum())}개교")
print(f"  score = 0.1 (15km 이상)    : {int((new_score == 0.1).sum())}개교")
print(f"  score = NaN (거리 산출 불가): {int(new_score.isna().sum())}개교")
print(f"  점수 평균             : {new_score.mean():.4f}")
print("=" * 58)
print()
print("[갱신 파일]")
print("  data/processed/gyeongnam_general_high_schools_with_wee_access_score_2025.csv")
print("  outputs/tables/wee_center_access_score_summary_2025.csv")
print("[신규 파일]")
print("  outputs/tables/wee_center_access_score_revised_summary_2025.csv")
print("[갱신 문서]")
print("  docs/wee_center_access_score_methodology_2025.md")
