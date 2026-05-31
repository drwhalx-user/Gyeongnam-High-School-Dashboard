# 상담 이용 변수 매핑표 (데이터 시점 2023~2025)

## API 호출 정보

| 파라미터 | 값 | 설명 |
|---|---|---|
| apiType | 61 | 학생·학부모 상담계획 및 실시 현황 |
| pbanYr | 2024, 2025, 2026 | 공시연도 (실제 데이터 시점: 2023, 2024, 2025) |
| sidoCode | 48 | 경상남도 |
| sggCode | 시군구별 코드 | 시도시군구코드.xlsx 참조 |
| schulKndCode | 04 | 고등학교 |
| 요청 URL | `http://www.schoolinfo.go.kr/openApi.do?apiKey={KEY}&apiType=61&pbanYr={PBAN}&sidoCode=48&sggCode={SGG}&schulKndCode=04` | |

## 공시연도 → 실제 데이터 시점 매핑

| pbanYr (API 파라미터) | data_year (실제 데이터 시점) |
|---|---|
| 2024 | 2023 |
| 2025 | 2024 |
| 2026 | 2025 |

> 파일명·변수명은 모두 실제 데이터 시점(2023~2025) 기준으로 표기한다.

## 상담 건수 원본 필드

| 표준 변수명 | 원본 API 필드 | 설명 |
|---|---|---|
| teacher_counseling_count | COSE_CNSL_TLGM_TCR_FGR | 전담교사 상담 건수 (연도별) |
| external_counseling_count | COSE_CNSL_EXTRL_SPLST_FGR | 외부전문가 상담 건수 (연도별) |
| total_counseling_count | (파생) | 전담+외부 합계 (연도별) |

> apiType=61에는 학생·학부모 상담 건수 별도 구분 필드 없음.
> 전담교사 상담 건수와 외부전문가 상담 건수를 합산하여 통합 상담 건수로 사용한다.

## 3개년 평균 변수

| 표준 변수명 | 설명 |
|---|---|
| avg_teacher_counseling_count_3yr | 2023·2024·2025 전담교사 상담 건수 평균 |
| avg_external_counseling_count_3yr | 2023·2024·2025 외부전문가 상담 건수 평균 |
| avg_total_counseling_count_3yr | 2023·2024·2025 통합 상담 건수 평균 |
| counseling_years_available | 평균 산출에 사용된 연도 수 (최대 3) |

- 일부 연도만 존재하는 경우 존재하는 연도의 평균 산출
- 3개년 모두 결측이면 평균값 NaN 유지

## 학생 수 대비 통합 상담 건수

- `counseling_count_per_student = avg_total_counseling_count_3yr / student_count`
- student_count가 0 또는 결측이면 NaN (0나누기 방지)

## Min-Max 정규화

- 공식: `normalized = (x - min) / (max - min)`
- 대상: `avg_total_counseling_count_3yr`, `counseling_count_per_student`
- max == min인 경우: 비결측값을 0으로 처리 (정규화 불능 명시)
- 정규화 변수명: `norm_avg_total_counseling_count_3yr`, `norm_counseling_count_per_student`

## counseling_use_score 산출 (실제 상담 이용 점수)

- `counseling_use_score = mean(norm_avg_total_counseling_count_3yr, norm_counseling_count_per_student)`
- 한 지표만 존재하면 해당 지표 단독 사용
- 두 지표 모두 결측이면 NaN
- `counseling_use_components_available`: 사용된 구성 지표 수 (0~2)

## CDI 구조 (향후 참고)

- CDI = (상담 수요 규모 점수 + 실제 상담 이용 점수 + 학교폭력 위험 점수) / 3
- 현 단계 산출 항목: **실제 상담 이용 점수 (`counseling_use_score`)**
- 상담 수요 규모 점수: 추후 학생 수 기준 구간화 또는 Min-Max 정규화 방식 결정
- 학교폭력 위험 점수: 추후 학교폭력 관련 공시자료 확보 후 산출

## 결측 처리 기준

- 상담 건수 0: 실제 0 (미운영) 유지
- 공란·변환 불가 값: NaN 처리
- 3개년 모두 결측: 평균값 NaN 유지

## 병합 기준

- KEDI 학교코드(KESS, 480XXXXX 형식)와 학교알리미 코드(API, SXXXXXXXXX 형식)는
  코드 체계가 상이하여 직접 병합 불가
- **school_name 정규화 병합**: 공백·괄호·특수문자 제거 후 소문자 매칭
- 미매칭 학교는 상담 건수 관련 변수 NaN 처리, 실행 시 검증 로그 출력

## API 전체 응답 필드 목록

| 필드명 | 예시값 |
|---|---|
| ADRCD_CD | 4812110700 |
| ADRCD_NM | 경상남도 창원시 의창구 |
| ATPT_OFCDC_ORG_CODE | S100000001 |
| ATPT_OFCDC_ORG_NM | 경상남도교육청 |
| BNHH_YN | N |
| COSE_CNSL_EXTRL_SPLST_FGR | 0 |
| COSE_CNSL_TLGM_TCR_FGR | 39 |
| EXTRL_CNSL_SPLST_OPER_YN | N |
| FOND_SC_CODE | 사립 |
| HS_KND_SC_NM | 일반고등학교 |
| INNER_CNSL_SPLST_OPER_YN | Y |
| JU_ORG_CODE | S100000001 |
| JU_ORG_NM | 경상남도교육청 |
| LCTN_SC_CODE | 17 |
| PBAN_EXCP_YN | N |
| SCHUL_CODE | S160000492 |
| SCHUL_CRSE_SC_VALUE | 4 |
| SCHUL_CRSE_SC_VALUE_NM | 고 |
| SCHUL_KND_SC_CODE | 04 |
| SCHUL_NM | 경상고등학교 |
| WEE_CINSTL_YN | Y |

## 한계 및 확인 필요 사항

- 2023년 이전 공시연도 데이터는 API에서 제공하지 않음 (최근 3년 제한, 현재 기준 2024~2026)
- school_name 정규화 병합으로 동명이교 오매핑 가능 → 결과 수동 검토 권장
- 창원시 시 단위(sggCode=48120)는 API 데이터 없음 → 구 단위 코드로 대체 호출