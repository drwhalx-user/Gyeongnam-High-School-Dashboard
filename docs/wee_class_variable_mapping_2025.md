# Wee클래스 변수 매핑표 (2025)

## API 호출 정보

| 파라미터 | 값 | 설명 |
|---|---|---|
| apiType | 61 | 학생·학부모 상담계획 및 실시 현황 |
| pbanYr | 2025 | 공시 연도 |
| sidoCode | 48 | 경상남도 |
| sggCode | 시군구별 코드 | 시도시군구코드.xlsx 참조 |
| schulKndCode | 04 | 고등학교 |
| 요청 URL 형식 | `http://www.schoolinfo.go.kr/openApi.do?apiKey={KEY}&apiType=61&pbanYr=2025&sidoCode=48&sggCode={SGG}&schulKndCode=04` | |

## Wee클래스 변수

| 표준 변수명 | 원본 API 필드 | 설명 |
|---|---|---|
| wee_class_raw | WEE_CINSTL_YN | API 원본 응답값 그대로 보존 |
| wee_class | (파생) | 이진화 결과 (1=운영, 0=미운영, NaN=확인불가) |

## wee_class 이진화 기준

| 원본값 | wee_class | 비고 |
|---|---|---|
| Y, y, 예, 설치, 운영, 있음, 유, O, o, 1 | 1 | Wee클래스 운영 |
| N, n, 아니오, 미설치, 미운영, 없음, 무, X, x, 0 | 0 | Wee클래스 미운영 |
| 공란, 기타 | NaN | 결측 처리 — 실제 0과 구분 |

## 결측 처리 기준

- `wee_class == 0` : API 응답값 N 확인 → 실제 미운영
- `wee_class == NaN` : API 미매칭 또는 원본값 불명확 → 0과 구분하여 보존

## 병합 기준

- KEDI 학교코드(KESS)와 학교알리미 학교코드(API)는 코드 체계가 상이하여 직접 병합 불가
  - KESS: 480041002 형식 (KEDI 부여 8자리 숫자)
  - API : S160000464 형식 (학교알리미 자체 코드)
- **1차: school_name 정규화 병합** (공백·괄호·특수문자 제거 후 소문자 매칭)
- 미매칭 학교는 wee_class = NaN 보존, 실행 시 검증 로그 출력

## API 전체 응답 필드 목록

| 필드명 | 예시값 |
|---|---|
| ADRCD_CD | 4812110700 |
| ADRCD_NM | 경상남도 창원시 의창구 |
| ATPT_OFCDC_ORG_CODE | S100000001 |
| ATPT_OFCDC_ORG_NM | 경상남도교육청 |
| BNHH_YN | N |
| COSE_CNSL_EXTRL_SPLST_FGR | 20 |
| COSE_CNSL_TLGM_TCR_FGR | 79 |
| EXTRL_CNSL_SPLST_OPER_YN | Y |
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

- 창원시 시 단위(sggCode=48120)는 API 데이터 없음 → 구 단위 코드(48121·48123·48125·48127·48129)로 대체 호출
- 공시 예외 학교(`PBAN_EXCP_YN=Y`)는 응답에 포함되나 실제 운영 여부 별도 확인 필요
- school_name 기준 병합으로 동명이교 존재 시 오매핑 가능 → 결과 수동 검토 권장
- API 시군구별 호출 로그: `outputs/tables/wee_class_api_call_log_2025.csv`