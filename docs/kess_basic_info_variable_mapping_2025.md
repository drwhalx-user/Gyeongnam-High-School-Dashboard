# KESS 기본 정보 변수 매핑표 (2025)

- 원본 파일: `kess_school_statistics_2025.xlsx`
- 시트: `학교별 주요 통계`
- 조사기준일: 2025-10-01

| 표준 변수명 | 원본 열 이름 | 설명 |
|---|---|---|
| school_code | KEDI학교코드 | KEDI 부여 고유 학교 식별 코드 |
| school_name | 학교명 | 학교 공식 명칭 |
| sido | 시도 | 광역 시·도 (경남) |
| sigungu | 행정구 | 시·군·구 단위 행정구역 |
| education_office | 교육(지원)청 | 소속 교육(지원)청 |
| school_type | 고등학교 유형 | 일반고/특성화고/특목고/자율고 구분 |
| postcode | 우편번호 | 5자리 우편번호 |
| address | 주소 | 도로명 주소 |
| student_count | 학생수_총계_계 | 전체 재학생 수(명) |
| teacher_count | 교원수_총계_계 | 전체 교원 수(정규+기간제, 명) |
| counselor_count | 교원수_정규_상담_계 | 정규직 전문상담교사 수(명) |
| students_per_counselor | (파생) | student_count / counselor_count; 상담교사 0명 또는 결측이면 NaN |

## 필터링 조건

| 조건 | 원본 열 | 값 |
|---|---|---|
| 경상남도 | 시도 | 경남 |
| 고등학교 | 학교급 | 고등학교 |
| 일반고 | 고등학교 유형 | 일반고 |
| 활성 학교 | 상태 | 폐(원)교·휴(원)교 제외 |
| 본교 | 본분교 | 본교 |
| 각종학교 제외 | 학교 세부유형 | 각종학교(고교) 제외 |

## 상담교사 0값 처리 방침

- `counselor_count == 0`: 조사기준일 기준 전문상담교사 미배치 (실제 0)
- `counselor_count` 결측: 미응답 또는 해당 없음
- `students_per_counselor`: 위 두 경우 모두 `NaN`으로 저장 (무한대·0나누기 방지)