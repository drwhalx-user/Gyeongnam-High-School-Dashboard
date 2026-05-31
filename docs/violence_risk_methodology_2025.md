# 학교폭력 피해 위험 점수 산출 방법론

## 1. 개요
2023~2025학년도 학교폭력 실태조사 결과를 바탕으로 경남 일반고 146개교의
학교폭력 피해 위험 점수를 산출한다.

## 2. 입력 데이터
| 파일 | 설명 |
|------|------|
| `data/raw/school_violence_survey_2023_2025.xlsx` | 연도별 학교폭력 실태조사 결과 (3개년 438행) |
| `data/processed/gyeongnam_high_schools_master_table.xlsx` | 2025년 total_student_count 보충용 |

## 3. 변수 정의

### 3.1 연도별 계산
| 변수 | 산식 |
|------|------|
| `victim_rate` | 피해학생수 ÷ 설문참여자수 |
| `survey_participation_rate` | 설문참여자수 ÷ 전체학생수 |

### 3.2 3개년 집계
| 변수 | 산식 |
|------|------|
| `avg_victim_rate_3yr` | 3개년 victim_rate 산술평균 (결측 연도 제외) |
| `violence_years_available` | 유효 연도 수 (1~3) |
| `avg_survey_participation_rate_3yr` | 3개년 survey_participation_rate 산술평균 |

### 3.3 정규화
| 변수 | 산식 |
|------|------|
| `school_violence_risk_score` | (avg_victim_rate_3yr − min) ÷ (max − min) |

- **정규화 기준**: 경남 일반고 146개교 전체
- **해석**: 0에 가까울수록 피해율 낮음, 1에 가까울수록 피해율 높음

## 4. 결측 처리
- 2025년 total_student_count 결측: master_table의 student_count로 보충
- victim_student_count 또는 survey_participant_count 결측 → victim_rate = NaN
- 유효 연도가 0인 학교 → avg_victim_rate_3yr = NaN, school_violence_risk_score = NaN

## 5. 산출 결과 요약
- 대상 학교: 146개교
- school_violence_risk_score 비결측: 146개교
- avg_victim_rate_3yr: 최솟값=0.000000, 최댓값=0.055556
- school_violence_risk_score: 평균=0.0877

## 6. 출력 파일
| 파일 | 설명 |
|------|------|
| `data/processed/gyeongnam_high_schools_master_table.xlsx` | master_table 시트 갱신 |
| `data/processed/gyeongnam_high_schools_violence_risk.xlsx` | 위험 점수 요약 및 연도별 상세 |
| `outputs/tables/violence_risk_summary_2025.csv` | 학교별 요약 |
| `outputs/tables/violence_risk_missing_check_2025.csv` | 결측 학교 목록 |

## 7. 주의 사항
- 이 단계에서는 CDI 최종 점수를 산출하지 않는다.
- CDI = (demand_size_score + counseling_use_score + school_violence_risk_score) / 3 는 별도 스크립트에서 산출 예정.
