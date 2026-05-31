# 상담인력 공급 점수 산출 방법론 (2025)

## 목적

상담공급지수(CSI)의 구성요소 중 **상담인력 공급 점수**를 산출한다.
전문상담교사 배치 여부와 1인당 담당 학생 수를 기준으로
학교별 상담인력 공급 여건을 0.0~1.0 사이의 점수로 수치화한다.

## 사용 변수

| 변수명 | 설명 | 출처 |
|---|---|---|
| counselor_count | 전문상담교사 수 (명) | KESS 교육통계 2025 |
| student_count | 총 학생 수 (명) | KESS 교육통계 2025 |
| students_per_counselor | 상담교사 1인당 학생 수 (= student_count / counselor_count) | 파생 변수 |
| counseling_staff_supply_score | 상담인력 공급 점수 (0.0 ~ 1.0) | 본 단계 산출 |

## 점수 구간화 기준

| 점수 | 조건 | 해석 |
|---|---|---|
| 0.0 | counselor_count = 0 | 전문상담교사 미배치 |
| 0.4 | counselor_count ≥ 1 AND students_per_counselor ≥ 500 | 배치되어 있으나 1인 담당 학생 과다 |
| 0.7 | counselor_count ≥ 1 AND 250 ≤ students_per_counselor < 500 | 적정 수준 미만 |
| 1.0 | counselor_count ≥ 1 AND students_per_counselor < 250 | 공급 여건 양호 |
| NaN | counselor_count 결측, 또는 counselor_count ≥ 1이나 students_per_counselor 결측 | 확인 불가 |

## 실제 산출 결과 분포

| 점수 | 조건 요약 | 학교 수 |
|---|---|---|
| 0.0 | 전문상담교사 미배치 | 75개교 (51.4%) |
| 0.4 | 1인당 학생 500명 이상 | 43개교 (29.5%) |
| 0.7 | 1인당 학생 250~499명 | 18개교 (12.3%) |
| 1.0 | 1인당 학생 250명 미만 | 10개교 (6.8%) |

## counselor_count 0값과 결측값 처리 기준

- `counselor_count = 0` : KESS 원본에서 실제 0으로 확인 → **미배치 확정, 0.0점 부여**
  - `students_per_counselor`가 NaN이어도 counselor_count=0이 명확하므로 0.0 처리
- `counselor_count` 결측 : 미배치로 단정할 수 없으므로 → **NaN 유지**
- `students_per_counselor` inf/-inf : 0나누기 결과로 발생 가능 → **NaN으로 변환 후 처리**

## students_per_counselor 해석 방식

- 값이 **낮을수록** 상담교사 1인이 담당하는 학생 수가 적음 → 공급 여건이 좋음
- 값이 **높을수록** 1인당 부담이 과중 → 공급 여건이 나쁨
- 따라서 점수는 students_per_counselor가 낮을수록 높게 부여한다 (역방향 구간화)

## CSI 구성요소로서의 위치

상담인력 공급 점수(`counseling_staff_supply_score`)는
**상담공급지수(CSI)** 의 구성요소 중 하나로 사용된다.

```
CSI = (상담인력 공급 점수 + 실제 상담 이용 점수 + ...) / 구성요소 수
```

- 실제 상담 이용 점수(`counseling_use_score`)는 scripts/03에서 산출 완료
- 추가 구성요소는 추후 학교폭력 관련 공시자료 등을 활용하여 확장 예정

## 한계

- 상담교사 **수**와 **학생 수**만 반영하며, 다음 항목은 직접 반영하지 못함:
  - 상담의 질 (상담교사 역량, 수련 이력 등)
  - 상담교사의 실질적 업무 부담 (행정 업무, 담임 병행 등)
  - 비정규 상담 인력(계약직 상담사, 외부 연계 인력) 여부
  - Wee클래스 운영 여부와의 상호작용
- 구간화 기준(250명, 500명)은 선행 연구 및 정책 기준을 참고한 설정이며,
  추후 근거 문헌 인용 또는 민감도 분석이 필요하다.
- 분교장 및 폐교·휴교 학교는 이미 필터링되어 있으나,
  특수 학급 통합 운영 학교의 학생 수 왜곡 가능성은 별도 검토가 필요하다.