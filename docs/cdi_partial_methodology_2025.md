# CDI 구성요소 부분 산출 방법론 문서 (2025)

## CDI 전체 구조

```
CDI = (demand_size_score + counseling_use_score + school_violence_risk_score) / 3
```

상담수요지수(CDI, Counseling Demand Index)는 학교의 상담 수요 수준을 세 가지
구성요소로 측정하는 복합 지수이다.

**이번 단계에서는 `demand_size_score`와 `counseling_use_score`만 산출하였다.**
`school_violence_risk_score`는 별도 자료 수집 후 추후 산출 예정이며,
CDI 최종 점수는 이번 단계에서 계산하지 않는다.

## 1. demand_size_score (상담수요규모점수)

### 목적

학교의 재학생 수(student_count)를 기준으로 잠재적 상담 수요 규모를 점수화한다.
규모가 클수록 절대적인 상담 수요가 많을 것으로 가정한다.

### 구간화 기준

| student_count | demand_size_score |
|---------------|-------------------|
| 250 미만 | 0.3 |
| 250 이상 ~ 500 미만 | 0.6 |
| 500 이상 | 1.0 |
| 결측 | NaN |
| 0 (실제 0값) | 확인 필요 |

### 해석 시 주의

- student_count는 상담 수요 가능성을 나타내는 규모 지표일 뿐,
  실제 위기 수준이나 미충족 상담 수요를 직접 의미하지 않는다.
- student_count는 CSI(상담공급지수) 산출 과정에서
  students_per_counselor에 간접 반영된 바 있으나,
  CDI에서는 수요 규모 지표로 별도 사용한다.

## 2. counseling_use_score (실제상담이용점수)

### 목적

학생·학부모가 실제로 학교 상담을 이용한 수준을 점수화한다.
상담 이용이 많을수록 상담 수요가 실제로 표출된 것으로 해석한다.

### 사용 지표

| 지표 | 변수명 | 의미 |
|------|--------|------|
| 3개년 평균 상담 건수 | `avg_total_counseling_count_3yr` | 2023~2025 학생·학부모 통합 상담 건수 3개년 평균 |
| 학생 수 대비 상담 건수 | `counseling_count_per_student` | 재학생 1인당 상담 건수 (규모 보정 지표) |

### Min-Max 정규화를 각각 적용한 이유

두 지표는 단위가 다르므로(건수 vs 비율) 직접 평균할 수 없다.
각 지표를 0~1 사이로 정규화한 후 평균하여 동등하게 반영한다.
두 지표 모두 값이 높을수록 상담 이용 수준이 높은 정방향 지표이므로
역정규화(1-x)하지 않는다.

### 정규화 공식

```
normalized = (x - min) / (max - min)
```

- 결측값은 정규화 계산에서 제외한다.
- max = min인 경우 해당 정규화 변수는 NaN으로 처리한다.

### counseling_use_score 산출식

```
counseling_use_score = mean(
    norm_avg_total_counseling_count_3yr,
    norm_counseling_count_per_student
)
```

### 결측 처리 기준

| 상황 | counseling_use_score |
|------|----------------------|
| 두 지표 모두 존재 | 두 지표 평균 |
| 한 지표만 존재 | 존재 지표만 사용 |
| 둘 다 결측 | NaN |

## 한계

- **규모와 위기 수준의 혼동**: 학생 수가 많다고 해서 실제 상담 위기 수준이
  반드시 높은 것은 아니다. demand_size_score는 잠재 규모 지표에 한정된다.
- **상담 건수의 다중 원인**: 상담 건수는 실제 수요뿐 아니라 학교의 상담 접근성,
  상담교사 역량, 기록 방식, 운영 체계의 영향을 받을 수 있다.
- **미충족 수요 미반영**: `counseling_use_score`는 실제 이용된 상담 수준을
  나타내며, 상담을 받지 못한 잠재적 수요(미충족 수요)는 직접 측정하지 못한다.
- **school_violence_risk_score 미산출**: 학교폭력 위험 점수는
  별도 공시 자료 수집 후 추후 산출 예정이며, CDI 최종 점수는 이후 단계에서 산출한다.