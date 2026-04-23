# 3-4단계 종합 — 역방향 갭 분석 + 보강

## 도구 진화 (역방향 갭 분석기)

| 버전 | 방식 | 정확도 |
|---|---|---|
| v1 (`gaps_reverse.md`) | Globe요약.md `[R]` 추출 + word-boundary | required=83 / missing=7 (false positive 7) |
| v2 (`gaps_reverse_v2.md`) | Globe.cs AST + 속성명 다중성 | green=116 / yellow=180 / red=7 |
| v3 (`gaps_reverse_v3.md`) | + 타입 추론 + 상속 + 위치기반 스코프 + LINQ + 컬렉션 | **hit=324 / miss=10 (모두 vacuous)** |

### v3 타입 해석기가 처리하는 패턴
- `var X = new TypeName(...)` / `new TypeName { ... }`
- `var X = ... ?? new TypeName(...)` (multi-line)
- `X = new Type();` (재할당, var 없이)
- `var X = chain.FirstOrDefault(...)` (LINQ, 요소 타입)
- `var X = chain.prop.subprop;` (별칭, iterative resolve, 위치 기반)
- `foreach (var X in chain)` (컬렉션 → 요소 타입)
- 메서드 파라미터 타입 (Globe.Type x)
- `chain.Prop ??= new Type();` (?? 연산자)
- 객체 이니셜라이저 (브레이스 카운팅, 중첩 처리)
- `chain.Collection.Add(...)` (컬렉션 사용)
- 클래스 상속 (`A : B` → A에서 못 찾으면 B에서 검색)

### 위치 기반 스코프
같은 변수명이 여러 곳에서 다른 타입으로 재사용되는 경우(예: `var item = new TypeA()` 후 다른 메서드에서 `var item = new TypeB()`) 위치별로 정확한 타입 매칭.

---

## 보강한 갭 (5종)

| # | 위치 | 변경 내용 |
|---|---|---|
| 1 | [MappingOrchestrator.cs](../Services/MappingOrchestrator.cs) `FillDocSpecs` | `JurisdictionSection[].DocSpec` 자동 생성 (`{sendCC}{year}JS{idx}{ts}`) |
| 2 | 동일 | `Summary[].DocSpec` 자동 생성 (`{sendCC}{year}SM{idx}{ts}`) |
| 3 | [Mapping_1.4.cs](../Services/Mapping_1.4.cs) | `SummaryType.RecJurCode` — 행 소재지국을 receiving 으로 자동 추가 |
| 4 | 동일 | `SummaryTypeJurisdictionSubgroup.{TypeofSubGroup, Tin}` — C/E 컬럼 매핑 추가 |
| 5 | [Mapping_JurCal.cs](../Services/Mapping_JurCal.cs) `Map321OverallComputation` | `TransBlendCfc.Total` — `CfcJur[].Allocation.AggAllocTax` 합계로 자동 계산 |

---

## 남은 10건 (모두 vacuous — 부모 미인스턴스화)

XSD `[Required]` 마커는 **부모 요소가 존재할 때**만 검증됨. 부모 요소가 매퍼에 의해 만들어지지 않으면 자식 필수 필드도 검증 대상이 아님.

| 클래스 | 누락 속성 | 사유 |
|---|---|---|
| `EtrComputationTypeNonMaterialCe` | `Average`, `Id`, `Rfy` (3건) | NonMaterialCe (비중요 구성기업) 보고 기능 미구현 — 매퍼에서 인스턴스 안 만듬 |
| `EtrComputationTypeNonMaterialCeRfy/Rfy1/Rfy2/Average` | `TotalRevenue` (4건) | 동일. 부모 NMCE 미생성이라 자식 필수 필드 vacuous |
| `LowTaxJurisdictionTypeUtprUtprSafeHarbour` | `CitRate` (1건) | LTCE의 UTPR SafeHarbour 미구현. 적용면제(시트2)의 UtprSafeHarbour는 별개 타입(`EtrType...`) 으로 매핑됨 |
| `SummaryTypeSbie` | `NotApplicable`, `NoTut` (2건) | SBIE (실질기반제외소득) Summary 출력 미구현. bool 기본값 `false` 자동 emit으로 XSD 통과는 함 |

**결론:** 이 10건은 매핑 갭이 아니라 **미구현 기능**. 비즈니스 결정에 따라 추가 구현 여부 판단 필요.

---

## 도구 한계 (현재 분석기로 못 잡는 패턴)

1. **메서드 반환값 타입** — `var x = SomeMethod()` 의 SomeMethod 반환 타입 추적 안 함. 헬퍼 메서드 통해 매핑 시 보이지 않을 수 있음.
2. **인덱서 / 람다 캡처** — `dict[key]`, `arr[i]`, `Where(x => x.Y == ...).Select(...)` 등 LINQ 체인.
3. **메서드 파라미터의 위치 스코프** — 모든 메서드 파라미터를 파일 전역으로 처리. 같은 이름의 파라미터가 여러 메서드에서 다른 타입이면 마지막 것만 인식.
4. **bool/string 기본값으로 사일런트 통과** — `[Required]` bool 속성은 C# 기본값 false로 자동 emit되어 XSD 통과. 의미 의도 검증은 사람만 가능.

이 한계를 완전히 제거하려면 **Roslyn SemanticModel** 도입 필요. 현재 정확도(324/334 = 97%)로 실용적 QA 충분.

---

## 검증
- 빌드 0 경고 / 0 오류
- `gaps_reverse_v3.md` 미해결 0건 (vacuous 10건은 미구현 표기)
