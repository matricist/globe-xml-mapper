# 미구현 기능 — 향후 작업 권고

AST 분석에서 vacuous로 분류된 10건 중, 실제 구현이 필요한 경우 아래 권고 절차를 따릅니다.

## 1. Sbie — ✅ 구현 완료 (1-1 단계)

- 위치: [Mapping_1.4.cs](../Services/Mapping_1.4.cs) — M열 "7. 실질기반제외소득 적용결과 추가세액 발생여부"
- 매핑: `summary.Sbie.NoTut = !ParseBool(M열값)` ("여" → NoTut=false, "부" → NoTut=true)
- `NotApplicable` 은 별도 입력칸이 없어 기본값 false (SBIE 적용 대상)
- 별도 "SBIE 미적용" 입력이 필요하면 서식에 새 열 추가 + 매퍼 보강 필요

## 2. NMCE (비중요 구성기업) — 사양 결정 필요 ⚠️

### XSD 구조
```
EtrComputationType.NonMaterialCe[] (Collection)
  - Id : IdType [R]
  - Rfy : {TotalRevenue [R], AggregateSimplified} [R]
  - Rfy1 : {TotalRevenue [R]}
  - Rfy2 : {TotalRevenue [R]}
  - Average : {TotalRevenue [R]} [R]
```

### 현 서식 상태
main_template.xlsx에 NMCE 입력칸 없음. `적용면제` 시트의 `EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc` 는 **다른 XSD 타입** (적용면제의 간소화 계산) — NMCE 와 별개.

### 필요 작업
1. 서식에 새 시트 또는 섹션 추가: "비중요 구성기업" — CE별 TIN + 국가 + 3개년(RFY/RFY-1/RFY-2) + 평균의 TotalRevenue, AggregateSimplified 입력
2. `TemplateMeta.SheetMap` 에 ("NMCE", "비중요 구성기업") 추가
3. `Mapping_NMCE.cs` 작성 — JurisdictionSection(국가별)에 연결하여 EtrComputation.NonMaterialCe 추가
4. 작성요령 시트에 안내 추가

### 우선순위
**저**. 국내 다국적기업 중 NMCE 적용 대상이 실제로 존재할 때 구현. 한국 시행령상 필수 보고 여부 확인 필요.

## 3. LowTaxJurisdiction UTPR Safe Harbour — 사양 결정 필요 ⚠️

### XSD 구조
```
LowTaxJurisdictionType.Utpr.UtprSafeHarbour
  - CitRate : decimal [R] (0~1, %)
```

### 현 서식 상태
- `적용면제` 시트에서 이미 다른 타입 `EtrTypeEtrStatusEtrExceptionUtprSafeHarbour.CitRate` 매핑 중 ([Mapping_2.cs:267](../Services/Mapping_2.cs))
- LTCE(저세율국 구성기업) 차원의 UTPR Safe Harbour 는 별도 XSD 경로. 입력칸 미정의.

### 필요 작업
1. `구성기업 계산` 시트 3.4.2 UTPR 섹션에 CIT Rate 입력 행 추가
2. `Mapping_EntityCe.cs` 의 Map342Utpr 메서드에 CIT Rate 읽기 + `js.LowTaxJurisdiction.Utpr.UtprSafeHarbour = new ... { CitRate = ... }` 할당
3. 작성요령 보강

### 우선순위
**중**. UTPR Safe Harbour 적용 국가가 있으면 CIT Rate 보고 의무가 XSD 상 필수. 비즈니스 규칙 확인 후 구현.

## 구현 판단 기준

- **당장 구현**: 입력칸이 이미 서식에 있고 매퍼만 보강하면 되는 경우 (예: Sbie)
- **보류**: 서식에 새 시트/섹션 설계가 필요한 경우 — 실제 보고 의무 발생 시점에 맞춰 진행
- **영구 미구현**: 한국 시행령상 해당 필드 보고 의무 없으면 vacuous로 남겨도 XSD 통과 (부모 요소 미생성)
