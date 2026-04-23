# 역방향 갭 리포트 v3 (타입 해석기)

매퍼 코드의 LHS 체인을 Globe.cs 타입 트리로 해석하여 (Class, Prop) 단위 정밀 매칭.

- ✓ : (Class, Prop) 가 매퍼에서 직접 또는 간접 할당됨
- ✗ : 미해결 — 진짜 갭이거나 해석기가 못 따라간 패턴

## 결과

- ✓ 매칭됨: **326** / 334
- ✗ 미해결: **8** / 334

---

## ✗ 미해결

| 클래스 | 속성 | 타입 | XmlName |
|---|---|---|---|
| EtrComputationTypeNonMaterialCe | `Average` | EtrComputationTypeNonMaterialCeAverage | Average |
| EtrComputationTypeNonMaterialCe | `Id` | IdType | ID |
| EtrComputationTypeNonMaterialCe | `Rfy` | EtrComputationTypeNonMaterialCeRfy | RFY |
| EtrComputationTypeNonMaterialCeAverage | `TotalRevenue` | string | TotalRevenue |
| EtrComputationTypeNonMaterialCeRfy | `TotalRevenue` | string | TotalRevenue |
| EtrComputationTypeNonMaterialCeRfy1 | `TotalRevenue` | string | TotalRevenue |
| EtrComputationTypeNonMaterialCeRfy2 | `TotalRevenue` | string | TotalRevenue |
| LowTaxJurisdictionTypeUtprUtprSafeHarbour | `CitRate` | decimal | CITRate |
