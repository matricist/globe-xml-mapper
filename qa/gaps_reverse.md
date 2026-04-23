# 역방향 갭 리포트 (XSD [R] → 코드)

`Globe요약.md`에서 `[R]`로 표시된 필수 속성이 코드(JSON 매퍼 + Services/*.cs) 어디에 등장하는지 확인.

- ✓: `.PropName` 패턴이 코드 또는 JSON target 어딘가에 존재
- ✗: 코드/JSON 어디에도 등장 없음 → 매핑 누락 의심 (또는 부모가 통째로 미사용)

> 한계: 동일 이름이 다른 타입에서 쓰이면 false-positive 가능. 부모 컨텍스트 함께 검토 필요.

| 타입 | 필수 속성 | 타입참조 | 상태 |
|---|---|---|---|
|  | `GlobeBody` | GlobeBodyType -> "GLOBEBody" | ✓ |
|  | `MessageSpec` | MessageSpecType -> "MessageSpec" | ✓ |
| DocSpecType | `DocRefId` | string | ✓ |
| DocSpecType | `DocTypeIndic` | OecdDocTypeIndicEnumType | ✓ |
| FilingInfo | `CfSofUpe` | FilingCeCofUpeEnumType | ✓ |
| FilingInfo | `CorporateStructure` | CorporateStructureType | ✓ |
| FilingInfo | `Currency` | CurrCodeType | ✓ |
| FilingInfo | `End` | DateTime | ✓ |
| FilingInfo | `ExcludedUpeStatus` | ExcludedUpeEnumType | ✓ |
| FilingInfo | `Fas` | string | ✓ |
| FilingInfo | `Id` | ExcludedUpeIdType | ✓ |
| FilingInfo | `Id` | IdType | ✓ |
| FilingInfo | `Name` | string | ✓ |
| FilingInfo | `NameMne` | string | ✓ |
| FilingInfo | `RecJurCode` | Coll<CountryCodeType> | ✓ |
| FilingInfo | `ResCountryCode` | CountryCodeType | ✓ |
| FilingInfo | `Role` | FilingCeRoleEnumType | ✓ |
| FilingInfo | `Start` | DateTime | ✓ |
| FilingInfo | `Tin` | TinType | ✓ |
| FilingInfo | `Upe` | Coll<UPE> | ✓ |
| GlobeBodyType | `FilingInfo` | GlobeBodyTypeFilingInfo | ✓ |
| GlobeTax | `Article251TopUpTax` | string | ✓ |
| GlobeTax | `Basis` | DeminimisSimpleBasisEnumType | ✓ |
| GlobeTax | `CitRate` | decimal | ✓ |
| GlobeTax | `Currency` | CurrCodeType | ✓ |
| GlobeTax | `EtrRate` | decimal | ✓ |
| GlobeTax | `ExcessProfits` | string | ✓ |
| GlobeTax | `Id` | IdType | ✓ |
| GlobeTax | `IirOffSet` | string | ✓ |
| GlobeTax | `InclusionRatio` | decimal | ✓ |
| GlobeTax | `IncomeTaxExpense` | string | ✓ |
| GlobeTax | `OtherOwnershipAllocation` | string | ✓ |
| GlobeTax | `ResCountryCode` | CountryCodeType | ✓ |
| GlobeTax | `StartDate` | DateTime | ✓ |
| GlobeTax | `Tin` | TinType | ✓ |
| GlobeTax | `Tin` | TinType | ✓ |
| GlobeTax | `Tin` | TinType | ✓ |
| GlobeTax | `TopUpTax` | string | ✓ |
| GlobeTax | `TopUpTax` | string | ✓ |
| GlobeTax | `TopUpTax` | string | ✓ |
| GlobeTax | `TopUpTaxAmount` | string | ✓ |
| GlobeTax | `TopUpTaxPercentage` | decimal | ✓ |
| GlobeTax | `TopUpTaxShare` | string | ✓ |
| GlobeTax | `Total` | string | ✓ |
| GlobeTax | `Total` | string | ✓ |
| GlobeTax | `Total` | string | ✓ |
| GlobeTax | `TotalUtprTopUpTax` | string | ✓ |
| IdType | `Art93` | bool | ✓ |
| IdType | `Change` | bool | ✓ |
| IdType | `ChangeDate` | DateTime | ✓ |
| IdType | `GlobeStatus` | Coll<IdTypeGloBeStatusEnumType> | ✓ |
| IdType | `GLoBeTax` | GlobeTax | ✓ |
| IdType | `Id` | IdType | ✓ |
| IdType | `Jurisdiction` | CountryCodeType | ✓ |
| IdType | `JurisdictionName` | CountryCodeType | ✓ |
| IdType | `Name` | string | ✓ |
| IdType | `Name` | string | ✓ |
| IdType | `OwnershipPercentage` | decimal | ✓ |
| IdType | `OwnershipType` | OwnershipTypeEnumType | ✓ |
| IdType | `PopeIpe` | PopeipeEnumType | ✓ |
| IdType | `PreGlobeStatus` | Coll<IdGlobeStatusEnumType> | ✓ |
| IdType | `RecJurCode` | Coll<CountryCodeType> | ✓ |
| IdType | `RecJurCode` | Coll<CountryCodeType> | ✓ |
| IdType | `ResCountryCode` | Coll<CountryCodeType> | ✓ |
| IdType | `Rules` | Coll<IdTypeRulesEnumType> | ✓ |
| IdType | `Tin` | Coll<TinType> | ✓ |
| IdType | `Tin` | TinType | ✓ |
| IdType | `Tin` | TinType | ✓ |
| IdType | `Type` | ExcludedEntityEnumType | ✓ |
| MessageSpecType | `MessageRefId` | string | ✓ |
| MessageSpecType | `MessageType` | MessageTypeEnumType | ✓ |
| MessageSpecType | `MessageTypeIndic` | MessageTypeIndicEnumType | ✓ |
| MessageSpecType | `ReceivingCountry` | CountryCodeType | ✓ |
| MessageSpecType | `ReportingPeriod` | DateTime | ✓ |
| MessageSpecType | `Timestamp` | DateTime | ✓ |
| MessageSpecType | `TransmittingCountry` | CountryCodeType | ✓ |
| UtprAttributionType | `AddCashTaxExpense` | string | ✓ |
| UtprAttributionType | `RecJurCode` | Coll<CountryCodeType> | ✓ |
| UtprAttributionType | `ResCountryCode` | CountryCodeType | ✓ |
| UtprAttributionType | `UtprPercentage` | decimal | ✓ |
| UtprAttributionType | `UtprTopUpTaxAttributed` | string | ✓ |
| UtprAttributionType | `UtprTopUpTaxCarriedForward` | string | ✓ |
| UtprAttributionType | `UtprTopUpTaxCarryForward` | string | ✓ |

**총 필수 속성: 83개 / 누락 의심: 0개**
