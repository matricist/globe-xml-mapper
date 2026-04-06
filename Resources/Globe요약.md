# Globe.cs Compact Summary

**원본**: `taxcompliacnehub.Lib.Model\Globe\KR\Models\Globe.cs` (13,231 lines, 564KB+)
**생성기**: XmlSchemaClassGenerator v3.0.1215.0 (`xscgen -o ./ -n Globe ../xsd/*.xsd`)
**C# Namespace**: `Globe`

## 프로젝트 경로

| 구분 | 경로 |
|------|------|
| Globe.cs (XML 바인딩 모델) | `taxcompliacnehub.Lib.Model\Globe\KR\Models\Globe.cs` |
| Step4 API Controllers | `taxcompliancehub.Api\Areas\Globe\Controllers\Step4` |
| Filing DTOs | `taxcompliancehub.Api\Areas\Globe\Models\DTO\Filing` |
| Step4 Services | `taxcompliancehub.Api\Areas\Globe\Services\Step4` |
| GIR Entity Models | `taxcompliacnehub.Lib.Model\DBContext\GlobeEntityModel` (GIR_* 테이블) |

---

## XML Namespaces

| Prefix | URI | 용도 |
|--------|-----|------|
| isoglobetypes:v1 | `urn:oecd:ties:isoglobetypes:v1` | ISO 코드 (Country, Currency, Language) |
| globestf:v5 | `urn:oecd:ties:globestf:v5` | STF 타입 (DocSpec, y-n, DocTypeIndic, Name, Address) |
| globe:v2 | `urn:oecd:ties:globe:v2` | GloBE 본체 (모든 GIR enum/class) |

---

## 표기법

- `[R]` = Required, `[O]` = Optional
- `Coll<T>` = `Collection<T>` (생성자에서 초기화)
- `?Specified` = 동반 `XxxSpecified` 프로퍼티 존재 (`[XmlIgnore]`)
- `[date]` / `[dateTime]` = DataType
- `%` = decimal Range(0,1), 4 fraction digits
- 금액 타입은 `string` (arbitrary precision)

---

## ROOT

```
[XmlRoot("GLOBE_OECD", NS=globe:v2)]
GlobeOecd
  - MessageSpec : MessageSpecType -> "MessageSpec" [R]
  - GlobeBody : GlobeBodyType -> "GLOBEBody" [R]
  - Version : string [XmlAttribute("version")] MaxLen=10
```

---

## ENUMS (41개)

### ISO Globe Types (NS: isoglobetypes:v1)

| Enum | XmlType | 값 |
|------|---------|-----|
| CountryCodeType | CountryCode_Type | ISO-3166 Alpha-2 ~250개 (C#: PascalCase `Kr`, XML: `KR`) |
| CurrCodeType | CurrCode_Type | ISO-4217 ~170개 (C#: `Krw`, XML: `KRW`) |
| LanguageCodeType | LanguageCode_Type | ISO-639 ~180개 |
| MsCountryCodeType | MSCountryCode_Type | EU 28개국 |

### STF Types (NS: globestf:v5)

| Enum | XmlType | 값 |
|------|---------|-----|
| YN | y-n | `Yes`="yes", `No`="no" |
| OecdDocTypeIndicEnumType | OECDDocTypeIndic_EnumType | Oecd0~3(실데이터), Oecd10~13(테스트) |
| OecdNameTypeEnumType | OECDNameType_EnumType | Oecd201~208 |
| OecdLegalAddressTypeEnumType | OECDLegalAddressType_EnumType | Oecd301~305 |

### GloBE Types (NS: globe:v2)

| # | Enum | XmlType | 값 (C# → XML) |
|---|------|---------|----------------|
| 9 | MessageTypeEnumType | MessageType_EnumType | `Gir`="GIR" |
| 10 | MessageTypeIndicEnumType | MessageTypeIndic_EnumType | Gir101~103 (new/corrections/no data) |
| 11 | **IdTypeRulesEnumType** | IDTypeRules_EnumType | Gir201=QIIR-other, 202=QIIR-both, 203=QUTPR, 204=QDMTT, 205=N/A |
| 12 | **IdTypeGloBeStatusEnumType** | IDTypeGloBEStatus_EnumType | Gir301=CE, 302=FTE-TT, 303=FTE-RH, 304=Hybrid, 305=PE, 306=MainEntity, 307=MOPE, 308=MOSubsidiary, 309=MOCE, 310=Investment, 311=InsuranceInv, 312=Securitisation, 313=JV, 314=JVSubsidiary, 315=NonMaterialCE, 316=ExcludedEntity, 317=ParentQIIR10.3.5, 318=NonGroupMember |
| 13 | **FilingCeRoleEnumType** | FilingCERole_EnumType | Gir401=UPE, 402=DesignatedFiling, 403=DesignatedLocal, 404=DesignatedLocal-QDMTT, 405=DesignatedLocal-Art10.1.1 |
| 14 | **FilingCeCofUpeEnumType** | FilingCECofUPE_EnumType | Gir501=IFRS, 502=US-GAAP, 503=Other, 504=NotAvailable |
| 15 | **ExcludedUpeEnumType** | ExcludedUPE_EnumType | Gir601=Gov, 602=IntlOrg, 603=NonProfit, 604=Pension, 605=InvFund-UPE, 606=REIV-UPE |
| 16 | IdGlobeStatusEnumType | IDGloBEStatus_EnumType | Gir701~721 (21값, 301~318과 유사+GovtEntity,IntlOrg,NonProfit) |
| 17 | **OwnershipTypeEnumType** | OwnershipType_EnumType | Gir801=Parent, 802=POPE, 803=IPE, 804=MOCE, 805=JV-Group, 806=JV-External |
| 18 | **PopeipeEnumType** | POPEIPE_EnumType | Gir901=POPE, 902=IPE, 903=N/A |
| 19 | **ExcludedEntityEnumType** | ExcludedEntity_EnumType | Gir1001~1008 (Gov,IntlOrg,NonProfit,Pension,InvFund,REIV,95%owned,85%Entity) |
| 20 | TypeofSubGroupEnumType | TypeofSubGroup_EnumType | Gir1101~1106 (Standard,MinOwned,JV,Investment,InsInv,Securit) |
| 21 | **SafeHarbourEnumType** | SafeHarbour_EnumType | Gir1201~1209 |
| 22 | EtrRangeEnumType | ETRRange_EnumType | Gir1301~1314 (0%,0-5%,5-10%,...,>100%,negative,zero denom,loss) |
| 23 | QdmtTuTEnumType | QDMTTuT_EnumType | Gir1401~1409 |
| 24 | GlobeTuTEnumType | GLoBETuT_EnumType | Gir1501~1509 |
| 25 | **EtrTypeofSubGroupEnumType** | ETRTypeofSubGroup_EnumType | Gir1601~1609 |
| 26 | MainEntityPEandFteBasisEnumType | MainEntityPEandFTEBasis_EnumType | Gir1701~1704 |
| 27 | CrossBorderAdjustmentsEnumType | CrossBorderAdjustments_EnumType | Gir1801~1802 |
| 28 | UpeAdjustmentsBasisEnumType | UPEAdjustmentsBasis_EnumType | Gir1901~1910, 1999 (11값) |
| 29 | **AdjustmentItemEnumType** | AdjustmentItem_EnumType | Gir2001~2026 (26값, a~z에 대응) |
| 30 | IntShipCategoryEnumType | IntShipCategory_EnumType | Gir2101~2106 |
| 31 | AncShipCategoryEnumType | AncShipCategory_EnumType | Gir2201~2205 |
| 32 | AdjustedBasisEnumType | AdjustedBasis_EnumType | Gir2301~2309 |
| 33 | CurrentAdjustedTaxEnumType | CurrentAdjustedTax_EnumType | Gir2401~2417 (17값) |
| 34 | DeferredAdjustedTaxEnumType | DeferredAdjustedTax_EnumType | Gir2501~2516 (16값) |
| 35 | NonArt415EnumType | NONArt4.1.5_EnumType | Gir2601~2606 |
| 36 | FinalAdjustedTaxEnumType | FinalAdjustedTax_EnumType | Gir2701~2720 (20값) |
| 37 | ExTypeOfEntityEnumType | ExTypeOfEntity_EnumType | Gir2801~2805 |
| 38 | **DeminimisSimpleBasisEnumType** | DeminimisSimpleBasis_EnumType | Gir2901~2902 |
| 39 | **TinEnumType** | TIN_EnumType | Gir3001=Domestic, 3002=LEI, 3003=EIN, 3004=NotRequired |
| 40 | CurrencyEnumType | Currency_EnumType | Gir3101=Local, 3102=CFS |

---

## CLASSES (104개)

### 공통 타입

```
TinType [TIN_Type]
  - Value : string [XmlText] MaxLen=200
  - IssuedBy : CountryCodeType [XmlAttr("issuedBy")] ?Specified
  - Unknown : bool [XmlAttr("unknown")] ?Specified
  - TypeOfTin : TinEnumType [XmlAttr("TypeOfTIN")] ?Specified

DocSpecType [DocSpec_Type, NS=globestf:v5]
  - DocTypeIndic : OecdDocTypeIndicEnumType [R] -> "DocTypeIndic"
  - DocRefId : string [R] MaxLen=200 -> "DocRefId"
  - CorrDocRefId : string [O] MaxLen=200 -> "CorrDocRefId"
```

### MessageSpec

```
MessageSpecType [MessageSpec_Type]
  - SendingEntityIn : string [O] MaxLen=200 -> "SendingEntityIN"
  - TransmittingCountry : CountryCodeType [R] -> "TransmittingCountry"
  - ReceivingCountry : CountryCodeType [R] -> "ReceivingCountry"
  - MessageType : MessageTypeEnumType [R] -> "MessageType"
  - Warning : string [O] MaxLen=4000 -> "Warning"
  - Contact : string [O] MaxLen=4000 -> "Contact"
  - MessageRefId : string [R] MaxLen=170 -> "MessageRefId"
  - MessageTypeIndic : MessageTypeIndicEnumType [R] -> "MessageTypeIndic"
  - ReportingPeriod : DateTime [R][date] -> "ReportingPeriod"
  - Timestamp : DateTime [R][dateTime] -> "Timestamp"
```

### GlobeBodyType

```
GlobeBodyType [GLOBEBody_Type]
  - FilingInfo : GlobeBodyTypeFilingInfo [R] -> "FilingInfo"
  - GeneralSection : GlobeBodyTypeGeneralSection [O] -> "GeneralSection"
  - Summary : Coll<GlobeBodyTypeSummary> [O] -> "Summary" ?Specified
  - JurisdictionSection : Coll<GlobeBodyTypeJurisdictionSection> [O] -> "JurisdictionSection" ?Specified
  - UtprAttribution : Coll<GlobeBodyTypeUtprAttribution> [O] -> "UTPRAttribution" ?Specified
```

### Filing Info

```
GlobeBodyTypeFilingInfo : FilingInfo  (+DocSpec)

FilingInfo [FilingInfo]
  - FilingCe -> "FilingCE" [R]
      - ResCountryCode : CountryCodeType [R]
      - Name : string [R] MaxLen=200
      - KName : string [O] MaxLen=200
      - Tin : TinType [R] -> "TIN"
      - Role : FilingCeRoleEnumType [R]
  - AccountingInfo -> "AccountingInfo" [R]
      - CfSofUpe : FilingCeCofUpeEnumType [R] -> "CFSofUPE"
      - Fas : string [R] MaxLen=200 -> "FAS"
      - Currency : CurrCodeType [R]
  - Period -> "Period" [R]
      - Start : DateTime [R][date]
      - End : DateTime [R][date]
  - NameMne : string [R] MaxLen=200 -> "NameMNE"
  - KNameMne : string [O] MaxLen=200 -> "KNameMNE"
  - AdditionalInfo : string [O] MaxLen=4000
```

### General Section / Corporate Structure

```
GlobeBodyTypeGeneralSection : GeneralSectionType  (+DocSpec)

GeneralSectionType
  - RecJurCode : Coll<CountryCodeType> [R]
  - CorporateStructure : CorporateStructureType [R]
  - AdditionalDataPoint : Coll<AdditionalDataPointType> [O] ?Specified

CorporateStructureType
  - Upe : Coll<UPE> [R] -> "UPE"
  - Ce : Coll<CE> [O] -> "CE" ?Specified
  - ExcludedEntity : Coll<ExcludedEntity> [O] ?Specified
  - UnreportChangeCorpStr : bool [O] ?Specified

UPE (CorporateStructureTypeUpe)
  - ExcludedUpe [O] -> "ExcludedUPE"
      - ExcludedUpeStatus : ExcludedUpeEnumType [R]
      - Art1035 : CountryCodeType [O] -> "Art10.3.5" ?Specified
      - Id : ExcludedUpeIdType [R] -> "ID"
          - Name, KName, ResCountryCode[], Tin[], Rules[], GlobeStatus[]
  - OtherUpe [O] -> "OtherUPE"
      - Id : IdType [R] -> "ID"
      - Art1035 : CountryCodeType [O] -> "Art10.3.5" ?Specified

IdType [ID_Type]
  - Name : string [R] MaxLen=200
  - KName : string [O] MaxLen=200
  - ResCountryCode : Coll<CountryCodeType> [R]
  - Tin : Coll<TinType> [R] -> "TIN"
  - Rules : Coll<IdTypeRulesEnumType> [R]
  - GlobeStatus : Coll<IdTypeGloBeStatusEnumType> [R]

CE (CorporateStructureTypeCe)
  - Id : IdType [R] -> "ID"
  - OwnershipChange[] [O]
      - ChangeDate : DateTime [R][date]
      - PreGlobeStatus : Coll<IdGlobeStatusEnumType> [R]
      - PreOwnership[] [O] (OwnershipType, Tin, PreOwnershipPercentage%)
  - Ownership[] [R]
      - OwnershipType : OwnershipTypeEnumType [R]
      - Tin : TinType [R] -> "TIN"
      - OwnershipPercentage : decimal [R][%]
  - Qiir [O] -> "QIIR"
      - PopeIpe : PopeipeEnumType [R] -> "POPE-IPE"
      - Exception [O]
          - ExceptionRule (Art213:bool[O], Art215:bool[O])
          - Tin : TinType [R]
  - Qutpr [O] -> "QUTPR"
      - Art93 : bool [R] -> "Art9.3"
      - AggOwnership : decimal [O][%] ?Specified
      - UpeOwnership : bool [O] ?Specified -> "UPEOwnership"

ExcludedEntity
  - Name : string [R] MaxLen=200
  - KName : string [O] MaxLen=200
  - Type : ExcludedEntityEnumType [R]
  - Change : bool [R]
```

### Summary

```
GlobeBodyTypeSummary : SummaryType  (+DocSpec)

SummaryType
  - RecJurCode : Coll<CountryCodeType> [R]
  - Jurisdiction [R]
      - JurisdictionName : CountryCodeType [O] ?Specified
      - Subgroup[] [O] (Tin:TinType[R], TypeofSubGroup[]:TypeofSubGroupEnumType[R])
  - JurWithTaxingRights[] [O] ?Specified
      - JurisdictionName : CountryCodeType [O] ?Specified
      - DiffDomesticTut : GlobeTuTEnumType [O] ?Specified
  - SafeHarbour : Coll<SafeHarbourEnumType> [O] ?Specified
  - EtrRange : EtrRangeEnumType [O] -> "ETRRange" ?Specified
  - Sbie [O] -> "SBIE" (NotApplicable:bool[R], NoTut:bool[R])
  - QdmtTut : QdmtTuTEnumType [O] -> "QDMTTut" ?Specified
  - GLoBeTut : GlobeTuTEnumType [O] ?Specified
  - AdditionalDataPoint[] [O] ?Specified
```

### Jurisdiction Section

```
GlobeBodyTypeJurisdictionSection : JurisdictionSectionType  (+DocSpec)

JurisdictionSectionType
  - RecJurCode : Coll<CountryCodeType> [R]
  - Jurisdiction : CountryCodeType [R]
  - JurWithTaxingRights[] [O] ?Specified
      - JurisdictionName : CountryCodeType [R]
      - Subgroup[] [O] (Tin, TypeofSubGroup[])
      - ReportDifference [O]
          - EtrDifference : decimal [O][%] -> "ETRDifference" ?Specified
          - AdjCoveredTaxDifference [O] (AggCurrentTaxExpense, QrtcExpense, OtherTaxCredits, DeferTaxExpense)
          - NetGLoBEDifference, SbieDifference, AddCurrentTuTDifference, TuTDifference : string [O]
          - ElectionsDifference, KElectionsDifference : string [O] MaxLen=4000
          - QrtcIncome, ExcessNegTaxCarryForw : string [O]
          - TransitionDifference : bool [O] ?Specified
  - LocalCurrency : CurrCodeType [O] ?Specified
  - GLoBeTax : GlobeTax [R] -> "GLoBETax"
  - LowTaxJurisdiction : LowTaxJurisdictionType [O]
  - AdditionalDataPoint[] [O] ?Specified
```

### GlobeTax / ETR

```
GlobeTax [GLOBETax]
  - Etr : Coll<EtrType> [O] -> "ETR" ?Specified
  - InitialIntActivity : InitialIntActivityType [O]

EtrType
  - SubGroup [O] (Tin:TinType[R], TypeofSubGroup[]:EtrTypeofSubGroupEnumType[R])
  - EtrStatus [R] -> "ETRStatus"
      - EtrException [O] -> "ETRException"
          - DeminimisSimplifiedNmceCalc [O] -> "Deminimis-SimplifiedNMCECalc"
              - Basis : DeminimisSimpleBasisEnumType [R]
              - FinancialData[] [R] (Year[date], Revenue[O], GlobeRevenue, NetGlobeIncome, Fanil->"FANIL")
              - Average [R] (Revenue[O], GlobeRevenue, NetGlobeIncome, Fanil->"FANIL")
          - TransitionalCbCrSafeHarbour [O] -> "TransitionalCbCRSafeHarbour"
              - Revenue[O], Profit[R], IncomeTax[O]
          - UtprSafeHarbour [O] -> "UTPRSafeHarbour"
              - CitRate : decimal [R][%] -> "CITRate"
      - EtrComputation : EtrComputationType [O] -> "ETRComputation"
  - Election [O] (see below)
```

### ETR Computation (per-CE + Overall)

```
EtrComputationType
  - CeComputation[] [O] -> "CEComputation" ?Specified
  - OverallComputation [O]
  - NonMaterialCe[] [O] -> "Non-MaterialCE" ?Specified

--- Per-CE Computation ---
CEComputation
  - Tin : TinType [R] -> "TIN"
  - OtherFas : string [O] MaxLen=200 -> "OtherFAS"
  - AdjustedFanil -> "AdjustedFANIL" [R]
      - Total, Fanil->"FANIL" : string [R]
      - Adjustment [O]
          - MainEntityPEandFte[] [O] -> "MainEntityPEandFTE"
              (Basis:MainEntityPEandFteBasisEnumType, OtherTin->"OtherTIN", ResCountryCode[O], Additions, Reductions)
          - CrossBorderAdjustments[] [O]
              (Basis:CrossBorderAdjustmentsEnumType, OtherTin, ResCountryCode[O], Additions[O], Reductions[O])
          - UpeAdjustments[] [O] -> "UPEAdjustments"
              (Basis:UpeAdjustmentsBasisEnumType, Reductions[O](Amount,Exception), IdentificationOfOwners[](OwnershipPercentage%, IndOwners[O](NumOfOwners,ResCountryCode,TaxRate%), EntityOwner[O](Tin,ResCountryCode,TaxRate%,ExTypeOfEntity)))
  - NetGlobeIncome [R]
      - Total : string [R]
      - Adjustments[] [O] (Amount[]:string, AdjustmentItem:AdjustmentItemEnumType)
      - IntShippingIncome [O]
          - InternationalShipIncome (Total, Category[]:IntShipCategoryEnumType, Revenue, Costs)
          - QualifiedAncShipIncome (Total, Category:AncShipCategoryEnumType, Revenue, Costs)
          - SubstanceExclusion (PayrollCosts, TangibleAssets)
          - CoveredTaxes : string
  - AdjustedIncomeTax [R]
      - Total, IncomeTax : string [R]
      - CrossAllocation[] [O] (Basis[]:AdjustedBasisEnumType, OtherTin, ResCountryCode, Additions[O], Reductions[O])
  - AdjustedCoveredTax [R]
      - Total : string [R]
      - Adjustments[] [O] (Amount[]:string, AdjustmentItem:CurrentAdjustedTaxEnumType)
      - DeferTaxAdjustAmt [R]
          - Total, DeferTaxExpense : string [R]
          - Adjustment[] [O] (Amount[], AdjustmentItem:DeferredAdjustedTaxEnumType, Recast[O](Higher,Lower))
  - Elections [O] (see CE-level Elections below)

--- Overall Computation ---
OverallComputation
  - Fanil, AdjustedFanil -> "FANIL","AdjustedFANIL" : string [R]
  - NetGlobeIncome [R]
      - Total : string [R]
      - Adjustments[] [O] (Amount:string, AdjustmentItem:AdjustmentItemEnumType)
      - IntShippingIncome [O] (Total, TotalIntShipIncome, FiftyPercentCap, TotalQualifiedAncIncome, ExcessOfCap[O])
  - IncomeTaxExpense : string [R]
  - EtrRate : decimal [R][%] -> "ETRRate"
  - TopUpTaxPercentage : decimal [R][%]
  - AdjustedCoveredTax [O]
      - Total, AggregrateCurrentTax : string [R]
      - Adjustments[] [O] (Amount, AdjustmentItem:FinalAdjustedTaxEnumType)
      - PostFilingAdjust [O] -> DeferTaxAsset(Total, AmountAttributed[](Year,Amount)), CoveredTaxRefund(same)
      - DeemedDistTax [O] (Total, Election[O] -> Recapture[](Year,StartAmount,DDTYear-0~3,TotalDDT,EndAmount), Reduction, IncrementalTopUpTax, Ratio%)
      - DeferTaxAdjustAmt [O]
          - Total, DefTaxAmt, DiffCarryValue, GloBEValue, BefRecastAdjust, TotalAdjust, PreRecast : string [R]
          - Recast [O] (Higher, Lower)
          - Adjustments[] [O] (Amount, AdjustmentItem:DeferredAdjustedTaxEnumType, RecaptureDeferred(DTLRFYMinus5,RecapDTLRFYMinus5,DTLRFY,AggregateDTL(ReportingFiscalYear,PriorFiscalYear each: AmountPreTransition,AmountOutBalance,AmountUnjustified)))
          - Transition[] [O] (Year, DeferredTaxLiabilityStart, DeferredTaxLiabilityRecast, DeferredTaxAssets(Total,Start,Recast,Excluded), Disposal[](ResCountryCode,NetDTADTL,CarryingValue,TaxPaid,DTADTL), AltJurisdiction[O])
      - TransBlendCfc [O] -> "TransBlendCFC" (CFCJur[](Jurisdiction:CountryCodeType, Allocation(SubGroupTIN,AggAllocTax)), Total)
  - SubstanceExclusion [O] -> "SubstanceExclusion"
      - Total, PayrollCost, PayrollMarkUp%, TangibleAssetValue, TangibleAssetMarkup% : [R]
      - PeAllocation[] [O] -> "PEAllocation" (JurOfOwners(ResCountryCode,Upe,NotApplicable), PayrollCost(Total,Allocation), TangibleAssetValue(Total,Allocation))
      - FteAllocation[] [O] -> "FTEAllocation" (same structure as PE)
  - ExcessProfits : string [R]
  - AdditionalTopUpTax [O]
      - NonArt415[] [O] -> "NONArt4.1.5" (Articles[]:NonArt415EnumType, Year[date], Previous/Recalculated each: NetGlobeIncome,AdjustedCoveredTax,EtrRate%,ExcessProfits,TopUpTaxPercentage%,TopUpTax, AdditionalTopUpTax)
      - Art415 [O] -> "Art4.1.5" (AdjustedCoveredTax, GlobeLoss, ExpectedAdjustedCoveredTax, AdditionalTopUpTax)
  - Qdmtt [O] -> "QDMTT"
      - Fas [R] MaxLen=200, Amount [R], MinRate% [O], BasisforBlending [O], KBasisforBlending [O]
      - SbieAvailable, DeMinAvailable : bool [R]
      - Currency : CurrCodeType [R]
      - CurrencyElection [O] (Status:bool, ElectionYear[date], RevocationYear[O][date], Currency:CurrencyEnumType)
  - TopUpTax : string [R]
  - ExcessNegTaxExpense [R] (PriorYearBalance, GeneratedInRFY, UtilizedInRFY, Remaining)

--- Non-Material CE ---
NonMaterialCe
  - Rfy [R] -> "RFY" (TotalRevenue, AggregateSimplified[O])
  - Rfy1 [O] -> "RFY-1" (TotalRevenue)
  - Rfy2 [O] -> "RFY-2" (TotalRevenue)
  - Average [R] (TotalRevenue)
  - Id : IdType [R]
```

### CE-level Elections

```
Elections (CeComputationElections)
  - Art153 [O] -> "Art1.5.3" (Status, ElectionYear[date], RevocationYear[date])
  - SimplCalculations : bool [O] ?Specified
  - Art321 : bool [O] -> "Art3.2.1" ?Specified
  - KArt447C : bool [O] -> "KArt4.4.7.c" ?Specified
  - Art321B [O] -> "Art3.2.1b" (Status, ElectionYear, RevocationYear)
  - Art321C [O] -> "Art3.2.1c" (Status, ElectionYear, RevocationYear)
  - Art634[] [O] -> "Art6.3.4" (FyTriggerEvent[date], Inclusion(Art634CI[O],Art634CIi[O]))
  - AggregatedReporting [O] (TaxConsolGroupTin->"TaxConsolGroupTIN", EntityTin[]->"EntityTIN")
  - Art447 [O] -> "Art4.4.7" (Status, ElectionYear, RevocationYear)
  - Art456 [O] -> "Art4.5.6" (Status, ElectionYear, RevocationYear)
  - Art75[] [O] -> "Art7.5" (Status, ElectionYear, RevocationYear)
  - Art76[] [O] -> "Art7.6" (Status, ElectionYear, RevocationYear, ActualDeemedDist, LocalCreditableTaxGross, ShareOfUndistNetGlobeInc%, InvestmentEntityTin->"InvestmentEntityTIN")
```

### Jurisdiction-level Election

```
EtrTypeElection
  - Art326 : bool [O] -> "Art3.2.6" ?Specified
  - Art415 : bool [O] -> "Art4.1.5" ?Specified
  - Art461 : bool [O] -> "Art4.6.1" ?Specified
  - Art531 : bool [O] -> "Art5.3.1" ?Specified
  - Art322 [O] -> "Art3.2.2" (Status, ElectionYear[date], RevocationYear[date])
  - Art325 [O] -> "Art3.2.5" (Status, ElectionYear, RevocationYear)
  - Art328 [O] -> "Art3.2.8" (Status, ElectionYear, RevocationYear)
  - NoDefTaxAllocation [O] (Status, ElectionYear, RevocationYear[O])
  - Art45 [O] -> "Art4.5" (Status, ElectionYear, RevocationYear)
  - Art321C [O] -> "Art3.2.1.c" (Status, ElectionYear, RevocationYear[O], KEquityInvestmentInclusionElection, QualOwnerIntentBalance, Additions, Reductions, OutstandingBalance)
  - SimplifiedReporting : bool [O] ?Specified
```

### Initial International Activity

```
InitialIntActivityType
  - StartDate : DateTime [R][date]
  - ReferenceJurisdiction [R] (ResCountryCode:CountryCodeType[R], TangibleAssetValue:string[R])
  - OtherJurisdiction[] [R] (ResCountryCode[]:CountryCodeType[R], TangibleAssetValue:string[R])
  - RfyNumberOfJurisdictions : string [O] -> "RFYNumberOfJurisdictions"
  - RfySumTangibleAssetValue : string [O] -> "RFYSumTangibleAssetValue"
```

### Low Tax Jurisdiction

```
LowTaxJurisdictionType
  - TopUpTaxAmount : string [R]
  - Ltce[] [O] -> "LTCE" ?Specified
      - Tin : TinType [R]
      - Iir[] [R] -> "IIR" (not "Iir")
          - NetGlobeIncome : string [O]
          - TopUpTax : string [R]
          - ParentEntity[] [O] ?Specified
              - Tin : TinType [R]
              - ResCountryCode : CountryCodeType [R]
              - OtherOwnershipAllocation : string [R]
              - InclusionRatio : decimal [R][%]
              - TopUpTaxShare : string [R]
              - IirOffSet : string [R] -> "IIROffSet"
              - TopUpTax : string [R]
  - Utpr [O] -> "UTPR"
      - UtprSafeHarbour [O] -> "UTPRSafeHarbour" (CitRate%)
      - UtprCalculation [O] -> "UTPRCalculation"
          - TotalUtprTopUpTax : string [R] -> "TotalUTPRTopUpTax"
          - Article251TopUpTax : string [R] -> "Article2.5.1TopUpTax"
```

### UTPR Attribution

```
GlobeBodyTypeUtprAttribution : UtprAttributionType  (+DocSpec)

UtprAttributionType [UTPRAttributionType]
  - RecJurCode : Coll<CountryCodeType> [R]
  - Attribution[] [R]
      - ResCountryCode : CountryCodeType [R]
      - UtprTopUpTaxCarryForward : string [R] -> "UTPRTopUpTaxCarryForward"
      - Employees : string [O]
      - TangibleAssetValue : string [O]
      - UtprPercentage : decimal [R][%] -> "UTPRPercentage"
      - UtprTopUpTaxAttributed : string [R] -> "UTPRTopUpTaxAttributed"
      - AddCashTaxExpense : string [R]
      - UtprTopUpTaxCarriedForward : string [R] -> "UTPRTopUpTaxCarriedForward"
  - AdditionalDataPoint[] [O] ?Specified
```

### AdditionalDataPointType

```
AdditionalDataPointType
  - Description : string [O] MaxLen=170
  - Amount : string [O]
  - Percentage : decimal [O][%] ?Specified
  - Text : string [O] MaxLen=4000
  - Boolean : bool [O] ?Specified
```

---

## XML 구조 트리

```
GLOBE_OECD (root)
  +-- MessageSpec
  +-- GLOBEBody
        +-- FilingInfo (+DocSpec)
        |     +-- FilingCE (Name, KName, TIN, Role, ResCountryCode)
        |     +-- AccountingInfo (CFSofUPE, FAS, Currency)
        |     +-- Period (Start, End)
        |     +-- NameMNE, KNameMNE
        +-- GeneralSection (+DocSpec)
        |     +-- RecJurCode[]
        |     +-- CorporateStructure
        |           +-- UPE[] (ExcludedUPE | OtherUPE -> ID)
        |           +-- CE[] (ID, Ownership[], OwnershipChange[], QIIR, QUTPR)
        |           +-- ExcludedEntity[] (Name, Type, Change)
        +-- Summary[] (+DocSpec)
        |     +-- RecJurCode[], Jurisdiction, JurWithTaxingRights[]
        |     +-- SafeHarbour[], ETRRange, SBIE, QDMTTut, GLoBETut
        +-- JurisdictionSection[] (+DocSpec)
        |     +-- RecJurCode[], Jurisdiction, LocalCurrency
        |     +-- JurWithTaxingRights[] (ReportDifference)
        |     +-- GLoBETax
        |     |     +-- ETR[] (SubGroup, ETRStatus, Election)
        |     |     |     +-- ETRException (Deminimis | CbCR SafeHarbour | UTPR SafeHarbour)
        |     |     |     +-- ETRComputation
        |     |     |           +-- CEComputation[] (AdjustedFANIL, NetGlobeIncome, AdjustedIncomeTax, AdjustedCoveredTax, Elections)
        |     |     |           +-- OverallComputation (ETRRate, TopUpTaxPercentage, SubstanceExclusion, AdditionalTopUpTax, QDMTT, TopUpTax)
        |     |     |           +-- Non-MaterialCE[]
        |     |     +-- InitialIntActivity
        |     +-- LowTaxJurisdiction
        |           +-- TopUpTaxAmount
        |           +-- LTCE[] (TIN, IIR[] -> ParentEntity[])
        |           +-- UTPR (SafeHarbour | Calculation)
        +-- UTPRAttribution[] (+DocSpec)
              +-- RecJurCode[]
              +-- Attribution[] (ResCountryCode, Employees, TangibleAssetValue, UTPRPercentage, amounts)
```

## 직렬화 패턴 요약

1. **상속+DocSpec**: FilingInfo, GeneralSection, Summary, JurisdictionSection, UTPRAttribution은 base type 상속 후 DocSpec 추가
2. **?Specified 패턴**: Optional value type에 동반 `XxxSpecified` 프로퍼티 (`[XmlIgnore]`)
3. **Collection 패턴**: `Collection<T>` 생성자 초기화, private setter, `Count != 0`으로 Specified 판단
4. **Article 네이밍**: XML element에 점 포함 (e.g., "Art3.2.6", "Art4.1.5", "DDTYear-0", "NONArt4.1.5")
5. **금액 = string**: Monetary amounts는 string 타입 (arbitrary precision)
6. **비율 = decimal**: Range(0,1), 4 fraction digits
7. **한국어 로컬라이제이션**: K 접두사 프로퍼티 (KName, KNameMNE, KBasisforBlending, KElectionsDifference, KArt4.4.7.c, KEquityInvestmentInclusionElection)
