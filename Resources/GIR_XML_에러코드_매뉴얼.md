# GIR XML 에러코드 매뉴얼 (부록)
# GIR XML Error Code Manual (Appendix)

> **출처**: XML_에러코드_매뉴얼_부록_영문병기_250911_worked.pdf
> **용도**: OECD GIR (GloBE Information Return) XML 스키마 유효성 검증 에러코드 참조

---

## 목차 (Table of Contents)

1. [파일 에러 (File Error) - 50000번대](#1-파일-에러-file-error---50000번대)
2. [심각한 에러 (Critical Error) - 60000번대](#2-심각한-에러-critical-error---60000번대)
3. [기타 레코드 에러 (Other Record Error) - 70000번대](#3-기타-레코드-에러-other-record-error---70000번대)

---

## 1. 파일 에러 (File Error) - 50000번대

파일 수준의 전송/보안/스키마 검증 관련 에러

| Error Code | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|
| **50001** | — | The receiving Competent Authority could not download the referenced file. Please resubmit the file with a new unique MessageRefID. | 수신 관할당국이 해당 파일을 다운로드 받을 수 없음. 고유의 MessageRefID를 첨부하여 다시 제출 요망 |
| **50002** | — | The receiving Competent Authority could not decrypt the referenced file. Please re-encrypt the file with a valid key and resubmit. | 수신 관할당국이 해당 파일을 복호화할 수 없음. 유효한 key로 다시 암호화하여 재제출 요망 |
| **50003** | — | The receiving Competent Authority could not decompress the referenced file. Please compress the file (before encrypting) and resubmit with a new unique MessageRefID. | 수신 관할당국이 해당 파일의 압축을 풀 수 없음. (암호화 전) 파일을 압축한 뒤 고유의 MessageRefID를 첨부하여 다시 제출 요망 |
| **50004** | — | The receiving Competent Authority could not validate the digital signature. Please re-sign the file with the owner's private key using CTS procedures. | 수신 관할당국이 해당 파일의 전자서명의 유효성을 확인할 수 없음. CTS 규정 절차에 따라 파일 소유자의 private key로 다시 서명 요망 |
| **50005** | — | Potential security threats detected within the decrypted file (hyperlinks, JavaScript, executables, etc.). Scan for threats and viruses, remove all prior to encryption. | 복호화된 파일에서 하나 이상의 보안 위협 발견 (하이퍼링크, 자바스크립트, 실행 파일 등). 위험 및 바이러스를 검사·제거 후 암호화하여 재제출 요망 |
| **50006** | — | Known viruses detected within the decrypted file. Scan, remove all threats/viruses prior to encryption, re-encrypt and resubmit. | 복호화된 파일에서 알려진 바이러스 감지. 위험 및 바이러스를 검사·제거 후 암호화하여 재제출 요망 |
| **50007** | — | The file failed validation against the GIR XML Schema. Re-validate, resolve errors, re-encrypt and resubmit. | 해당 파일이 GIR XML Schema 유효성 검증을 통과하지 못함. 유효성 오류를 해결하고 암호화한 뒤 재제출 요망 |
| **50008** | `.//DocTypeIndic` | File received in test environment with DocTypeIndic value OECD0-OECD3 (production values). Test files must use OECD10-OECD13. | 테스트 환경에서 수신된 파일에 OECD0~OECD3 범위의 DocTypeIndic 값이 포함됨. 테스트 파일은 OECD10~OECD13 범위만 사용해야 함 |
| **50009** | `.//DocTypeIndic` | File contains records with DocTypeIndic OECD10-OECD13 (test data). If intended as valid GIR file, resubmit with OECD0-OECD3. | 파일에 OECD10~OECD13 범위의 DocTypeIndic 값이 포함되어 테스트 데이터임을 나타냄. 유효한 GIR 파일로 제출하려면 OECD0~OECD3으로 재제출 요망 |
| **50010** | `MessageSpec/ReceivingCountry` | Records not meant for the receiving Competent Authority. File must be deleted by erroneous receiver; notify sending CA via GIR Status Message. | Payload 파일이 수신 관할당국이 아닌 다른 관할권에 제공되어야 하는 데이터임. 잘못 수신한 관할당국이 즉시 삭제하고 GIR 상태 메시지로 통지 |
| **50011** | — | Encryption errors detected: ECB cipher mode, missing IV in Key File, incorrect AES key size (not 48 bytes), or missing concatenated Key and IV. Resend with correct encryption. | 암호화 오류 감지: ECB 방식, Key 파일 내 IV 누락, 키 크기 48바이트 불일치 등. 올바른 AES 키 크기를 적용하여 재전송 |

---

## 2. 심각한 에러 (Critical Error) - 60000번대

메시지 구조, 문서 유형, 참조 무결성 등 핵심 유효성 검증 에러

### 2.1 MessageSpec

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60001** | `MessageSpec/MessageRefId` | MessageRefId | Must be in the correct format: `[발신국가코드][보고기간][수신국가코드][고유식별번호]` | MessageRefId는 올바른 형식이어야 함: `[발신국가코드][보고기간][수신국가코드][고유식별번호]` |
| **60002** | `MessageSpec/MessageRefId` | MessageRefId | Duplicate MessageRefID value received on a previous file. Replace with new unique value. | 이전 파일에서 수신된 중복된 MessageRefID 값이 포함됨. 고유한 새 값으로 교체하여 재제출 |
| **60003** | `MessageSpec/ReportingPeriod` | ReportingPeriod | The YYYY value must be ≤ current year. | ReportingPeriod의 연도(YYYY)는 현재 연도보다 작거나 같아야 함 |

### 2.2 GLOBEBody - DocTypeIndic / DocSpec

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60004** | `.//DocTypeIndic` | DocTypeIndic | A message can contain either new records (OECD1) or corrections (OECD2/OECD3), but not a mixture of both. | 하나의 메시지에는 신규(OECD1) 또는 정정(OECD2/OECD3) 중 하나만 포함해야 하며 혼합 불가 |
| **60005** | `.//DocTypeIndic` | DocTypeIndic | When OECD2 or OECD3, the record must concern the same sub section as the CorrDocRefId. | OECD2 또는 OECD3인 경우, CorrDocRefId가 참조하는 동일한 하위 섹션에 관한 것이어야 함 |
| **60006** | `.//CorrDocRefID` | CorrDocRefId | The same DocRefID cannot be corrected or deleted twice in the same message. Each CorrDocRefID must be unique. | 동일한 DocRefID는 하나의 메시지 내에서 두 번 정정/삭제할 수 없음. CorrDocRefID는 고유해야 함 |
| **60007** | `.//DocRefId` | DocRefId | The DocRefID is already used for another record. Must be 'unique in time and space'. | 해당 DocRefID는 이미 다른 레코드에 사용됨. '시간과 공간에서 고유한' 값이어야 함 |
| **60008** | `.//CorrDocRefID` | CorrDocRefId | The CorrDocRefId refers to an unknown record. Must match an existing DocRefId. | CorrDocRefId가 알 수 없는 레코드를 참조함. 기존의 DocRefId와 정확히 일치해야 함 |
| **60009** | `.//CorrDocRefID` | CorrDocRefID | Must relate to the latest instance of the DocRefID. Cannot reference an invalidated or outdated version. | CorrDocRefID는 해당 DocRefID의 최신 인스턴스를 참조해야 함. 무효화/이전 정정된 버전 참조 불가 |
| **60010** | `GLOBEBody/FilingInfo/DocSpec/DocTypeIndic` | DocTypeIndic | FilingInfo cannot be deleted without deleting all related GeneralSection, Summary, JurisdictionSection and UTPRAttribution records. | FilingInfo는 관련된 모든 GeneralSection, Summary, JurisdictionSection 및 UTPRAttribution 레코드를 함께 삭제하지 않고는 삭제 불가 |
| **60011** | `.//DocRefId` | DocRefId | Must be in the correct format: `[발신관할권국가코드][보고연도][고유식별번호]` | DocRefId는 올바른 형식이어야 함: `[발신관할권국가코드][보고연도][고유식별번호]` |
| **60012** | `.//CorrDocRefID` | CorrDocRefID | When DocTypeIndic is OECD1 or OECD0, the CorrDocRefId field must be omitted. | DocTypeIndic이 OECD1 또는 OECD0인 경우, CorrDocRefId 필드는 생략되어야 함 |
| **60013** | GeneralSection/Summary/JurisdictionSection/UTPRAttribution DocTypeIndic | DocTypeIndic | OECD0 (Resend) may only be used for FilingInfo. Not valid for GeneralSection, Summary, JurisdictionSection, UTPRAttribution. | 재제출(OECD0)은 FilingInfo에 대해서만 사용 가능. 다른 요소에는 사용 불가 |
| **60014** | `GLOBEBody/FilingInfo/DocSpec/DocRefId` | DocRefId | When DocTypeIndic is OECD0 (Resend), DocRefID must match the latest version of the FilingInfo. | OECD0인 경우, DocRefID는 FilingInfo의 최신 버전에 사용된 DocRefID와 동일해야 함 |
| **60015** | `.//CorrDocRefId` | CorrDocRefId | When DocTypeIndic is OECD2 or OECD3, CorrDocRefId must be provided. | OECD2 또는 OECD3인 경우, CorrDocRefId를 반드시 제공해야 함 |
| **60016** | `FilingInfo/DocTypeIndic`, `GeneralSection/DocTypeIndic` | — | If FilingInfo DocTypeIndic is OECD0, GeneralSection DocTypeIndic must not be OECD1 (can't file new if one already exists). | FilingInfo가 OECD0인 경우, GeneralSection의 DocTypeIndic은 OECD1이 될 수 없음 (이미 존재하는 경우 정정/삭제로 처리) |
| **60017** | `FilingInfo/DocTypeIndic`, `GeneralSection` | — | If FilingInfo DocTypeIndic is OECD1, GeneralSection must be provided. | FilingInfo가 OECD1인 경우, GeneralSection이 반드시 포함되어야 함 |

### 2.3 RecJurCode

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60018** | `.//RecJurCode` | RecJurCode | At least one Country ISO Code in RecJurCode must match the ReceivingCountry. | RecJurCode에 포함된 국가 ISO 코드 중 적어도 하나는 ReceivingCountry와 동일해야 함 |
| **60019** | `.//RecJurCode` | RecJurCode | When FilingCE Role is GIR403, GIR404, GIR405, the GIR is only a local lodgement and should not be exchanged. | FilingCE 역할이 GIR403, GIR404, GIR405인 경우, 로컬 제출(Local Lodgement)이며 교환 대상 아님 |

### 2.4 FilingInfo

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60020** | `FilingInfo/Period/Start` vs `Period/End` | Start | Period start date must not be later than the Period end date. | 기간 시작일은 기간 종료일보다 늦어서는 안 됨 |
| **60021** | `FilingInfo/Period/End` vs `MessageSpec/ReportingPeriod` | End | Reporting Period End Date must not be later than the ReportingPeriod in the Message Header. | Reporting Period 종료일은 메시지 헤더의 ReportingPeriod 날짜보다 늦어서는 안 됨 |
| **60022** | `FilingInfo/FilingCE/TIN`, `FilingCE/Role` | Role | When role is GIR401, FilingCE TIN should match at least one TIN in UPE element (OtherUPE or ExcludedUPE). | 역할이 GIR401인 경우, FilingCE의 TIN은 UPE 요소의 TIN 중 적어도 하나와 일치해야 함 |
| **60023** | `FilingInfo/FilingCE/ResCountryCode` | ResCountryCode | The ResCountryCode of the FilingCE must match to the TransmittingCountry. | FilingCE의 ResCountryCode는 TransmittingCountry와 일치해야 함 |

### 2.5 Summary

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60024** | `Summary/JurWithTaxingRights` | JurWithTaxingRights | If Summary contains SafeHarbour, ETRRange, SBIE, QDMTTut, or GLoBETut, then JurWithTaxingRights/JurisdictionName must be provided. | Summary에 SafeHarbour, ETRRange, SBIE, QDMTTut, GLoBETut 중 하나라도 포함 시 JurWithTaxingRights/JurisdictionName 필수 |

### 2.6 JurisdictionSection - OverallComputation (ETR Calculation)

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60025** | `OverallComputation/ETRRate` | ETRRate | ETRRate = AdjustedCoveredTax/Total ÷ NetGlobeIncome/Total. Not applicable if NetGloBEIncome is zero or negative. | ETRRate = AdjustedCoveredTax/Total ÷ NetGlobeIncome/Total. NetGloBEIncome이 0 또는 음수인 경우 미적용 |

### 2.7 JurisdictionSection - OverallComputation (Top Up Tax Calculation)

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60026** | `OverallComputation/TopUpTax` | TopUpTax | TopUpTax = (TopUpTaxPercentage × ExcessProfits) + (NONArt4.1.5/AdditionalTopUpTax + Art4.1.5/AdditionalTopUpTax) - QDMTT/Amount. Missing elements treated as 0. | TopUpTax = (TopUpTaxPercentage × ExcessProfits) + (NONArt4.1.5/AdditionalTopUpTax + Art4.1.5/AdditionalTopUpTax) - QDMTT/Amount. 미제공 요소는 0으로 간주 |

### 2.8 JurisdictionSection - IIR / UTPR

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60027** | `IIR/ParentEntity/TopUpTax` | TopUpTax | IIR/ParentEntity/TopUpTax = TopUpTaxShare − IIROffset | IIR/ParentEntity/TopUpTax = TopUpTaxShare − IIROffset |

### 2.9 CEComputation - AdjustedFANIL

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **60028** | `CEComputation/AdjustedFANIL/Total` | Total | Total = AdjustedFANIL/FANIL + Σ(MainEntityPEandFTE/Additions) − Σ(MainEntityPEandFTE/Reductions) | Total = AdjustedFANIL/FANIL + Σ(MainEntityPEandFTE/Additions) − Σ(MainEntityPEandFTE/Reductions) |

---

## 3. 기타 레코드 에러 (Other Record Error) - 70000번대

개별 레코드 수준의 유효성 검증 에러

### 3.1 TIN (Taxpayer Identification Number)

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70001** | `.//TIN` | TypeOfTIN | If TypeOfTIN is GIR3004 → TIN must contain 'NOTIN', Unknown must be TRUE, IssuedBy must not be provided. | TypeOfTIN이 GIR3004인 경우: TIN='NOTIN', Unknown=TRUE, IssuedBy 미제공 |
| **70002** | `.//TIN` | TypeOfTIN | If TIN value is 'NOTIN' → TypeOfTIN must be GIR3004, Unknown must be TRUE, IssuedBy must not be provided. | TIN이 'NOTIN'인 경우: TypeOfTIN=GIR3004, Unknown=TRUE, IssuedBy 미제공 |
| **70003** | `.//TIN` | Unknown | If Unknown is TRUE → TIN must be NOTIN, TypeOfTIN must be GIR3004, IssuedBy must not be provided. | Unknown=TRUE인 경우: TIN='NOTIN', TypeOfTIN=GIR3004, IssuedBy 미제공 |
| **70004** | `.//TIN` | IssuedBy | Where the TIN has IssuedBy value, TIN validation tool applies. | TIN의 IssuedBy 값에 대한 TIN 검증 도구 적용 |
| **70005** | `.//TIN` | TypeOfTIN | Certain TIN elements must NOT have TypeOfTIN=GIR3004 or Unknown=TRUE (UPE TINs, CE TINs except GIR316/GIR318, QIIR/Exception/TIN, etc.) | 특정 TIN 요소들은 TypeOfTIN=GIR3004 또는 Unknown=TRUE가 허용되지 않음 (UPE TIN, CE TIN(GIR316/GIR318 제외), QIIR/Exception/TIN 등) |
| **70007** | `.//TIN` | TIN | When TypeOfTIN is GIR3003, reference must be in format: `P2JJYYYYMMDDCCCXXX` (P2=constant, JJ=ISO country code, YYYYMMDD=creation date, CCC=3-letter group code, XXX=unique number). | TypeOfTIN이 GIR3003인 경우 형식: `P2JJYYYYMMDDCCCXXX` |
| **70008** | `UTPRAttribution/RecJurCode` | RecJurCode | Must be the UPE jurisdiction or one of the jurisdictions in JurWithTaxingRights/JurisdictionName. | RecJurCode는 UPE 국가이거나 JurWithTaxingRights/JurisdictionName에 보고된 국가 중 하나여야 함 |

### 3.2 General Section - UPE

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70009** | `UPE/ExcludedUPE/ID/GloBEStatus`, `UPE/OtherUPE/ID/GloBEStatus` | GloBEStatus | UPE should not accept: GIR305, GIR307, GIR308, GIR309, GIR312, GIR313, GIR314, GIR315, GIR317, GIR318. Must align with GIR Note 1.3.1.6. | UPE의 GloBEStatus에는 GIR305, GIR307~GIR309, GIR312~GIR315, GIR317, GIR318 불가. GIR Note 1.3.1.6 참조 |
| **70010** | `UPE/OtherUPE/ID/ResCountryCode` | ResCountryCode | Only one value is allowed for the ResCountryCode of UPE/OtherUPE. ISO 3166-1 Alpha 2 standard. | UPE/OtherUPE의 ResCountryCode에는 하나의 값만 허용 (ISO 3166-1 Alpha 2) |

### 3.3 General Section - CE/ID

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70011** | `CE/ID/ResCountryCode` | ResCountryCode | Only one value allowed for CE ResCountryCode. ISO 3166-1 Alpha 2. | CE의 ResCountryCode에는 하나의 값만 허용 (ISO 3166-1 Alpha 2) |
| **70012** | `ExcludedUPE/ID/Rules`, `OtherUPE/ID/Rules`, `CE/ID/Rules` | Rules | All Entities in same jurisdiction must have same Rules, unless Rules contains GIR204 (QDMTT). | 동일 국가 소재 모든 기업은 동일한 Rules를 가져야 함 (GIR204/QDMTT 제외) |
| **70013** | `CE/ID/GloBEStatus` | GloBEStatus | GIR313 (JV) and GIR314 (JV Subsidiary) must not both be reported for the same CE. | GIR313(공동기업)과 GIR314(공동기업그룹 기업)는 동일 CE에 동시 보고 불가 |
| **70014** | `CE/ID/GloBEStatus` | GloBEStatus | GIR307 (Minority-Owned Parent) and GIR308 (Minority-Owned Subsidiary) must not both be reported for the same CE. | GIR307(소수지분모기업)과 GIR308(소수지분 자회사)는 동일 CE에 동시 보고 불가 |
| **70015** | `CE/ID/GloBEStatus` | GloBEStatus | When GIR308 exists, there must be another CE with GIR307 in the corporate structure. | GIR308이 있으면 기업구조 내 별도 CE에 GIR307이 존재해야 함 |
| **70016** | `CE/ID/GloBEStatus` | GloBEStatus | When GIR307 exists, GIR309 should also be reported for the same CE. | GIR307이 있으면 동일 CE에 GIR309도 함께 보고해야 함 |
| **70017** | `CE/ID/GloBEStatus` | GloBEStatus | When GIR308 exists, GIR309 should also be reported for the same CE. | GIR308이 있으면 동일 CE에 GIR309도 함께 보고해야 함 |
| **70018** | `CE/ID/GloBEStatus` | GloBEStatus | GIR305 (PE) and GIR306 (Main Entity) must not both be reported for the same CE. | GIR305(고정사업장)과 GIR306(본점)은 동일 CE에 동시 보고 불가 |
| **70019** | `CE/ID/GloBEStatus` | GloBEStatus | When GIR305 (PE) exists, there must be another CE with GIR306 (Main Entity). | GIR305(고정사업장) 존재 시, GIR306(본점)을 포함하는 다른 CE가 있어야 함 |
| **70020** | `CE/ID/GloBEStatus` | GloBEStatus | GIR316 or GIR318 can only be the sole value in GloBEStatus (no other values allowed). | GIR316 또는 GIR318은 GloBEStatus에서 유일한 값이어야 함 (다른 값 동시 불가) |
| **70021** | `CE/ID/GloBEStatus`, `CE/OwnershipChange` | GloBEStatus | GIR316 or GIR318 can only be set when there is a completed OwnershipChange for the CE. | GIR316 또는 GIR318은 해당 CE에 대해 OwnershipChange가 있어야만 설정 가능 |

### 3.4 General Section - CE/Ownership

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70026** | `CE/Ownership/OwnershipPercentage` | OwnershipPercentage | When GloBEStatus = GIR305 (PE), OwnershipPercentage must equal 100%. | GIR305(고정사업장)인 경우 OwnershipPercentage = 100% |
| **70027** | `CE/Ownership/OwnershipPercentage` | OwnershipPercentage | When GloBEStatus = GIR318 (Non-Group Member), OwnershipPercentage = 0%, TIN = 'NOTIN', OwnershipType = GIR806. | GIR318(그룹 외 기업)인 경우: OwnershipPercentage=0%, TIN='NOTIN', OwnershipType=GIR806 |
| **70028** | `CE/Ownership/OwnershipPercentage` | OwnershipPercentage | Unless GloBEStatus = GIR318, OwnershipPercentage should not be 0%. | GIR318이 아닌 경우 OwnershipPercentage는 0%가 될 수 없음 |
| **70030** | `CE/Ownership/TIN` | TIN | When OwnershipType is CE/JV/JV Subsidiary, TIN must match a TIN reported in the CorporateStructure. | 소유 구성기업 유형이 CE/JV/JV Subsidiary인 경우, TIN은 기업구조에 보고된 TIN과 일치해야 함 |
| **70031** | `CE/Ownership/TIN` | TIN | When GloBEStatus = GIR305, Ownership/TIN must match at least one TIN of an entity with GIR306 status. | GIR305인 경우, Ownership/TIN은 GIR306 상태의 법인 TIN 중 하나와 일치해야 함 |

### 3.5 General Section - CE/QIIR

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70032** | `CE/QIIR` | QIIR | When QIIR element is provided, CE Rules must contain GIR201 or GIR202. | QIIR 제공 시, CE의 Rules에 GIR201 또는 GIR202 포함 필수 |
| **70033** | `CE/QIIR/Exception/TIN` | TIN | TIN must match with a TIN reported for any other CE in CorporateStructure. | TIN은 CorporateStructure 내 다른 CE의 TIN과 일치해야 함 |
| **70034** | `CE/QIIR/Exception/Art2.1.3` | Art2.1.3 | If POPE-IPE = "GIR902 - IPE" and Exception is completed, Art2.1.3 exception should be selected. | POPE-IPE가 "GIR902 – IPE"이고 Exception이 작성된 경우, Art2.1.3 예외 선택 필요 |

### 3.6 Summary - SafeHarbour

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70037** | `Summary/Subgroup/TIN` | Subgroup | When Subgroup is completed in Summary, JurisdictionSection/Subgroup must also be completed with matching TIN. | Summary에 Subgroup이 작성된 경우, JurisdictionSection에도 일치하는 TIN의 Subgroup 필수 |
| **70038** | `Summary/SafeHarbour` | SafeHarbour | When ReportingPeriod or Period End > 2028-06-30, SafeHarbour cannot have GIR1203, GIR1204 or GIR1205 (Transitional CbCR Safe Harbour expired). | 보고기간 또는 종료일이 2028.6.30 이후면 SafeHarbour에 GIR1203/GIR1204/GIR1205 불가 (전환기 CbCR 적용면제 만료) |
| **70039** | `Summary/SafeHarbour` | SafeHarbour | When ReportingPeriod or Period End > 2026-12-31, SafeHarbour cannot have GIR1206 (Transitional UTPR Safe Harbour expired). | 보고기간 또는 종료일이 2026.12.31 이후면 SafeHarbour에 GIR1206 불가 (UTPR 적용면제 만료) |
| **70040** | `Summary/SafeHarbour` | SafeHarbour | GIR1206 can only be input in the UPE jurisdiction (JurisdictionName = UPE ResCountryCode). | GIR1206은 UPE 국가에서만 입력 가능 |
| **70041** | `Summary/SafeHarbour` | SafeHarbour | When CFSofUPE = GIR502 or GIR504, SafeHarbour cannot have GIR1207/GIR1208/GIR1209 (NMCE Simplified). | CFSofUPE가 GIR502 또는 GIR504인 경우 SafeHarbour에 GIR1207/GIR1208/GIR1209 불가 |
| **70042** | `Summary/ETRRange`, `Summary/SBIE`, `Summary/QDMTTut`, `Summary/GloBETut` | ETRRange, SBIE, QDMTTut, GloBETut | If JurWithTaxingRights is completed and SafeHarbour is not completed or only GIR1206, then ETRRange, SBIE, QDMTTut and GloBETut must all be completed. | JurWithTaxingRights 작성 + SafeHarbour 미작성(또는 GIR1206만) → ETRRange, SBIE, QDMTTut, GloBETut 모두 필수 |
| **70043** | `Summary/ETRRange`, `Summary/SBIE`, `Summary/QDMTTut` | ETRRange, SBIE, QDMTTut | If JurWithTaxingRights is completed and SafeHarbour = GIR1202, then ETRRange, SBIE and QDMTTut must be completed. | JurWithTaxingRights 작성 + SafeHarbour=GIR1202 → ETRRange, SBIE, QDMTTut 필수 |

### 3.7 JurisdictionSection

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70044** | `ETR/ETRStatus` | ETRStatus | When completed, must contain at least one of ETRException or ETRComputation. | ETRStatus 작성 시, ETRException 또는 ETRComputation 중 적어도 하나 필수 |

### 3.8 JurisdictionSection - ETRException

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70045** | `ETRException/TransitionalCbCRSafeHarbour` | TransitionalCbCRSafeHarbour | When SafeHarbour has GIR1203/1204/1205, the corresponding TransitionalCbCRSafeHarbour must be completed. | SafeHarbour가 GIR1203/1204/1205이면 해당 국가의 TransitionalCbCRSafeHarbour 작성 필수 |
| **70047** | `TransitionalCbCRSafeHarbour/Revenue` | Revenue | When SafeHarbour = GIR1203, the Revenue element must be completed for that jurisdiction. | SafeHarbour=GIR1203이면 해당 국가의 Revenue 요소 작성 필수 |
| **70048** | `TransitionalCbCRSafeHarbour/IncomeTax` | IncomeTax | When SafeHarbour = GIR1204, the IncomeTax element must be completed. | SafeHarbour=GIR1204이면 해당 국가의 IncomeTax 요소 작성 필수 |
| **70049** | `ETRException/UTPRSafeHarbour` | UTPRSafeHarbour | When SafeHarbour = GIR1206, UTPRSafeHarbour and CITRate must be completed. | SafeHarbour=GIR1206이면 UTPRSafeHarbour 및 CITRate 작성 필수 |
| **70050** | `ETRComputation/Non-MaterialCE` | Non-MaterialCE | When SafeHarbour = GIR1207/1208/1209, Non-MaterialCE must be completed. | SafeHarbour=GIR1207/1208/1209이면 Non-MaterialCE 작성 필수 |
| **70051** | `Non-MaterialCE/RFY/AggregateSimplified` | AggregateSimplified | When SafeHarbour = GIR1208, AggregateSimplified must be completed. | SafeHarbour=GIR1208이면 AggregateSimplified 작성 필수 |
| **70053** | `OverallComputation/SubstanceExclusion` | SubstanceExclusion | When SafeHarbour = GIR1205, SubstanceExclusion must be completed (unless Profit = 0 or negative). | SafeHarbour=GIR1205이면 SubstanceExclusion 필수 (Profit이 0 또는 음수인 경우 제외) |

### 3.9 JurisdictionSection - Elections

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70054** | `Election/*/RevocationYear` | RevocationYear | RevocationYear is only to be provided when Status is FALSE (election revoked). | RevocationYear는 Status가 FALSE(선택 철회)일 경우에만 제공 |
| **70055** | `Election/Art3.2.1.c/OutstandingBalance` | OutstandingBalance | OutstandingBalance = QualOwnerIntentBalance + Additions − Reductions | OutstandingBalance = QualOwnerIntentBalance + Additions − Reductions |
| **70056** | `Election/*/RevocationYear` | RevocationYear | Same as 70054 - RevocationYear only when Status is FALSE. | 70054와 동일 - RevocationYear는 Status=FALSE일 때만 |

### 3.10 OverallComputation - NetGlobeIncome

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70060** | `OverallComputation/NetGlobeIncome/IntShippingIncome` | IntShippingIncome | When AdjustmentItem is completed, IntShippingIncome must be completed. | AdjustmentItem 작성 시 IntShippingIncome 작성 필수 |

### 3.11 OverallComputation - AdjustedCoveredTax

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70061** | `AdjustedCoveredTax/Adjustments/AdjustmentItem` | AdjustmentItem | When Art4.6.1 election = TRUE, AdjustmentItem must contain GIR2711 and amount must be negative. | Art4.6.1 선택=TRUE이면 AdjustmentItem에 GIR2711 포함, amount는 음수 |
| **70062** | `AdjustedCoveredTax/Total` | Total | If AdjustmentItem = GIR2720, total AdjustedCoveredTax cannot be negative (Excess Negative Tax Expense cannot reduce below 0). | AdjustmentItem=GIR2720이면 AdjustedCoveredTax 총액은 음수 불가 |
| **70066** | `PostFilingAdjust/DeferTaxAsset/AmountAttributed/Year` | Year | Year should correspond to or be before Period Start Date YYYY value. | Year은 기간 시작일(Period Start Date) YYYY 값 이전이어야 함 |
| **70067** | `PostFilingAdjust/DeferTaxAsset/AmountAttributed/Year` | Year | If more than one AmountAttributed, years cannot be the same. | 둘 이상의 AmountAttributed 시, Year는 같을 수 없음 |
| **70068** | `PostFilingAdjust/CoveredTaxRefund/AmountAttributed/Year` | Year | Year should be ≤ Period Start Date YYYY value. | Year은 기간 시작일 YYYY 값 이전이어야 함 |
| **70069** | `PostFilingAdjust/CoveredTaxRefund/AmountAttributed/Year` | Year | If more than one AmountAttributed, years cannot be the same. | 둘 이상의 AmountAttributed 시, Year는 같을 수 없음 |

### 3.12 AdjustedCoveredTax - DeemedDistTax

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70070** | `DeemedDistTax/Election/Recapture/Year` | Year | Year cannot be after the Period End Date. | Year은 기간 종료일보다 이후 불가 |
| **70071** | `DeemedDistTax/Election/Recapture/Year` | Year | Year cannot be 4+ years before the Period End Date (Reporting FY + previous 3 FY). | Year은 기간 종료일보다 4년 이상 앞선 일자 불가 (보고 FY + 이전 3 FY) |
| **70072** | `DeemedDistTax/Election/Recapture/EndAmount` | EndAmount | EndAmount = StartAmount − TotalDDT | EndAmount = StartAmount − TotalDDT |
| **70073** | `DeemedDistTax/Election/Recapture/EndAmount` | EndAmount | EndAmount must not be negative. | EndAmount는 음수 불가 |
| **70074** | `DeemedDistTax/Election/Recapture/TotalDDT` | TotalDDT | TotalDDT = DDTYear-0 + DDTYear-1 + DDTYear-2 + DDTYear-3 | TotalDDT = DDTYear-0 + DDTYear-1 + DDTYear-2 + DDTYear-3 |
| **70075** | `DeemedDistTax/Election/Recapture/DDTYear-*` | Year | When Year = Period End Date YYYY, DDTYear-0/1/2/3 must all be "0". | Year이 기간 종료일 YYYY와 같으면 DDTYear-0/1/2/3 모두 "0" 기재 |

### 3.13 AdjustedCoveredTax - TransBlendCFC / DeferTaxAdjustAmt

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70076** | `TransBlendCFC/Total` | Total | TransBlendCFC Total = Σ(AggAllocTax) | TransBlendCFC 총액 = Σ(AggAllocTax) |
| **70077** | `DeferTaxAdjustAmt` | Total | Total = DeferredTaxAssetStart − DeferredTaxAssetExcluded OR DeferredTaxAssetRecast − DeferredTaxAssetExcluded | Total = DeferredTaxAssetStart − DeferredTaxAssetExcluded 또는 DeferredTaxAssetRecast − DeferredTaxAssetExcluded |
| **70082** | `DeferTaxAdjustAmt/Transition/DeferredTaxAssets` | DeferredTaxAsset | When DeferredTaxAssets provided, one of DeferredTaxAssetStart or DeferredTaxAssetRecast must be completed. | DeferredTaxAssets 제공 시, DeferredTaxAssetStart 또는 DeferredTaxAssetRecast 중 하나 필수 |

### 3.14 ExcessNegTaxExpense

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70084** | `AdjustedCoveredTax/Adjustments/Amount` | GeneratedInRFY | If AdjustmentItem = GIR2719, amount must equal GeneratedInRFY value. | AdjustmentItem=GIR2719이면, amount = GeneratedInRFY 값 |
| **70085** | `AdjustedCoveredTax/Adjustments/Amount` | UtilizedInRFY | If AdjustmentItem = GIR2720, amount must equal UtilizedInRFY value. | AdjustmentItem=GIR2720이면, amount = UtilizedInRFY 값 |

### 3.15 OverallComputation - Top Up Tax Calculation

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70086** | `OverallComputation/ExcessProfits` | ExcessProfits | ExcessProfits = NetGlobeIncome/Total − SubstanceExclusion/Total. If SubstanceExclusion not provided → 0. If result < 0 → ExcessProfits = 0. | ExcessProfits = NetGlobeIncome/Total − SubstanceExclusion/Total. 미제공 시 0. 결과 < 0이면 0 |
| **70087** | `SubstanceExclusion/Total` | Total | Total = (PayrollCost × PayrollMarkUp) + (TangibleAssetValue × TangibleAssetMarkup) | Total = (PayrollCost × PayrollMarkUp) + (TangibleAssetValue × TangibleAssetMarkup) |

### 3.16 AdditionalTopUpTax - Art4.1.5

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70088** | `AdditionalTopUpTax/Art4.1.5` | Art4.1.5 | Must be completed if NetGlobeIncome/Total is negative. | NetGlobeIncome/Total이 음수면 Art4.1.5 필수 |
| **70089** | `Art4.1.5/AdjustedCoveredTax` | AdjustedCoveredTax | Value should be negative. | 값은 음수이어야 함 |
| **70090** | `Art4.1.5/GlobeLoss` | GlobeLoss | Must equal NetGlobeIncome/Total value. | NetGlobeIncome/Total 값과 같아야 함 |
| **70091** | `Art4.1.5/ExpectedAdjustedCoveredTax` | ExpectedAdjustedCoveredTax | = GlobeLoss × 15% | = GlobeLoss × 15% |
| **70092** | `Art4.1.5/AdditionalTopUpTax` | AdditionalTopUpTax | = ExpectedAdjustedCoveredTax − AdjustedCoveredTax. If result < 0 → 0. | = ExpectedAdjustedCoveredTax − AdjustedCoveredTax. 결과 < 0이면 0 |

### 3.17 AdditionalTopUpTax - NONArt4.1.5

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70093** | `NONArt4.1.5/Year` | Year | Year must not be greater than Period End Date YYYY. | Year은 기간 종료일 YYYY 값보다 클 수 없음 |
| **70094** | `NONArt4.1.5/Year` | Year | When Articles = GIR2605, Year should be at least 4 years before Period End Date. | Articles=GIR2605이면, Year은 기간 종료일보다 최소 4년 이전 |
| **70095** | `NONArt4.1.5/Year` | Year | When Articles = GIR2602, Year should be the 5th FY preceding Period End Date. | Articles=GIR2602이면, Year은 기간 종료일보다 5번째 이전 회계연도 |
| **70096** | `NONArt4.1.5/AdditionalTopUpTax` | AdditionalTopUpTax | = Recalculated/TopUpTax − Previous/TopUpTax | = Recalculated/TopUpTax − Previous/TopUpTax |

### 3.18 JurisdictionSection - IIR

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70097** | `IIR/ParentEntity/InclusionRatio` | InclusionRatio | = (NetGlobeIncome − OtherOwnershipAllocation) ÷ NetGlobeIncome | = (NetGlobeIncome − OtherOwnershipAllocation) ÷ NetGlobeIncome |
| **70098** | `IIR/ParentEntity/TopUpTaxShare` | TopUpTaxShare | = IIR/TopUpTax × InclusionRatio | = IIR/TopUpTax × InclusionRatio |

### 3.19 JurisdictionSection - UTPR / UTPRAttribution

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70099** | `UTPRAttribution/Attribution/UTPRTopUpTaxAttributed` | TotalUTPRTopUpTax | Σ(UTPRTopUpTaxAttributed) should equal Σ(TotalUTPRTopUpTax) across all jurisdictions. | UTPRTopUpTaxAttributed 합계 = 모든 국가의 TotalUTPRTopUpTax 합계 |
| **70100** | `UTPRAttribution` | UTPRAttribution | When UTPRCalculation is provided and TotalUTPRTopUpTax > 0, UTPRAttribution must be completed. | UTPRCalculation 제공 + TotalUTPRTopUpTax > 0이면 UTPRAttribution 필수 |
| **70101** | `UTPRAttribution/Attribution/Employees` | Employees | Must be completed unless UTPRTopUpTaxCarryForward = 0. | UTPRTopUpTaxCarryForward=0이 아닌 한 Employees 필수 |
| **70102** | `UTPRAttribution/Attribution/TangibleAssetValue` | TangibleAssetValue | Must be completed unless UTPRTopUpTaxCarryForward = 0. | UTPRTopUpTaxCarryForward=0이 아닌 한 TangibleAssetValue 필수 |
| **70103** | `UTPRAttribution/Attribution/UTPRPercentage` | UTPRPercentage | Must be 0% when UTPRTopUpTaxCarryForward > 0. If CarryForward = 0, can only be 0% when all UTPR jurisdictions have 0%. | UTPRTopUpTaxCarryForward > 0이면 UTPRPercentage=0%. CarryForward=0이면 모든 국가가 0%일 때만 0% 가능 |
| **70104** | `UTPRAttribution/Attribution/UTPRTopUpTaxCarriedForward` | UTPRTopUpTaxCarriedForward | Cannot be negative. | 음수 불가 |
| **70105** | `UTPRAttribution/Attribution/UTPRTopUpTaxCarriedForward` | UTPRTopUpTaxCarriedForward | = UTPRTopUpTaxCarryForward + UTPRTopUpTaxAttributed − AddCashTaxExpense | = UTPRTopUpTaxCarryForward + UTPRTopUpTaxAttributed − AddCashTaxExpense |

### 3.20 CEComputation - CrossBorderAdjustments / UPEAdjustments

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70106** | `CrossBorderAdjustments/OtherTIN` | OtherTIN | Must not be the same as CEComputation/TIN. | OtherTIN은 CEComputation/TIN과 일치해서는 안 됨 |
| **70107** | `CrossBorderAdjustments` | CrossBorderAdjustments | When UPEAdjustments/Exception = TRUE, CrossBorderAdjustments must not be provided. | UPEAdjustments/Exception=TRUE이면 CrossBorderAdjustments 미제공 |
| **70109** | `UPEAdjustments/Basis` | Basis | When Basis = GIR1907, ResCountryCode must be completed (choice = IndOwners). | Basis=GIR1907이면 ResCountryCode 필수 (IndOwners 선택) |
| **70110** | `UPEAdjustments/IdentificationOfOwners/IndOwners` | IndOwners | When Basis = GIR1903 or GIR1908, IndOwners must be completed. | Basis=GIR1903 또는 GIR1908이면 IndOwners 필수 |
| **70111** | `UPEAdjustments/IdentificationOfOwners/EntityOwner/ExTypeOfEntity` | ExTypeOfEntity | When Basis = GIR1904 or GIR1909, ExTypeOfEntity must be completed (choice = EntityOwner). | Basis=GIR1904 또는 GIR1909이면 ExTypeOfEntity 필수 (EntityOwner 선택) |

### 3.21 CEComputation - NetGlobeIncome Adjustments

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70115** | `CEComputation/AdjustedFANIL/Adjustment/UPEAdjustments` | AdjustmentItem | When AdjustmentItem contains GIR2022 and/or GIR2023, UPEAdjustments must be completed. | AdjustmentItem에 GIR2022/GIR2023 포함 시 UPEAdjustments 필수 |
| **70116** | `CEComputation/NetGlobeIncome/IntShippingIncome` | AdjustmentItem | When AdjustmentItem = GIR2025, IntShippingIncome must be completed. | AdjustmentItem=GIR2025이면 IntShippingIncome 필수 |
| **70117** | `CEComputation/Elections/Art7.6` | AdjustmentItem | When AdjustmentItem = GIR2024, Art7.6 element must be completed. | AdjustmentItem=GIR2024이면 Art7.6 필수 |

### 3.22 CEComputation - AdjustedIncomeTax

| Error Code | Path (Target) | Element | Validation Rule (EN) | 유효성 규칙 (KR) |
|:---:|:---|:---|:---|:---|
| **70123** | `AdjustedIncomeTax/CrossAllocation/Additions` | Additions | Must not have a negative value. | Additions 요소는 음수 불가 |
| **70124** | `AdjustedIncomeTax/CrossAllocation/Reductions` | Reductions | Must not be a positive value. | Reductions 요소는 양수 불가 |

---

## 부록: GloBEStatus 코드 참조 (Quick Reference)

| GIR Code | Description (EN) | 설명 (KR) |
|:---|:---|:---|
| GIR305 | Permanent Establishment (PE) | 고정사업장 |
| GIR306 | Main Entity (ME) | 본점 |
| GIR307 | Minority-Owned Parent Entity | 소수지분모기업 |
| GIR308 | Minority-Owned Subsidiary | 소수지분 자회사 |
| GIR309 | Minority-Owned Constituent Entity | 소수지분 구성기업 |
| GIR310 | Investment Entity | 투자기업 |
| GIR312 | — | (기타 상태) |
| GIR313 | Joint Venture (JV) | 공동기업 |
| GIR314 | JV Subsidiary | 공동기업그룹 기업 |
| GIR315 | — | (기타 상태) |
| GIR316 | Excluded Entity | 제외기업 |
| GIR317 | — | (기타 상태) |
| GIR318 | Non-Group Member | 그룹 외 기업 |

---

## 부록: DocTypeIndic 코드 참조

| Code | Description (EN) | 설명 (KR) |
|:---|:---|:---|
| OECD0 | Resend (FilingInfo only) | 재제출 (FilingInfo에만 허용) |
| OECD1 | New Data | 신규 데이터 |
| OECD2 | Corrected Data | 정정 데이터 |
| OECD3 | Deletion | 삭제 |
| OECD10-13 | Test Data | 테스트 데이터 |

---

## 부록: SafeHarbour 코드 참조

| GIR Code | Description (EN) | 설명 (KR) |
|:---|:---|:---|
| GIR1202 | QDMTT Safe Harbour | 적격소재국추가세 적용면제 |
| GIR1203 | Transitional CbCR Safe Harbour - De Minimis | 전환기 CbCR 적용면제 - 소액요건 |
| GIR1204 | Transitional CbCR Safe Harbour - Simplified ETR Test | 전환기 CbCR 적용면제 - 간이실효세율요건 |
| GIR1205 | Transitional CbCR Safe Harbour - Routine Profits Test | 전환기 CbCR 적용면제 - 초과이익요건 |
| GIR1206 | Transitional UTPR Safe Harbour | 소득산입보완규칙 적용면제 |
| GIR1207 | NMCE Simplified Calculations | NMCE 간소화 계산 |
| GIR1208 | NMCE Simplified - Aggregate | NMCE 간소화 - 집계 |
| GIR1209 | NMCE Simplified | NMCE 간소화 |

---

> **Note**: 이 문서는 OECD GIR XML Schema 유효성 검증 에러코드의 정리본입니다.
> 실제 XML 제출 시에는 최신 OECD GIR User Guide 및 XML Schema를 반드시 참조하세요.
