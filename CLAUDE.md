# Globe XML Mapper - 작업 가이드

## 프로젝트 개요
국조 서식(.xlsx) → OECD GloBE XML 변환 WinForms 프로그램.
**단일 xlsx 파일** — ClosedXML로 파일 직접 읽어 XML 변환.

## UX 흐름
```
1. 메인 화면: [템플릿 다운로드] / [XML 변환]
2. 템플릿 다운로드 → main_template.xlsx 복사
3. XML 변환 → main_template.xlsx 선택 → 매핑 + 검증 → XML 저장
```

## 숨김시트 (_META)
```
blockCount:{시트명}   |  값(정수)
```
시트별 블록 수 보존용. `EnsureMetaSheet()`가 자동 생성/유지.

## 프로젝트 구조
```
mapper/
├── GlobeMapper.csproj
├── Program.cs
├── MainForm.cs                     # 메인 화면 (템플릿 다운로드 / XML 변환)
├── TermsDialog.cs
├── Services/
│   ├── ExcelController.cs          # Excel COM 래퍼 (블록 추가/삭제, _META)
│   ├── MappingBase.cs              # 공통 유틸 (ParseTin, ParseBool 등)
│   ├── EntityGroupMap.cs           # "기업매핑" 시트 리더 (entity TIN → 국가/하위그룹)
│   ├── Mapping_1.1~1.2.cs
│   ├── Mapping_1.3.1.cs            # UPE
│   ├── Mapping_1.3.2.1.cs          # CE (O11 통합 셀 소유지분)
│   ├── Mapping_1.3.2.2.cs          # 제외기업
│   ├── Mapping_1.3.3.cs            # 기업구조 변동 (XML 미포함)
│   ├── Mapping_1.4.cs              # Summary
│   ├── Mapping_2.cs                # 적용면제
│   ├── Mapping_Utpr.cs
│   ├── Mapping_JurCal.cs           # "국가별 계산" 시트 — 3.1~3.3 (세로 스택 블록)
│   ├── Mapping_EntityCe.cs         # "구성기업 계산" 시트 — 3.2.4 + 3.4 (세로 스택 블록)
│   ├── MappingOrchestrator.cs      # 섹션 순회 + MessageSpec/DocSpec 채움
│   ├── ValidationUtil.cs
│   └── XmlExportService.cs
├── Resources/
│   ├── Globe.cs / Globe요약.md     # XSD 생성 클래스 + 요약
│   ├── XSD/
│   ├── mappings/                   # 섹션별 매핑 JSON
│   ├── main_template.xlsx          # 단일 통합 템플릿
│   ├── terms.txt / expired_message.txt / activation_config.json
│   └── GIR_XML_에러코드_매뉴얼.md
└── Tools/                          # 개발/검증용 C# 콘솔 도구 (빌드에서 제외됨)
```

## 핵심 클래스

### ExcelController (Services/ExcelController.cs)
- `Open(path)` / `AttachToActive()` / `CreateNew(template, savePath)` — Excel 연결
- `AddRowBlock` / `RemoveRowBlock` / `GetRowBlockCount` — 헤더 키워드 기반 블록 추가/삭제
- `AddCeBlock` / `RemoveCeBlock` / `ResetCeSheet` / `GetCeBlockCount` — 그룹구조 CE 블록
- `AddSheet2Block` / `RemoveSheet2Block` / `ResetSheet2` — 적용면제 52행 복합 블록
- `GetFilePathForMapping()` — 저장 후 파일 경로 반환 (ClosedXML용)
- `CloseWithSavePrompt()` — 저장 확인 후 종료

### MappingOrchestrator
- `MapWorkbook(filePath, globe)` — 단일 main_template.xlsx 기반 매핑 (폴더/group/entity 개념 없음)

### 매핑 행 탐색 원칙 (Mapping_JurCal 등)
- **절대 행번호 금지** — 행 삽입/삭제에 취약하므로 반드시 `FindRow(ws, "헤더텍스트")` 로 위치를 동적 탐색
- 헤더 발견 후 바로 아래 데이터가 고정 구조일 때만 `rowHeader + N` 오프셋 허용 (예: "(a)" 바로 다음 1행이 항상 값 행인 경우)
- 여러 항목이 각자 레이블을 가질 때는 항목마다 별도 `FindRow`로 탐색 (예: "1. 모든 구성기업...", "2. 50% 한도" 등)
- 값 열(column)도 실제 템플릿 덤프로 확인 — O열(15)이 아닌 N열(14) 등 다를 수 있음

### CE 행 블록 반복 (그룹구조 시트)
- 그룹구조 시트에서 3~21행이 CE 1건 블록 (19행). [+]로 블록 복제
- ExcelController 상수: CE_BLOCK_START=3, CE_BLOCK_END=21, CE_BLOCK_GAP=2
- 소유지분은 블록 내 O11 통합 셀(offset 8, 병합 O11:R14)에 인라인 입력 — 별첨 시트 없음
  - 포맷: `유형,TIN,TIN유형,발급국가,지분` × N (주주 구분: `;`)
  - 예: `GIR801, 1234567890, GIR3001, KR, 1; GIR802, 987, GIR3001, KR, 0.5`
- Mapping_1.3.2.1이 blockCount 기반으로 N개 CE 순회, O12에서 Ownership 파싱

### 세로 스택 블록 (국가별 계산 / 구성기업 계산)
- `국가별 계산` 시트: "3.1 국가별" 헤더마다 새 합산단위 블록 (259행 단위)
- `구성기업 계산` 시트: "1. 구성기업 또는 공동기업그룹 기업의 납세자번호" 헤더마다 새 entity 블록 (167행 단위)
- Mapping_JurCal / Mapping_EntityCe 둘 다 `_blockStart/_blockEnd` 인스턴스 필드로 FindRow 스코핑
  - 블록 탐지 → 각 블록마다 `_blockStart/_blockEnd` 설정 → 하위 Map*() 메서드 호출
  - FindRow는 블록 범위 내에서만 검색 (전역 검색 방지)

### 기업매핑 시트 (main_template.xlsx)
- B: 기업 납세자번호 / C: 국가명 (ISO2) / D: 하위그룹 유형 / E: 하위그룹 최상위기업 TIN
- Mapping_EntityCe가 각 entity 블록의 TIN 값으로 조회 → 해당 JurisdictionSection(국가) + ETR(하위그룹 TIN) 찾기
- `EntityGroupMap.Load(workbook)`가 시트 읽기 캡슐화

## 새 섹션 추가 절차
1. 매퍼 클래스: `Services/Mapping_{섹션}.cs` (MappingBase 상속)
2. `MappingOrchestrator.MapperFactory`에 등록
3. `ExcelController.SheetMap`에 (section, sheetName) 추가
4. `ControlPanelForm.UpdateDynamicPanel`에 시트 분기 추가 (블록 +/- 버튼)
5. `ValidationUtil`에 검증 규칙 추가
6. CLAUDE.md 업데이트

## 구현 진행 상황 (main_template.xlsx 시트별)

### 기본 섹션 (1.x)
- [x] 1.1~1.2: 신고구성기업, 사업연도, 회계정보
- [x] 1.3.1: 최종모기업 (UPE) — 행 블록 반복
- [x] 1.3.2.1: 구성기업 (CE) — 행 블록 반복 + O11 통합 셀 소유지분 (첨부 시트 제거됨)
- [x] 1.3.2.2: 제외기업 — 행 블록 반복 (블록 범위 2~5행, 여/부 ParseBool, Name;KName 분리)
- [x] 1.3.3: 기업구조 변동 — 단순 행 추가 (XML 미포함)
- [x] 1.4: 정보 요약 (Summary) — 단순 행 추가

### 국가별 계산 시트 (Mapping_JurCal, main_template.xlsx 내부)
- [x] 2: 국가별 적용면제 — 52행 복합 블록 반복 + Summary 상세 (적용면제 첨부 제거됨, 2.3 item 5 OtherJurisdiction은 M46 통합 셀 `국가,값; 국가,값` 포맷)
- [x] 3.1: 기본사항 (국가명/하위그룹/JurWithTaxingRights)
  - 3.1.4~3.1.18 ReportDifference: 매핑 미구현 (XSD 상 full optional, 국가별 계산 첨부 시트 제거됨)
- [x] 3.2.1: 실효세율 계산 (FANIL, 글로벌최저한세소득/결손 26개 조정, 조정대상조세 20개, 이월, CFC)
- [x] 3.2.2.1: 이연법인세조정 요약/내역 16개
- [x] 3.2.2.1(c): 결손금 소급공제
- [x] 3.2.2.2: 환입금액 계산
- [x] 3.2.2.3: 최초적용연도 특례
- [x] 3.2.3.1: 국가별 선택 (매년/5년/그 밖의 + 필요정보)
- [x] 3.2.3.2: 간주분배세액
- [x] 3.2.4.4(b) 적격국제해운부수소득 (IntShippingIncome) — 국가별 계산 시트 내부로 통합됨 (FindRow "(b) 적격국제해운부수소득" 기반)
- [x] 3.3 추가세액 계산 (3.3.1 요약, 3.3.2 실질기반제외소득+PE/FTE배분, 3.3.3 당기추가세액가산액, 3.3.4 QDMTT) — 국가별 계산 시트 내부로 통합됨 (FindRow "3.3.1 추가세액" 등 기반)

### 구성기업 계산 시트 (Mapping_EntityCe.cs, main_template.xlsx 내부)
- [x] CE TIN 식별: 3.2.4.1(a) "납세자번호" 항목 M열 ("값,GIR300x,발급국가")
  - 기업매핑 시트에서 TIN으로 (국가, 하위그룹 TIN) 조회 → JurisdictionSection 연결
  - 폴백: 기업매핑에 없으면 TIN issuedBy 국가코드 사용
  - JurCal이 먼저 처리되어 JurisdictionSection이 선행 생성됨 (SheetMap 순서 보장)
- [x] 3.2.4(a): 전환기 국가별 간소화 신고체계 선택 K열 → Elections.SimplCalculations (여/부)
- [x] 3.2.4(b): 연결납세그룹 통합신고 → Elections.AggregatedReporting
  - 헤더 다음 행들: B열=GroupTIN, K열=EntityTINs (각 행마다 1개)
- [x] 3.2.4.1(a): 납세자번호 (AdjustedFANIL/FANIL, AdjustedIncomeTax, AdjustedCoveredTax, NetGlobeIncome 상세 포함)
- [x] 3.2.4.1(b): FANIL 조정 — MainEntityPEandFTE, CrossBorderAdjustments, UPEAdjustments (K열: "값,열거코드,상대방TIN,국가")
- [x] 3.2.4.1(c): 법인세조정 CrossAllocation (K열: "값,GIR230x,TIN,국가,issuedBy,TIN유형")
- [x] 3.2.4.1(d): UPEAdjustments IdentificationOfOwners (K열: "개인/법인,..." 유형별 파싱)
- [x] 3.2.4.2(a): NetGlobeIncome 조정 (K열: "값,GIR200x")
- [x] 3.2.4.2(b): AdjustedCoveredTax 조정 (K열: "값,GIR240x")
- [x] 3.2.4.2(c): DeferTaxAdjustAmt (K열: "값,GIR250x", Recast: Lower/Higher)
- [x] 3.2.4.3: 구성기업별 선택 (Elections a~k)
  - a~c: 매년선택 (O열 여/부) — SimplCalculations, Art321, KArt447C
  - d~i: 5년선택 (O=선택연도, Q=취소연도) — Art153, Art321B, Art321C, Art75, Art76, Art447
  - j: 기타선택 — Art456
  - k: 공정가액조정 (E=사업연도, J=i/ii 드롭다운) — Art634[]
- [x] 3.2.4.4: 국제해운소득·결손 제외 → NetGlobeIncome.IntShippingIncome
  - O(15)열 입력, GIR2101~2106 Category(복수, 정규식 추출), GIR2201~2205 부수소득 Category
- [x] 3.2.4.5: 과세분배방법 적용 선택 관련 정보 → Elections.Art76[]
  - B(2)=주주CE TIN, E(5)=투자기업TIN, G(7)=ActualDeemedDist, K(11)=LocalCreditableTaxGross, N(14)=ShareOfUndistNetGlobeInc
  - Status/ElectionYear/RevocationYear: 3.2.4.3 h행(r5Year+5)에서 공통 적용
- [x] 3.2.4.6: 그 밖의 회계기준 → CEComputation.OtherFas (K열, B열 TIN 매칭)
- [x] 3.4.1: 소득산입규칙(IIR) 적용 → JurisdictionSection.LowTaxJurisdiction.Ltce[]
  - 구성기업 계산 시트 내부로 통합됨 (FindRow "3.4.1" / "1. 추가세액 배분" / "2. 적격소득산입규칙" 기반)
  - 1.a,b,c: O(15)열 → Ltce.Tin / Iir.NetGlobeIncome / Iir.TopUpTax
  - 2+3: "2. 적격소득산입규칙" 행 O열 통합 셀 → IirParentEntity[]
    - 포맷: `TIN, TIN유형(GIR300x), 발급국가(ISO2), 소재지국(ISO2), OtherOwnershipAllocation, InclusionRatio, TopUpTaxShare, IirOffSet, TopUpTax` × 모기업 수 (구분: `;`)
- [x] 3.4.2: 소득산입보완규칙(UTPR) → LowTaxJurisdiction.Utpr.UtprCalculation
  - 구성기업 계산 시트 내부로 통합됨 (FindRow "3.4.2" 기반)
  - 2: O열 → Article251TopUpTax / 3: O열 → TotalUtprTopUpTax (1: TIN 정보성 — XML 미포함)

### ValidationUtil.cs 검증 커버리지
- **구현됨**: 60001/03/04/11~13/15~23, 70001~03/05/07/09~21/26~28/32~34/38~40/60, 70030/31, 70041~49/53/54, 70060~92, 70101~105
- **미구현**: 70037(Subgroup TIN 일치), 70077/82(DeferTaxAdjustAmt 구조 복잡), 70097~100(3.4 IIR/UTPR 계산 — 구조 확인 필요)

## 참고
Globe.cs 대신에 Globe요약.md 참고하면 편함
