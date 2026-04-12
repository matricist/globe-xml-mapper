# Globe XML Mapper - 작업 가이드

## 프로젝트 개요
국조 서식(.xlsx) → OECD GloBE XML 변환 WinForms 프로그램.
**단일 xlsx 파일** + **Control Panel** 방식 — Excel COM Interop으로 엑셀 직접 제어.

## UX 흐름
```
1. 메인 화면: [파일 편집] / [새 파일 만들기] / [템플릿 다운로드]
2. 파일 편집/생성 → Excel COM으로 xlsx 열기
3. Control Panel 표시 (TopMost, 드래그 이동, 접기/펼치기)
   ┌─────────────────────────────┐
   │ Globe XML Mapper       [─]  │
   ├─────────────────────────────┤
   │ 최종모기업 (UPE)  2개  [+][-]│
   │ 구성기업 (CE)     3개  [+][-]│
   │ 제외기업           1개  [+][-]│
   ├─────────────────────────────┤
   │      [XML 변환하기]          │
   └─────────────────────────────┘
4. [+] 시트 복제 / [-] 마지막 시트 삭제
5. [XML 변환하기] → 저장 후 ClosedXML로 매핑 + 검증
6. 엑셀 종료 → 메인 화면 복귀
7. 윈폼 종료 → 엑셀 저장 확인 후 함께 종료
```

## template.xlsx 시트 구조
- 국조53부표2 (1): 1.1~1.2 + 1.3.1 (기본정보 + 첫 번째 UPE)
- 국조53부표2 (2): 1.3.2.1 (첫 번째 CE)
- 국조53부표2 (3): 1.3.2.2 (첫 번째 제외기업)
- (4)~(25): 빈 시트 (추가 시 복제 대상)

## 숨김시트 (_META)
섹션→시트 매핑 메타정보. ExcelController가 자동 생성/관리.
```
section    | sheetName
1.1~1.2    | 국조53부표2 (1)
1.3.1      | 국조53부표2 (1)      ← (1)에 포함
1.3.2.1    | 국조53부표2 (2)
1.3.2.2    | 국조53부표2 (3)
```
시트 추가/삭제 시 메타도 자동 갱신.
XML 변환 시 MappingOrchestrator가 메타 참조.

## 프로젝트 구조
```
mapper/
├── GlobeMapper.csproj
├── Program.cs
├── MainForm.cs                  # 메인 화면 (3버튼)
├── ControlPanelForm.cs          # Control Panel (TopMost, 시트 추가/삭제, XML 변환)
├── TermsDialog.cs
├── Services/
│   ├── ExcelController.cs        # Excel COM Interop 래퍼
│   ├── MappingBase.cs            # 공통 유틸
│   ├── Mapping_1.1~1.2.cs       # 기본정보 매퍼
│   ├── Mapping_1.3.1.cs         # UPE 매퍼
│   ├── Mapping_1.3.2.1.cs       # CE 매퍼 (별첨 시트 포함)
│   ├── Mapping_1.3.2.2.cs       # 제외기업 매퍼
│   ├── Mapping_2.cs             # 합산단위 — 국가별 적용면제 (섹션 2)
│   ├── Mapping_JurCal.cs        # 합산단위 — 3.1~3.3 전체 (JurCalMapper)
│   ├── MappingOrchestrator.cs   # 메타 기반 매핑 + FillMessageSpec
│   ├── ValidationUtil.cs
│   └── XmlExportService.cs
├── Resources/
│   ├── Globe.cs                  # XSD 생성 클래스
│   ├── Globe요약.md              # Globe.cs 구조 요약 (코드 탐색 대신 참조)
│   ├── XSD/
│   ├── mappings/                 # 섹션별 매핑 JSON
│   ├── main_template.xlsx        # 메인 템플릿 (1.1~1.3.2.2)
│   ├── group_template.xlsx       # 합산단위 템플릿 (2, 3.1~3.3)
│   ├── entity_template.xlsx      # 구성기업 템플릿 (3.4 — 미구현)
│   ├── terms.txt
│   ├── Globe_Enum_정리.csv
│   └── GIR_XML_에러코드_매뉴얼.md
└── Tools/                        # 개발/검증용 C# 콘솔 도구
    ├── RunMapper.cs / .csproj    # 샘플 폴더 XML 변환 테스트 (net9.0-windows)
    ├── DumpAttach.cs / .csproj   # xlsx 시트 셀 덤프 (ClosedXML, net9.0)
    ├── DumpGroup.cs / .csproj    # 그룹구조 시트 덤프
    ├── DumpOutput.cs / .csproj   # 출력 검증 덤프
    └── FillFromJson.cs / .csproj # JSON → xlsx 자동 채우기
```

## 핵심 클래스

### ExcelController (Services/ExcelController.cs)
- `Open(path)` / `CreateNew(template, savePath)` — 엑셀 열기
- `AddSectionSheet(section)` — 템플릿 시트 복제 + 메타 등록
- `RemoveSectionSheet(section)` — 마지막 시트 삭제 + 메타 제거
- `GetSectionSheets(section)` / `GetSectionCounts()` — 메타 조회
- `GetFilePathForMapping()` — 저장 후 파일 경로 반환 (ClosedXML용)
- `CloseWithSavePrompt()` — 저장 확인 후 종료
- `WorkbookClosed` 이벤트 — 엑셀 종료 감지

### MappingOrchestrator
- `MapWorkbook(filePath, globe)` — _META 기반 매핑 (Control Panel 방식)
- `MapFolder(rootPath, globe)` — 디렉토리 기반 매핑 (하위 호환)

### 매핑 행 탐색 원칙 (Mapping_JurCal 등)
- **절대 행번호 금지** — 행 삽입/삭제에 취약하므로 반드시 `FindRow(ws, "헤더텍스트")` 로 위치를 동적 탐색
- 헤더 발견 후 바로 아래 데이터가 고정 구조일 때만 `rowHeader + N` 오프셋 허용 (예: "(a)" 바로 다음 1행이 항상 값 행인 경우)
- 여러 항목이 각자 레이블을 가질 때는 항목마다 별도 `FindRow`로 탐색 (예: "1. 모든 구성기업...", "2. 50% 한도" 등)
- 값 열(column)도 실제 템플릿 덤프로 확인 — O열(15)이 아닌 N열(14) 등 다를 수 있음

### CE 행 블록 반복 (그룹구조 시트)
- 그룹구조 시트에서 3~21행이 CE 1건 블록 (19행). [+]로 블록 복제, 각 블록의 O11셀에 "첨부N" 자동 갱신
- ExcelController 상수: CE_BLOCK_START=3, CE_BLOCK_END=21, CE_BLOCK_GAP=2, CE_ATTACH_REF_ROW_OFFSET=8
- 별첨 시트(`부표2 (2) 별첨`)에 별첨1, 별첨2... 섹션이 자동 추가/삭제
- Mapping_1.3.2.1이 blockCount 기반으로 N개 CE 순회, 별첨N에서 Ownership 읽기
- 별첨 시트에서 주주 행 추가/삭제는 Control Panel에서 별첨 번호 선택 후 [+][-]

## 새 섹션 추가 절차
1. mapping JSON: `Resources/mappings/mapping_{섹션}.json`
2. 매퍼 클래스: `Services/Mapping_{섹션}.cs` (MappingBase 상속)
3. MappingOrchestrator.MapperFactory에 등록
4. ExcelController.SectionTemplateIndex에 템플릿 시트 인덱스 등록
5. ControlPanelForm.Sections에 표시명 추가
6. ValidationUtil 검증 추가
7. CLAUDE.md 업데이트

## 구현 진행 상황

### main_template.xlsx (Mapping_1.1~1.2, Mapping_1.3.1, Mapping_1.3.2.1, Mapping_1.3.2.2)
- [x] 1.1~1.2: 신고구성기업, 사업연도, 회계정보
- [x] 1.3.1: 최종모기업 (UPE) — 행 블록 반복
- [x] 1.3.2.1: 구성기업 (CE) — 행 블록 반복 + 첨부 시트 소유지분
- [x] 1.3.2.2: 제외기업 — 행 블록 반복 (블록 범위 2~5행, 여/부 ParseBool, Name;KName 분리)
- [x] 1.3.3: 기업구조 변동 — 단순 행 추가 (XML 미포함)
- [x] 1.4: 정보 요약 (Summary) — 단순 행 추가

### 합산단위_N.xlsx (Mapping_2, Mapping_JurCal)
- [x] 2: 국가별 적용면제 — 52행 복합 블록 반복 + Summary 상세
- [x] 3.1: 기본사항 (국가명/하위그룹/JurWithTaxingRights) — 국가별 계산 첨부 시트 연동
  - 3.1.4~3.1.10: 국가별 계산 첨부 시트 B~O 열 (D열=조정대상조세 합계는 display only, XML 미포함)
  - 3.1.11~3.1.18: 국가별 계산 첨부 시트 I~P 열 (ElectionsDifference는 P열, 세미콜론으로 KElectionsDifference 구분)
  - TransitionDifference(O열): boolean, "여"/"부" 입력
- [x] 3.2.1: 실효세율 계산 (FANIL, 글로벌최저한세소득/결손 26개 조정, 조정대상조세 20개, 이월, CFC)
- [x] 3.2.2.1: 이연법인세조정 요약/내역 16개
- [x] 3.2.2.1(c): 결손금 소급공제
- [x] 3.2.2.2: 환입금액 계산
- [x] 3.2.2.3: 최초적용연도 특례
- [x] 3.2.3.1: 국가별 선택 (매년/5년/그 밖의 + 필요정보)
- [x] 3.2.3.2: 간주분배세액
- [x] 적격국제해운부수소득 시트 (IntShippingIncome)
- [x] 추가세액 계산 시트 (3.3.1 요약, 3.3.2 실질기반제외소득+PE/FTE배분, 3.3.3 당기추가세액가산액, 3.3.4 QDMTT)

### entity_template.xlsx (Mapping_EntityCe.cs — "구성기업 계산" 시트)
- [x] CE TIN 식별: 3.2.4.1(a) "납세자번호" 항목 K열 ("값,GIR300x,발급국가")
  - TIN issuedBy 국가코드로 JurisdictionSection 조회 → CEComputation 생성
  - MapFolder에서 group 파일 먼저 처리, entity 파일 나중 처리 (JurisdictionSection 선행 생성)
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
- [ ] 3.4: 추가세액 배분과 귀속 (IIR/UTPR — LowTaxJurisdiction.Ltce/Utpr)

### ValidationUtil.cs 검증 커버리지
- **구현됨**: 60001/03/04/11~13/15~23, 70001~03/05/07/09~21/26~28/32~34/38~40/60
- **미구현 (주요)**: 70030/31(소유지분 TIN 교차검증), 70037(Subgroup TIN 일치), 70041~53(SafeHarbour 세부조건), 70054/56(RevocationYear=Status FALSE만), 70061~105(수치계산 검증)

## 참고
Globe.cs 대신에 Globe요약.md 참고하면 편함
