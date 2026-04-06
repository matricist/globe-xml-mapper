# Globe XML Mapper - 작업 가이드

## 프로젝트 개요
국조 서식(.xlsx) → OECD GloBE XML 변환 WinForms 프로그램.
**디렉토리 기반** — 정해진 폴더 구조에 섹션별 xlsx를 넣고, 최상위 폴더를 선택하면 XML 생성.

## 폴더 구조 (사용자 입력)
```
사용자_폴더/
├── 아무이름.xlsx          ← 기본정보 (루트, 시트 "1.1~1.2")
├── 1.3.1/                ← UPE (필수, 파일 N개)
│   └── *.xlsx
├── 1.3.2.1/              ← CE (필수, 파일 N개)
│   └── *.xlsx
└── 1.3.2.2/              ← 제외기업 (필수, 파일 N개)
    └── *.xlsx
```
- 파일명은 자유, 시트 이름으로 인식
- 하위 디렉토리 3개 모두 필수 (없거나 xlsx 0개면 오류)
- 복수 파일 → Collection에 각각 Add

## 시트 이름 → 매퍼 매핑

| 시트 이름 | 위치 | 내용 | Globe 대상 | 매퍼 클래스 |
|---|---|---|---|---|
| `1.1~1.2` | 루트 | 신고구성기업, 사업연도, 회계정보 | FilingInfo | Mapping_1_1_1_2 |
| `1.3.1` | 1.3.1/ | 최종모기업 정보 | CorporateStructure.UPE | Mapping_1_3_1 |
| `1.3.2.1` | 1.3.2.1/ | 구성기업 정보 | CorporateStructure.CE | Mapping_1_3_2_1 |
| `1.3.2.2` | 1.3.2.2/ | 제외기업 정보 | CorporateStructure.ExcludedEntity | Mapping_1_3_2_2 |

## 프로젝트 구조
```
mapper/
├── GlobeMapper.csproj
├── Program.cs
├── MainForm.cs                  # UI (폴더선택 + XML생성 + 템플릿다운로드)
├── TermsDialog.cs
├── Services/
│   ├── MappingBase.cs            # 공통 유틸
│   ├── Mapping_1.1~1.2.cs       # 기본정보 매퍼
│   ├── Mapping_1.3.1.cs         # UPE 매퍼
│   ├── Mapping_1.3.2.1.cs       # CE 매퍼 (별첨 시트 소유지분 포함)
│   ├── Mapping_1.3.2.2.cs       # 제외기업 매퍼
│   ├── MappingOrchestrator.cs   # 디렉토리 기반 순회 + FillMessageSpec/DocSpecs
│   ├── ValidationUtil.cs
│   └── XmlExportService.cs
├── Resources/
│   ├── Globe.cs
│   ├── XSD/
│   ├── mappings/                # 섹션별 매핑 JSON
│   │   ├── mapping_1.1~1.2.json
│   │   ├── mapping_1.3.1.json
│   │   ├── mapping_1.3.2.1.json
│   │   └── mapping_1.3.2.2.json
│   ├── templates/               # 배포용 템플릿 xlsx (작성요령 시트 포함)
│   ├── sample/                  # 테스트용 (디렉토리 구조)
│   ├── terms.txt
│   ├── Globe_Enum_정리.csv
│   └── GIR_XML_에러코드_매뉴얼.md
└── Tools/                       # Python 스크립트
    ├── create_templates.py      # 원본 template.xlsx에서 템플릿 생성
    └── create_sample.py         # 샘플 데이터 생성 (디렉토리 구조)
```

## 처리 흐름
```
1. 사용자가 최상위 폴더 선택
2. MappingOrchestrator.MapFolder(rootPath)
   → 루트/*.xlsx → Mapping_1_1_1_2 (기본정보)
   → 1.3.1/*.xlsx → 파일마다 new Mapping_1_3_1() → UPE Add
   → 1.3.2.1/*.xlsx → 파일마다 new Mapping_1_3_2_1() → CE Add
   → 1.3.2.2/*.xlsx → 파일마다 new Mapping_1_3_2_2() → ExcludedEntity Add
   → 필수 디렉토리 누락/비어있음 → error
3. FillMessageSpec + FillDocSpecs
4. ValidationUtil.Validate
5. XML 저장 + .errors.txt
```

## 새 섹션 추가 절차

### Step 1: 엑셀 시트 구조 파악 (Python 스크립트로)
### Step 2: mapping JSON 추가 → `Resources/mappings/mapping_{시트이름}.json`
### Step 3: 매퍼 클래스 추가 → `Services/Mapping_{시트이름}.cs` (MappingBase 상속)
### Step 4: MappingOrchestrator.SubDirMappers에 등록
### Step 5: ValidationUtil에 검증 추가
### Step 6: `Tools/create_templates.py`에 함수 추가 + 실행
### Step 7: `Tools/create_sample.py`에 데이터 추가 + 실행
### Step 8: CLAUDE.md 업데이트

## 검증 규칙 (현재 구현)
**MessageSpec/DocSpec:** 60001, 60003, 60004, 60011, 60012, 60015, 60017, 60018
**FilingInfo:** 60020, 60021, 60022, 60023
**UPE:** 70009, 70010
**CE:** 70011, 70013~70015, 70018~70020, 70026, 70032
**TIN:** 70001~70003, 70005

## 구현 진행 상황
- [x] 1.1~1.2: 신고구성기업, 사업연도, 회계정보
- [x] 1.3.1: 최종모기업 (UPE)
- [x] 1.3.2.1: 구성기업 (CE) + 별첨 시트 소유지분
- [x] 1.3.2.2: 제외기업 (ExcludedEntity)
- [ ] 1.3.3: 추가 데이터포인트 (AdditionalDataPoint)
- [ ] 2: Summary
- [ ] 3: JurisdictionSection / ETR
- [ ] 4: UTPRAttribution
