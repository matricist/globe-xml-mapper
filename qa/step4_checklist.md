# 4단계 — 특수 케이스 수동 점검 체크리스트

[qa/inline_formats.md](inline_formats.md)와 [qa/block_constants.md](block_constants.md)를 보조 자료로 사용.

---

## A. 인라인 복합 포맷 점검 (총 64건)

`코드의 Split/Parse 위치` ↔ `작성요령 시트 안내문` ↔ `CLAUDE.md 포맷 명세` 3자 일치 확인.

### A-1. 시트별 핵심 인라인 포맷

| 시트 | 셀/위치 | 포맷 | 코드 위치 |
|---|---|---|---|
| 그룹구조 | O11 통합 셀 (CE 소유지분) | `유형,TIN,TIN유형,발급국가,지분` × N (구분 `;`) | Mapping_1.3.2.1.cs:218,225 |
| 적용면제 | M46 통합 셀 (2.3 OtherJurisdiction) | `국가,값` × N (구분 `;`) | Mapping_2.cs:178,326 |
| 적용면제 | (2.2 SafeHarbour codes) | `,` 다중코드 | Mapping_2.cs:87 |
| 요약 | (1.4.1 TaxJur 코드들) | `,` 다중코드 | Mapping_1.4.cs:41,58 |
| 그룹구조 변동 | (1.3.3 TIN 참조) | `값,GIR300x,발급국가` | Mapping_1.3.3.cs:51,68,109 |
| 최종모기업 | (1.3.1 ResCountry 다중) | `,` 다중코드 | Mapping_1.3.1.cs:55,98 |
| 구성기업 계산 | K열 다양 (3.2.4.1(b)~(d)) | `값,열거코드,상대방TIN,국가[,issuedBy,TIN유형]` | Mapping_EntityCe.cs:795,1329 |
| 구성기업 계산 | O열 (3.4.1 IirParentEntity) | `TIN, TIN유형, 발급국가, 소재지국, OtherOwnAlloc, InclusionRatio, TopUpTaxShare, IirOffSet, TopUpTax` × N (구분 `;`) | Mapping_EntityCe.cs:1329 |
| 국가별 계산 | (3.2.4.4 IntShipping Category) | 정규식 추출 GIR2101~2106, GIR2201~2205 | Mapping_JurCal.cs (Regex) |

### A-2. 점검 절차

1. main_template.xlsx 열기 → `Main 작성요령` / `Entity 작성요령` / `Group 작성요령` 시트 확인
2. 각 인라인 포맷 셀 위치 옆/아래 안내문이 위 표의 포맷과 일치하는지 대조
3. 불일치 시: 작성요령 텍스트 또는 코드 둘 중 무엇이 정답인지 결정 후 수정

### A-3. 수동 확인 필요 (FindRow 헤더 자동 추출 실패한 항목)

[qa/inline_formats.md](inline_formats.md) 표에서 "최근 FindRow 헤더 = `-`" 인 항목들. helper 메서드 안에서 호출되어 자동 매칭 안 됨. 코드를 직접 읽어 어느 항목인지 확인 필요. (총 약 20여 건, 대부분 `ParseTin`/`ParseBool`/`Split` 단순 케이스)

---

## B. 블록 반복 검증

상수 기반 매퍼 5종 + 동적 헤더 매퍼 2종.

### B-1. 상수 기반 매퍼 (블록 #2 좌표 손계산)

[qa/block_constants.md](block_constants.md) 참고.

| 시트 | 매퍼 | 블록 #2 시작 행 (계산식) | 검증 방법 |
|---|---|---|---|
| 제외기업 | Mapping_1.3.2.2 | `2 + (5-2+1) + 2 = 8` | 8행에 헤더 있는지 확인 |
| 적용면제 | Mapping_2 | `2 + 52 + 2 = 56` | 56행이 새 국가 블록 시작인지 |
| 그룹구조 (CE) | Mapping_1.3.2.1 | `BLOCK_HEADER` 텍스트 기반 (상수 X) | 다음 "1.3.2.1" 헤더까지 |
| 최종모기업 (UPE) | Mapping_1.3.1 | `BLOCK_HEADER` 텍스트 기반 (상수 X) | 다음 "1.3.1" 헤더까지 |
| 그룹구조 변동 | Mapping_1.3.3 | `DATA_START_ROW=6` 부터 EOF까지 단순 행 반복 | 빈 행 만나면 종료 |
| 요약 | Mapping_1.4 | `DATA_START_ROW=4` 부터 EOF까지 단순 행 반복 | 빈 행 만나면 종료 |
| UTPR 배분 | Mapping_Utpr | `DATA_START_ROW=4` 부터 EOF까지 단순 행 반복 | 빈 행 만나면 종료 |

### B-2. 동적 매퍼 (헤더 텍스트 기반)

| 시트 | 매퍼 | 블록 헤더 |
|---|---|---|
| 국가별 계산 | Mapping_JurCal | `"3.1 국가별"` |
| 구성기업 계산 | Mapping_EntityCe | `"1. 구성기업 또는 공동기업그룹 기업의 납세자번호"` |

### B-3. 검증 절차 (E2E)

1. main_template.xlsx 사본 만들기 → `_META` 시트에 `blockCount:최종모기업 = 2` 추가
2. (Excel 또는 코드로) 1.3.1 블록을 한 번 더 복제하고 두 번째 블록에 다른 TIN/Name 입력
3. 변환 실행 → 결과 XML 의 `Upe` 가 2개인지 확인
4. 적용면제 시트도 동일 절차 (blockCount:적용면제 = 2, 두 국가 입력)
5. 국가별 계산: "3.1 국가별" 헤더를 한 번 더 복제하면 자동으로 새 블록으로 인식되는지

⇒ B-3 은 5단계(골든 샘플 E2E)로 흡수 가능.

---

## C. 통합 셀 anchor 점검

ClosedXML은 통합 셀에서 anchor 셀(첫 셀)에만 값이 있고 나머지는 빈 값. 코드가 anchor 셀에서 읽는지 확인.

### C-1. 알려진 통합 셀 입력칸

| 시트 | anchor | 통합 범위 | 코드 |
|---|---|---|---|
| 그룹구조 | O11 | O11:R14 | Mapping_1.3.2.1.cs (offset 8) |
| 적용면제 | M46 | (M46:R46?) | Mapping_2.cs |
| 다국적기업그룹 정보 | C2/C5 | C2:D2 / C5:D5 | mapping_1.1~1.2.json |
| 다국적기업그룹 정보 | E2 | E2:M5 (라벨) | - |

### C-2. 점검 절차

1. [qa/template_cells.csv](template_cells.csv)에서 `mergeRange` 컬럼이 비어있지 않은 셀 추출
2. 그 중 `kind=value` 또는 `kind=candidate` 인 것 → 입력 가능 통합 셀 후보
3. 코드가 해당 anchor 주소(`mergeAnchor`)와 동일한 (row, col) 로 읽는지 확인
4. 다른 셀로 읽으면 빈 값이 들어와 매핑이 사일런트 실패

---

## D. 진행 상황

- [ ] A-1, A-2 완료
- [ ] A-3 (helper 내부 Split 수동 확인)
- [ ] B-1, B-2 손계산 완료
- [ ] B-3 → 5단계로 이관
- [ ] C-1, C-2 완료
