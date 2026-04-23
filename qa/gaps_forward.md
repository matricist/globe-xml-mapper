# 순방향 갭 리포트 (서식 → 코드)

`value` 또는 `candidate` 셀 중 코드 매핑에 미연결로 추정되는 셀 목록.

- **Tier A** (JSON 시트): `(sheet, cell)` 정확 매칭 — 미매칭 = 확실한 갭
- **Tier B** (C# 시트, 블록 반복): 컬럼 단위 매칭만 가능 — 같은 컬럼이 매퍼에서 한 번도 안 읽히면 갭, 그 외는 수동 확인 필요

## 구성기업 계산

- 매퍼: `Mapping_EntityCe.cs`
- 처리 방식: C# (블록 반복)
- 입력 셀 후보: 5개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| M44 | 13 | value | △ col 사용 (수동) | 0 |
| O88 | 15 | value | △ col 사용 (수동) | 0 |
| O116 | 15 | value | △ col 사용 (수동) | 0 |
| O145 | 15 | value | △ col 사용 (수동) | 0 |
| O149 | 15 | value | △ col 사용 (수동) | 0 |

## 국가별 계산

- 매퍼: `Mapping_JurCal.cs`
- 처리 방식: C# (블록 반복)
- 입력 셀 후보: 25개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| E27 | 5 | value | △ col 사용 (수동) | 0 |
| L27 | 12 | value | ✗ col 미사용 | 0 |
| O27 | 15 | candidate | △ col 사용 (수동) | #DIV/0! |
| O58 | 15 | value | △ col 사용 (수동) | 0 |
| O80 | 15 | value | △ col 사용 (수동) | 0 |
| O84 | 15 | value | △ col 사용 (수동) | 0 |
| O90 | 15 | value | △ col 사용 (수동) | 0 |
| B96 | 2 | candidate | △ col 사용 (수동) | 합계 |
| O104 | 15 | value | △ col 사용 (수동) | 0 |
| O105 | 15 | value | △ col 사용 (수동) | 0 |
| O106 | 15 | value | △ col 사용 (수동) | 0 |
| O109 | 15 | value | △ col 사용 (수동) | 0 |
| O129 | 15 | value | △ col 사용 (수동) | 0 |
| B137 | 2 | candidate | △ col 사용 (수동) | 합계 |
| O194 | 15 | value | △ col 사용 (수동) | 0 |
| O197 | 15 | candidate | △ col 사용 (수동) | □ |
| N215 | 14 | value | △ col 사용 (수동) | 0 |
| N217 | 14 | value | △ col 사용 (수동) | 0 |
| C222 | 3 | value | △ col 사용 (수동) | 0 |
| I222 | 9 | value | △ col 사용 (수동) | 0 |
| L222 | 12 | value | ✗ col 미사용 | 0 |
| N226 | 14 | candidate | △ col 사용 (수동) | 합계 |
| M247 | 13 | value | △ col 사용 (수동) | 0 |
| M248 | 13 | value | △ col 사용 (수동) | 0 |
| K256 | 11 | candidate | △ col 사용 (수동) | 통화 |

## 그룹구조

- 매퍼: `Mapping_1.3.2.1.cs`
- 처리 방식: JSON
- 입력 셀 후보: 2개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| O4 | 15 | candidate | ✓ JSON | 부 |
| O18 | 15 | candidate | ✓ JSON | 부 |

## 그룹구조 변동

- 매퍼: `Mapping_1.3.3.cs`
- 처리 방식: C# (블록 반복)
- 입력 셀 후보: 1개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| B3 | 2 | candidate | △ col 사용 (수동) | a. |

## 다국적기업그룹 정보

- 매퍼: `Mapping_1.1~1.2.cs`
- 처리 방식: JSON
- 입력 셀 후보: 12개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| C2 | 3 | value | ✓ JSON | 2024-01-01 오전 12:00:00 |
| C3 | 3 | candidate | ✗ JSON 미매핑 | ~ |
| C5 | 3 | value | ✓ JSON | 2024-12-31 오전 12:00:00 |
| B9 | 2 | candidate | ✗ JSON 미매핑 | 여 |
| K9 | 11 | candidate | ✓ JSON | GIR401 |
| M9 | 13 | candidate | ✓ JSON | KR |
| F14 | 6 | value | ✓ JSON | 2024-01-01 오전 12:00:00 |
| K14 | 11 | value | ✓ JSON | 2024-12-31 오전 12:00:00 |
| O14 | 15 | candidate | ✓ JSON | 부 |
| B17 | 2 | candidate | ✓ JSON | GIR501 |
| H17 | 8 | candidate | ✓ JSON | K-IFRS |
| M17 | 13 | candidate | ✓ JSON | KRW |

## 제외기업

- 매퍼: `Mapping_1.3.2.2.cs`
- 처리 방식: JSON
- 입력 셀 후보: 1개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| O3 | 15 | candidate | ✓ JSON | 부 |

## 최종모기업

- 매퍼: `Mapping_1.3.1.cs`
- 처리 방식: JSON
- 입력 셀 후보: 5개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| J4 | 10 | candidate | ✓ JSON | KR |
| J5 | 10 | candidate | ✓ JSON | GIR201 |
| J6 | 10 | value | ✓ JSON | 0 |
| J7 | 10 | value | ✓ JSON | 0 |
| J9 | 10 | candidate | ✓ JSON | GIR301 |

## UTPR 배분

- 매퍼: `Mapping_Utpr.cs`
- 처리 방식: C# (블록 반복)
- 입력 셀 후보: 4개

| address | col | kind | 매핑상태 | text(요약) |
|---|---|---|---|---|
| C4 | 3 | value | △ col 사용 (수동) | 0 |
| C5 | 3 | value | △ col 사용 (수동) | 0 |
| B6 | 2 | candidate | △ col 사용 (수동) | 합계 |
| C6 | 3 | value | △ col 사용 (수동) | 0 |

