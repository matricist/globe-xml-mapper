# 블록 반복 상수 인벤토리

각 매퍼의 블록 크기/시작/간격 상수. blockCount≥2 시나리오에서 N번째 블록의 행 좌표를 확인하는 기준.

## 최종모기업 · `Mapping_1.3.1.cs`

| 상수 | 값 |
|---|---|
| `BLOCK_HEADER` | `"1.3.1"` |

## 그룹구조 · `Mapping_1.3.2.1.cs`

| 상수 | 값 |
|---|---|
| `BLOCK_HEADER` | `"1.3.2.1"` |
| `OWNERSHIP_ROW_OFFSET` | `8` |

## 제외기업 · `Mapping_1.3.2.2.cs`

| 상수 | 값 |
|---|---|
| `BLOCK_START` | `2` |
| `BLOCK_END` | `5` |
| `BLOCK_GAP` | `2` |
| `SET_SIZE` | `BLOCK_END - BLOCK_START + 1 + BLOCK_GAP` |

## 그룹구조 변동 · `Mapping_1.3.3.cs`

| 상수 | 값 |
|---|---|
| `DATA_START_ROW` | `6` |

## 요약 · `Mapping_1.4.cs`

| 상수 | 값 |
|---|---|
| `DATA_START_ROW` | `4` |

## 적용면제 · `Mapping_2.cs`

| 상수 | 값 |
|---|---|
| `BLOCK1_START` | `2` |
| `BLOCK1_SIZE` | `21` |
| `GAP` | `2` |
| `SET_SIZE` | `52` |
| `SET_GAP` | `2` |

## UTPR 배분 · `Mapping_Utpr.cs`

| 상수 | 값 |
|---|---|
| `DATA_START_ROW` | `4` |

---

## JurCal / EntityCe — 동적 블록

이 두 매퍼는 상수 기반이 아닌 **헤더 텍스트**로 블록 경계 탐지:

- `Mapping_JurCal`: `"3.1 국가별"` 헤더 등장 행마다 새 블록 시작, `_blockStart`/`_blockEnd`를 다음 헤더 직전까지로 스코핑
- `Mapping_EntityCe`: `"1. 구성기업 또는 공동기업그룹 기업의 납세자번호"` 헤더 등장 행마다 새 블록

⇒ blockCount 기반이 아니므로 _META 의 blockCount 와 무관하게 동작. 단, `Mapping_2`/`Mapping_1.3.x`/`Mapping_1.4`/`Mapping_Utpr`/`Mapping_1.3.3` 는 _META 의 blockCount 필요.
