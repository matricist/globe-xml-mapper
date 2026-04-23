# 인라인 복합 포맷 셀 인벤토리

코드에서 `.Split(delimiter)` 로 여러 필드를 파싱하는 셀 위치 목록.
작성요령 시트의 안내문과 포맷이 일치하는지 수동 대조 필요.

| 시트 | 파일 | 라인 | 종류 | delimiter | 최근 FindRow 헤더 | 최근 col | 스니펫 |
|---|---|---|---|---|---|---|---|
| 다국적기업그룹 정보 | Mapping_1.1~1.2.cs | 41 | TryParseDate | - | - | - | `if (TryParseDate(val, out var sd))` |
| 다국적기업그룹 정보 | Mapping_1.1~1.2.cs | 49 | TryParseDate | - | - | - | `if (TryParseDate(val, out var ed))` |
| 다국적기업그룹 정보 | Mapping_1.1~1.2.cs | 57 | ParseNameKName | - | - | - | `var (ceName, ceKName) = ParseNameKName(val);` |
| 다국적기업그룹 정보 | Mapping_1.1~1.2.cs | 65 | ParseTin | - | - | - | `fi.FilingCe.Tin = string.IsNullOrWhiteSpace(val) ? NoTin() : ParseTin(val);` |
| 다국적기업그룹 정보 | Mapping_1.1~1.2.cs | 86 | ParseNameKName | - | - | - | `var (mneName, mneKName) = ParseNameKName(val);` |
| 최종모기업 | Mapping_1.3.1.cs | 140 | ParseNameKName | - | - | 2 | `var (n, kn) = ParseNameKName(value);` |
| 최종모기업 | Mapping_1.3.1.cs | 149 | ParseTin | - | - | 2 | `id.Tin.Add(ParseTin(value)); break;` |
| 최종모기업 | Mapping_1.3.1.cs | 151 | ParseTin | - | - | 2 | `id.Tin.Add(ParseTin(value)); break;` |
| 최종모기업 | Mapping_1.3.1.cs | 55 | Split | `,` | - | 10 | `? cellValue.Split(',', System.StringSplitOptions.RemoveEmptyEntries \| System.St…` |
| 최종모기업 | Mapping_1.3.1.cs | 98 | Split | `,` | - | 10 | `var art1035Val = art1035Raw.Split(',')[0].Trim();` |
| 그룹구조 | Mapping_1.3.2.1.cs | 113 | ParseBool | - | - | 15 | `cs.UnreportChangeCorpStr = ParseBool(val);` |
| 그룹구조 | Mapping_1.3.2.1.cs | 135 | ParseNameKName | - | - | 15 | `var (ceName, ceKName) = ParseNameKName(val);` |
| 그룹구조 | Mapping_1.3.2.1.cs | 140 | ParseTin | - | - | 15 | `ce.Id.Tin.Add(ParseTin(val));` |
| 그룹구조 | Mapping_1.3.2.1.cs | 143 | ParseTin | - | - | 15 | `ce.Id.Tin.Add(ParseTin(val));` |
| 그룹구조 | Mapping_1.3.2.1.cs | 171 | ParseTin | - | - | 15 | `ce.Qiir.Exception.Tin = ParseTin(val);` |
| 그룹구조 | Mapping_1.3.2.1.cs | 180 | ParseTin | - | - | 15 | `ce.Qiir.Exception.Tin = ParseTin(val);` |
| 그룹구조 | Mapping_1.3.2.1.cs | 184 | ParseBool | - | - | 15 | `ce.Qutpr.Art93 = ParseBool(val);` |
| 그룹구조 | Mapping_1.3.2.1.cs | 199 | ParseBool | - | - | 15 | `ce.Qutpr.UpeOwnership = ParseBool(val);` |
| 그룹구조 | Mapping_1.3.2.1.cs | 73 | Split | `,` | - | 15 | `? cellValue.Split(` |
| 그룹구조 | Mapping_1.3.2.1.cs | 218 | Split | `;` | - | 15 | `var shareholders = cellValue.Split(` |
| 그룹구조 | Mapping_1.3.2.1.cs | 225 | Split | `,` | - | 15 | `var parts = shareholders[i].Split(` |
| 제외기업 | Mapping_1.3.2.2.cs | 57 | ParseBool | - | - | 15 | `entity.Change = ParseBool(val); break;` |
| 제외기업 | Mapping_1.3.2.2.cs | 59 | ParseNameKName | - | - | 15 | `var (eName, eKName) = ParseNameKName(val);` |
| 그룹구조 변동 | Mapping_1.3.3.cs | 81 | ParseTin | - | - | 16 | `Tin = ParseTin(ownerTinRaw)` |
| 그룹구조 변동 | Mapping_1.3.3.cs | 51 | Split | `,` | - | 16 | `errors.Add($"[{fileName}] 행{row}: 1.3.2.1에 TIN '{tinRaw.Split(',')[0].Trim()}'인 …` |
| 그룹구조 변동 | Mapping_1.3.3.cs | 68 | Split | `,` | - | 16 | `foreach (var code in preTypeRaw.Split(',', StringSplitOptions.TrimEntries \| Str…` |
| 그룹구조 변동 | Mapping_1.3.3.cs | 109 | Split | `,` | - | 16 | `var tinValue = tinRaw.Split(',')[0].Trim();` |
| 요약 | Mapping_1.4.cs | 41 | Split | `,` | - | 7 | `foreach (var code in taxJurRaw.Split(',', StringSplitOptions.TrimEntries \| Stri…` |
| 요약 | Mapping_1.4.cs | 58 | Split | `,` | - | 9 | `foreach (var code in safeHarbour.Split(',', StringSplitOptions.TrimEntries \| St…` |
| 적용면제 | Mapping_2.cs | 87 | Split | `,` | - | - | `foreach (var code in safeHarbourRaw.Split(',', StringSplitOptions.TrimEntries \|…` |
| 적용면제 | Mapping_2.cs | 178 | Split | `;` | - | - | `foreach (var entry in taxJurRaw.Split(';', StringSplitOptions.TrimEntries \| Str…` |
| 적용면제 | Mapping_2.cs | 321 | Split | `;` | - | - | `var entries = cellValue.Split(` |
| 적용면제 | Mapping_2.cs | 326 | Split | `,` | - | - | `var parts = entries[i].Split(',', StringSplitOptions.TrimEntries);` |
| 적용면제 | Mapping_2.cs | 391 | Split | `,` | - | - | `var parts = subgroupPart.Split(',', StringSplitOptions.TrimEntries);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 86 | ParseTin | - | 1. 구성기업 또는 공동기업그룹 기업의 납세자번호 | 13 | `var ceTin = ParseTin(ceTinRaw);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 209 | ParseBool | - | 1. 귀 다국적기업그룹은 | 15 | `ceComp.Elections.SimplCalculations = ParseBool(v);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 241 | ParseTin | - | 1. 연결납세그룹(납세자번호) | 11 | `entityTins.Add(ParseTin(kVal));` |
| 구성기업 계산 | Mapping_EntityCe.cs | 249 | ParseTin | - | 1. 연결납세그룹(납세자번호) | 11 | `TaxConsolGroupTin = ParseTin(groupTinRaw)` |
| 구성기업 계산 | Mapping_EntityCe.cs | 380 | ParseTin | - | 대상조세 국가간 배분 | 3 | `if (!string.IsNullOrEmpty(otherTinRaw)) item.OtherTin = ParseTin(otherTinRaw);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 646 | ParseTin | - | 배분 후 회계상 순손익 | 14 | `OtherTin = ParseTin(otherTinRaw ?? ""),` |
| 구성기업 계산 | Mapping_EntityCe.cs | 699 | ParseTin | - | 국가간 손익 조정 | 16 | `OtherTin = ParseTin(otherTinRaw ?? ""),` |
| 구성기업 계산 | Mapping_EntityCe.cs | 829 | ParseTin | - | 최종모기업의 글로벌최저한세소득 감액 | 14 | `var tin = ParseTin(tinRaw);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 886 | ParseBool | - | 2. 매년 선택 | 15 | `el.SimplCalculations = ParseBool(vA);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 893 | ParseBool | - | 2. 매년 선택 | 15 | `el.Art321 = ParseBool(vB);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 900 | ParseBool | - | 2. 매년 선택 | 15 | `el.KArt447C = ParseBool(vC);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1006 | TryParseDate | - | k. 공정가액조정 | 9 | `if (!TryParseDate(eyRaw, out var fy)) continue;` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1208 | ParseTin | - | 5년 선택 | 14 | `var invTin = ParseTin(eRaw);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1303 | ParseTin | - | 적격소득산입규칙 | 15 | `var ltce = new Globe.LowTaxJurisdictionTypeLtce { Tin = ParseTin(tinRaw) };` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1433 | TryParseDate | - | 소득산입보완규칙 추가세액 합계 | 17 | `TryParseDate(eyRaw, out var ey);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1434 | TryParseDate | - | 소득산입보완규칙 추가세액 합계 | 17 | `TryParseDate(ryRaw, out var ry);` |
| 구성기업 계산 | Mapping_EntityCe.cs | 795 | Split | `,` | 최종모기업의 글로벌최저한세소득 감액 | 14 | `var parts = ownerRaw.Split(',');` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1324 | Split | `;` | 적격소득산입규칙 | 15 | `var shareholders = cellValue.Split(` |
| 구성기업 계산 | Mapping_EntityCe.cs | 1329 | Split | `,` | 적격소득산입규칙 | 15 | `var parts = shareholders[i].Split(',', System.StringSplitOptions.TrimEntries);` |
| 국가별 계산 | Mapping_JurCal.cs | 221 | ParseTin | - | 3.1 국가별 | 2 | `subGroup.Tin = ParseTin(` |
| 국가별 계산 | Mapping_JurCal.cs | 409 | ParseTin | - | (c) 통합형피지배외국법인 | 10 | `: ParseTin(tinRaw),` |
| 국가별 계산 | Mapping_JurCal.cs | 1017 | ParseBool | - | 결손취급특례 | 13 | `el.Art326 = ParseBool(v);` |
| 국가별 계산 | Mapping_JurCal.cs | 1026 | ParseBool | - | 결손취급특례 | 13 | `el.Art461 = ParseBool(v);` |
| 국가별 계산 | Mapping_JurCal.cs | 1035 | ParseBool | - | 결손취급특례 | 13 | `el.Art531 = ParseBool(v);` |
| 국가별 계산 | Mapping_JurCal.cs | 1044 | ParseBool | - | 결손취급특례 | 13 | `el.Art415 = ParseBool(v);` |
| 국가별 계산 | Mapping_JurCal.cs | 1212 | ParseBool | - | 간주분배세액 선택 | 15 | `elected = v == "■" \|\| ParseBool(v);` |
| 국가별 계산 | Mapping_JurCal.cs | 1804 | ParseBool | - | 8. 최소적용제외 특례 | 16 | `q.SbieAvailable = ParseBool(sbieRaw);` |
| 국가별 계산 | Mapping_JurCal.cs | 1806 | ParseBool | - | 8. 최소적용제외 특례 | 16 | `q.DeMinAvailable = ParseBool(deMinRaw);` |
| 국가별 계산 | Mapping_JurCal.cs | 227 | Split | `,` | 3.1 국가별 | 2 | `var code in subGroupTypeRaw.Split(` |
| 국가별 계산 | Mapping_JurCal.cs | 1631 | Split | `,` | 3.3.3.1 | 3 | `var code in articlesRaw.Split(` |

**총 64건**
