using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 시트 2: 국가별 적용면제 및 제외.
    /// 블록1(2~22) + 간격(2행) + 블록2(25~53) = 52행 세트.
    /// blockCount 기반으로 N개 국가 순회.
    ///
    /// 검증된 행 오프셋 (b1=2, b2=25 기준):
    ///   2.1  : O(b1+3)=국가, O(b1+6)=과세권국가
    ///   2.2.1: O(b1+12)=SafeHarbour
    ///   2.2.1.2: H/N(b1+16~b1+19)=간소화계산
    ///   2.2.1.3: O(b2+1)=수익, O(b2+2)=세전손익, O(b2+3)=간이세액, O(b2+6)=CIT율
    ///   2.2.2: B(b2+9)=GIR2901체크, B(b2+10)=GIR2902체크
    ///          E/I/L/O(b2+12~b2+15)=FinancialData(신고/직전/직전전/평균)
    ///   2.3  : M(b2+18)=개시일, M(b2+19)=준거국가, M(b2+20)=유형자산, M(b2+21)=국가수
    /// </summary>
    public class Mapping_2 : MappingBase
    {
        private const int BLOCK1_START = 2;
        private const int BLOCK1_SIZE = 21;  // rows 2~22 (inclusive)
        private const int GAP = 2;           // rows 23~24
        private const int SET_SIZE = 52;     // 21+2+29
        private const int SET_GAP = 2;       // 세트 간 간격

        private const string ATTACH_SHEET = "적용면제 첨부";
        private const int ATTACH_HEADER_ROWS = 1; // 헤더 행(소재지국/유형자산) 1개

        public Mapping_2() : base("mapping_2.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            var blockCount = 1;
            if (ws.Workbook.TryGetWorksheet(ExcelController.MetaSheetName, out var metaWs))
                blockCount = ExcelController.ReadBlockCount(metaWs, ws.Name);

            // "적용면제 첨부" 시트 참조 (있으면)
            ws.Workbook.TryGetWorksheet(ATTACH_SHEET, out var attachWs);

            for (int idx = 0; idx < blockCount; idx++)
            {
                var b1 = BLOCK1_START + idx * (SET_SIZE + SET_GAP);
                var b2 = b1 + BLOCK1_SIZE + GAP;
                MapOneCountry(ws, attachWs, globe, errors, fileName, b1, b2, idx + 1);
            }
        }

        private void MapOneCountry(IXLWorksheet ws, IXLWorksheet attachWs,
            Globe.GlobeOecd globe, List<string> errors, string fileName,
            int b1, int b2, int blockNum)
        {
            // ─── 2.1 소재지국 (O5 = b1+3) ─────────────────────────────────
            var jurCode = ws.Cell(b1 + 3, 15).GetString()?.Trim();
            if (string.IsNullOrEmpty(jurCode)) return;

            if (!TryParseEnum<Globe.CountryCodeType>(jurCode, out var countryCode))
            {
                errors.Add($"[{fileName}] 적용면제 블록{blockNum}: 소재지국 코드 '{jurCode}' 파싱 실패");
                return;
            }

            var loc = $"2 적용면제/블록{blockNum}('{jurCode}')";

            // ─── Summary 찾기 또는 생성 ───────────────────────────────────
            var summary = globe.GlobeBody.Summary
                .FirstOrDefault(s => s.Jurisdiction?.JurisdictionNameSpecified == true
                                  && s.Jurisdiction.JurisdictionName == countryCode);
            if (summary == null)
            {
                summary = new Globe.GlobeBodyTypeSummary
                {
                    Jurisdiction = new Globe.SummaryTypeJurisdiction
                    {
                        JurisdictionName = countryCode,
                        JurisdictionNameSpecified = true
                    }
                };
                globe.GlobeBody.Summary.Add(summary);
            }

            // 과세권 국가 (O8=b1+6) — JurisdictionSection.JurWithTaxingRights로 이동 (js 생성 후 처리)
            var taxJurRaw = ws.Cell(b1 + 6, 15).GetString()?.Trim();

            // ─── 2.2.1 적용면제 (O14=b1+12) ──────────────────────────────
            var safeHarbourRaw = ws.Cell(b1 + 12, 15).GetString()?.Trim();
            if (!string.IsNullOrEmpty(safeHarbourRaw))
            {
                foreach (var code in safeHarbourRaw.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
                    SetEnum<Globe.SafeHarbourEnumType>(code, v =>
                    {
                        if (!summary.SafeHarbour.Contains(v))
                            summary.SafeHarbour.Add(v);
                    }, errors, fileName, new MappingEntry { Cell = $"O{b1 + 12}", Label = $"[{loc}] 적용면제" });
            }

            // ─── 2.2.1.2 간소화 데이터 (H18-N21 = b1+16 ~ b1+19) ────────
            var s1Rev = ws.Cell(b1 + 16, 8).GetString()?.Trim();   // H18
            var s1Tax = ws.Cell(b1 + 16, 14).GetString()?.Trim();  // N18
            var s2Rev = ws.Cell(b1 + 17, 8).GetString()?.Trim();   // H19
            var s2Tax = ws.Cell(b1 + 17, 14).GetString()?.Trim();  // N19
            var s3Rev = ws.Cell(b1 + 18, 8).GetString()?.Trim();   // H20
            var s3Tax = ws.Cell(b1 + 18, 14).GetString()?.Trim();  // N20
            var saRev = ws.Cell(b1 + 19, 8).GetString()?.Trim();   // H21
            var saTax = ws.Cell(b1 + 19, 14).GetString()?.Trim();  // N21

            // ─── 2.2.1.3 전환기 데이터 (b2+1 ~ b2+6) ─────────────────────
            // O26=b2+1: 총수익, O27=b2+2: 세전손익, O28=b2+3: 간이대상조세
            // O31=b2+6: 법인세 명목세율
            var cbcrRev  = ws.Cell(b2 + 1, 15).GetString()?.Trim();  // O26
            var cbcrPl   = ws.Cell(b2 + 2, 15).GetString()?.Trim();  // O27
            var cbcrTax  = ws.Cell(b2 + 3, 15).GetString()?.Trim();  // O28
            var utprRate = ws.Cell(b2 + 6, 15).GetString()?.Trim();  // O31

            // ─── 2.2.2 체크박스 (b2+8=섹션헤더, b2+9=GIR2901, b2+10=GIR2902) ──
            // R34=b2+9: □/■ 신고대상 사업연도 → GIR2901
            // R35=b2+10: □/■ 중요성이 낮은 구성기업 → GIR2902
            Globe.DeminimisSimpleBasisEnumType? deminiBasis = null;
            if (RowContains(ws, b2 + 9, "■"))
                deminiBasis = Globe.DeminimisSimpleBasisEnumType.Gir2901;
            else if (RowContains(ws, b2 + 10, "■"))
                deminiBasis = Globe.DeminimisSimpleBasisEnumType.Gir2902;

            // ─── 2.2.2 상세 재무 데이터 (b2+12 ~ b2+15) ──────────────────
            // R37=b2+12: 신고대상, R38=b2+13: 직전, R39=b2+14: 직전전, R40=b2+15: 3년평균
            // 열: E(5)=회계매출, I(9)=GloBE매출, L(12)=회계순이익, O(15)=GloBE소득
            var f1AcRev = ws.Cell(b2 + 12, 5).GetString()?.Trim();   // E37
            var f1GbRev = ws.Cell(b2 + 12, 9).GetString()?.Trim();   // I37
            var f1AcPl  = ws.Cell(b2 + 12, 12).GetString()?.Trim();  // L37
            var f1GbPl  = ws.Cell(b2 + 12, 15).GetString()?.Trim();  // O37
            var f2AcRev = ws.Cell(b2 + 13, 5).GetString()?.Trim();   // E38
            var f2GbRev = ws.Cell(b2 + 13, 9).GetString()?.Trim();   // I38
            var f2AcPl  = ws.Cell(b2 + 13, 12).GetString()?.Trim();  // L38
            var f2GbPl  = ws.Cell(b2 + 13, 15).GetString()?.Trim();  // O38
            var f3AcRev = ws.Cell(b2 + 14, 5).GetString()?.Trim();   // E39
            var f3GbRev = ws.Cell(b2 + 14, 9).GetString()?.Trim();   // I39
            var f3AcPl  = ws.Cell(b2 + 14, 12).GetString()?.Trim();  // L39
            var f3GbPl  = ws.Cell(b2 + 14, 15).GetString()?.Trim();  // O39
            var faAcRev = ws.Cell(b2 + 15, 5).GetString()?.Trim();   // E40
            var faGbRev = ws.Cell(b2 + 15, 9).GetString()?.Trim();   // I40
            var faAcPl  = ws.Cell(b2 + 15, 12).GetString()?.Trim();  // L40
            var faGbPl  = ws.Cell(b2 + 15, 15).GetString()?.Trim();  // O40

            // ─── 2.3 해외진출 초기 특례 (b2+18 ~ b2+21) ──────────────────
            // R43=b2+18: 개시일, R44=b2+19: 준거국가, R45=b2+20: 유형자산, R46=b2+21: 국가수
            var initStartRaw = ws.Cell(b2 + 18, 13).GetString()?.Trim(); // M43
            var initRefJur   = ws.Cell(b2 + 19, 13).GetString()?.Trim(); // M44
            var initRefAsset = ws.Cell(b2 + 20, 13).GetString()?.Trim(); // M45
            var initNumJur   = ws.Cell(b2 + 21, 13).GetString()?.Trim(); // M46

            // ─── ETR / InitialIntActivity 데이터 유무 판단 ────────────────
            bool hasDemini  = deminiBasis.HasValue || !string.IsNullOrEmpty(f1GbRev) || !string.IsNullOrEmpty(f1AcRev)
                           || !string.IsNullOrEmpty(s1Rev);
            bool hasCbcr    = !string.IsNullOrEmpty(cbcrRev) || !string.IsNullOrEmpty(cbcrPl) || !string.IsNullOrEmpty(cbcrTax);
            bool hasUtpr    = !string.IsNullOrEmpty(utprRate);
            bool hasEtrData              = hasDemini || hasCbcr || hasUtpr;
            bool hasInit                 = !string.IsNullOrEmpty(initStartRaw);
            bool hasJurWithTaxingRights  = !string.IsNullOrEmpty(taxJurRaw);

            if (!hasEtrData && !hasInit && !hasJurWithTaxingRights) return;

            // ─── JurisdictionSection 찾기 또는 생성 ──────────────────────
            var js = globe.GlobeBody.JurisdictionSection
                .FirstOrDefault(s => s.Jurisdiction == countryCode);
            if (js == null)
            {
                js = new Globe.GlobeBodyTypeJurisdictionSection();
                js.Jurisdiction = countryCode;
                js.RecJurCode.Add(countryCode);
                js.GLoBeTax = new Globe.GlobeTax();
                globe.GlobeBody.JurisdictionSection.Add(js);
            }

            // ─── 과세권 국가 → JurisdictionSection.JurWithTaxingRights ──────
            // 복수 항목은 세미콜론으로 구분: "KR; JP" 또는 "KR, (GIR1101,...); JP"
            if (hasJurWithTaxingRights)
            {
                foreach (var entry in taxJurRaw.Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
                {
                    var jwr = ParseJwrEntry(entry, errors, fileName, loc);
                    if (jwr == null) continue;
                    js.JurWithTaxingRights.Add(jwr);
                }
            }

            // ─── ETR 항목 생성 ────────────────────────────────────────────
            if (hasEtrData)
            {
                var etr = new Globe.EtrType
                {
                    EtrStatus = new Globe.EtrTypeEtrStatus()
                };
                var exception = new Globe.EtrTypeEtrStatusEtrException();
                etr.EtrStatus.EtrException = exception;

                // DeminimisSimplifiedNmceCalc
                if (hasDemini)
                {
                    var dmCalc = new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc
                    {
                        Basis = deminiBasis ?? Globe.DeminimisSimpleBasisEnumType.Gir2901
                    };

                    var periodEnd = globe.GlobeBody.FilingInfo?.Period?.End ?? new DateTime(DateTime.Today.Year, 12, 31);

                    // 2.2.2 상세 데이터 우선, 없으면 2.2.1.2 간소화 데이터
                    bool useFullData = !string.IsNullOrEmpty(f1GbRev) || !string.IsNullOrEmpty(f1AcRev);

                    if (useFullData)
                    {
                        TryAddFinancialData(dmCalc, periodEnd,              f1AcRev, f1GbRev, f1GbPl, f1AcPl);
                        TryAddFinancialData(dmCalc, periodEnd.AddYears(-1), f2AcRev, f2GbRev, f2GbPl, f2AcPl);
                        TryAddFinancialData(dmCalc, periodEnd.AddYears(-2), f3AcRev, f3GbRev, f3GbPl, f3AcPl);

                        if (!string.IsNullOrEmpty(faGbRev) || !string.IsNullOrEmpty(faAcRev))
                        {
                            dmCalc.Average = new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcAverage
                            {
                                Revenue        = faAcRev,
                                GlobeRevenue   = faGbRev ?? "",
                                NetGlobeIncome = faGbPl  ?? "",
                                Fanil          = faAcPl  ?? ""
                            };
                        }
                    }
                    else
                    {
                        TryAddSimpleFinancialData(dmCalc, periodEnd,              s1Rev, s1Tax);
                        TryAddSimpleFinancialData(dmCalc, periodEnd.AddYears(-1), s2Rev, s2Tax);
                        TryAddSimpleFinancialData(dmCalc, periodEnd.AddYears(-2), s3Rev, s3Tax);

                        if (!string.IsNullOrEmpty(saRev) || !string.IsNullOrEmpty(saTax))
                        {
                            dmCalc.Average = new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcAverage
                            {
                                Revenue        = saRev,
                                GlobeRevenue   = "",
                                NetGlobeIncome = saTax ?? "",
                                Fanil          = ""
                            };
                        }
                    }

                    exception.DeminimisSimplifiedNmceCalc = dmCalc;
                }

                // TransitionalCbCrSafeHarbour
                if (hasCbcr)
                {
                    exception.TransitionalCbCrSafeHarbour =
                        new Globe.EtrTypeEtrStatusEtrExceptionTransitionalCbCrSafeHarbour
                        {
                            Revenue   = cbcrRev,
                            Profit    = cbcrPl ?? "",
                            IncomeTax = cbcrTax
                        };
                }

                // UtprSafeHarbour
                if (hasUtpr)
                {
                    if (decimal.TryParse(utprRate.TrimEnd('%').Trim(),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out var citRate))
                    {
                        exception.UtprSafeHarbour = new Globe.EtrTypeEtrStatusEtrExceptionUtprSafeHarbour
                        {
                            CitRate = citRate > 1m ? citRate / 100m : citRate
                        };
                    }
                    else
                    {
                        errors.Add($"[{fileName}] [{loc}/UTPR] CIT율 파싱 실패: '{utprRate}'");
                    }
                }

                js.GLoBeTax.Etr.Add(etr);
            }

            // ─── InitialIntActivity (2.3) ─────────────────────────────────
            if (hasInit && DateTime.TryParse(initStartRaw, out var startDate))
            {
                var init = new Globe.InitialIntActivityType { StartDate = startDate };

                if (!string.IsNullOrEmpty(initRefJur))
                {
                    if (TryParseEnum<Globe.CountryCodeType>(initRefJur, out var refCode))
                    {
                        init.ReferenceJurisdiction = new Globe.InitialIntActivityTypeReferenceJurisdiction
                        {
                            ResCountryCode     = refCode,
                            TangibleAssetValue = initRefAsset ?? ""
                        };
                    }
                    else
                        errors.Add($"[{fileName}] [{loc}/2.3] 준거국가 코드 파싱 실패: '{initRefJur}'");
                }

                if (!string.IsNullOrEmpty(initNumJur))
                    init.RfyNumberOfJurisdictions = initNumJur;

                // 적용면제 첨부 시트에서 OtherJurisdiction 읽기 (flat 구조: 헤더 1행 + 데이터)
                if (attachWs != null)
                    ReadOtherJurisdictions(attachWs, blockNum, init, errors, fileName, loc);

                js.GLoBeTax.InitialIntActivity = init;
            }
        }

        /// <summary>
        /// "국가코드[, (하위그룹유형, TIN, TIN유형, 발급국가)]" 형식 파싱.
        /// 예: "KR, (GIR1101, 123456790, GIR3001, KR)" 또는 "KR"
        /// </summary>
        private Globe.JurisdictionSectionTypeJurWithTaxingRights ParseJwrEntry(
            string entry, List<string> errors, string fileName, string loc)
        {
            string countryPart;
            string subgroupPart = null;

            // 쉼표는 있지만 괄호가 없으면 → 잘못된 형식 (복수 항목은 세미콜론으로 구분해야 함)
            var parenIdx = entry.IndexOf('(');
            if (parenIdx < 0 && entry.Contains(','))
            {
                errors.Add($"[{fileName}] [{loc}] 과세권 국가 형식 오류: '{entry}' — 복수 항목은 세미콜론(;)으로 구분하세요. 예) KR; JP");
                return null;
            }

            if (parenIdx >= 0)
            {
                countryPart = entry[..parenIdx].Trim().TrimEnd(',').Trim();
                var closeIdx = entry.IndexOf(')');
                if (closeIdx > parenIdx)
                    subgroupPart = entry[(parenIdx + 1)..closeIdx].Trim();
            }
            else
            {
                countryPart = entry.Trim();
            }

            if (!TryParseEnum<Globe.CountryCodeType>(countryPart, out var countryCode))
            {
                errors.Add($"[{fileName}] [{loc}] 과세권 국가 코드 '{countryPart}' 파싱 실패");
                return null;
            }

            var jwr = new Globe.JurisdictionSectionTypeJurWithTaxingRights
            {
                JurisdictionName = countryCode
            };

            if (!string.IsNullOrEmpty(subgroupPart))
            {
                var parts = subgroupPart.Split(',', StringSplitOptions.TrimEntries);
                var subgroup = new Globe.JurisdictionSectionTypeJurWithTaxingRightsSubgroup();

                if (parts.Length >= 1 && !string.IsNullOrEmpty(parts[0]))
                    SetEnum<Globe.TypeofSubGroupEnumType>(parts[0], v => subgroup.TypeofSubGroup.Add(v),
                        errors, fileName, new MappingEntry { Label = $"[{loc}] 과세권/하위그룹유형" });

                var tinVal     = parts.Length >= 2 ? parts[1] : null;
                var tinTypeStr = parts.Length >= 3 ? parts[2] : null;
                var issuedBy   = parts.Length >= 4 ? parts[3] : null;

                if (!string.IsNullOrEmpty(tinVal))
                {
                    var tin = new Globe.TinType { Value = tinVal };
                    if (!string.IsNullOrEmpty(tinTypeStr) && TryParseEnum<Globe.TinEnumType>(tinTypeStr, out var tinEnum))
                    {
                        tin.TypeOfTin = tinEnum;
                        tin.TypeOfTinSpecified = true;
                    }
                    if (!string.IsNullOrEmpty(issuedBy) && TryParseEnum<Globe.CountryCodeType>(issuedBy, out var issuedByCode))
                    {
                        tin.IssuedBy = issuedByCode;
                        tin.IssuedBySpecified = true;
                    }
                    subgroup.Tin = tin;
                }
                else
                {
                    subgroup.Tin = NoTin();
                }

                jwr.Subgroup.Add(subgroup);
            }

            return jwr;
        }

        /// <summary>
        /// "적용면제 첨부" 시트에서 OtherJurisdiction 데이터 읽기.
        /// 현재 구조: 헤더 1행(소재지국 / 유형자산) + 데이터 행.
        /// 블록번호에 해당하는 섹션(첨부N) 탐색 → 없으면 전체 flat 읽기.
        /// </summary>
        private void ReadOtherJurisdictions(IXLWorksheet attachWs, int blockNum,
            Globe.InitialIntActivityType init, List<string> errors, string fileName, string loc)
        {
            var lastRow = attachWs.LastRowUsed()?.RowNumber() ?? 1;

            // 첨부N 섹션 존재 여부 확인
            int sectionStart = -1;
            var target = $"첨부{blockNum}";
            for (int r = 1; r <= lastRow; r++)
            {
                if (attachWs.Cell(r, 2).GetString()?.Trim() == target)
                {
                    sectionStart = r + ATTACH_HEADER_ROWS; // 헤더 행 이후부터 데이터
                    break;
                }
            }

            // 첨부N 섹션이 없으면 스킵 (flat 구조에서는 각 블록이 따로 기록하지 않음)
            if (sectionStart < 0) return;

            for (int r = sectionStart; r <= lastRow; r++)
            {
                var col2 = attachWs.Cell(r, 2).GetString()?.Trim();
                if (col2?.StartsWith("첨부") == true) break; // 다음 섹션
                var assetRaw = attachWs.Cell(r, 3).GetString()?.Trim();
                if (string.IsNullOrEmpty(col2) && string.IsNullOrEmpty(assetRaw)) break;

                if (TryParseEnum<Globe.CountryCodeType>(col2 ?? "", out var otherCode))
                {
                    var other = new Globe.InitialIntActivityTypeOtherJurisdiction
                    {
                        TangibleAssetValue = assetRaw ?? ""
                    };
                    other.ResCountryCode.Add(otherCode);
                    init.OtherJurisdiction.Add(other);
                }
                else if (!string.IsNullOrEmpty(col2))
                    errors.Add($"[{fileName}] [{loc}/2.3/첨부{blockNum}] 국가코드 파싱 실패: '{col2}'");
            }
        }

        private static bool RowContains(IXLWorksheet ws, int row, string text)
        {
            for (int col = 2; col <= 6; col++)
            {
                var val = ws.Cell(row, col).GetString();
                if (val?.Contains(text) == true) return true;
            }
            return false;
        }

        private static void TryAddFinancialData(
            Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc dmCalc,
            DateTime year, string revenue, string globeRevenue, string netGlobeIncome, string fanil)
        {
            if (string.IsNullOrEmpty(globeRevenue) && string.IsNullOrEmpty(revenue)
                && string.IsNullOrEmpty(netGlobeIncome) && string.IsNullOrEmpty(fanil)) return;

            dmCalc.FinancialData.Add(
                new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcFinancialData
                {
                    Year           = year,
                    Revenue        = revenue,
                    GlobeRevenue   = globeRevenue   ?? "",
                    NetGlobeIncome = netGlobeIncome ?? "",
                    Fanil          = fanil          ?? ""
                });
        }

        private static void TryAddSimpleFinancialData(
            Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc dmCalc,
            DateTime year, string revenue, string simplifiedTax)
        {
            if (string.IsNullOrEmpty(revenue) && string.IsNullOrEmpty(simplifiedTax)) return;

            dmCalc.FinancialData.Add(
                new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcFinancialData
                {
                    Year           = year,
                    Revenue        = revenue,
                    GlobeRevenue   = "",
                    NetGlobeIncome = simplifiedTax ?? "",
                    Fanil          = ""
                });
        }
    }
}
