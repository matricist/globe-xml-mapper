using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 시트 2: 국가별 적용면제 및 제외.
    /// 블록 반복: 블록1(3~23) + gap(2) + 블록2(26~54) = 52행 세트.
    /// blockCount 기반으로 N개 국가 순회.
    /// </summary>
    public class Mapping_2 : MappingBase
    {
        private const int BLOCK1_START = 2;
        private const int BLOCK1_SIZE = 21;  // 2~22
        private const int GAP = 2;           // 24~25
        private const int BLOCK2_SIZE = 29;  // 26~54
        private const int SET_SIZE = 52;     // 21+2+29
        private const int SET_GAP = 2;       // 세트 간 간격

        public Mapping_2() : base("mapping_2.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            var blockCount = 1;
            if (ws.Workbook.TryGetWorksheet(ExcelController.MetaSheetName, out var metaWs))
                blockCount = ExcelController.ReadBlockCount(metaWs, ws.Name);

            for (int idx = 0; idx < blockCount; idx++)
            {
                // 각 세트의 시작 행 계산
                var b1Start = BLOCK1_START + idx * (SET_SIZE + SET_GAP);
                var b2Start = b1Start + BLOCK1_SIZE + GAP;

                MapOneCountry(ws, globe, errors, fileName, b1Start, b2Start);
            }
        }

        private void MapOneCountry(IXLWorksheet ws, Globe.GlobeOecd globe,
            List<string> errors, string fileName, int b1, int b2)
        {
            // b1 기준 오프셋 (원본: 2행 기준)
            // O5 = b1+3, O6 = b1+4, O7 = b1+5, O8 = b1+6, O9 = b1+7
            // O14 = b1+12
            // H18 = b1+16, N18 = b1+16, H19 = b1+17 ...
            // b2 기준 (원본: 25행 기준)
            // O27 = b2+2, O28 = b2+3, O29 = b2+4, O32 = b2+7
            // E38 = b2+13, I38 = b2+13, ...
            // M44 = b2+19, M45 = b2+20, M46 = b2+21, M47 = b2+22

            var jurCode = ws.Cell(b1 + 3, 15).GetString()?.Trim(); // O열=15
            if (string.IsNullOrEmpty(jurCode)) return;

            // 기존 Summary에서 같은 JurisdictionName 찾기 또는 새로 생성
            Globe.GlobeBodyTypeSummary summary = null;
            if (TryParseEnum<Globe.CountryCodeType>(jurCode, out var countryCode))
            {
                summary = globe.GlobeBody.Summary
                    .FirstOrDefault(s => s.Jurisdiction?.JurisdictionNameSpecified == true
                                      && s.Jurisdiction.JurisdictionName == countryCode);
            }

            if (summary == null)
            {
                summary = new Globe.GlobeBodyTypeSummary();
                summary.Jurisdiction = new Globe.SummaryTypeJurisdiction();
                if (TryParseEnum<Globe.CountryCodeType>(jurCode, out var cc))
                {
                    summary.Jurisdiction.JurisdictionName = cc;
                    summary.Jurisdiction.JurisdictionNameSpecified = true;
                }
                globe.GlobeBody.Summary.Add(summary);
            }

            // === 2.1 기본사항 (블록1) ===

            // 하위그룹
            var subGroupType = ws.Cell(b1 + 4, 15).GetString()?.Trim();
            var subGroupTin = ws.Cell(b1 + 5, 15).GetString()?.Trim();
            if (!string.IsNullOrEmpty(subGroupType) || !string.IsNullOrEmpty(subGroupTin))
            {
                var subgroup = new Globe.SummaryTypeJurisdictionSubgroup();
                if (!string.IsNullOrEmpty(subGroupTin))
                    subgroup.Tin = new Globe.TinType { Value = subGroupTin };
                summary.Jurisdiction.Subgroup.Add(subgroup);
            }

            // 과세권 국가
            var taxJur = ws.Cell(b1 + 6, 15).GetString()?.Trim();
            if (!string.IsNullOrEmpty(taxJur))
            {
                var jwr = new Globe.SummaryTypeJurWithTaxingRights();
                SetEnum<Globe.CountryCodeType>(taxJur, v =>
                {
                    jwr.JurisdictionName = v;
                    jwr.JurisdictionNameSpecified = true;
                }, errors, fileName, new MappingEntry { Cell = $"O{b1 + 6}", Label = "과세권 국가" });
                summary.JurWithTaxingRights.Add(jwr);
            }

            // === 2.2.1 적용면제 (블록1) ===
            var safeHarbour = ws.Cell(b1 + 12, 15).GetString()?.Trim();
            if (!string.IsNullOrEmpty(safeHarbour))
                SetEnum<Globe.SafeHarbourEnumType>(safeHarbour, v =>
                {
                    if (!summary.SafeHarbour.Contains(v))
                        summary.SafeHarbour.Add(v);
                }, errors, fileName, new MappingEntry { Cell = $"O{b1 + 12}", Label = "적용면제 유형" });

            // === 2.2.1.3 전환기 (블록2) ===
            // O28=b2+2, O29=b2+3, O30=b2+4, O33=b2+7
            // 현재 Globe Summary에 직접 대응 필드 없음 — 추후 JurisdictionSection에서 처리

            // === 2.3 해외진출 초기 특례 (블록2) ===
            // M45=b2+19, M46=b2+20
            // 현재 Globe Summary에 직접 대응 필드 없음 — 추후 처리
        }
    }
}
