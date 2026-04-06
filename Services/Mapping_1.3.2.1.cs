using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_1 : MappingBase
    {
        public Mapping_1_3_2_1() : base("mapping_1.3.2.1.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();
            var cs = globe.GlobeBody.GeneralSection.CorporateStructure;

            var ce = new Globe.CorporateStructureTypeCe { Id = new Globe.IdType() };
            cs.Ce.Add(ce);

            // 메인 시트 매핑 (소유지분 제외)
            foreach (var (_, section) in Mapping.Sections)
            {
                foreach (var m in section.Mappings)
                {
                    ForEachValue(ws, m, fileName, errors, val =>
                    {
                        switch (m.Target)
                        {
                            case "Ce.ChangeFlag":
                                cs.UnreportChangeCorpStr = ParseBool(val);
                                cs.UnreportChangeCorpStrSpecified = true; break;

                            case "Ce.Id.ResCountryCode":
                                SetEnum<Globe.CountryCodeType>(val, v => ce.Id.ResCountryCode.Add(v), errors, fileName, m); break;
                            case "Ce.Id.Rules":
                                SetEnum<Globe.IdTypeRulesEnumType>(val, v => ce.Id.Rules.Add(v), errors, fileName, m); break;
                            case "Ce.Id.Name": ce.Id.Name = val; break;
                            case "Ce.Id.Tin.Value":
                                ce.Id.Tin.Add(new Globe.TinType { Value = val }); break;
                            case "Ce.Id.ReceivingTin":
                                ce.Id.Tin.Add(new Globe.TinType { Value = val, IssuedBy = Globe.CountryCodeType.Kr, IssuedBySpecified = true }); break;
                            case "Ce.Id.GlobeStatus":
                                SetEnum<Globe.IdTypeGloBeStatusEnumType>(val, v => ce.Id.GlobeStatus.Add(v), errors, fileName, m); break;

                            case "Ce.Qiir.PopeIpe":
                                ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                                SetEnum<Globe.PopeipeEnumType>(val, v => ce.Qiir.PopeIpe = v, errors, fileName, m); break;
                            case "Ce.Qiir.Exception.Tin.Value":
                                ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                                ce.Qiir.Exception ??= new Globe.CorporateStructureTypeCeQiirException();
                                ce.Qiir.Exception.Tin = new Globe.TinType { Value = val }; break;
                            case "Ce.Qiir.MopeIpe.Tin.Value":
                                break;

                            case "Ce.Qutpr.Art93":
                                ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                                ce.Qutpr.Art93 = ParseBool(val); break;
                            case "Ce.Qutpr.AggOwnership":
                                ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                                if (decimal.TryParse(val, out var agg))
                                { ce.Qutpr.AggOwnership = agg >= 1 ? agg / 100m : agg; ce.Qutpr.AggOwnershipSpecified = true; }
                                break;
                            case "Ce.Qutpr.UpeOwnership":
                                ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                                ce.Qutpr.UpeOwnership = ParseBool(val);
                                ce.Qutpr.UpeOwnershipSpecified = true; break;

                            default:
                                errors.Add($"[{fileName}] 셀 {m.Cell}: CE 알 수 없는 매핑 대상 '{m.Target}'"); break;
                        }
                    });
                }
            }

            // 별첨 시트에서 소유지분 읽기 (B=유형, C=납세자번호, D=소유지분%)
            MapOwnershipFromAttachment(ws.Workbook, ce, errors, fileName);
        }

        private void MapOwnershipFromAttachment(
            IXLWorkbook workbook, Globe.CorporateStructureTypeCe ce,
            List<string> errors, string fileName)
        {
            const string attachSheetName = "별첨";
            if (!workbook.TryGetWorksheet(attachSheetName, out var attachWs))
                return;

            // 2행부터 데이터 (1행은 헤더: 유형, 납세자번호, 소유지분%)
            var row = 3;
            while (true)
            {
                var typeVal = attachWs.Cell(row, 2).GetString()?.Trim();  // B열: 유형
                var tinVal = attachWs.Cell(row, 3).GetString()?.Trim();   // C열: 납세자번호
                var pctVal = attachWs.Cell(row, 4).GetString()?.Trim();   // D열: 소유지분%

                if (string.IsNullOrEmpty(typeVal) && string.IsNullOrEmpty(tinVal) && string.IsNullOrEmpty(pctVal))
                    break;

                var ownership = new Globe.CorporateStructureTypeCeOwnership();

                if (!string.IsNullOrEmpty(typeVal))
                    SetEnum<Globe.OwnershipTypeEnumType>(typeVal, v => ownership.OwnershipType = v,
                        errors, fileName, new MappingEntry { Cell = $"별첨!B{row}", Label = "소유지분 유형" });

                if (!string.IsNullOrEmpty(tinVal))
                    ownership.Tin = new Globe.TinType { Value = tinVal };

                if (!string.IsNullOrEmpty(pctVal) && decimal.TryParse(pctVal, out var pct))
                    ownership.OwnershipPercentage = pct / 100m;

                ce.Ownership.Add(ownership);
                row++;
            }
        }
    }
}
