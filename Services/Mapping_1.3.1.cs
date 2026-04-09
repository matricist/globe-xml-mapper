using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_1 : MappingBase
    {
        private const string BLOCK_HEADER = "1.3.1";

        // 헤더 행 기준 상대 오프셋 → (offset, target)
        // 헤더=row, J(col 10) 기준:
        // +1=소재지국, +2=규칙, +3=상호, +4=TIN, +5=접수TIN, +6=기업유형,
        // +7=제외기업유형, +8=103조국가 (OtherUPE/ExcludedUPE 공통, 블록 처리 후 분기)
        private static readonly (int Offset, string Target)[] FieldMap =
        {
            (1, "Upe.OtherUpe.Id.ResCountryCode"),
            (2, "Upe.OtherUpe.Id.Rules"),
            (3, "Upe.OtherUpe.Id.Name"),
            (4, "Upe.OtherUpe.Id.Tin.Value"),
            (5, "Upe.OtherUpe.Id.ReceivingTin"),
            (6, "Upe.OtherUpe.Id.GlobeStatus"),
            (7, "Upe.ExcludedUpe.ExcludedUpeStatus"),
            (8, "Upe.Art1035"),   // OtherUPE/ExcludedUPE 공통 — 후처리에서 분기
        };

        public Mapping_1_3_1() : base("mapping_1.3.1.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();

            var blockStartRows = FindBlockStartRows(ws);

            foreach (var blockStartRow in blockStartRows)
            {
                var upe = new Globe.CorporateStructureTypeUpe();
                bool hasData = false;
                bool hasExcludedStatus = false;
                string art1035Raw = null;   // 후처리에서 OtherUpe/ExcludedUpe에 분기

                foreach (var (offset, target) in FieldMap)
                {
                    var row = blockStartRow + offset;
                    var cellValue = ws.Cell(row, 10).GetString()?.Trim();  // J열 = 10
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    hasData = true;

                    if (target == "Upe.Art1035") { art1035Raw = cellValue; continue; }
                    if (target == "Upe.ExcludedUpe.ExcludedUpeStatus") hasExcludedStatus = true;

                    var isMulti = target is "Upe.OtherUpe.Id.Rules" or "Upe.OtherUpe.Id.GlobeStatus";
                    var values = isMulti
                        ? cellValue.Split(',', System.StringSplitOptions.RemoveEmptyEntries | System.StringSplitOptions.TrimEntries)
                        : new[] { cellValue };

                    foreach (var val in values)
                    {
                        var entry = new MappingEntry { Cell = $"J{row}", Label = target };
                        if (target.StartsWith("Upe.OtherUpe."))
                        {
                            upe.OtherUpe ??= new Globe.CorporateStructureTypeUpeOtherUpe();
                            upe.OtherUpe.Id ??= new Globe.IdType();
                            SetIdValue(upe.OtherUpe.Id, target.Replace("Upe.OtherUpe.", ""), val, errors, fileName, entry);
                        }
                        else if (target.StartsWith("Upe.ExcludedUpe."))
                        {
                            upe.ExcludedUpe ??= new Globe.CorporateStructureTypeUpeExcludedUpe();
                            SetExcludedUpe(upe.ExcludedUpe, target.Replace("Upe.ExcludedUpe.", ""), val, errors, fileName, entry);
                        }
                    }
                }

                // ExcludedUpeStatus 있으면 OtherUpe → ExcludedUpe 전환
                if (hasExcludedStatus && upe.OtherUpe?.Id != null)
                {
                    var src = upe.OtherUpe.Id;
                    upe.ExcludedUpe.Id ??= new Globe.ExcludedUpeIdType();
                    var dst = upe.ExcludedUpe.Id;
                    if (src.Name  != null) dst.Name  = src.Name;
                    if (src.KName != null) dst.KName = src.KName;
                    foreach (var v in src.ResCountryCode) dst.ResCountryCode.Add(v);
                    foreach (var v in src.Tin)            dst.Tin.Add(v);
                    foreach (var v in src.Rules)          dst.Rules.Add(v);
                    foreach (var v in src.GlobeStatus)    dst.GlobeStatus.Add(v);
                    if (dst.Tin.Count == 0) dst.Tin.Add(NoTin());
                    upe.OtherUpe = null;
                }
                else if (!hasExcludedStatus && upe.ExcludedUpe != null)
                {
                    upe.ExcludedUpe = null;
                }

                // Art10.3.5 — ExcludedUPE면 ExcludedUpe에, OtherUPE면 OtherUpe에
                if (art1035Raw != null)
                {
                    var art1035Val = art1035Raw.Split(',')[0].Trim();
                    var entry = new MappingEntry { Cell = "J(Art10.3.5)", Label = "Art10.3.5" };
                    if (hasExcludedStatus)
                    {
                        upe.ExcludedUpe ??= new Globe.CorporateStructureTypeUpeExcludedUpe();
                        SetEnum<Globe.CountryCodeType>(art1035Val,
                            v => { upe.ExcludedUpe.Art1035 = v; upe.ExcludedUpe.Art1035Specified = true; },
                            errors, fileName, entry);
                    }
                    else
                    {
                        upe.OtherUpe ??= new Globe.CorporateStructureTypeUpeOtherUpe();
                        SetEnum<Globe.CountryCodeType>(art1035Val,
                            v => { upe.OtherUpe.Art1035 = v; upe.OtherUpe.Art1035Specified = true; },
                            errors, fileName, entry);
                    }
                }

                if (hasData)
                    globe.GlobeBody.GeneralSection.CorporateStructure.Upe.Add(upe);
            }
        }

        private static List<int> FindBlockStartRows(IXLWorksheet ws)
        {
            var result = new List<int>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 200;
            for (int r = 1; r <= lastRow; r++)
            {
                var val = ws.Cell(r, 2).GetString()?.Trim(); // B열
                if (!string.IsNullOrEmpty(val) && val.Contains(BLOCK_HEADER))
                    result.Add(r);
            }
            return result;
        }

        private void SetIdValue(Globe.IdType id, string path, string value,
            List<string> errors, string fileName, MappingEntry entry)
        {
            switch (path)
            {
                case "Id.Name":
                    var (n, kn) = ParseNameKName(value);
                    id.Name = n;
                    if (kn != null) id.KName = kn;
                    break;
                case "Id.ResCountryCode":
                    SetEnum<Globe.CountryCodeType>(value, v => id.ResCountryCode.Add(v), errors, fileName, entry); break;
                case "Id.Rules":
                    SetEnum<Globe.IdTypeRulesEnumType>(value, v => id.Rules.Add(v), errors, fileName, entry); break;
                case "Id.Tin.Value":
                    id.Tin.Add(ParseTin(value)); break;
                case "Id.ReceivingTin":
                    id.Tin.Add(ParseTin(value)); break;
                case "Id.GlobeStatus":
                    SetEnum<Globe.IdTypeGloBeStatusEnumType>(value, v => id.GlobeStatus.Add(v), errors, fileName, entry); break;
                default:
                    errors.Add($"[{fileName}] 셀 {entry.Cell}: UPE 알 수 없는 경로 '{path}'"); break;
            }
        }

        private void SetExcludedUpe(Globe.CorporateStructureTypeUpeExcludedUpe excUpe,
            string path, string value, List<string> errors, string fileName, MappingEntry entry)
        {
            switch (path)
            {
                case "ExcludedUpeStatus":
                    SetEnum<Globe.ExcludedUpeEnumType>(value, v => excUpe.ExcludedUpeStatus = v, errors, fileName, entry); break;
                case "Art1035":
                    // Art10.3.5는 Map() 후처리에서 직접 처리됨 (여기 도달 안 함)
                    break;
                default:
                    errors.Add($"[{fileName}] 셀 {entry.Cell}: ExcludedUPE 알 수 없는 경로 '{path}'"); break;
            }
        }
    }
}
