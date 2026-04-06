using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_1 : MappingBase
    {
        public Mapping_1_3_1() : base("mapping_1.3.1.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();

            var upe = new Globe.CorporateStructureTypeUpe();
            globe.GlobeBody.GeneralSection.CorporateStructure.Upe.Add(upe);

            foreach (var (_, section) in Mapping.Sections)
            {
                foreach (var m in section.Mappings)
                {
                    ForEachValue(ws, m, fileName, errors, val =>
                    {
                        if (m.Target.StartsWith("Upe.OtherUpe."))
                        {
                            upe.OtherUpe ??= new Globe.CorporateStructureTypeUpeOtherUpe();
                            upe.OtherUpe.Id ??= new Globe.IdType();
                            SetIdValue(upe.OtherUpe.Id, m.Target.Replace("Upe.OtherUpe.", ""),
                                val, errors, fileName, m);
                        }
                        else if (m.Target.StartsWith("Upe.ExcludedUpe."))
                        {
                            upe.ExcludedUpe ??= new Globe.CorporateStructureTypeUpeExcludedUpe();
                            SetExcludedUpe(upe.ExcludedUpe, m.Target.Replace("Upe.ExcludedUpe.", ""),
                                val, errors, fileName, m);
                        }
                    });
                }
            }
        }

        private void SetIdValue(Globe.IdType id, string path, string value,
            List<string> errors, string fileName, MappingEntry entry)
        {
            switch (path)
            {
                case "Id.Name": id.Name = value; break;
                case "Id.ResCountryCode":
                    SetEnum<Globe.CountryCodeType>(value, v => id.ResCountryCode.Add(v), errors, fileName, entry); break;
                case "Id.Rules":
                    SetEnum<Globe.IdTypeRulesEnumType>(value, v => id.Rules.Add(v), errors, fileName, entry); break;
                case "Id.Tin.Value":
                    id.Tin.Add(new Globe.TinType { Value = value }); break;
                case "Id.ReceivingTin":
                    id.Tin.Add(new Globe.TinType { Value = value, IssuedBy = Globe.CountryCodeType.Kr, IssuedBySpecified = true }); break;
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
                    SetEnum<Globe.CountryCodeType>(value, v => { excUpe.Art1035 = v; excUpe.Art1035Specified = true; }, errors, fileName, entry); break;
                default:
                    errors.Add($"[{fileName}] 셀 {entry.Cell}: ExcludedUPE 알 수 없는 경로 '{path}'"); break;
            }
        }
    }
}
