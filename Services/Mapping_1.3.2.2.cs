using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_2 : MappingBase
    {
        public Mapping_1_3_2_2() : base("mapping_1.3.2.2.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();

            var entity = new Globe.CorporateStructureTypeExcludedEntity();
            bool hasData = false;

            foreach (var (_, section) in Mapping.Sections)
            {
                foreach (var m in section.Mappings)
                {
                    ForEachValue(ws, m, fileName, errors, val =>
                    {
                        hasData = true;
                        switch (m.Target)
                        {
                            case "ExcludedEntity.Change":
                                entity.Change = ParseBool(val); break;
                            case "ExcludedEntity.Name":
                                entity.Name = val; break;
                            case "ExcludedEntity.Type":
                                SetEnum<Globe.ExcludedEntityEnumType>(val, v => entity.Type = v, errors, fileName, m); break;
                            default:
                                errors.Add($"[{fileName}] 셀 {m.Cell}: ExcludedEntity 알 수 없는 매핑 대상 '{m.Target}'"); break;
                        }
                    });
                }
            }

            if (hasData)
                globe.GlobeBody.GeneralSection.CorporateStructure.ExcludedEntity.Add(entity);
        }
    }
}
