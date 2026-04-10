using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_2 : MappingBase
    {
        // BlockService / ControlPanelForm의 EX_BLOCK 상수와 동일하게 유지
        private const int BLOCK_START = 2;  // 헤더 포함 블록 첫 행
        private const int BLOCK_END   = 5;  // 블록 마지막 행
        private const int BLOCK_GAP   = 2;  // 블록 간 구분 행 수
        private const int SET_SIZE    = BLOCK_END - BLOCK_START + 1 + BLOCK_GAP; // = 6

        // 블록 내 상대 오프셋 (BLOCK_START 기준, O열 = 15)
        // +0 = 헤더행("1.3.2.2 제외기업"), 데이터 없음
        // +1 = 1. 변동 여부
        // +2 = 2. 제외기업 상호
        // +3 = 3. 제외기업 유형
        private static readonly (int Offset, string Target)[] FieldMap =
        {
            (1, "Change"),
            (2, "Name"),
            (3, "Type"),
        };

        public Mapping_1_3_2_2() : base("mapping_1.3.2.2.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int n = 0; ; n++)
            {
                var blockStart = BLOCK_START + n * SET_SIZE;
                if (blockStart > lastRow) break;

                var entity = new Globe.CorporateStructureTypeExcludedEntity();
                bool hasData = false;

                foreach (var (offset, target) in FieldMap)
                {
                    var row = blockStart + offset;
                    var val = ws.Cell(row, 15).GetString()?.Trim();
                    if (string.IsNullOrEmpty(val)) continue;

                    hasData = true;
                    var entry = new MappingEntry { Cell = $"O{row}", Label = target };

                    switch (target)
                    {
                        case "Change":
                            entity.Change = ParseBool(val); break;
                        case "Name":
                            var (eName, eKName) = ParseNameKName(val);
                            entity.Name = eName;
                            if (eKName != null) entity.KName = eKName;
                            break;
                        case "Type":
                            SetEnum<Globe.ExcludedEntityEnumType>(val, v => entity.Type = v, errors, fileName, entry); break;
                    }
                }

                if (hasData)
                    globe.GlobeBody.GeneralSection.CorporateStructure.ExcludedEntity.Add(entity);
            }
        }
    }
}
