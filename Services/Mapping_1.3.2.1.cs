using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_1 : MappingBase
    {
        // 헤더 검색 키워드 (B열에서 이 문자열을 포함하는 행 = 블록 시작)
        private const string BLOCK_HEADER = "1.3.2.1";

        // 블록 내 상대 오프셋 (헤더 행 기준)
        // +1=변동, +2=소재지국, +3=규칙, +4=상호, +5=TIN,
        // +6=접수TIN, +7=기업유형, +12=모기업유형, +13=QIIR TIN,
        // +14=부분소유TIN, +15=QUTPR초기, +16=소유합계, +17=UPE소유
        private static readonly (int Offset, string Target)[] FieldMap =
        {
            (1, "Ce.ChangeFlag"),
            (2, "Ce.Id.ResCountryCode"),
            (3, "Ce.Id.Rules"),
            (4, "Ce.Id.Name"),
            (5, "Ce.Id.Tin.Value"),
            (6, "Ce.Id.ReceivingTin"),
            (7, "Ce.Id.GlobeStatus"),
            (12, "Ce.Qiir.PopeIpe"),
            (13, "Ce.Qiir.Exception.Art213.Tin"),  // O16: Art2.1.3 해당 모기업 TIN
            (14, "Ce.Qiir.Exception.Art215.Tin"),  // O17: Art2.1.5 해당 부분소유모기업 TIN
            (15, "Ce.Qutpr.Art93"),
            (16, "Ce.Qutpr.AggOwnership"),
            (17, "Ce.Qutpr.UpeOwnership"),
        };

        // 별첨 시트 이름
        private const string ATTACH_SHEET = "그룹구조 첨부";

        public Mapping_1_3_2_1()
            : base("mapping_1.3.2.1.json") { }

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??=
                new Globe.CorporateStructureType();
            var cs = globe.GlobeBody.GeneralSection.CorporateStructure;

            // 헤더 행 위치를 동적으로 검색
            var blockStartRows = FindBlockStartRows(ws);

            for (int blockIdx = 0; blockIdx < blockStartRows.Count; blockIdx++)
            {
                var blockStartRow = blockStartRows[blockIdx];
                var ce = new Globe.CorporateStructureTypeCe { Id = new Globe.IdType() };
                bool hasData = false;

                // 블록 내 필드 매핑
                foreach (var (offset, target) in FieldMap)
                {
                    var row = blockStartRow + offset;
                    var cellValue = ws.Cell(row, 15).GetString()?.Trim(); // O열 = 15
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    hasData = true;
                    var isMulti = target is "Ce.Id.Rules" or "Ce.Id.GlobeStatus";
                    var values = isMulti
                        ? cellValue.Split(
                            ',',
                            System.StringSplitOptions.RemoveEmptyEntries
                                | System.StringSplitOptions.TrimEntries
                        )
                        : new[] { cellValue };

                    foreach (var val in values)
                        SetCeValue(ce, cs, target, val, errors, fileName, row);
                }

                // 별첨에서 소유지분 읽기 (블록 순서 = 별첨 번호)
                MapOwnershipFromAttach(ws.Workbook, ce, blockIdx + 1, errors, fileName);

                if (hasData)
                    cs.Ce.Add(ce);
            }
        }

        private void SetCeValue(
            Globe.CorporateStructureTypeCe ce,
            Globe.CorporateStructureType cs,
            string target,
            string val,
            List<string> errors,
            string fileName,
            int row
        )
        {
            var entry = new MappingEntry { Cell = $"O{row}", Label = target };

            switch (target)
            {
                case "Ce.ChangeFlag":
                    cs.UnreportChangeCorpStr = ParseBool(val);
                    cs.UnreportChangeCorpStrSpecified = true;
                    break;
                case "Ce.Id.ResCountryCode":
                    SetEnum<Globe.CountryCodeType>(
                        val,
                        v => ce.Id.ResCountryCode.Add(v),
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Id.Rules":
                    SetEnum<Globe.IdTypeRulesEnumType>(
                        val,
                        v => ce.Id.Rules.Add(v),
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Id.Name":
                    var (ceName, ceKName) = ParseNameKName(val);
                    ce.Id.Name = ceName;
                    if (ceKName != null) ce.Id.KName = ceKName;
                    break;
                case "Ce.Id.Tin.Value":
                    ce.Id.Tin.Add(ParseTin(val));
                    break;
                case "Ce.Id.ReceivingTin":
                    ce.Id.Tin.Add(ParseTin(val));
                    break;
                case "Ce.Id.GlobeStatus":
                    SetEnum<Globe.IdTypeGloBeStatusEnumType>(
                        val,
                        v => ce.Id.GlobeStatus.Add(v),
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Qiir.PopeIpe":
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    SetEnum<Globe.PopeipeEnumType>(
                        val,
                        v => ce.Qiir.PopeIpe = v,
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Qiir.Exception.Art213.Tin":
                    // Art2.1.3: 최종모기업/중간모기업에 QIIR 적용 시 해당 모기업 TIN
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    ce.Qiir.Exception ??= new Globe.CorporateStructureTypeCeQiirException();
                    ce.Qiir.Exception.ExceptionRule ??= new Globe.CorporateStructureTypeCeQiirExceptionExceptionRule();
                    ce.Qiir.Exception.ExceptionRule.Art213 = true;
                    ce.Qiir.Exception.ExceptionRule.Art213Specified = true;
                    ce.Qiir.Exception.Tin = ParseTin(val);
                    break;
                case "Ce.Qiir.Exception.Art215.Tin":
                    // Art2.1.5: 부분소유모기업에 QIIR 적용 시 해당 모기업 TIN
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    ce.Qiir.Exception ??= new Globe.CorporateStructureTypeCeQiirException();
                    ce.Qiir.Exception.ExceptionRule ??= new Globe.CorporateStructureTypeCeQiirExceptionExceptionRule();
                    ce.Qiir.Exception.ExceptionRule.Art215 = true;
                    ce.Qiir.Exception.ExceptionRule.Art215Specified = true;
                    ce.Qiir.Exception.Tin = ParseTin(val);
                    break;
                case "Ce.Qutpr.Art93":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    ce.Qutpr.Art93 = ParseBool(val);
                    break;
                case "Ce.Qutpr.AggOwnership":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    if (decimal.TryParse(val.TrimEnd('%').Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var agg))
                    {
                        ce.Qutpr.AggOwnership = agg > 1m ? agg / 100m : agg;
                        ce.Qutpr.AggOwnershipSpecified = true;
                    }
                    break;
                case "Ce.Qutpr.UpeOwnership":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    ce.Qutpr.UpeOwnership = ParseBool(val);
                    ce.Qutpr.UpeOwnershipSpecified = true;
                    break;
            }
        }

        /// <summary>
        /// 별첨 시트에서 별첨N의 주주 데이터를 읽어 Ownership에 추가.
        /// </summary>
        private void MapOwnershipFromAttach(
            IXLWorkbook workbook,
            Globe.CorporateStructureTypeCe ce,
            int attachNum,
            List<string> errors,
            string fileName
        )
        {
            if (!workbook.TryGetWorksheet(ATTACH_SHEET, out var attachWs))
                return;

            // "별첨N" 제목 행 찾기
            var startRow = FindAttachStart(attachWs, attachNum);
            if (startRow < 0)
                return;

            // 제목(1) + 빈행(1) + 헤더(1) = 3행 뒤부터 데이터
            var dataRow = startRow + 3;
            while (true)
            {
                var typeVal = attachWs.Cell(dataRow, 2).GetString()?.Trim();
                var tinVal = attachWs.Cell(dataRow, 3).GetString()?.Trim();
                var pctVal = attachWs.Cell(dataRow, 4).GetString()?.Trim();

                if (
                    string.IsNullOrEmpty(typeVal)
                    && string.IsNullOrEmpty(tinVal)
                    && string.IsNullOrEmpty(pctVal)
                )
                    break;

                // 다음 별첨 제목이면 종료
                if (typeVal != null && typeVal.StartsWith("첨부"))
                    break;

                var ownership = new Globe.CorporateStructureTypeCeOwnership();

                if (!string.IsNullOrEmpty(typeVal))
                    SetEnum<Globe.OwnershipTypeEnumType>(
                        typeVal,
                        v => ownership.OwnershipType = v,
                        errors,
                        fileName,
                        new MappingEntry
                        {
                            Cell = $"첨부!B{dataRow}",
                            Label = $"첨부{attachNum} 유형",
                        }
                    );

                // TIN 없으면 NOTIN 처리 (required 필드)
                ownership.Tin = !string.IsNullOrEmpty(tinVal) ? ParseTin(tinVal) : NoTin();

                if (!string.IsNullOrEmpty(pctVal))
                {
                    // "90%" → "90", "0.9" → "0.9" 정규화 후 파싱
                    var pctClean = pctVal.TrimEnd('%').Trim();
                    if (decimal.TryParse(pctClean,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var pct))
                    {
                        // >1이면 퍼센트 단위(예: 90 → 0.9), ≤1이면 이미 소수(예: 0.9 → 그대로)
                        ownership.OwnershipPercentage = pct > 1m ? pct / 100m : pct;
                    }
                }

                ce.Ownership.Add(ownership);
                dataRow++;
            }
        }

        /// <summary>
        /// B열에서 BLOCK_HEADER 문자열을 포함하는 행을 모두 찾아 반환.
        /// 각 행이 CE 블록의 시작점(헤더 행).
        /// </summary>
        private static List<int> FindBlockStartRows(IXLWorksheet ws)
        {
            var result = new List<int>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 200;
            for (int r = 1; r <= lastRow; r++)
            {
                var val = ws.Cell(r, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(val) && val.Contains(BLOCK_HEADER))
                    result.Add(r);
            }
            return result;
        }

        private static int FindAttachStart(IXLWorksheet ws, int attachNum)
        {
            var target = $"첨부{attachNum}";
            for (int r = 1; r <= 500; r++)
            {
                var val = ws.Cell(r, 2).GetString()?.Trim();
                if (val == target)
                    return r;
            }
            return -1;
        }
    }
}
