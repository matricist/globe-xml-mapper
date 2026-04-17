using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class MappingOrchestrator
    {
        // 섹션키 → 매퍼 생성 팩토리
        private static readonly Dictionary<string, Func<MappingBase>> MapperFactory = new()
        {
            { "1.1~1.2",  () => new Mapping_1_1_1_2() },
            { "1.3.1",    () => new Mapping_1_3_1() },
            { "1.3.2.1",  () => new Mapping_1_3_2_1() },
            { "1.3.2.2",  () => new Mapping_1_3_2_2() },
            { "1.3.3",    () => new Mapping_1_3_3() },
            { "1.4",      () => new Mapping_1_4() },
            { "2",        () => new Mapping_2() },
            { "UTPR",     () => new Mapping_Utpr() },
            { "JurCal",   () => new Mapping_JurCal() },
            { "EntityCe", () => new Mapping_EntityCe() },
        };

        /// <summary>
        /// 단일 Workbook + _META 숨김시트 기반 매핑.
        /// ControlPanelForm에서 호출.
        /// </summary>
        public List<string> MapWorkbook(string filePath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            using var workbook = new XLWorkbook(filePath);

            foreach (var (section, sheetName) in ExcelController.SheetMap)
            {
                if (!MapperFactory.TryGetValue(section, out var createMapper))
                    continue; // 매퍼 없는 섹션은 스킵 (1.3.3 등 XML 미포함)

                if (!workbook.TryGetWorksheet(sheetName, out var ws))
                    continue; // 시트가 없으면 해당 섹션은 건너뜀

                var mapper = createMapper();
                mapper.Map(ws, globe, errors, sheetName);
            }

            FillMessageSpec(globe);
            return errors;
        }

        #region MessageSpec / DocSpec

        private void FillMessageSpec(Globe.GlobeOecd globe)
        {
            var spec = globe.MessageSpec;
            var fi = globe.GlobeBody?.FilingInfo;

            if (fi?.FilingCe != null)
                spec.TransmittingCountry = fi.FilingCe.ResCountryCode;

            spec.ReceivingCountry = spec.TransmittingCountry;
            spec.MessageType = Globe.MessageTypeEnumType.Gir;

            if (fi?.Period != null && fi.Period.End != default)
                spec.ReportingPeriod = fi.Period.End;

            spec.Timestamp = DateTime.Now;

            if (string.IsNullOrEmpty(spec.MessageRefId))
            {
                var sendCC = spec.TransmittingCountry.ToString().ToUpper();
                var recvCC = spec.ReceivingCountry.ToString().ToUpper();
                var uid = spec.Timestamp.ToString("yyyyMMddHHmmss");
                spec.MessageRefId = $"{sendCC}{spec.ReportingPeriod:yyyy}{recvCC}{uid}";
            }

            FillDocSpecs(globe);
        }

        private void FillDocSpecs(Globe.GlobeOecd globe)
        {
            var sendCC = globe.MessageSpec.TransmittingCountry.ToString().ToUpper();
            var year = globe.MessageSpec.ReportingPeriod.ToString("yyyy");
            var ts = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            if (globe.GlobeBody.FilingInfo != null)
            {
                globe.GlobeBody.FilingInfo.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}FI{ts}"
                };
            }

            if (globe.GlobeBody.GeneralSection != null)
            {
                globe.GlobeBody.GeneralSection.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}GS{ts}"
                };
            }

            int utprIdx = 0;
            foreach (var ua in globe.GlobeBody.UtprAttribution)
            {
                ua.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}UA{utprIdx++}{ts}"
                };
            }
        }

        #endregion
    }
}
