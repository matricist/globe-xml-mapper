using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// UTPR 배분 시트 → GlobeBody.UtprAttribution 매핑.
    /// 4행부터 데이터, 합계 행(국가코드 파싱 불가)은 자동 스킵.
    /// </summary>
    public class Mapping_Utpr : MappingBase
    {
        private const int DATA_START_ROW = 4;

        public Mapping_Utpr() : base(null) { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? DATA_START_ROW;

            var utpr = new Globe.GlobeBodyTypeUtprAttribution();

            // RecJurCode: 신고구성기업 소재지국
            var filingCountry = globe.GlobeBody.FilingInfo?.FilingCe?.ResCountryCode;
            if (filingCountry.HasValue)
                utpr.RecJurCode.Add(filingCountry.Value);

            bool hasData = false;

            for (int row = DATA_START_ROW; row <= lastRow; row++)
            {
                var resCodeRaw = ws.Cell(row, 2).GetString()?.Trim(); // B
                if (string.IsNullOrEmpty(resCodeRaw)) continue;

                // 합계 행 등 국가코드 파싱 불가 → 스킵
                if (!TryParseEnum<Globe.CountryCodeType>(resCodeRaw, out var country)) continue;

                var attr = new Globe.UtprAttributionTypeAttribution
                {
                    ResCountryCode          = country,
                    UtprTopUpTaxCarryForward = ws.Cell(row,  3).GetString()?.Trim() ?? "", // C
                    Employees               = ws.Cell(row,  5).GetString()?.Trim(),         // E
                    TangibleAssetValue      = ws.Cell(row,  7).GetString()?.Trim(),         // G
                    UtprTopUpTaxAttributed  = ws.Cell(row, 12).GetString()?.Trim() ?? "", // L
                    AddCashTaxExpense       = ws.Cell(row, 14).GetString()?.Trim() ?? "", // N
                    UtprTopUpTaxCarriedForward = ws.Cell(row, 17).GetString()?.Trim() ?? "", // Q
                };

                // J(10): 배분율 (0~1 decimal; 퍼센트 입력 시 /100)
                var pctRaw = ws.Cell(row, 10).GetString()?.Trim();
                if (!string.IsNullOrEmpty(pctRaw) &&
                    decimal.TryParse(pctRaw.TrimEnd('%').Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var pct))
                {
                    attr.UtprPercentage = pct > 1m ? pct / 100m : pct;
                }

                utpr.Attribution.Add(attr);
                hasData = true;
            }

            if (hasData)
                globe.GlobeBody.UtprAttribution.Add(utpr);
        }
    }
}
