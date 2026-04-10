using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        /// <summary>
        /// 프로젝트 폴더 기반 매핑.
        /// 구조:
        ///   root/MNE*.xlsx              → 1.x, 2 섹션 (최종모기업 파일)
        ///   root/합산단위_N/합산단위_N.xlsx → 합산단위별 매핑
        ///   root/구성기업_N.xlsx          → 구성기업별 매핑
        /// </summary>
        public List<string> MapFolder(string rootPath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            // ── MNE 파일 (루트의 xlsx 중 임시파일 제외, 첫 번째) ─────────
            var mneFiles = Directory.GetFiles(rootPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => f)
                .ToList();

            if (mneFiles.Count == 0)
            {
                errors.Add($"루트 폴더에 xlsx 파일이 없습니다. ({rootPath})");
                return errors;
            }
            if (mneFiles.Count > 1)
                errors.Add($"루트 폴더에 xlsx 파일이 여러 개입니다. 첫 번째 파일만 사용합니다: {Path.GetFileName(mneFiles[0])}");

            errors.AddRange(MapWorkbook(mneFiles[0], globe));

            // ── 합산단위_N 하위 디렉터리 ──────────────────────────────────
            var groupDirs = Directory.GetDirectories(rootPath, "합산단위_*")
                .OrderBy(d => NaturalOrder(Path.GetFileName(d)))
                .ToList();

            foreach (var dir in groupDirs)
            {
                var dirName = Path.GetFileName(dir);
                var xlsxFiles = Directory.GetFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                    .OrderBy(f => f)
                    .ToList();
                if (xlsxFiles.Count == 0)
                { errors.Add($"[{dirName}] xlsx 파일 없음"); continue; }
                foreach (var f in xlsxFiles)
                    MapFileBySheets(f, globe, errors);
            }

            // ── 구성기업_N.xlsx ───────────────────────────────────────────
            var ceFiles = Directory.GetFiles(rootPath, "구성기업_*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => NaturalOrder(Path.GetFileName(f)))
                .ToList();

            foreach (var f in ceFiles)
                MapFileBySheets(f, globe, errors);

            FillMessageSpec(globe);
            return errors;
        }

        private static int NaturalOrder(string name)
        {
            var digits = System.Text.RegularExpressions.Regex.Match(name, @"\d+$");
            return digits.Success ? int.Parse(digits.Value) : 0;
        }

        /// <summary>
        /// 파일 내 시트 이름을 MapperFactory에 직접 조회하여 매핑.
        /// _META 없이 동작 (Group.xlsx, CE_N.xlsx용).
        /// </summary>
        private void MapFileBySheets(string filePath, Globe.GlobeOecd globe, List<string> errors)
        {
            var fileName = Path.GetFileName(filePath);
            try
            {
                using var workbook = new XLWorkbook(filePath);
                foreach (var ws in workbook.Worksheets)
                {
                    if (!MapperFactory.TryGetValue(ws.Name, out var createMapper))
                        continue; // 아직 매퍼 없는 시트는 스킵
                    createMapper().Map(ws, globe, errors, fileName);
                }
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 파일 읽기 오류: {ex.Message}");
            }
        }

        #region 내부 유틸



        private static void ProcessFile(string filePath, MappingBase mapper, Globe.GlobeOecd globe, List<string> errors)
        {
            var fileName = Path.GetFileName(filePath);
            try
            {
                using var workbook = new XLWorkbook(filePath);
                foreach (var ws in workbook.Worksheets)
                {
                    if (ws.Name == mapper.SheetName)
                    {
                        mapper.Map(ws, globe, errors, fileName);
                        return;
                    }
                }
                errors.Add($"[{fileName}] 시트 '{mapper.SheetName}' 없음");
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 파일 읽기 오류: {ex.Message}");
            }
        }

        private static List<string> GetXlsxFiles(string dirPath)
        {
            return Directory.GetFiles(dirPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => f)
                .ToList();
        }

        #endregion

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
