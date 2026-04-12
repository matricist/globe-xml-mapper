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

        /// <summary>
        /// 프로젝트 폴더 기반 매핑.
        ///   루트 디렉터리: main 파일 1개 (fileType=main 또는 xlsx 유일 파일)
        ///   하위 디렉터리: group / entity 파일 (fileType으로 판별, 재귀)
        /// </summary>
        public List<string> MapFolder(string rootPath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            // ── 루트의 main 파일 ───────────────────────────────────────────
            var rootFiles = Directory.GetFiles(rootPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => f)
                .ToList();

            if (rootFiles.Count == 0)
            {
                errors.Add($"루트 폴더에 xlsx 파일이 없습니다. ({rootPath})");
                return errors;
            }

            // main 파일은 루트에 1개여야 함
            var mainFile = rootFiles.Count == 1
                ? rootFiles[0]
                : rootFiles.FirstOrDefault(f => ReadFileTypeFromXlsx(f, errors) == "main");

            if (mainFile == null)
            {
                errors.Add($"루트 폴더에서 main 파일(fileType=main)을 찾을 수 없습니다.");
                return errors;
            }
            if (rootFiles.Count > 1)
                foreach (var f in rootFiles.Where(f => f != mainFile))
                    errors.Add($"루트에 main 외 xlsx 파일 있음 (건너뜀): {Path.GetFileName(f)}");

            errors.AddRange(MapWorkbook(mainFile, globe));

            // ── 하위 디렉터리별 처리: 같은 폴더의 group → entity 순서 보장 ──
            var subDirs = Directory.GetDirectories(rootPath, "*", SearchOption.AllDirectories)
                .OrderBy(d => d)
                .ToList();

            foreach (var dir in subDirs)
            {
                var filesInDir = GetXlsxFiles(dir);
                if (filesInDir.Count == 0) continue;

                var typedFiles = filesInDir
                    .Select(f => (path: f, type: ReadFileTypeFromXlsx(f, errors)))
                    .ToList();

                // group 먼저 처리 → 해당 폴더의 JurisdictionSection 생성
                string groupFile = null;
                foreach (var (path, _) in typedFiles.Where(x => x.type == "group"))
                {
                    MapFileBySheets(path, "group", globe, errors);
                    groupFile = path;
                }

                // entity 파일: 같은 폴더 group 파일의 JurisdictionSection + ETR(SubGroup) 기준
                var (jurCode, subGroupTin) = groupFile != null
                    ? ReadGroupJurisdiction(groupFile, errors)
                    : ((Globe.CountryCodeType?)null, (string)null);
                foreach (var (path, _) in typedFiles.Where(x => x.type == "entity"))
                    MapEntityFile(path, jurCode, subGroupTin, globe, errors);

                foreach (var (path, type) in typedFiles.Where(x => x.type != "group" && x.type != "entity"))
                    errors.Add($"[{Path.GetFileName(path)}] fileType='{type}' — 건너뜀 (group 또는 entity여야 함)");
            }

            FillMessageSpec(globe);
            return errors;
        }

        /// <summary>
        /// xlsx에서 _META.fileType 읽기 (ClosedXML). 읽기 실패 시 "unknown" 반환.
        /// </summary>
        private static string ReadFileTypeFromXlsx(string filePath, List<string> errors)
        {
            try
            {
                using var wb = new XLWorkbook(filePath);
                if (!wb.TryGetWorksheet(ExcelController.MetaSheetName, out var meta))
                    return "unknown";
                return ExcelController.ReadFileType(meta);
            }
            catch (Exception ex)
            {
                errors.Add($"[{Path.GetFileName(filePath)}] _META 읽기 오류: {ex.Message}");
                return "unknown";
            }
        }

        private static int NaturalOrder(string name)
        {
            var digits = System.Text.RegularExpressions.Regex.Match(name, @"\d+$");
            return digits.Success ? int.Parse(digits.Value) : 0;
        }

        /// <summary>
        /// group / entity 파일 매핑.
        /// fileType으로 코드에 정의된 FileTypeSheetMap을 사용.
        /// </summary>
        private void MapFileBySheets(string filePath, string fileType, Globe.GlobeOecd globe, List<string> errors)
        {
            var fileName = Path.GetFileName(filePath);
            if (!ExcelController.FileTypeSheetMap.TryGetValue(fileType, out var sheetMappings))
            {
                errors.Add($"[{fileName}] fileType='{fileType}'에 대한 시트 매핑 없음");
                return;
            }
            try
            {
                using var workbook = new XLWorkbook(filePath);
                foreach (var (section, sheetName) in sheetMappings)
                {
                    if (!MapperFactory.TryGetValue(section, out var createMapper)) continue;
                    if (!workbook.TryGetWorksheet(sheetName, out var ws)) continue;
                    createMapper().Map(ws, globe, errors, fileName);
                }
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 파일 읽기 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// group 파일("국가별 계산" 시트)에서 국가코드 + SubGroup TIN을 읽어 반환.
        /// </summary>
        private static (Globe.CountryCodeType? jurCode, string subGroupTin) ReadGroupJurisdiction(string filePath, List<string> errors)
        {
            try
            {
                using var wb = new XLWorkbook(filePath);
                if (!wb.TryGetWorksheet("국가별 계산", out var ws)) return (null, null);
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
                for (int r = 1; r <= lastRow; r++)
                {
                    var cell = ws.Cell(r, 2).GetString();
                    if (cell != null && cell.Contains("3.1 국가별"))
                    {
                        var jurRaw = ws.Cell(r + 1, 15).GetString()?.Trim(); // +1: 국가코드
                        var subGroupTin = ws.Cell(r + 3, 15).GetString()?.Trim(); // +3: SubGroup TIN
                        Globe.CountryCodeType? code = null;
                        if (!string.IsNullOrEmpty(jurRaw)
                            && System.Enum.TryParse<Globe.CountryCodeType>(jurRaw, true, out var parsed))
                            code = parsed;
                        return (code, string.IsNullOrEmpty(subGroupTin) ? null : subGroupTin);
                    }
                }
                return (null, null);
            }
            catch { return (null, null); }
        }

        /// <summary>
        /// entity 파일을 처리. 같은 디렉토리 group 파일의 JurisdictionSection + ETR(SubGroup)에 CEComputation 추가.
        /// </summary>
        private void MapEntityFile(string filePath, Globe.CountryCodeType? jurCode, string subGroupTin, Globe.GlobeOecd globe, List<string> errors)
        {
            var fileName = Path.GetFileName(filePath);
            if (!ExcelController.FileTypeSheetMap.TryGetValue("entity", out var sheetMappings)) return;
            try
            {
                using var workbook = new XLWorkbook(filePath);
                foreach (var (section, sheetName) in sheetMappings)
                {
                    if (!MapperFactory.TryGetValue(section, out var createMapper)) continue;
                    if (!workbook.TryGetWorksheet(sheetName, out var ws)) continue;
                    var mapper = createMapper();
                    if (mapper is Mapping_EntityCe entityMapper)
                        entityMapper.MapWithJur(ws, globe, errors, fileName, jurCode, subGroupTin);
                    else
                        mapper.Map(ws, globe, errors, fileName);
                }
            }
            catch (Exception ex) { errors.Add($"[{fileName}] 파일 읽기 오류: {ex.Message}"); }
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
