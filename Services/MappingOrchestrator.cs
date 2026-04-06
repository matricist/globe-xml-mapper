using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class MappingOrchestrator
    {
        // 루트에 배치되는 매퍼 (기본정보)
        private readonly Mapping_1_1_1_2 _filingInfoMapper = new();

        // 하위 디렉토리별 매퍼 (디렉토리명 → 매퍼)
        private static readonly (string DirName, Func<MappingBase> CreateMapper)[] SubDirMappers =
        {
            ("1.3.1",   () => new Mapping_1_3_1()),
            ("1.3.2.1", () => new Mapping_1_3_2_1()),
            ("1.3.2.2", () => new Mapping_1_3_2_2()),
        };

        public List<string> MapFolder(string rootPath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            // 1. 루트: 기본정보 (1.1~1.2)
            var rootFiles = GetXlsxFiles(rootPath);
            if (rootFiles.Count == 0)
                errors.Add("루트 폴더에 기본정보(1.1~1.2) xlsx 파일이 없습니다.");

            foreach (var filePath in rootFiles)
            {
                ProcessFile(filePath, _filingInfoMapper, globe, errors);
            }

            // 2. 하위 디렉토리별 처리
            foreach (var (dirName, createMapper) in SubDirMappers)
            {
                var subDir = Path.Combine(rootPath, dirName);
                if (!Directory.Exists(subDir))
                {
                    errors.Add($"필수 디렉토리 '{dirName}'이(가) 없습니다.");
                    continue;
                }

                var files = GetXlsxFiles(subDir);
                if (files.Count == 0)
                {
                    errors.Add($"'{dirName}' 디렉토리에 xlsx 파일이 없습니다.");
                    continue;
                }

                foreach (var filePath in files)
                {
                    var mapper = createMapper();
                    ProcessFile(filePath, mapper, globe, errors);
                }
            }

            FillMessageSpec(globe);
            return errors;
        }

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
                errors.Add($"[{fileName}] 시트 '{mapper.SheetName}'을(를) 찾을 수 없습니다.");
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
        }
    }
}
