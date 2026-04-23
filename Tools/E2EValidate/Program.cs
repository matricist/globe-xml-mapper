using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using GlobeMapper.Services;

// E2E 검증:
//   xlsx → XML 변환 (MappingOrchestrator)
//   + 매핑 오류 + ValidationUtil 오류 + XSD 검증 오류
//   → qa/e2e_report.md

class Program
{
    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        // args[1]: 입력 xlsx 경로 (생략 시 main_template.xlsx)
        // args[2]: 출력 리포트 파일명 (생략 시 e2e_report.md)
        var xlsxPath = args.Length > 1
            ? Path.GetFullPath(args[1])
            : Path.Combine(repo, "Resources", "main_template.xlsx");
        var reportName = args.Length > 2 ? args[2] : "e2e_report.md";
        var xmlOutName = Path.GetFileNameWithoutExtension(reportName).Replace("_report", "_output") + ".xml";

        var xsdPath  = Path.Combine(repo, "Resources", "XSD", "GLOBEXML_v1.0_KR.xsd");
        var qaDir    = Path.Combine(repo, "qa");
        Directory.CreateDirectory(qaDir);

        if (!File.Exists(xlsxPath)) { Console.Error.WriteLine($"Not found: {xlsxPath}"); return 1; }
        if (!File.Exists(xsdPath))  { Console.Error.WriteLine($"Not found: {xsdPath}"); return 1; }

        var sb = new StringBuilder();
        sb.AppendLine("# E2E 검증 리포트");
        sb.AppendLine();
        sb.AppendLine($"- 입력 파일: `{Path.GetRelativePath(repo, xlsxPath).Replace('\\', '/')}`");
        sb.AppendLine($"- XSD: `Resources/XSD/GLOBEXML_v1.0_KR.xsd`");
        sb.AppendLine($"- 실행 시각: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        // 1) Map
        var globe = new Globe.GlobeOecd
        {
            Version     = "2.0",
            MessageSpec = new Globe.MessageSpecType(),
            GlobeBody   = new Globe.GlobeBodyType(),
        };

        List<string> mappingErrors;
        try
        {
            var orchestrator = new MappingOrchestrator();
            mappingErrors = orchestrator.MapWorkbook(xlsxPath, globe);
        }
        catch (Exception ex)
        {
            sb.AppendLine("## ❌ 매핑 실행 실패");
            sb.AppendLine();
            sb.AppendLine("```");
            sb.AppendLine(ex.ToString());
            sb.AppendLine("```");
            File.WriteAllText(Path.Combine(qaDir, "e2e_report.md"), sb.ToString(), new UTF8Encoding(true));
            Console.Error.WriteLine("매핑 실패: " + ex.Message);
            return 2;
        }

        sb.AppendLine("## 1. 매핑 오류");
        sb.AppendLine();
        sb.AppendLine($"총 **{mappingErrors.Count}** 건");
        if (mappingErrors.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("```");
            foreach (var e in mappingErrors.Take(50)) sb.AppendLine(e);
            if (mappingErrors.Count > 50) sb.AppendLine($"... (+{mappingErrors.Count - 50} more)");
            sb.AppendLine("```");
        }
        sb.AppendLine();

        // 2) ValidationUtil
        var validationErrors = ValidationUtil.Validate(globe);
        sb.AppendLine("## 2. ValidationUtil 오류");
        sb.AppendLine();
        sb.AppendLine($"총 **{validationErrors.Count}** 건");
        if (validationErrors.Count > 0)
        {
            // 에러코드별 그루핑
            var byCode = validationErrors
                .Select(e => new { Raw = e, Code = ExtractCode(e) })
                .GroupBy(e => e.Code)
                .OrderBy(g => g.Key)
                .ToList();
            sb.AppendLine();
            sb.AppendLine("### 코드별 합계");
            sb.AppendLine();
            sb.AppendLine("| 코드 | 건수 | 샘플 |");
            sb.AppendLine("|---|---|---|");
            foreach (var g in byCode)
                sb.AppendLine($"| `{g.Key}` | {g.Count()} | {Truncate(g.First().Raw)} |");
            sb.AppendLine();
            sb.AppendLine("<details><summary>전체 목록 (처음 100건)</summary>");
            sb.AppendLine();
            sb.AppendLine("```");
            foreach (var e in validationErrors.Take(100)) sb.AppendLine(e);
            if (validationErrors.Count > 100) sb.AppendLine($"... (+{validationErrors.Count - 100} more)");
            sb.AppendLine("```");
            sb.AppendLine("</details>");
        }
        sb.AppendLine();

        // 3) Serialize XML
        string xml;
        try
        {
            xml = XmlExportService.Serialize(globe);
        }
        catch (Exception ex)
        {
            sb.AppendLine("## ❌ XML 직렬화 실패");
            sb.AppendLine();
            sb.AppendLine("```");
            sb.AppendLine(ex.ToString());
            sb.AppendLine("```");
            File.WriteAllText(Path.Combine(qaDir, "e2e_report.md"), sb.ToString(), new UTF8Encoding(true));
            return 3;
        }

        var outXmlPath = Path.Combine(qaDir, xmlOutName);
        File.WriteAllText(outXmlPath, xml, new UTF8Encoding(true));
        sb.AppendLine("## 3. XML 직렬화");
        sb.AppendLine();
        sb.AppendLine($"- 저장: [qa/{xmlOutName}]({xmlOutName})");
        sb.AppendLine($"- 크기: {new FileInfo(outXmlPath).Length:N0} bytes");
        sb.AppendLine();

        // 4) XSD 검증
        var xsdErrors = new List<string>();
        try
        {
            var settings = new XmlReaderSettings
            {
                ValidationType = ValidationType.Schema,
                ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings
                                 | XmlSchemaValidationFlags.ProcessInlineSchema
                                 | XmlSchemaValidationFlags.ProcessSchemaLocation,
            };
            // XSD 로드 — GLOBEXML 메인 + iso + oecd + stf 참조
            settings.Schemas.XmlResolver = new XmlUrlResolver();
            settings.Schemas.Add(null, xsdPath);
            settings.ValidationEventHandler += (s, e) =>
                xsdErrors.Add($"[{e.Severity}] L{e.Exception?.LineNumber ?? 0}: {e.Message}");

            using var sr = new StringReader(xml);
            using var xr = XmlReader.Create(sr, settings);
            while (xr.Read()) { }
        }
        catch (Exception ex)
        {
            xsdErrors.Add($"[Fatal] {ex.Message}");
        }

        sb.AppendLine("## 4. XSD 검증 오류");
        sb.AppendLine();
        sb.AppendLine($"총 **{xsdErrors.Count}** 건");
        if (xsdErrors.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("```");
            foreach (var e in xsdErrors.Take(80)) sb.AppendLine(e);
            if (xsdErrors.Count > 80) sb.AppendLine($"... (+{xsdErrors.Count - 80} more)");
            sb.AppendLine("```");
        }
        sb.AppendLine();

        // 5) 요약
        int total = mappingErrors.Count + validationErrors.Count + xsdErrors.Count;
        sb.AppendLine("## 종합");
        sb.AppendLine();
        sb.AppendLine($"- 매핑 오류: {mappingErrors.Count}");
        sb.AppendLine($"- Validation 오류: {validationErrors.Count}");
        sb.AppendLine($"- XSD 검증 오류: {xsdErrors.Count}");
        sb.AppendLine($"- **총 {total}건**");
        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine();
        sb.AppendLine("## 해석 가이드");
        sb.AppendLine();
        sb.AppendLine("- **매핑 오류**: 서식 셀에 필수 데이터가 비어 있음 (사용자 입력 필요)");
        sb.AppendLine("- **Validation 오류**: 코드의 70xxx 비즈니스 룰 위반 (ValidationUtil 구현 범위)");
        sb.AppendLine("- **XSD 검증 오류**: 생성된 XML 구조가 스키마 위반 — 종류별:");
        sb.AppendLine("  - `has invalid child element 'X'. List of possible elements expected: 'Y'` → Y([R])가 비어 있어 emit되지 않음 → **서식 샘플 미완성**");
        sb.AppendLine("  - `is invalid. The value 'Z' is not valid` → 값 타입/포맷 위반 → **코드 또는 서식 포맷 버그**");
        sb.AppendLine("  - `Schema location` 관련 → XSD 참조 설정 이슈 (`XmlUrlResolver` 사용 중)");

        var outReport = Path.Combine(qaDir, reportName);
        File.WriteAllText(outReport, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"E2E 리포트 → {outReport}");
        Console.WriteLine($"  매핑={mappingErrors.Count}  Validation={validationErrors.Count}  XSD={xsdErrors.Count}");
        return 0;
    }

    static string ExtractCode(string msg)
    {
        // [70001] ... → "70001"
        if (msg.StartsWith("[") && msg.Length > 7)
        {
            var end = msg.IndexOf(']');
            if (end > 0) return msg.Substring(1, end - 1);
        }
        return "(기타)";
    }

    static string Truncate(string s)
    {
        s = s.Replace("\r", " ").Replace("\n", " ").Replace("|", "/");
        return s.Length > 80 ? s.Substring(0, 80) + "…" : s;
    }
}
