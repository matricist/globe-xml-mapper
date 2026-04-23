using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// 특수 케이스 점검:
//   1) 인라인 복합 포맷 셀 인벤토리 (.Split(';'), .Split(',') 등)
//      → qa/inline_formats.md
//   2) 블록 반복 상수 인벤토리
//      → qa/block_constants.md

class Program
{
    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var srcDir = Path.Combine(repo, "Services");
        var qaDir = Path.Combine(repo, "qa");
        Directory.CreateDirectory(qaDir);

        DumpInlineFormats(srcDir, Path.Combine(qaDir, "inline_formats.md"));
        DumpBlockConstants(srcDir, Path.Combine(qaDir, "block_constants.md"));
        Console.WriteLine("Done.");
        return 0;
    }

    // ───────────────────────── 인라인 포맷 ─────────────────────────

    // 시트명 매퍼 파일 매핑 (역)
    static readonly Dictionary<string, string> FileToSheet = new()
    {
        { "Mapping_1.1~1.2.cs",  "다국적기업그룹 정보" },
        { "Mapping_1.3.1.cs",    "최종모기업" },
        { "Mapping_1.3.2.1.cs",  "그룹구조" },
        { "Mapping_1.3.2.2.cs",  "제외기업" },
        { "Mapping_1.3.3.cs",    "그룹구조 변동" },
        { "Mapping_1.4.cs",      "요약" },
        { "Mapping_2.cs",        "적용면제" },
        { "Mapping_JurCal.cs",   "국가별 계산" },
        { "Mapping_EntityCe.cs", "구성기업 계산" },
        { "Mapping_Utpr.cs",     "UTPR 배분" },
    };

    // .Split("x") / .Split('x') / .Split(new[] { 'x' }) / multi-line — 첫 delimiter 추출
    // Singleline 으로 줄바꿈 허용
    static readonly Regex SplitRegex = new(
        @"\.Split\s*\(\s*(?:new(?:\s+\w+)?\s*\[\s*\]\s*\{\s*)?[""']([^""']+)[""']",
        RegexOptions.Compiled | RegexOptions.Singleline);

    // 매핑 헬퍼: ParseTin / ParseNameKName / ParseBool / ParseOwnership 등 복합 파서
    static readonly Regex ParseHelperRegex = new(
        @"\b(Parse(?:Tin|NameKName|Ownership|Bool|Date)|TryParseDate)\s*\(",
        RegexOptions.Compiled);

    static readonly Regex FindRowCallRegex = new(
        @"FindRow\s*\(\s*ws\s*,\s*""([^""]+)""",
        RegexOptions.Compiled);

    static readonly Regex CellColRegex = new(
        @"ws\.Cell\s*\(\s*\w+\s*,\s*(\d+)\s*\)",
        RegexOptions.Compiled);

    static void DumpInlineFormats(string srcDir, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# 인라인 복합 포맷 셀 인벤토리");
        sb.AppendLine();
        sb.AppendLine("코드에서 `.Split(delimiter)` 로 여러 필드를 파싱하는 셀 위치 목록.");
        sb.AppendLine("작성요령 시트의 안내문과 포맷이 일치하는지 수동 대조 필요.");
        sb.AppendLine();
        sb.AppendLine("| 시트 | 파일 | 라인 | 종류 | delimiter | 최근 FindRow 헤더 | 최근 col | 스니펫 |");
        sb.AppendLine("|---|---|---|---|---|---|---|---|");

        int count = 0;
        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs").OrderBy(p => p))
        {
            var lines = File.ReadAllLines(file);
            var text = File.ReadAllText(file);
            var name = Path.GetFileName(file);
            var sheet = FileToSheet.TryGetValue(name, out var sh) ? sh : "?";

            // 라인 단위 스캔 (FindRow/Cell context + ParseHelper)
            string lastHeader = "";
            int lastCol = 0;
            var matchedLines = new HashSet<int>();
            var contextAtLine = new Dictionary<int, (string header, int col)>();

            for (int i = 0; i < lines.Length; i++)
            {
                var hm = FindRowCallRegex.Match(lines[i]);
                if (hm.Success) lastHeader = hm.Groups[1].Value;

                var cm = CellColRegex.Match(lines[i]);
                if (cm.Success) lastCol = int.Parse(cm.Groups[1].Value);

                contextAtLine[i] = (lastHeader, lastCol);

                foreach (Match hm2 in ParseHelperRegex.Matches(lines[i]))
                {
                    EmitRow(sb, sheet, name, i + 1, hm2.Groups[1].Value, "-", lastHeader, lastCol, lines[i]);
                    count++;
                }
            }

            // 파일 단위 multi-line Split 스캔
            foreach (Match sm in SplitRegex.Matches(text))
            {
                int charIdx = sm.Index;
                int lineNo = text.Substring(0, charIdx).Count(c => c == '\n');
                if (!contextAtLine.TryGetValue(lineNo, out var ctx)) ctx = ("", 0);
                var delim = sm.Groups[1].Value;
                var lineText = lineNo < lines.Length ? lines[lineNo] : "";
                EmitRow(sb, sheet, name, lineNo + 1, "Split", $"`{EscapePipe(delim)}`", ctx.header, ctx.col, lineText);
                count++;
            }
        }

        sb.AppendLine();
        sb.AppendLine($"**총 {count}건**");

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Inline formats: {count} rows → {outPath}");
    }

    static void EmitRow(StringBuilder sb, string sheet, string file, int line, string kind, string delim, string header, int col, string rawLine)
    {
        var snippet = rawLine.Trim().Replace("|", "\\|");
        if (snippet.Length > 80) snippet = snippet.Substring(0, 80) + "…";
        sb.Append("| ").Append(sheet)
          .Append(" | ").Append(file)
          .Append(" | ").Append(line)
          .Append(" | ").Append(kind)
          .Append(" | ").Append(delim)
          .Append(" | ").Append(header.Length == 0 ? "-" : header)
          .Append(" | ").Append(col == 0 ? "-" : col.ToString())
          .Append(" | `").Append(snippet).Append("`")
          .AppendLine(" |");
    }

    static string EscapePipe(string s) => s.Replace("|", "\\|");

    // ───────────────────────── 블록 상수 ─────────────────────────

    static readonly Regex ConstRegex = new(
        @"private\s+(?:const|static\s+readonly)\s+\w+\s+(\w+)\s*=\s*([^;]+?);",
        RegexOptions.Compiled);

    static readonly Regex FieldMapHeaderRegex = new(
        @"private\s+static\s+readonly\s+\(int\s+Offset,\s+string\s+Target\)\[\]\s+FieldMap",
        RegexOptions.Compiled);

    static void DumpBlockConstants(string srcDir, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# 블록 반복 상수 인벤토리");
        sb.AppendLine();
        sb.AppendLine("각 매퍼의 블록 크기/시작/간격 상수. blockCount≥2 시나리오에서 N번째 블록의 행 좌표를 확인하는 기준.");
        sb.AppendLine();

        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs").OrderBy(p => p))
        {
            var name = Path.GetFileName(file);
            var sheet = FileToSheet.TryGetValue(name, out var sh) ? sh : "?";
            var text = File.ReadAllText(file);

            var consts = ConstRegex.Matches(text)
                .Select(m => (Name: m.Groups[1].Value, Value: m.Groups[2].Value.Trim()))
                // FieldMap 등 배열 리터럴은 너무 길어 제외
                .Where(t => t.Value.Length < 80
                            && !t.Value.StartsWith("new")
                            && !t.Value.Contains("["))
                .ToList();

            if (consts.Count == 0) continue;

            sb.AppendLine($"## {sheet} · `{name}`");
            sb.AppendLine();
            sb.AppendLine("| 상수 | 값 |");
            sb.AppendLine("|---|---|");
            foreach (var (n, v) in consts)
                sb.AppendLine($"| `{n}` | `{v}` |");
            sb.AppendLine();
        }

        sb.AppendLine("---");
        sb.AppendLine();
        sb.AppendLine("## JurCal / EntityCe — 동적 블록");
        sb.AppendLine();
        sb.AppendLine("이 두 매퍼는 상수 기반이 아닌 **헤더 텍스트**로 블록 경계 탐지:");
        sb.AppendLine();
        sb.AppendLine("- `Mapping_JurCal`: `\"3.1 국가별\"` 헤더 등장 행마다 새 블록 시작, `_blockStart`/`_blockEnd`를 다음 헤더 직전까지로 스코핑");
        sb.AppendLine("- `Mapping_EntityCe`: `\"1. 구성기업 또는 공동기업그룹 기업의 납세자번호\"` 헤더 등장 행마다 새 블록");
        sb.AppendLine();
        sb.AppendLine("⇒ blockCount 기반이 아니므로 _META 의 blockCount 와 무관하게 동작. 단, `Mapping_2`/`Mapping_1.3.x`/`Mapping_1.4`/`Mapping_Utpr`/`Mapping_1.3.3` 는 _META 의 blockCount 필요.");

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Block constants → {outPath}");
    }
}
