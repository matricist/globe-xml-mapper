using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

// 매핑 인벤토리 추출:
//   1) Resources/mappings/*.json → qa/code_mappings_json.csv
//      columns: section_file, sheet, section, cell, target, type, label, multi
//   2) Services/Mapping_{JurCal,EntityCe,Utpr}.cs 의 FindRow("...") 호출 → qa/code_mappings_findrow.csv
//      columns: file, line, header_text, context_method
//   3) 같은 C# 파일들의 ws.Cell(_,_) 호출 → qa/code_mappings_cells.csv (참조 컬럼 추출용)
//      columns: file, line, row_expr, col, snippet

class Program
{
    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var jsonDir = Path.Combine(repo, "Resources", "mappings");
        var srcDir  = Path.Combine(repo, "Services");
        var qaDir   = Path.Combine(repo, "qa");
        Directory.CreateDirectory(qaDir);

        DumpJsonMappings(jsonDir, Path.Combine(qaDir, "code_mappings_json.csv"));
        DumpFindRowCalls(srcDir,  Path.Combine(qaDir, "code_mappings_findrow.csv"));
        DumpCellCalls(srcDir,     Path.Combine(qaDir, "code_mappings_cells.csv"));

        Console.WriteLine("Done.");
        return 0;
    }

    // ───────────────────────── JSON 매퍼 평탄화 ─────────────────────────

    static void DumpJsonMappings(string dir, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("file,sheet,section,cell,target,type,label,multi");

        int total = 0;
        foreach (var path in Directory.EnumerateFiles(dir, "mapping_*.json").OrderBy(p => p))
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(path));
            var root = doc.RootElement;
            var sheet = root.TryGetProperty("sheetName", out var s) ? s.GetString() ?? "" : "";

            if (!root.TryGetProperty("sections", out var sections)) continue;

            foreach (var sec in sections.EnumerateObject())
            {
                var sectionKey = sec.Name;
                if (!sec.Value.TryGetProperty("mappings", out var maps)) continue;

                foreach (var m in maps.EnumerateArray())
                {
                    var cell   = Get(m, "cell");
                    var target = Get(m, "target");
                    var type   = Get(m, "type");
                    var label  = Get(m, "label");
                    var multi  = m.TryGetProperty("multi", out var mu) && mu.ValueKind == JsonValueKind.True ? "true" : "";

                    sb.Append(Csv(Path.GetFileName(path))).Append(',')
                      .Append(Csv(sheet)).Append(',')
                      .Append(Csv(sectionKey)).Append(',')
                      .Append(Csv(cell)).Append(',')
                      .Append(Csv(target)).Append(',')
                      .Append(Csv(type)).Append(',')
                      .Append(Csv(label)).Append(',')
                      .Append(multi).AppendLine();
                    total++;
                }
            }
        }

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"JSON mappings: {total} rows → {outPath}");
    }

    static string Get(JsonElement obj, string key)
        => obj.TryGetProperty(key, out var v) ? (v.GetString() ?? "") : "";

    // ───────────────────────── FindRow 호출 인벤토리 ─────────────────────────

    static readonly Regex FindRowRegex = new(@"FindRow\s*\(\s*ws\s*,\s*""([^""]+)""", RegexOptions.Compiled);
    static readonly Regex MethodRegex  = new(@"^\s*(public|private|protected|internal)[^\(]*\b(\w+)\s*\(", RegexOptions.Compiled);

    static void DumpFindRowCalls(string srcDir, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("file,line,header,method");
        int total = 0;

        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs"))
        {
            var lines = File.ReadAllLines(file);
            string currentMethod = "";
            for (int i = 0; i < lines.Length; i++)
            {
                var mm = MethodRegex.Match(lines[i]);
                if (mm.Success) currentMethod = mm.Groups[2].Value;

                foreach (Match m in FindRowRegex.Matches(lines[i]))
                {
                    sb.Append(Csv(Path.GetFileName(file))).Append(',')
                      .Append(i + 1).Append(',')
                      .Append(Csv(m.Groups[1].Value)).Append(',')
                      .Append(Csv(currentMethod)).AppendLine();
                    total++;
                }
            }
        }

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"FindRow calls: {total} rows → {outPath}");
    }

    // ───────────────────────── ws.Cell(_,_) 호출 인벤토리 ─────────────────────────

    // ws.Cell(rowExpr, colNumber) 의 col이 정수인 경우만 추출 (변수일 땐 정확도가 떨어져 별도 처리)
    static readonly Regex CellRegex = new(
        @"ws\.Cell\s*\(\s*([^,)]+?)\s*,\s*(\d+)\s*\)",
        RegexOptions.Compiled);

    static void DumpCellCalls(string srcDir, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("file,line,row_expr,col,snippet");
        int total = 0;

        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs"))
        {
            var lines = File.ReadAllLines(file);
            for (int i = 0; i < lines.Length; i++)
            {
                foreach (Match m in CellRegex.Matches(lines[i]))
                {
                    var rowExpr = m.Groups[1].Value.Trim();
                    var col     = m.Groups[2].Value;
                    var snippet = lines[i].Trim();
                    if (snippet.Length > 200) snippet = snippet.Substring(0, 200) + "...";

                    sb.Append(Csv(Path.GetFileName(file))).Append(',')
                      .Append(i + 1).Append(',')
                      .Append(Csv(rowExpr)).Append(',')
                      .Append(col).Append(',')
                      .Append(Csv(snippet)).AppendLine();
                    total++;
                }
            }
        }

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Cell calls: {total} rows → {outPath}");
    }

    // ───────────────────────── CSV 헬퍼 ─────────────────────────

    static string Csv(string? s)
    {
        if (s == null) return "";
        if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        return s;
    }
}
