using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// 갭 분석:
//   순방향 (서식 → 코드): qa/gaps_forward.md
//     value/candidate 셀 중 매핑 인벤토리에 없는 셀
//   역방향 (XSD → 코드): qa/gaps_reverse.md
//     Globe요약.md의 [R] 속성 중 코드 어디에도 등장하지 않는 것

class Program
{
    // 시트명 → 시트를 처리하는 C# 매퍼 파일명
    static readonly Dictionary<string, string> SheetToFile = new()
    {
        { "다국적기업그룹 정보", "Mapping_1.1~1.2.cs" },
        { "최종모기업",          "Mapping_1.3.1.cs" },
        { "그룹구조",            "Mapping_1.3.2.1.cs" },
        { "제외기업",            "Mapping_1.3.2.2.cs" },
        { "그룹구조 변동",       "Mapping_1.3.3.cs" },
        { "요약",                "Mapping_1.4.cs" },
        { "적용면제",            "Mapping_2.cs" },
        { "국가별 계산",         "Mapping_JurCal.cs" },
        { "구성기업 계산",       "Mapping_EntityCe.cs" },
        { "UTPR 배분",           "Mapping_Utpr.cs" },
    };

    // 작성요령/메타: 매핑 대상 아님
    static readonly HashSet<string> ExcludeSheets = new()
    {
        "Main 작성요령", "Entity 작성요령", "Group 작성요령", "_META", "기업매핑"
    };

    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var qaDir = Path.Combine(repo, "qa");
        ForwardGap(repo, qaDir);
        ReverseGap(repo, qaDir);
        Console.WriteLine("Done.");
        return 0;
    }

    // ───────────────────────── 순방향 ─────────────────────────

    static void ForwardGap(string repo, string qaDir)
    {
        var cells = ReadCsv(Path.Combine(qaDir, "template_cells.csv"));
        var jsonMaps = ReadCsv(Path.Combine(qaDir, "code_mappings_json.csv"));
        var codeCells = ReadCsv(Path.Combine(qaDir, "code_mappings_cells.csv"));

        // 1) JSON 시트별 매핑된 (sheet, cell) 셋
        var jsonCovered = jsonMaps
            .Select(r => (sheet: r["sheet"], cell: r["cell"]))
            .ToHashSet();

        // 2) C# 매퍼 파일별로 호출된 col 셋 (동적 row라 cell-단위 매칭은 불가)
        var fileCols = codeCells
            .GroupBy(r => r["file"])
            .ToDictionary(g => g.Key, g => g.Select(r => int.Parse(r["col"])).ToHashSet());

        // 시트별 그루핑
        var sb = new StringBuilder();
        sb.AppendLine("# 순방향 갭 리포트 (서식 → 코드)");
        sb.AppendLine();
        sb.AppendLine("`value` 또는 `candidate` 셀 중 코드 매핑에 미연결로 추정되는 셀 목록.");
        sb.AppendLine();
        sb.AppendLine("- **Tier A** (JSON 시트): `(sheet, cell)` 정확 매칭 — 미매칭 = 확실한 갭");
        sb.AppendLine("- **Tier B** (C# 시트, 블록 반복): 컬럼 단위 매칭만 가능 — 같은 컬럼이 매퍼에서 한 번도 안 읽히면 갭, 그 외는 수동 확인 필요");
        sb.AppendLine();

        var sheets = cells.Select(r => r["sheet"]).Distinct().Where(s => !ExcludeSheets.Contains(s));

        foreach (var sheet in sheets.OrderBy(s => s))
        {
            var sheetCells = cells.Where(r => r["sheet"] == sheet
                                              && (r["kind"] == "value" || r["kind"] == "candidate"))
                                  .ToList();
            if (sheetCells.Count == 0) continue;

            // 시트가 JSON 매퍼인지 확인 (JSON 매핑이 1개라도 있으면 JSON 처리 시트)
            bool hasJson = jsonMaps.Any(r => r["sheet"] == sheet);
            string fileForSheet = SheetToFile.TryGetValue(sheet, out var f) ? f : null;
            var colsInFile = fileForSheet != null && fileCols.TryGetValue(fileForSheet, out var c) ? c : new HashSet<int>();

            sb.AppendLine($"## {sheet}");
            sb.AppendLine();
            sb.AppendLine($"- 매퍼: `{fileForSheet ?? "(미정)"}`");
            sb.AppendLine($"- 처리 방식: {(hasJson ? "JSON" : "C# (블록 반복)")}");
            sb.AppendLine($"- 입력 셀 후보: {sheetCells.Count}개");
            sb.AppendLine();
            sb.AppendLine("| address | col | kind | 매핑상태 | text(요약) |");
            sb.AppendLine("|---|---|---|---|---|");

            foreach (var c2 in sheetCells.OrderBy(r => int.Parse(r["row"])).ThenBy(r => int.Parse(r["col"])))
            {
                var addr = c2["address"];
                var col = int.Parse(c2["col"]);
                var kind = c2["kind"];
                var text = TruncMd(c2["text"]);

                string status;
                if (hasJson && jsonCovered.Contains((sheet, addr)))
                    status = "✓ JSON";
                else if (hasJson)
                    status = "✗ JSON 미매핑";
                else if (colsInFile.Contains(col))
                    status = "△ col 사용 (수동)";
                else
                    status = "✗ col 미사용";

                sb.AppendLine($"| {addr} | {col} | {kind} | {status} | {text} |");
            }
            sb.AppendLine();
        }

        var outPath = Path.Combine(qaDir, "gaps_forward.md");
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Forward gap → {outPath}");
    }

    // ───────────────────────── 역방향 ─────────────────────────

    // Globe요약.md 의 [R] 마커가 붙은 라인을 추출:
    //   "  - PropName : Type [R] ..."
    static readonly Regex RequiredFieldRegex = new(
        @"^\s*-\s+(\w+)\s*:\s*([^\[]+?)\s*\[R\]",
        RegexOptions.Compiled);

    // 타입 헤더:  "TypeName [DefName]" (들여쓰기 없는 행)
    static readonly Regex TypeHeaderRegex = new(
        @"^([A-Z]\w+(?:Type)?)\s*\[",
        RegexOptions.Compiled);

    static void ReverseGap(string repo, string qaDir)
    {
        var sumPath = Path.Combine(repo, "Resources", "Globe요약.md");
        if (!File.Exists(sumPath))
        {
            Console.Error.WriteLine($"Not found: {sumPath}");
            return;
        }

        // [R] 필드 추출 (현재 타입 컨텍스트 함께)
        var required = new List<(string type, string prop, string typeRef)>();
        string currentType = "";
        foreach (var line in File.ReadAllLines(sumPath))
        {
            var th = TypeHeaderRegex.Match(line);
            if (th.Success && !line.StartsWith("  ") && !line.StartsWith("\t"))
                currentType = th.Groups[1].Value;

            var rm = RequiredFieldRegex.Match(line);
            if (rm.Success)
                required.Add((currentType, rm.Groups[1].Value, rm.Groups[2].Value.Trim()));
        }

        // 코드 + JSON 매퍼에서 등장하는 식별자 모두 모아 단순 텍스트 검색
        var corpus = new StringBuilder();
        foreach (var f in Directory.EnumerateFiles(Path.Combine(repo, "Services"), "*.cs"))
            corpus.AppendLine(File.ReadAllText(f));
        foreach (var f in Directory.EnumerateFiles(Path.Combine(repo, "Resources", "mappings"), "*.json"))
            corpus.AppendLine(File.ReadAllText(f));
        var corpusText = corpus.ToString();

        // 각 [R] 속성에 대해 ".PropName" 패턴이 코드/JSON에 등장하는지 확인
        var sb = new StringBuilder();
        sb.AppendLine("# 역방향 갭 리포트 (XSD [R] → 코드)");
        sb.AppendLine();
        sb.AppendLine("`Globe요약.md`에서 `[R]`로 표시된 필수 속성이 코드(JSON 매퍼 + Services/*.cs) 어디에 등장하는지 확인.");
        sb.AppendLine();
        sb.AppendLine("- ✓: `.PropName` 패턴이 코드 또는 JSON target 어딘가에 존재");
        sb.AppendLine("- ✗: 코드/JSON 어디에도 등장 없음 → 매핑 누락 의심 (또는 부모가 통째로 미사용)");
        sb.AppendLine();
        sb.AppendLine("> 한계: 동일 이름이 다른 타입에서 쓰이면 false-positive 가능. 부모 컨텍스트 함께 검토 필요.");
        sb.AppendLine();
        sb.AppendLine("| 타입 | 필수 속성 | 타입참조 | 상태 |");
        sb.AppendLine("|---|---|---|---|");

        int missing = 0;
        foreach (var (type, prop, typeRef) in required.OrderBy(r => r.type).ThenBy(r => r.prop))
        {
            // word boundary 매칭 — ".Prop", "{ Prop = ", ", Prop = ", JSON target "Prop" 모두 커버
            var hit = Regex.IsMatch(corpusText, @"\b" + Regex.Escape(prop) + @"\b");
            var status = hit ? "✓" : "✗ **누락**";
            if (!hit) missing++;
            sb.AppendLine($"| {type} | `{prop}` | {EscapeMd(typeRef)} | {status} |");
        }
        sb.AppendLine();
        sb.AppendLine($"**총 필수 속성: {required.Count}개 / 누락 의심: {missing}개**");

        var outPath = Path.Combine(qaDir, "gaps_reverse.md");
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Reverse gap → {outPath}  (required={required.Count}, missing-suspect={missing})");
    }

    // ───────────────────────── 유틸 ─────────────────────────

    static List<Dictionary<string, string>> ReadCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        if (lines.Length == 0) return new();
        // BOM 제거
        var header = lines[0].TrimStart('\uFEFF').Split(',');
        var rows = new List<Dictionary<string, string>>();
        for (int i = 1; i < lines.Length; i++)
        {
            var fields = ParseCsvLine(lines[i]);
            var d = new Dictionary<string, string>();
            for (int j = 0; j < header.Length && j < fields.Count; j++)
                d[header[j]] = fields[j];
            rows.Add(d);
        }
        return rows;
    }

    static List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        var sb = new StringBuilder();
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (inQuotes)
            {
                if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                else if (c == '"') inQuotes = false;
                else sb.Append(c);
            }
            else
            {
                if (c == ',') { result.Add(sb.ToString()); sb.Clear(); }
                else if (c == '"') inQuotes = true;
                else sb.Append(c);
            }
        }
        result.Add(sb.ToString());
        return result;
    }

    static string TruncMd(string s)
    {
        s = s.Replace("|", "/").Replace("\r", " ").Replace("\n", " ");
        return s.Length > 60 ? s.Substring(0, 60) + "…" : s;
    }
    static string EscapeMd(string s) => s.Replace("|", "\\|");
}
