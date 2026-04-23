using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// ws.Cell(row_expr, col) 호출 옆에 해결된 셀 주소 주석(// B15, // O30 등) 자동 추가.
//   - OffsetCheck 와 동일한 행 해석 규칙 사용 (상수 + FindRow + b1/b2 seed + 루프변수 회피)
//   - 이미 정확한 주석이 있으면 건너뛰고, 불일치 주석이 있으면 수정
//   - 해결 불가(루프변수 / 다블록 / 비-리터럴 col)는 주석 없이 원본 유지
// 출력: 파일 in-place 수정 + 콘솔 요약

class Program
{
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

    static readonly Dictionary<(string, int, int), (string kind, string text)> CellIndex = new();
    static readonly Dictionary<(string, int), string> RowLabels = new();

    static readonly Regex CellRegex = new(
        @"ws\.Cell\s*\(\s*([^,)]+?)\s*,\s*(\d+)\s*\)",
        RegexOptions.Compiled);
    static readonly Regex FindRowAssignRegex = new(
        @"\b(?:var|int)\s+(\w+)\s*=\s*FindRow\s*\(\s*ws\s*,\s*""([^""]+)""",
        RegexOptions.Compiled);

    static int Main(string[] args)
    {
        bool dryRun = args.Contains("--dry-run");
        var posArgs = args.Where(a => !a.StartsWith("--")).ToArray();
        var repo = posArgs.Length > 0
            ? Path.GetFullPath(posArgs[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var templateCsv = Path.Combine(repo, "qa", "template_cells.csv");
        if (!File.Exists(templateCsv))
        {
            Console.Error.WriteLine("먼저 Tools/DumpTemplate 을 실행하세요: qa/template_cells.csv 필요");
            return 1;
        }
        LoadTemplateCsv(templateCsv);

        var srcDir = Path.Combine(repo, "Services");
        int totalAdded = 0, totalFixed = 0, totalSkipped = 0, totalAlready = 0;

        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs").OrderBy(p => p))
        {
            var name = Path.GetFileName(file);
            if (!FileToSheet.TryGetValue(name, out var sheet)) continue;

            var result = AnnotateFile(file, name, sheet, dryRun);
            Console.WriteLine($"  {name}: 추가={result.added}, 수정={result.fixed_}, 유지={result.already}, 스킵={result.skipped}");
            totalAdded += result.added;
            totalFixed += result.fixed_;
            totalAlready += result.already;
            totalSkipped += result.skipped;
        }

        Console.WriteLine();
        Console.WriteLine($"합계: 추가 {totalAdded} / 수정 {totalFixed} / 유지 {totalAlready} / 해결불가 스킵 {totalSkipped}");
        if (dryRun) Console.WriteLine("(--dry-run: 실제 파일 변경 없음)");
        return 0;
    }

    static (int added, int fixed_, int already, int skipped) AnnotateFile(
        string file, string fileName, string sheet, bool dryRun)
    {
        var lines = File.ReadAllLines(file);
        var text = string.Join("\n", lines);

        // 변수 해석: 상수 + FindRow + 파일별 seed + 루프변수 검출
        var decls = new List<(int Pos, string Name, int Row)>();
        var globals = new Dictionary<string, int>();

        foreach (Match m in Regex.Matches(text, @"private\s+const\s+int\s+(\w+)\s*=\s*(\d+)\s*;"))
            globals[m.Groups[1].Value] = int.Parse(m.Groups[2].Value);

        foreach (Match m in FindRowAssignRegex.Matches(text))
        {
            var row = FindHeaderRowInTemplate(sheet, m.Groups[2].Value);
            if (row > 0) decls.Add((m.Index, m.Groups[1].Value, row));
        }

        if (fileName == "Mapping_2.cs")
        {
            globals["b1"] = 2;
            globals["b2"] = 25;
        }
        else if (fileName == "Mapping_1.3.1.cs" || fileName == "Mapping_1.3.2.1.cs"
              || fileName == "Mapping_1.3.2.2.cs")
        {
            var bh = Regex.Match(text, @"BLOCK_HEADER\s*=\s*""([^""]+)""");
            if (bh.Success)
            {
                var row = FindHeaderRowInTemplate(sheet, bh.Groups[1].Value);
                if (row > 0) globals["blockStartRow"] = row;
            }
        }
        decls.Sort((a, b) => a.Pos.CompareTo(b.Pos));

        // for/foreach 루프 변수 (30행 윈도우)
        var loopVars = new List<(int StartLine, int EndLine, string Name)>();
        foreach (Match m in Regex.Matches(text, @"\bfor\s*\(\s*(?:var|int)\s+(\w+)\s*="))
        {
            var start = text.Substring(0, m.Index).Count(c => c == '\n');
            loopVars.Add((start, start + 30, m.Groups[1].Value));
        }
        foreach (Match m in Regex.Matches(text, @"\bforeach\s*\(\s*(?:var|int)\s+(\w+)\s+in\b"))
        {
            var start = text.Substring(0, m.Index).Count(c => c == '\n');
            loopVars.Add((start, start + 30, m.Groups[1].Value));
        }

        bool IsLoopVar(string name, int lineIdx)
        {
            foreach (var lv in loopVars)
                if (lv.Name == name && lineIdx >= lv.StartLine && lineIdx <= lv.EndLine) return true;
            return false;
        }

        int LookupAt(string name, int pos)
        {
            int best = 0;
            foreach (var ev in decls)
            {
                if (ev.Pos > pos) break;
                if (ev.Name == name) best = ev.Row;
            }
            if (best > 0) return best;
            return globals.TryGetValue(name, out var v) ? v : 0;
        }

        int ResolveRow(string expr, int pos)
        {
            if (int.TryParse(expr, out var n)) return n;
            var m = Regex.Match(expr, @"^(\w+)\s*([+\-])\s*(\d+)$");
            if (m.Success)
            {
                var baseRow = LookupAt(m.Groups[1].Value, pos);
                if (baseRow <= 0) return 0;
                var off = int.Parse(m.Groups[3].Value);
                return m.Groups[2].Value == "+" ? baseRow + off : baseRow - off;
            }
            return LookupAt(expr, pos);
        }

        // 라인 시작 인덱스
        var lineStart = new List<int> { 0 };
        for (int k = 0; k < text.Length; k++)
            if (text[k] == '\n') lineStart.Add(k + 1);

        // 각 라인 처리
        int added = 0, fixed_ = 0, already = 0, skipped = 0;
        bool changed = false;
        for (int i = 0; i < lines.Length; i++)
        {
            int charIdx = i < lineStart.Count ? lineStart[i] : 0;
            var matches = CellRegex.Matches(lines[i]);
            if (matches.Count == 0) continue;

            // 한 라인의 첫 ws.Cell 만 주석화 (멀티콜 라인은 너무 복잡)
            if (matches.Count > 1)
            {
                skipped += matches.Count;
                continue;
            }

            var m = matches[0];
            var rowExpr = m.Groups[1].Value.Trim();
            var col = int.Parse(m.Groups[2].Value);

            var firstVar = Regex.Match(rowExpr, @"^(\w+)").Groups[1].Value;
            if (IsLoopVar(firstVar, i)) { skipped++; continue; }

            var resolvedRow = ResolveRow(rowExpr, charIdx);
            if (resolvedRow <= 0) { skipped++; continue; }

            var expected = $"{ColLetter(col)}{resolvedRow}";

            // 라인 끝의 기존 주석 확인
            var commentIdx = FindLineCommentStart(lines[i]);
            if (commentIdx < 0)
            {
                // 주석 없음 → 추가
                lines[i] = lines[i].TrimEnd() + $" // {expected}";
                added++;
                changed = true;
            }
            else
            {
                var existing = lines[i].Substring(commentIdx);
                // 기존 주석에 셀주소 패턴이 있으면 비교
                var cellInExisting = Regex.Match(existing, @"\b([A-Z]+)(\d+)\b");
                if (cellInExisting.Success)
                {
                    var existingAddr = cellInExisting.Groups[0].Value;
                    if (existingAddr == expected) { already++; }
                    else
                    {
                        // 불일치 — 단순 치환 (동일 패턴만)
                        lines[i] = lines[i].Substring(0, commentIdx) +
                                    existing.Replace(existingAddr, expected);
                        fixed_++;
                        changed = true;
                    }
                }
                else
                {
                    // 주석은 있지만 셀주소 패턴 없음 → 기존 뒤에 이어붙이기
                    lines[i] = lines[i].TrimEnd() + $" ({expected})";
                    added++;
                    changed = true;
                }
            }
        }

        if (changed && !dryRun)
            File.WriteAllLines(file, lines, new UTF8Encoding(true));

        return (added, fixed_, already, skipped);
    }

    // 라인 내에서 최초 // 주석 시작 위치 (문자열 리터럴 내부 // 는 제외)
    static int FindLineCommentStart(string line)
    {
        bool inStr = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"') inStr = !inStr;
            else if (!inStr && c == '/' && i + 1 < line.Length && line[i + 1] == '/')
                return i;
        }
        return -1;
    }

    static string ColLetter(int col)
    {
        var sb = new StringBuilder();
        while (col > 0)
        {
            col--;
            sb.Insert(0, (char)('A' + col % 26));
            col /= 26;
        }
        return sb.ToString();
    }

    static void LoadTemplateCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = ParseCsvLine(lines[i]);
            if (parts.Count < 8) continue;
            var sheet = parts[0];
            if (!int.TryParse(parts[2], out var row)) continue;
            if (!int.TryParse(parts[3], out var col)) continue;
            CellIndex[(sheet, row, col)] = (parts[6], parts[7]);
            if (col == 2 && parts[6] == "label" && !RowLabels.ContainsKey((sheet, row)))
                RowLabels[(sheet, row)] = parts[7];
        }
    }

    static int FindHeaderRowInTemplate(string sheet, string header)
    {
        foreach (var kv in RowLabels)
            if (kv.Key.Item1 == sheet && kv.Value.Contains(header)) return kv.Key.Item2;
        foreach (var kv in CellIndex)
            if (kv.Key.Item1 == sheet && kv.Value.kind == "label" && kv.Value.text.Contains(header))
                return kv.Key.Item2;
        return 0;
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
}
