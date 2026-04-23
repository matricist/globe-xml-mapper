using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// Off-by-one 행 오프셋 버그 자동 탐지:
//
//   각 Mapping_*.cs 의 `ws.Cell(row_expr, col)` 호출을 정적 해석:
//     1) 시트를 추론 (파일 ↔ 시트 매핑)
//     2) row_expr 를 구체 행번호로 해석
//          - 상수 기반: BLOCK1_START, DATA_START_ROW 등
//          - FindRow 기반: `var r = FindRow(ws, "...")` → 해당 텍스트를 template 에서 찾아 해결
//          - 블록 헤더 기반: `blockStartRow` 는 BLOCK_HEADER 첫 출현 행
//     3) template_cells.csv 에서 (sheet, row, col) kind 조회
//     4) kind=label 이면 "라벨을 값으로 읽는 의심" 경고
//
//   출력: qa/offset_check.md

class Program
{
    // 파일명 → 시트명 (TemplateMeta.SheetMap 과 동일)
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

    // 시트별 (row, col) → (kind, text)
    static readonly Dictionary<(string, int, int), (string kind, string text)> CellIndex = new();
    // 시트별 row → (any row에서 col=2 라벨 텍스트)
    static readonly Dictionary<(string, int), string> RowLabels = new();

    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var qaDir = Path.Combine(repo, "qa");
        var templateCsv = Path.Combine(qaDir, "template_cells.csv");
        if (!File.Exists(templateCsv))
        {
            Console.Error.WriteLine("먼저 Tools/DumpTemplate 을 실행해 qa/template_cells.csv 를 생성하세요.");
            return 1;
        }

        LoadTemplateCsv(templateCsv);

        var srcDir = Path.Combine(repo, "Services");
        var findings = new List<Finding>();

        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs").OrderBy(p => p))
        {
            var name = Path.GetFileName(file);
            if (!FileToSheet.TryGetValue(name, out var sheet)) continue;
            AnalyzeMapper(file, name, sheet, findings);
        }

        var outPath = Path.Combine(qaDir, "offset_check.md");
        WriteReport(findings, outPath);
        Console.WriteLine($"Offset check → {outPath}  (suspicious={findings.Count})");
        return 0;
    }

    // ───────────────────────── Template CSV 로드 ─────────────────────────
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
            var kind = parts[6];
            var text = parts[7];
            CellIndex[(sheet, row, col)] = (kind, text);
            if (col == 2 && kind == "label" && !RowLabels.ContainsKey((sheet, row)))
                RowLabels[(sheet, row)] = text;
        }
    }

    // ───────────────────────── 매퍼 분석 ─────────────────────────

    record Finding(string File, int Line, string Sheet, int ResolvedRow, int Col, string RowExpr, string LabelAtCell, string Note);

    // ws.Cell(row_expr, literal_col)
    static readonly Regex CellRegex = new(
        @"ws\.Cell\s*\(\s*([^,)]+?)\s*,\s*(\d+)\s*\)",
        RegexOptions.Compiled);

    // var r = FindRow(ws, "header");   / var r = FindRow(ws, "header", startRow);
    static readonly Regex FindRowAssignRegex = new(
        @"\b(?:var|int)\s+(\w+)\s*=\s*FindRow\s*\(\s*ws\s*,\s*""([^""]+)""",
        RegexOptions.Compiled);

    static void AnalyzeMapper(string file, string fileName, string sheet, List<Finding> findings)
    {
        var text = File.ReadAllText(file);
        var lines = File.ReadAllLines(file);

        // 위치 기반 변수 선언: (pos, varName, row)
        var decls = new List<(int Pos, string Name, int Row)>();
        // 전역 상수 (위치 무관) — const 키워드 + 특수 seed
        var globals = new Dictionary<string, int>();

        // 1) 상수 (파일 전역)
        var constRegex = new Regex(@"private\s+const\s+int\s+(\w+)\s*=\s*(\d+)\s*;", RegexOptions.Compiled);
        foreach (Match m in constRegex.Matches(text))
            globals[m.Groups[1].Value] = int.Parse(m.Groups[2].Value);

        // 2) FindRow 기반 변수 선언 (위치 기록)
        foreach (Match m in FindRowAssignRegex.Matches(text))
        {
            var varName = m.Groups[1].Value;
            var header = m.Groups[2].Value;
            var row = FindHeaderRowInTemplate(sheet, header);
            if (row > 0) decls.Add((m.Index, varName, row));
        }

        // 3) 매퍼별 seed
        if (fileName == "Mapping_2.cs")
        {
            // 첫 블록 기준 (idx=0)
            globals["b1"] = 2;
            globals["b2"] = 25;
        }
        else if (fileName == "Mapping_1.3.1.cs" || fileName == "Mapping_1.3.2.1.cs"
              || fileName == "Mapping_1.3.2.2.cs")
        {
            var blockHeaderRegex = new Regex(@"BLOCK_HEADER\s*=\s*""([^""]+)""", RegexOptions.Compiled);
            var bh = blockHeaderRegex.Match(text);
            if (bh.Success)
            {
                var row = FindHeaderRowInTemplate(sheet, bh.Groups[1].Value);
                if (row > 0) globals["blockStartRow"] = row;
            }
        }

        decls.Sort((a, b) => a.Pos.CompareTo(b.Pos));

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

        // 4) ws.Cell 호출 분석 (위치 기반)
        //   라인 단위지만 각 라인의 시작 인덱스를 charIdx 로 계산해 decls lookup
        var lineStart = new List<int> { 0 };
        for (int k = 0; k < text.Length; k++)
            if (text[k] == '\n') lineStart.Add(k + 1);

        // for-loop / foreach / out 변수 수집 — 이 변수는 FindRow 값이 아님을 표시
        // 루프 변수가 쓰이는 라인 범위는 루프 시작 이후 20 라인 정도로 휴리스틱
        var loopVarRegex = new Regex(@"\bfor\s*\(\s*(?:var|int)\s+(\w+)\s*=", RegexOptions.Compiled);
        var foreachVarRegex = new Regex(@"\bforeach\s*\(\s*(?:var|int)\s+(\w+)\s+in\b", RegexOptions.Compiled);
        var loopVars = new List<(int StartLine, int EndLine, string Name)>();
        foreach (Match m in loopVarRegex.Matches(text))
        {
            var name = m.Groups[1].Value;
            var startLine = text.Substring(0, m.Index).Count(c => c == '\n');
            loopVars.Add((startLine, startLine + 30, name));
        }
        foreach (Match m in foreachVarRegex.Matches(text))
        {
            var name = m.Groups[1].Value;
            var startLine = text.Substring(0, m.Index).Count(c => c == '\n');
            loopVars.Add((startLine, startLine + 30, name));
        }

        bool IsLoopVar(string name, int lineIdx)
        {
            foreach (var lv in loopVars)
                if (lv.Name == name && lineIdx >= lv.StartLine && lineIdx <= lv.EndLine) return true;
            return false;
        }

        for (int i = 0; i < lines.Length; i++)
        {
            int charIdx = i < lineStart.Count ? lineStart[i] : 0;
            foreach (Match m in CellRegex.Matches(lines[i]))
            {
                var rowExpr = m.Groups[1].Value.Trim();
                var col = int.Parse(m.Groups[2].Value);

                // 루프 변수가 포함된 식이면 해석 포기
                var firstVar = Regex.Match(rowExpr, @"^(\w+)").Groups[1].Value;
                if (IsLoopVar(firstVar, i)) continue;

                var resolvedRow = ResolveRowScoped(rowExpr, n => LookupAt(n, charIdx));
                if (resolvedRow <= 0) continue;

                if (CellIndex.TryGetValue((sheet, resolvedRow, col), out var info)
                    && info.kind == "label")
                {
                    var nearbyLabel = NearbyLabel(sheet, resolvedRow);
                    findings.Add(new Finding(
                        fileName, i + 1, sheet, resolvedRow, col, rowExpr,
                        info.text, "[라벨셀] " + nearbyLabel));
                }

                // 주석 vs 해결된 행 불일치 탐지
                //   예: ws.Cell(b2 + 1, 15).Trim(); // O26   → 해결된 행이 26 이 아니면 경고
                var commentM = Regex.Match(lines[i], @"//\s*([A-Z]+)(\d+)");
                if (commentM.Success)
                {
                    var commentCol = ColLetterToNumber(commentM.Groups[1].Value);
                    var commentRow = int.Parse(commentM.Groups[2].Value);
                    if (commentCol == col && commentRow != resolvedRow)
                    {
                        findings.Add(new Finding(
                            fileName, i + 1, sheet, resolvedRow, col, rowExpr,
                            $"주석:{commentM.Groups[0].Value.TrimStart('/', ' ')}  해결:{ColLetter(col)}{resolvedRow}",
                            "[주석 불일치]"));
                    }
                }
            }
        }
    }

    static int ResolveRowScoped(string expr, Func<string, int> lookup)
    {
        if (int.TryParse(expr, out var n)) return n;
        var m = Regex.Match(expr, @"^(\w+)\s*([+\-])\s*(\d+)$");
        if (m.Success)
        {
            var baseRow = lookup(m.Groups[1].Value);
            if (baseRow <= 0) return 0;
            var off = int.Parse(m.Groups[3].Value);
            return m.Groups[2].Value == "+" ? baseRow + off : baseRow - off;
        }
        return lookup(expr);
    }

    static int ResolveRow(string expr, Dictionary<string, int> varToRow)
    {
        // 순수 숫자
        if (int.TryParse(expr, out var n)) return n;

        // X + N 또는 X - N (공백 허용)
        var m = Regex.Match(expr, @"^(\w+)\s*([+\-])\s*(\d+)$");
        if (m.Success)
        {
            if (!varToRow.TryGetValue(m.Groups[1].Value, out var baseRow)) return 0;
            var off = int.Parse(m.Groups[3].Value);
            return m.Groups[2].Value == "+" ? baseRow + off : baseRow - off;
        }

        // 단일 변수
        if (varToRow.TryGetValue(expr, out var v)) return v;

        return 0;
    }

    static int FindHeaderRowInTemplate(string sheet, string header)
    {
        // 시트 내 col=2 라벨 중 header 문자열을 포함하는 첫 행
        foreach (var kv in RowLabels)
        {
            if (kv.Key.Item1 != sheet) continue;
            if (kv.Value.Contains(header)) return kv.Key.Item2;
        }
        // 다른 컬럼에도 존재할 수 있음 (예: D9 등) — 포괄 검색
        foreach (var kv in CellIndex)
        {
            if (kv.Key.Item1 != sheet) continue;
            if (kv.Value.kind != "label") continue;
            if (kv.Value.text.Contains(header)) return kv.Key.Item2;
        }
        return 0;
    }

    static string NearbyLabel(string sheet, int row)
    {
        // 해당 행의 col=2 라벨 우선
        if (RowLabels.TryGetValue((sheet, row), out var t)) return t;
        // 바로 이전 행의 col=2 라벨
        for (int r = row; r >= Math.Max(1, row - 3); r--)
            if (RowLabels.TryGetValue((sheet, r), out var t2)) return $"(~{r}): {t2}";
        return "";
    }

    // ───────────────────────── 리포트 ─────────────────────────

    static void WriteReport(List<Finding> findings, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Off-by-one 행 오프셋 자동 점검");
        sb.AppendLine();
        sb.AppendLine("각 매퍼의 `ws.Cell(row_expr, col)` 호출에서 해결된 행 위치가 **template 의 라벨 셀** 에 해당하면 보고.");
        sb.AppendLine("라벨 셀을 값으로 읽고 있으면 거의 확실한 off-by-one 버그.");
        sb.AppendLine();
        sb.AppendLine("> 한계: b1/b2 는 첫 블록 기준으로만 해석. FindRow 는 템플릿의 해당 헤더 첫 출현으로 해석.");
        sb.AppendLine("> 단일 변수(변수명만 있는 `ws.Cell(r, col)`) 는 +0 오프셋으로 간주.");
        sb.AppendLine();

        if (findings.Count == 0)
        {
            sb.AppendLine("## ✓ 의심 사례 없음");
            sb.AppendLine();
            File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
            return;
        }

        sb.AppendLine($"## ⚠ 총 {findings.Count}건 의심");
        sb.AppendLine();
        sb.AppendLine("| 파일 | 라인 | 시트 | 읽는 셀 | 행표현 | 라벨 텍스트 | 인근 라벨 |");
        sb.AppendLine("|---|---|---|---|---|---|---|");
        foreach (var f in findings.OrderBy(f => f.File).ThenBy(f => f.Line))
        {
            sb.Append("| ").Append(f.File)
              .Append(" | ").Append(f.Line)
              .Append(" | ").Append(f.Sheet)
              .Append(" | ").Append(ColLetter(f.Col)).Append(f.ResolvedRow)
              .Append(" | `").Append(EscapeMd(f.RowExpr)).Append("`")
              .Append(" | ").Append(EscapeMd(Truncate(f.LabelAtCell, 50)))
              .Append(" | ").Append(EscapeMd(Truncate(f.Note, 50)))
              .AppendLine(" |");
        }

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
    }

    static int ColLetterToNumber(string letters)
    {
        int col = 0;
        foreach (var c in letters)
        {
            if (c < 'A' || c > 'Z') return 0;
            col = col * 26 + (c - 'A' + 1);
        }
        return col;
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

    // ───────────────────────── 유틸 ─────────────────────────
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

    static string Truncate(string? s, int len)
    {
        if (string.IsNullOrEmpty(s)) return "";
        s = s.Replace("\r", " ").Replace("\n", " ").Replace("|", "/");
        return s.Length > len ? s.Substring(0, len) + "…" : s;
    }
    static string EscapeMd(string s) => s.Replace("|", "\\|");
}
