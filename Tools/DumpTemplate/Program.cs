using System;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

// main_template.xlsx 의 모든 비어있지 않은 셀을 CSV로 덤프.
// 출력: qa/template_cells.csv
//   columns: sheet, address, row, col, mergeRange, mergeAnchor, kind, text
//     kind ∈ { label, value, candidate }
//     - label   : 한글 라벨/번호 패턴 (예: "1. ...", "(a) ...", "가. ...")
//     - value   : 숫자/소수점/날짜 등 명백한 데이터
//     - candidate: 그 외 (사용자가 채워야 할 위치 후보 포함, 비어있는 셀은 스캔 대상 아님)

class Program
{
    static int Main(string[] args)
    {
        var srcPath = args.Length > 0
            ? args[0]
            : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "Resources", "main_template.xlsx");
        var outPath = args.Length > 1
            ? args[1]
            : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "qa", "template_cells.csv");

        srcPath = Path.GetFullPath(srcPath);
        outPath = Path.GetFullPath(outPath);

        if (!File.Exists(srcPath))
        {
            Console.Error.WriteLine($"Not found: {srcPath}");
            return 1;
        }
        Directory.CreateDirectory(Path.GetDirectoryName(outPath)!);

        var sb = new StringBuilder();
        sb.AppendLine("sheet,address,row,col,mergeRange,mergeAnchor,kind,text");

        using var wb = new XLWorkbook(srcPath);
        int sheets = 0, cells = 0;

        foreach (var ws in wb.Worksheets)
        {
            if (ws.Name == "_META") continue;
            sheets++;

            var used = ws.RangeUsed();
            if (used == null) continue;

            foreach (var cell in used.CellsUsed())
            {
                var raw = cell.GetString() ?? "";
                var text = raw.Replace("\r", " ").Replace("\n", " ").Trim();
                if (string.IsNullOrEmpty(text)) continue;

                var merged = cell.MergedRange();
                string mergeRange = "";
                string mergeAnchor = "";
                if (merged != null)
                {
                    mergeRange = merged.RangeAddress.ToStringRelative();
                    mergeAnchor = merged.FirstCell().Address.ToString();
                    // merged 영역에서 anchor 셀이 아니면 스킵 (중복 출력 방지)
                    if (cell.Address.ToString() != mergeAnchor) continue;
                }

                var kind = ClassifyKind(text);

                sb.Append(Csv(ws.Name)).Append(',')
                  .Append(Csv(cell.Address.ToString())).Append(',')
                  .Append(cell.Address.RowNumber).Append(',')
                  .Append(cell.Address.ColumnNumber).Append(',')
                  .Append(Csv(mergeRange)).Append(',')
                  .Append(Csv(mergeAnchor)).Append(',')
                  .Append(kind).Append(',')
                  .Append(Csv(text)).AppendLine();
                cells++;
            }
        }

        // UTF-8 BOM 으로 저장 (Excel에서 한글 깨짐 방지)
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Dumped {cells} cells across {sheets} sheets to:");
        Console.WriteLine($"  {outPath}");
        return 0;
    }

    // 매우 간단한 휴리스틱. label / value / candidate 셋으로만 분류.
    // - 숫자나 날짜 패턴으로 시작 → value
    // - 라벨 패턴 (1./2./(a)/가./① 등) 또는 한글 4자 이상 → label
    // - 그 외 → candidate
    static string ClassifyKind(string text)
    {
        if (text.Length == 0) return "candidate";

        // 명백한 숫자/날짜
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[-+]?\d[\d,\.]*%?$")) return "value";
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d{4}[-./]\d{1,2}[-./]\d{1,2}")) return "value";

        // 라벨 패턴
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\(?\s*[a-zA-Z가-힣]\)\s*")) return "label";  // (a), (가)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+(\.\d+)*\s*[\.\)]?\s*")) return "label"; // 1., 1.1, 1)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[가-힣]\s*[\.\)]\s*")) return "label";       // 가., 나)
        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[①-⑳㉠-㉯]")) return "label";              // ①, ㉠

        // 한글 비중 + 길이로 라벨 추정
        int hangul = text.Count(c => c >= '가' && c <= '힣');
        if (hangul >= 2 && text.Length >= 3) return "label";

        return "candidate";
    }

    static string Csv(string? s)
    {
        if (s == null) return "";
        if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        return s;
    }
}
