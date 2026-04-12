/// 국가별 계산 시트 구조 확인용 덤프
using ClosedXML.Excel;

var path = args.Length > 0 ? args[0]
    : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Resources", "group_template.xlsx");

var wb = new XLWorkbook(path);

Console.WriteLine("=== 시트 목록 ===");
foreach (var ws0 in wb.Worksheets)
    Console.WriteLine($"  [{ws0.Position}] \"{ws0.Name}\"  visible={ws0.Visibility}");

// 덤프할 시트 이름 (2번째 인자로 지정 가능)
var sheetName = args.Length > 1 ? args[1] : "국가별 계산";

if (!wb.TryGetWorksheet(sheetName, out var ws))
{
    Console.WriteLine($"\n'{sheetName}' 시트 없음");
    return;
}

int maxRow = ws.LastRowUsed()?.RowNumber() ?? 300;
Console.WriteLine($"\n=== {sheetName} (총 {maxRow}행) ===");

for (int r = 1; r <= maxRow; r++)
{
    var b  = ws.Cell(r,  2).GetString().Trim();
    var c  = ws.Cell(r,  3).GetString().Trim();
    var d  = ws.Cell(r,  4).GetString().Trim();
    var f  = ws.Cell(r,  6).GetString().Trim();
    var g  = ws.Cell(r,  7).GetString().Trim();
    var i  = ws.Cell(r,  9).GetString().Trim();
    var k  = ws.Cell(r, 11).GetString().Trim();
    var m  = ws.Cell(r, 13).GetString().Trim();
    var n  = ws.Cell(r, 14).GetString().Trim();
    var o  = ws.Cell(r, 15).GetString().Trim();
    var p  = ws.Cell(r, 16).GetString().Trim();

    if (string.IsNullOrEmpty(b + c + d + f + g + i + k + m + n + o + p)) continue;

    Console.WriteLine($"R{r,3}: B={b,-44} C={c,-18} D={d,-12} F={f,-14} G={g,-14} I={i,-14} K={k,-18} M={m,-18} N={n,-12} O={o,-20} P={p}");
}

// 적격국제해운부수소득 시트 진단
if (wb.TryGetWorksheet("적격국제해운부수소득", out var shipWs))
{
    Console.WriteLine("\n=== 적격국제해운부수소득 FindRow 진단 ===");
    int lastRow = shipWs.LastRowUsed()?.RowNumber() ?? 20;
    Console.WriteLine($"LastUsedRow={lastRow}");

    int rowHdr = -1;
    for (int r = 1; r <= lastRow; r++)
    {
        var txt = shipWs.Cell(r, 2).GetString()?.Trim() ?? "";
        Console.WriteLine($"  R{r}: B=\"{txt}\" | N={shipWs.Cell(r,14).GetString()} | O={shipWs.Cell(r,15).GetString()}");
        if (rowHdr < 0 && txt.Contains("(b) 적격국제해운부수소득"))
            rowHdr = r;
    }
    Console.WriteLine($"\n  rowHdr={(rowHdr >= 0 ? rowHdr.ToString() : "NOT FOUND")}");

    if (rowHdr >= 0)
    {
        static int FindRow(ClosedXML.Excel.IXLWorksheet w, string s, int from)
        {
            var last = w.LastRowUsed()?.RowNumber() ?? 300;
            for (int r = from; r <= last; r++)
                if ((w.Cell(r, 2).GetString()?.Trim() ?? "").Contains(s)) return r;
            return -1;
        }
        int r1 = FindRow(shipWs, "1. 모든 구성기업", rowHdr);
        int r2 = FindRow(shipWs, "2. 50% 한도", rowHdr);
        int r3 = FindRow(shipWs, "3. 모든 구성기업", rowHdr);
        int r4 = FindRow(shipWs, "4. B가 A의 50%", rowHdr);
        Console.WriteLine($"  r1={r1} val={shipWs.Cell(Math.Max(r1,1),14).GetString()}");
        Console.WriteLine($"  r2={r2} val={shipWs.Cell(Math.Max(r2,1),14).GetString()}");
        Console.WriteLine($"  r3={r3} val={shipWs.Cell(Math.Max(r3,1),14).GetString()}");
        Console.WriteLine($"  r4={r4} val={shipWs.Cell(Math.Max(r4,1),14).GetString()}");
    }
}
