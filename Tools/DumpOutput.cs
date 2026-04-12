using ClosedXML.Excel;

var path = args.Length > 0 ? args[0] : "output_1321_test.xlsx";
var wb = new XLWorkbook(path);

Console.WriteLine("=== 시트 목록 ===");
foreach (var ws in wb.Worksheets)
    Console.WriteLine($"  [{ws.Position}] \"{ws.Name}\"  visible={ws.Visibility}");

// 모든 시트에서 O열(15) 비어있지 않은 값 출력
foreach (var ws in wb.Worksheets)
{
    var lines = new List<string>();
    for (int r = 1; r <= 60; r++)
    {
        var val = ws.Cell(r, 15).GetString();
        if (!string.IsNullOrWhiteSpace(val))
            lines.Add($"  O{r,-3} = {val}");
    }
    if (lines.Count > 0)
    {
        Console.WriteLine($"\n=== {ws.Name} — O열 ===");
        lines.ForEach(Console.WriteLine);
    }
}

// 첨부 포함 시트 B~D 출력
foreach (var ws in wb.Worksheets)
{
    if (!ws.Name.Contains("첨부")) continue;
    Console.WriteLine($"\n=== {ws.Name} — B/C/D열 ===");
    for (int r = 1; r <= 40; r++)
    {
        var b = ws.Cell(r, 2).GetString();
        var c = ws.Cell(r, 3).GetString();
        var d = ws.Cell(r, 4).GetString();
        if (!string.IsNullOrWhiteSpace(b + c + d))
            Console.WriteLine($"  행{r,-3}: B={b,-30} C={c,-20} D={d}");
    }
}
