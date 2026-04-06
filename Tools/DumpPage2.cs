using ClosedXML.Excel;

var filePath = @"c:\Users\jaewo\Desktop\work\mapper\Resources\국조53, 국조55, 국조56 서식(2026.3.20 개정분)-260402.xlsx";
var sheetName = "국조53부표2 (2)";

using var wb = new XLWorkbook(filePath);
var ws = wb.Worksheet(sheetName);

Console.OutputEncoding = System.Text.Encoding.UTF8;

Console.WriteLine($"=== Sheet: {sheetName} ===");
Console.WriteLine();

// Dump all non-empty cells in rows 1-50, columns A(1) - R(18)
Console.WriteLine("--- Non-empty cells (rows 1-50, cols A-R) ---");
for (int row = 1; row <= 50; row++)
{
    for (int col = 1; col <= 18; col++)
    {
        var cell = ws.Cell(row, col);
        // Check both Value and rich text
        if (!cell.IsEmpty())
        {
            var addr = cell.Address.ToString();
            var val = cell.GetFormattedString();
            if (string.IsNullOrEmpty(val))
                val = cell.Value.ToString();
            if (!string.IsNullOrWhiteSpace(val))
                Console.WriteLine($"  {addr} = {val}");
        }
    }
}

Console.WriteLine();
Console.WriteLine("--- Merged ranges ---");
foreach (var mr in ws.MergedRanges)
{
    Console.WriteLine($"  {mr.RangeAddress}");
}

Console.WriteLine();
Console.WriteLine("Done.");
