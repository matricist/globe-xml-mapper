/// 국가별 계산 첨부 시트 전체 셀 덤프
using ClosedXML.Excel;

var path = args.Length > 0 ? args[0]
    : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Resources", "group_template.xlsx");
var sheetName = args.Length > 1 ? args[1] : "국가별 계산 첨부";

var wb = new XLWorkbook(path);
if (!wb.TryGetWorksheet(sheetName, out var ws))
{
    Console.WriteLine($"'{sheetName}' 시트 없음");
    return;
}

int maxRow = ws.LastRowUsed()?.RowNumber() ?? 10;
int maxCol = ws.LastColumnUsed()?.ColumnNumber() ?? 20;
Console.WriteLine($"=== {sheetName} ({maxRow}행 {maxCol}열) ===");

for (int r = 1; r <= maxRow; r++)
{
    var sb = new System.Text.StringBuilder($"R{r,3}: ");
    bool hasAny = false;
    for (int c = 1; c <= maxCol; c++)
    {
        var v = ws.Cell(r, c).GetString()?.Trim() ?? "";
        if (!string.IsNullOrEmpty(v))
        {
            var colLetter = ws.Cell(r, c).Address.ColumnLetter;
            sb.Append($"{colLetter}({c})='{v}'  ");
            hasAny = true;
        }
    }
    if (hasAny) Console.WriteLine(sb.ToString());
}
