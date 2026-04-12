/// <summary>
/// 사용법: dotnet run --project Tools/FillFromJson.csproj -- ce_data.json [output.xlsx]
///
/// HTML 프로토타입(prototype_1.3.2.1.html)에서 내려받은 ce_data.json을
/// template.xlsx의 1.3.2.1 시트에 기입하여 output.xlsx를 생성합니다.
/// </summary>

using System.Text.Json;
using ClosedXML.Excel;

// ── 인자 처리 ──────────────────────────────────────────────────────────────
if (args.Length < 1)
{
    Console.WriteLine("사용법: dotnet run -- <ce_data.json> [output.xlsx]");
    return 1;
}

var jsonPath    = args[0];
var outputPath  = args.Length >= 2 ? args[1] : "output_1321.xlsx";
var templatePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Resources", "main_template.xlsx");

if (!File.Exists(jsonPath))     { Console.WriteLine($"JSON 파일 없음: {jsonPath}"); return 1; }
if (!File.Exists(templatePath)) { Console.WriteLine($"main_template.xlsx 없음: {templatePath}"); return 1; }

// ── JSON 파싱 ──────────────────────────────────────────────────────────────
var json = await File.ReadAllTextAsync(jsonPath);
var ceArray = JsonSerializer.Deserialize<JsonElement[]>(json);
if (ceArray == null || ceArray.Length == 0)
{
    Console.WriteLine("CE 데이터가 없습니다.");
    return 1;
}

Console.WriteLine($"CE {ceArray.Length}개 읽음.");

// ── template.xlsx 복사 후 열기 ─────────────────────────────────────────────
File.Copy(templatePath, outputPath, overwrite: true);
using var wb = new XLWorkbook(outputPath);

// ── 1.3.2.1 시트 ──────────────────────────────────────────────────────────
const string CE_SHEET    = "1.3.2.1";
const string ATTACH_SHEET = "1.3.2.1 첨부";
const int BLOCK_START    = 4;   // Mapping_1.3.2.1 기준
const int BLOCK_SIZE     = 18;
const int BLOCK_GAP      = 2;
const int COL_O          = 15;  // O열

if (!wb.TryGetWorksheet(CE_SHEET, out var ws))
{
    Console.WriteLine($"시트 '{CE_SHEET}'를 찾을 수 없습니다.");
    return 1;
}

// 2번째 이상 CE: 행 블록 삽입 후 첫 블록에서 서식 복사
for (int i = 1; i < ceArray.Length; i++)
{
    var insertAt = BLOCK_START + i * (BLOCK_SIZE + BLOCK_GAP);   // 24, 44, …
    var gapStart = BLOCK_START + (i - 1) * (BLOCK_SIZE + BLOCK_GAP) + BLOCK_SIZE + 1; // 22, …

    // gap 2행 + 블록 BLOCK_SIZE행 = BLOCK_SIZE + BLOCK_GAP 행 삽입
    ws.Row(gapStart).InsertRowsAbove(BLOCK_SIZE + BLOCK_GAP);

    // 첫 블록(rows 4~21)에서 서식만 복사 → 새 블록 위치로
    for (int r = 0; r < BLOCK_SIZE; r++)
    {
        var srcRow = ws.Row(BLOCK_START + r);
        var dstRow = ws.Row(insertAt + r);
        foreach (var cell in srcRow.Cells(1, 20))
        {
            var dst = dstRow.Cell(cell.Address.ColumnNumber);
            dst.Style = cell.Style;
            // 레이블(A~N열) 값도 복사 (서식 참고용)
            if (cell.Address.ColumnNumber < COL_O)
                dst.Value = cell.Value;
        }
        dstRow.Height = srcRow.Height;
    }
}

// ── CE 값 기입 ────────────────────────────────────────────────────────────
for (int i = 0; i < ceArray.Length; i++)
{
    var ce     = ceArray[i];
    var bStart = BLOCK_START + i * (BLOCK_SIZE + BLOCK_GAP);

    void SetO(int offset, object? val)
    {
        if (val == null) return;
        ws.Cell(bStart + offset, COL_O).Value = val switch
        {
            bool b   => (XLCellValue)b.ToString().ToLower(),
            decimal d => (XLCellValue)d,
            double  d => (XLCellValue)d,
            _         => (XLCellValue)(val?.ToString() ?? ""),
        };
    }

    // offset 1: ChangeFlag
    if (ce.TryGetProperty("changeFlag", out var cf))
        SetO(1, cf.GetBoolean() ? "true" : "false");

    // offset 2: ResCountryCode (첫 번째)
    if (ce.TryGetProperty("id", out var id))
    {
        if (id.TryGetProperty("resCountryCode", out var rcc) && rcc.GetArrayLength() > 0)
            SetO(2, rcc[0].GetString());

        // offset 3: Rules (쉼표 구분)
        if (id.TryGetProperty("rules", out var rules) && rules.GetArrayLength() > 0)
            SetO(3, string.Join(",", rules.EnumerateArray().Select(r => r.GetString())));

        // offset 4: Name
        if (id.TryGetProperty("name", out var name))
            SetO(4, name.GetString());

        // offset 5: TIN (첫 번째)
        if (id.TryGetProperty("tin", out var tins) && tins.GetArrayLength() > 0)
            SetO(5, tins[0].GetProperty("value").GetString());

        // offset 6: ReceivingTin
        if (id.TryGetProperty("receivingTin", out var rTin))
            SetO(6, rTin.GetString());

        // offset 7: GlobeStatus (쉼표 구분)
        if (id.TryGetProperty("globeStatus", out var gs) && gs.GetArrayLength() > 0)
            SetO(7, string.Join(",", gs.EnumerateArray().Select(g => g.GetString())));
    }

    // 별첨 참조: 템플릿은 ExcelController 기준 row 3+10=13에 "첨부1" 기입됨
    // block 0은 이미 있으므로 건드리지 않고, block 1 이상만 직접 기입
    if (i > 0)
    {
        const int CE_BLOCK_START_EC = 3;    // ExcelController 기준
        const int CE_ATTACH_OFFSET  = 10;
        var attachRefRow = CE_BLOCK_START_EC + i * (BLOCK_SIZE + BLOCK_GAP) + CE_ATTACH_OFFSET;
        ws.Cell(attachRefRow, COL_O).Value = $"첨부{i + 1}";
    }

    // QIIR (offset 12~14)
    if (ce.TryGetProperty("qiir", out var qiir))
    {
        if (qiir.TryGetProperty("popeIpe", out var pope))
            SetO(12, pope.GetString());
        if (qiir.TryGetProperty("exception", out var ex) &&
            ex.TryGetProperty("tin", out var exTin))
            SetO(13, exTin.GetProperty("value").GetString());
        if (qiir.TryGetProperty("mopeIpe", out var mope) &&
            mope.TryGetProperty("tin", out var mopeTin))
            SetO(14, mopeTin.GetProperty("value").GetString());
    }

    // QUTPR (offset 15~17)
    if (ce.TryGetProperty("qutpr", out var qutpr))
    {
        if (qutpr.TryGetProperty("art93", out var art93))
            SetO(15, art93.GetBoolean() ? "true" : "false");
        if (qutpr.TryGetProperty("aggOwnership", out var agg))
            SetO(16, Math.Round(agg.GetDecimal() * 100, 2).ToString());
        if (qutpr.TryGetProperty("upeOwnership", out var upeOw))
            SetO(17, upeOw.GetBoolean() ? "true" : "false");
    }
}

// ── 별첨 시트 (소유지분) ──────────────────────────────────────────────────
if (wb.TryGetWorksheet(ATTACH_SHEET, out var attachWs))
{
    for (int i = 0; i < ceArray.Length; i++)
    {
        var ce = ceArray[i];
        if (!ce.TryGetProperty("ownership", out var owList) || owList.GetArrayLength() == 0)
            continue;

        var attachNum = i + 1;

        // "첨부N" 헤더 행 찾기
        int headerRow = -1;
        for (int r = 1; r <= 500; r++)
        {
            var val = attachWs.Cell(r, 2).GetString().Trim();
            if (val == $"첨부{attachNum}") { headerRow = r; break; }
        }
        if (headerRow < 0)
        {
            // 첨부1 구조를 복사하여 새 섹션 추가
            headerRow = AppendAttachSection(attachWs, attachNum);
            if (headerRow < 0)
            {
                Console.WriteLine($"첨부{attachNum} 섹션 생성 실패.");
                continue;
            }
        }

        // 데이터 행: headerRow + 3 (제목 + 빈행 + 헤더)
        int dataRow = headerRow + 3;

        foreach (var ow in owList.EnumerateArray())
        {
            if (ow.TryGetProperty("ownershipType", out var owType))
                attachWs.Cell(dataRow, 2).Value = owType.GetString();
            if (ow.TryGetProperty("tin", out var owTin))
                attachWs.Cell(dataRow, 3).Value = owTin.GetString();
            if (ow.TryGetProperty("ownershipPercentage", out var owPct))
                attachWs.Cell(dataRow, 4).Value = Math.Round(owPct.GetDecimal() * 100, 2);
            dataRow++;
        }
    }
}
else
{
    Console.WriteLine($"경고: '{ATTACH_SHEET}' 시트 없음 — 소유지분 데이터 생략");
}

// ── 별첨 섹션 추가 헬퍼 ───────────────────────────────────────────────────
static int AppendAttachSection(IXLWorksheet ws, int attachNum)
{
    // 첨부1 헤더 위치 찾기
    int ref1Row = -1;
    for (int r = 1; r <= 500; r++)
    {
        if (ws.Cell(r, 2).GetString().Trim() == "첨부1") { ref1Row = r; break; }
    }
    if (ref1Row < 0) return -1;

    // 첨부1 섹션 크기 파악 (헤더3행 + 데이터1행 + 구분1행 = 5행 기본)
    // 다음 "첨부" 시작 또는 빈 구간까지 측정
    int sec1End = ref1Row + 3; // 최소한 헤더+빈+컬럼헤더+데이터 = 4행
    for (int r = ref1Row + 1; r <= ref1Row + 30; r++)
    {
        var v = ws.Cell(r, 2).GetString().Trim();
        if (v.StartsWith("첨부") && v != "첨부1") { sec1End = r - 1; break; }
        sec1End = r;
    }

    int sectionRows = sec1End - ref1Row + 1 + 1; // +1 구분 빈행

    // 마지막 행 다음에 새 섹션 삽입
    int insertAt = sec1End + 2; // 구분행 이후

    // 첨부1 섹션 복사 — 헤더 3행은 서식+레이블 복사, 데이터 행은 서식만 복사 (값 제외)
    for (int r = 0; r < sectionRows - 1; r++)
    {
        var srcRow = ws.Row(ref1Row + r);
        var dstRow = ws.Row(insertAt + r);
        bool isHeaderRow = r < 3; // 행0=제목, 행1=빈, 행2=컬럼헤더

        foreach (var cell in srcRow.Cells(1, 10))
        {
            var dst = dstRow.Cell(cell.Address.ColumnNumber);
            dst.Style = cell.Style;

            if (isHeaderRow)
            {
                // 제목 행의 "첨부1"만 "첨부N"으로 교체, 나머지 헤더 값은 그대로 복사
                if (r == 0 && cell.Address.ColumnNumber == 2 && cell.GetString().Trim() == "첨부1")
                    dst.Value = $"첨부{attachNum}";
                else if (cell.Address.ColumnNumber <= 5)
                    dst.Value = cell.Value;
            }
            // 데이터 행(r >= 3): 서식만 복사, 값은 비워둠
        }
        dstRow.Height = srcRow.Height;
    }

    return insertAt; // 새 헤더 행 위치
}

// ── 저장 ──────────────────────────────────────────────────────────────────
wb.Save();
Console.WriteLine($"저장 완료: {Path.GetFullPath(outputPath)}");
return 0;
