using ClosedXML.Excel;

string templatePath = Path.Combine(
    Directory.GetParent(AppContext.BaseDirectory)!.Parent!.Parent!.Parent!.FullName,
    "..", "Resources", "template.xlsx");
templatePath = Path.GetFullPath(templatePath);

// Allow override via environment or just hardcode the known path
templatePath = @"c:\Users\jaewo\Desktop\work\mapper\Resources\template.xlsx";

Console.WriteLine($"Opening: {templatePath}");

using var workbook = new XLWorkbook(templatePath);

// 1. List all existing sheet names
Console.WriteLine("\n=== Existing Sheets ===");
foreach (var ws in workbook.Worksheets)
{
    Console.WriteLine($"  - {ws.Name}");
}

// 2. Add or get the guide sheet
const string sheetName = "작성요령";
IXLWorksheet sheet;
if (workbook.Worksheets.TryGetWorksheet(sheetName, out var existing))
{
    Console.WriteLine($"\nSheet '{sheetName}' already exists. Clearing and updating...");
    existing.Clear();
    sheet = existing;
}
else
{
    Console.WriteLine($"\nAdding new sheet '{sheetName}'...");
    sheet = workbook.Worksheets.Add(sheetName);
}

// Column widths
sheet.Column("A").Width = 10;
sheet.Column("B").Width = 10;
sheet.Column("C").Width = 30;
sheet.Column("D").Width = 10;
sheet.Column("E").Width = 20;
sheet.Column("F").Width = 12;
sheet.Column("G").Width = 60;
sheet.Column("H").Width = 40;

// Wrap text on G and H
sheet.Column("G").Style.Alignment.WrapText = true;
sheet.Column("H").Style.Alignment.WrapText = true;

// --- Row 1: Header ---
var headers = new[] { "섹션", "항목번호", "항목명", "셀 위치", "입력 형식", "복수기재", "허용값 (Enum)", "비고" };
for (int i = 0; i < headers.Length; i++)
{
    var cell = sheet.Cell(1, i + 1);
    cell.Value = headers[i];
    cell.Style.Font.Bold = true;
    cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
}

// Row 2: empty separator (just skip)

// Helper to fill a row
void SetRow(int row, string section, string itemNo, string itemName, string cellPos,
            string format, string multi, string enumVal, string note)
{
    sheet.Cell(row, 1).Value = section;
    sheet.Cell(row, 2).Value = itemNo;
    sheet.Cell(row, 3).Value = itemName;
    sheet.Cell(row, 4).Value = cellPos;
    sheet.Cell(row, 5).Value = format;
    sheet.Cell(row, 6).Value = multi;
    sheet.Cell(row, 7).Value = enumVal;
    sheet.Cell(row, 8).Value = note;
}

// 헤더 섹션
SetRow(3, "헤더", "-", "사업연도 시작", "C3", "날짜 (YYYY-MM-DD)", "X", "-", "-");
SetRow(4, "헤더", "-", "사업연도 종료", "C6", "날짜 (YYYY-MM-DD)", "X", "-", "-");
SetRow(5, "헤더", "-", "법인명(상호)", "P3", "텍스트", "X", "-", "-");
SetRow(6, "헤더", "-", "사업자등록번호", "P5", "텍스트", "X", "-", "-");

// Row 7: empty

// 1.1 신고구성기업 정보
SetRow(8, "1.1", "1", "최종모기업 여부", "B10", "Enum", "X",
    "GIR401=Ultimate Parent Entity, GIR402=Designated Filing Entity, GIR403=Designated Local Entity, GIR404=Constituent Entity, GIR405=Other",
    "FilingCeRoleEnumType");
SetRow(9, "1.1", "2", "신고구성기업 상호", "D10", "텍스트", "X", "-", "-");
SetRow(10, "1.1", "3", "납세자번호", "H10", "텍스트", "X", "-", "-");
SetRow(11, "1.1", "4", "신고구성기업 지위", "K10", "Enum", "X",
    "GIR401=Ultimate Parent Entity, GIR402=Designated Filing Entity, GIR403=Designated Local Entity, GIR404=Constituent Entity, GIR405=Other",
    "FilingCeRoleEnumType");
SetRow(12, "1.1", "5", "신고구성기업 소재지국", "M10", "ISO 국가코드", "X",
    "KR, JP, US 등 ISO 3166-1 Alpha-2", "CountryCodeType");
SetRow(13, "1.1", "6", "연례 자동정보 교환당사국", "O10", "ISO 국가코드", "O (쉼표구분)",
    "KR, JP, US 등 ISO 3166-1 Alpha-2", "CountryCodeType, 복수 시 쉼표로 구분 (예: KR, JP, US)");

// Row 14: empty

// 1.2.1 다국적기업그룹과 신고대상 사업연도
SetRow(15, "1.2.1", "1", "다국적기업그룹 명칭", "B15", "텍스트", "X", "-", "-");
SetRow(16, "1.2.1", "2", "신고대상 사업연도 개시일", "F15", "날짜 (YYYY-MM-DD)", "X", "-", "-");
SetRow(17, "1.2.1", "3", "신고대상 사업연도 종료일", "K15", "날짜 (YYYY-MM-DD)", "X", "-", "-");
SetRow(18, "1.2.1", "4", "수정신고서 여부", "O15", "Enum", "X",
    "GIR101=신규, GIR102=정정, GIR103=보고대상 없음", "MessageTypeIndicEnumType");

// Row 19: empty

// 1.2.2 다국적기업그룹 일반 회계정보
SetRow(20, "1.2.2", "1", "최종모기업 연결재무제표(유형)", "B18", "Enum", "X",
    "GIR501=Subparagraph a, GIR502=Subparagraph b, GIR503=Subparagraph c, GIR504=Subparagraph d",
    "FilingCeCofUpeEnumType");
SetRow(21, "1.2.2", "2", "최종모기업 연결재무제표 회계기준", "H18", "텍스트", "X", "-",
    "예: K-IFRS, IFRS, US-GAAP 등");
SetRow(22, "1.2.2", "3", "최종모기업 연결재무제표 통화", "M18", "ISO 통화코드", "X",
    "KRW, USD, EUR 등 ISO 4217", "CurrCodeType");

// Row 23: empty

// 1.3.1 최종모기업
SetRow(24, "1.3.1", "1", "최종모기업 소재지국", "J22", "ISO 국가코드", "X",
    "KR, JP, US 등 ISO 3166-1 Alpha-2", "CountryCodeType, 70010: 하나만 허용");
SetRow(25, "1.3.1", "2", "적용가능규칙", "J23", "Enum", "O (쉼표구분)",
    "GIR201=QIIR(타국만), GIR202=QIIR(자국+타국), GIR203=QUTPR, GIR204=QDMTT, GIR205=Not applicable",
    "IdTypeRulesEnumType");
SetRow(26, "1.3.1", "3", "최종모기업 상호", "J24", "텍스트", "X", "-", "-");
SetRow(27, "1.3.1", "4", "최종모기업 납세자번호", "J25", "텍스트", "O (쉼표구분)", "-",
    "60022: Role=GIR401이면 FilingCE TIN과 일치 필요");
SetRow(28, "1.3.1", "5", "신고접수국가 납세자번호", "J26", "텍스트", "X", "-",
    "issuedBy=KR 자동 부여");
SetRow(29, "1.3.1", "6", "글로벌최저한세 목적상 기업유형", "J27", "Enum", "O (쉼표구분)",
    "GIR301=Constituent Entity, GIR302=Flow-Through(Tax Transparent), GIR303=Flow-Through(Reverse Hybrid), GIR304=Hybrid Entity, GIR306=Main Entity, GIR310=Investment Entity, GIR311=Insurance Investment Entity",
    "IdTypeGloBeStatusEnumType, 70009: UPE에서 GIR305,307-309,312-315,317,318 불가");
SetRow(30, "1.3.1", "7", "제외기업 유형", "J28", "Enum", "X",
    "GIR601=Governmental Entity, GIR602=International Organisation, GIR603=Non-profit Organisation, GIR604=Pension Fund, GIR605=Investment Fund(UPE), GIR606=Real Estate Investment Vehicle(UPE)",
    "ExcludedUpeEnumType");
SetRow(31, "1.3.1", "8", "시행령 제103조제4항 적용 국가 해당 여부", "J29", "ISO 국가코드", "X",
    "KR, JP 등", "CountryCodeType");

// Freeze panes on row 1
sheet.SheetView.FreezeRows(1);

// Save
workbook.Save();
Console.WriteLine($"\nSaved successfully: {templatePath}");
Console.WriteLine("Done!");
