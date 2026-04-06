using ClosedXML.Excel;

var outputDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Resources", "sample"));
Directory.CreateDirectory(outputDir);

// ── File 1: sample_기본정보.xlsx — sheet "1.1~1.2" ──
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("1.1~1.2");

    ws.Cell("C3").Value = "2025-01-01";   // 사업연도 시작
    ws.Cell("C6").Value = "2025-12-31";   // 사업연도 종료
    ws.Cell("P3").Value = "테스트주식회사";  // 법인명
    ws.Cell("P5").Value = "123-45-67890"; // 사업자등록번호

    ws.Cell("B10").Value = "GIR401";          // 최종모기업 여부
    ws.Cell("D10").Value = "Test Corp";       // 신고구성기업 상호
    ws.Cell("H10").Value = "1234567890";      // 납세자번호
    ws.Cell("K10").Value = "GIR401";          // 신고구성기업 지위
    ws.Cell("M10").Value = "KR";              // 소재지국
    ws.Cell("O10").Value = "KR, JP, US";      // 교환당사국

    ws.Cell("B15").Value = "테스트그룹";       // 다국적기업그룹 명칭
    ws.Cell("F15").Value = "2025-01-01";      // 사업연도 개시일
    ws.Cell("K15").Value = "2025-12-31";      // 사업연도 종료일
    ws.Cell("O15").Value = "GIR101";          // 수정신고서 여부

    ws.Cell("B18").Value = "GIR501";          // 연결재무제표 유형
    ws.Cell("H18").Value = "K-IFRS";          // 회계기준
    ws.Cell("M18").Value = "KRW";             // 통화

    var path = Path.Combine(outputDir, "sample_기본정보.xlsx");
    wb.SaveAs(path);
    Console.WriteLine($"Created: {path}");
}

// ── File 2: sample_UPE.xlsx — sheet "1.3.1" ──
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("1.3.1");

    ws.Cell("J22").Value = "KR";                     // 소재지국
    ws.Cell("J23").Value = "GIR201";                 // 적용가능규칙
    ws.Cell("J24").Value = "Test Ultimate Parent";   // 상호
    ws.Cell("J25").Value = "1234567890";             // 납세자번호
    ws.Cell("J26").Value = "1234567890";             // 신고접수국가 납세자번호
    ws.Cell("J27").Value = "GIR301";                 // 기업유형
    ws.Cell("J28").Value = "GIR601";                 // 제외기업 유형
    ws.Cell("J29").Value = "KR";                     // 시행령

    var path = Path.Combine(outputDir, "sample_UPE.xlsx");
    wb.SaveAs(path);
    Console.WriteLine($"Created: {path}");
}

// ── File 3: sample_CE_001.xlsx — sheet "1.3.2.1" ──
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("1.3.2.1");

    ws.Cell("O5").Value  = "true";                    // 변동 여부
    ws.Cell("O6").Value  = "JP";                      // 소재지국
    ws.Cell("O7").Value  = "GIR201";                  // 적용가능규칙
    ws.Cell("O8").Value  = "Test Subsidiary Japan";   // 구성기업 상호
    ws.Cell("O9").Value  = "JP1234567890";            // 납세자번호
    ws.Cell("O10").Value = "JP1234567890";            // 신고접수국가 TIN
    ws.Cell("O11").Value = "GIR301";                  // 기업유형
    ws.Cell("O13").Value = "GIR802";                  // 소유지분 유형
    ws.Cell("O14").Value = "1234567890";              // 소유지분 보유기업 TIN
    ws.Cell("O15").Value = "100";                     // 소유지분 %
    ws.Cell("O16").Value = "";                        // 모기업 유형
    ws.Cell("O17").Value = "";                        // QIIR TIN
    ws.Cell("O18").Value = "";                        // 부분소유모기업 TIN
    ws.Cell("O19").Value = "false";                   // 해외진출 초기 특례
    ws.Cell("O20").Value = "";                        // 소유지분 합계
    ws.Cell("O21").Value = "";                        // UPE 소유지분 > 합계

    var path = Path.Combine(outputDir, "sample_CE_001.xlsx");
    wb.SaveAs(path);
    Console.WriteLine($"Created: {path}");
}

Console.WriteLine("Done — all sample files created.");
