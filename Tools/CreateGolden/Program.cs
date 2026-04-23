using System;
using System.IO;
using ClosedXML.Excel;

// main_template.xlsx → qa/golden_sample.xlsx
// XSD 통과를 위한 최소 필수 데이터 채움.
// 코드 변경 없이 데이터만 보강.

class Program
{
    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var src = Path.Combine(repo, "Resources", "main_template.xlsx");
        var qaDir = Path.Combine(repo, "qa");
        Directory.CreateDirectory(qaDir);
        var dst = Path.Combine(qaDir, "golden_sample.xlsx");

        if (!File.Exists(src))
        {
            Console.Error.WriteLine($"Not found: {src}");
            return 1;
        }

        File.Copy(src, dst, overwrite: true);
        Console.WriteLine($"Copied: {Path.GetFileName(src)} → qa/golden_sample.xlsx");

        using var wb = new XLWorkbook(dst);

        // ── 1. 다국적기업그룹 정보 시트 ──
        var ws1 = wb.Worksheet("다국적기업그룹 정보");

        // D9: 신고구성기업 상호 (FilingCe.Name) — 비어 있으면 채움
        FillIfEmpty(ws1, "D9", "테스트구성기업", "FilingCe.Name");
        // H9: 신고구성기업 TIN (FilingCe.Tin.Value)
        FillIfEmpty(ws1, "H9", "1234567890", "FilingCe.Tin.Value");
        // O9: 연례 자동정보 교환당사국 (GeneralSection.RecJurCode, multi)
        FillIfEmpty(ws1, "O9", "KR", "GeneralSection.RecJurCode");
        // B14: 다국적기업그룹 명칭 (NameMne)
        FillIfEmpty(ws1, "B14", "테스트다국적기업그룹", "NameMne");
        // O14: 수정신고서 여부 (MessageTypeIndic) — 잘못된 값 "부" 가 들어 있어 GIR101로 정정
        var o14 = ws1.Cell("O14").GetString().Trim();
        if (o14 == "부" || o14 == "여" || string.IsNullOrEmpty(o14))
        {
            ws1.Cell("O14").Value = "GIR101";
            Console.WriteLine($"  [수정] O14 (MessageTypeIndic): '{o14}' → 'GIR101'");
        }

        // ── 2. 그룹구조 시트 — 첫 CE 블록 ──
        // 블록 헤더 "1.3.2.1" 행 3 기준, O열(15) 입력:
        //   +4=Name(O7), +5=TIN(O8), +7=GlobeStatus(O10), +8=Ownership통합(O11)
        var wsCs = wb.Worksheet("그룹구조");
        FillIfEmpty(wsCs, "O5", "KR", "CE.Id.ResCountryCode");
        FillIfEmpty(wsCs, "O6", "GIR201", "CE.Id.Rules");
        FillIfEmpty(wsCs, "O7", "테스트구성기업CE", "CE.Id.Name");
        FillIfEmpty(wsCs, "O8", "9876543210", "CE.Id.Tin.Value");
        FillIfEmpty(wsCs, "O10", "GIR301", "CE.Id.GlobeStatus");
        // O11 (병합 O11:R14): "유형,TIN,TIN유형,발급국가,지분"
        FillIfEmpty(wsCs, "O11", "GIR801, 1234567890, GIR3001, KR, 1", "CE.Ownership");

        // ── 3. 최종모기업 시트 — 첫 UPE 블록 ──
        // 블록 헤더 "1.3.1" 행 3, J열(10) 입력
        var wsUpe = wb.Worksheet("최종모기업");
        FillIfEmpty(wsUpe, "J4", "KR", "UPE.Id.ResCountryCode");
        FillIfEmpty(wsUpe, "J5", "GIR201", "UPE.Id.Rules");
        FillIfEmpty(wsUpe, "J6", "테스트최종모기업", "UPE.Id.Name");
        FillIfEmpty(wsUpe, "J7", "1234567890", "UPE.Id.Tin.Value");
        FillIfEmpty(wsUpe, "J9", "GIR301", "UPE.Id.GlobeStatus");

        // 주의: 구성기업 계산 / 국가별 계산 / UTPR 배분 시트는 채우지 않음.
        //   채우면 그 안의 모든 [R] 필드 (FANIL, AdjustedFANIL, ETR 등 수십개) 도 채워야
        //   XSD 통과. 또한 Excel 계산식에서 #DIV/0! 같은 결과가 발생해 매핑 실패 가능.
        //   골든 샘플은 "최소 통과 케이스" — 완전한 사례는 사용자가 실제 데이터로 입력.

        wb.SaveAs(dst);
        Console.WriteLine($"Saved: {dst}");
        return 0;
    }

    static void FillIfEmpty(IXLWorksheet ws, string addr, string value, string label)
    {
        var cell = ws.Cell(addr);
        var current = cell.GetString().Trim();
        if (string.IsNullOrEmpty(current))
        {
            cell.Value = value;
            Console.WriteLine($"  [채움] {ws.Name}!{addr} ({label}): '{value}'");
        }
        else
        {
            Console.WriteLine($"  [유지] {ws.Name}!{addr} ({label}): 이미 '{current}'");
        }
    }
}
