using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// HTML 폼 JSON → template.xlsx 1.3.2.1 시트 역방향 기입.
    /// </summary>
    public static class XlsxFiller_1321
    {
        private const string CE_SHEET     = "그룹구조";
        private const string ATTACH_SHEET = "그룹구조 첨부";
        private const int    BLOCK_START  = 3;
        private const int    BLOCK_SIZE   = 18;
        private const int    BLOCK_GAP    = 2;
        private const int    COL_O        = 15;

        /// <summary>
        /// ce_data JSON 문자열을 받아 template.xlsx를 복사 후 기입, 저장 경로 반환.
        /// </summary>
        public static string Fill(string jsonData, string templatePath, string outputPath)
        {
            var ceArray = JsonSerializer.Deserialize<JsonElement[]>(jsonData)
                          ?? throw new ArgumentException("CE 데이터가 비어 있습니다.");

            File.Copy(templatePath, outputPath, overwrite: true);
            using var wb = new XLWorkbook(outputPath);

            if (!wb.TryGetWorksheet(CE_SHEET, out var ws))
                throw new InvalidOperationException($"시트 '{CE_SHEET}'를 찾을 수 없습니다.");

            // 2번째 CE 이상 — 행 블록 삽입 + 서식 복사
            for (int i = 1; i < ceArray.Length; i++)
            {
                var insertAt = BLOCK_START + i * (BLOCK_SIZE + BLOCK_GAP);
                var gapStart = BLOCK_START + (i - 1) * (BLOCK_SIZE + BLOCK_GAP) + BLOCK_SIZE + 1;
                ws.Row(gapStart).InsertRowsAbove(BLOCK_SIZE + BLOCK_GAP);

                for (int r = 0; r < BLOCK_SIZE; r++)
                {
                    var srcRow = ws.Row(BLOCK_START + r);
                    var dstRow = ws.Row(insertAt + r);
                    foreach (var cell in srcRow.Cells(1, 20))
                    {
                        var dst = dstRow.Cell(cell.Address.ColumnNumber);
                        dst.Style = cell.Style;
                        if (cell.Address.ColumnNumber < COL_O)
                            dst.Value = cell.Value;
                    }
                    dstRow.Height = srcRow.Height;
                }
            }

            // CE 값 기입
            for (int i = 0; i < ceArray.Length; i++)
                WriteCeBlock(ws, ceArray[i], i);

            // 별첨(소유지분) 기입
            if (wb.TryGetWorksheet(ATTACH_SHEET, out var attachWs))
            {
                for (int i = 0; i < ceArray.Length; i++)
                    WriteOwnership(attachWs, ceArray[i], i + 1);
            }

            wb.Save();
            return outputPath;
        }

        // ── CE 블록 값 기입 ───────────────────────────────────────────────────

        private static void WriteCeBlock(IXLWorksheet ws, JsonElement ce, int blockIdx)
        {
            var bStart = BLOCK_START + blockIdx * (BLOCK_SIZE + BLOCK_GAP);

            void SetO(int offset, string? val)
            {
                if (string.IsNullOrEmpty(val)) return;
                ws.Cell(bStart + offset, COL_O).Value = val;
            }

            if (ce.TryGetProperty("changeFlag", out var cf))
                SetO(1, cf.GetBoolean() ? "true" : "false");

            if (ce.TryGetProperty("id", out var id))
            {
                if (id.TryGetProperty("resCountryCode", out var rcc) && rcc.GetArrayLength() > 0)
                    SetO(2, rcc[0].GetString());
                if (id.TryGetProperty("rules", out var rules) && rules.GetArrayLength() > 0)
                    SetO(3, string.Join(",", EnumArray(rules)));
                if (id.TryGetProperty("name", out var name))
                    SetO(4, name.GetString());
                if (id.TryGetProperty("tin", out var tins) && tins.GetArrayLength() > 0)
                    SetO(5, tins[0].GetProperty("value").GetString());
                if (id.TryGetProperty("receivingTin", out var rTin))
                    SetO(6, rTin.GetString());
                if (id.TryGetProperty("globeStatus", out var gs) && gs.GetArrayLength() > 0)
                    SetO(7, string.Join(",", EnumArray(gs)));
            }

            // 별첨 참조: block 0은 template에 이미 있으므로 1 이상만 기입
            if (blockIdx > 0)
            {
                var attachRefRow = 3 + blockIdx * (BLOCK_SIZE + BLOCK_GAP) + 10;
                ws.Cell(attachRefRow, COL_O).Value = $"첨부{blockIdx + 1}";
            }

            if (ce.TryGetProperty("qiir", out var qiir))
            {
                SetO(12, GetStr(qiir, "popeIpe"));
                if (qiir.TryGetProperty("exception", out var ex) && ex.TryGetProperty("tin", out var exTin))
                    SetO(13, exTin.GetProperty("value").GetString());
                if (qiir.TryGetProperty("mopeIpe", out var mope) && mope.TryGetProperty("tin", out var mopeTin))
                    SetO(14, mopeTin.GetProperty("value").GetString());
            }

            if (ce.TryGetProperty("qutpr", out var qutpr))
            {
                if (qutpr.TryGetProperty("art93", out var art93))
                    SetO(15, art93.GetBoolean() ? "true" : "false");
                if (qutpr.TryGetProperty("aggOwnership", out var agg))
                    ws.Cell(bStart + 16, COL_O).Value = Math.Round(agg.GetDecimal() * 100, 2);
                if (qutpr.TryGetProperty("upeOwnership", out var upeOw))
                    SetO(17, upeOw.GetBoolean() ? "true" : "false");
            }
        }

        // ── 소유지분 기입 ─────────────────────────────────────────────────────

        private static void WriteOwnership(IXLWorksheet attachWs, JsonElement ce, int attachNum)
        {
            if (!ce.TryGetProperty("ownership", out var owList) || owList.GetArrayLength() == 0)
                return;

            int headerRow = FindAttachHeader(attachWs, attachNum);
            if (headerRow < 0)
            {
                headerRow = AppendAttachSection(attachWs, attachNum);
                if (headerRow < 0) return;
            }

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

        private static int FindAttachHeader(IXLWorksheet ws, int attachNum)
        {
            for (int r = 1; r <= 500; r++)
                if (ws.Cell(r, 2).GetString().Trim() == $"첨부{attachNum}") return r;
            return -1;
        }

        private static int AppendAttachSection(IXLWorksheet ws, int attachNum)
        {
            int ref1Row = FindAttachHeader(ws, 1);
            if (ref1Row < 0) return -1;

            int sec1End = ref1Row + 3;
            for (int r = ref1Row + 1; r <= ref1Row + 30; r++)
            {
                var v = ws.Cell(r, 2).GetString().Trim();
                if (v.StartsWith("첨부") && v != "첨부1") { sec1End = r - 1; break; }
                sec1End = r;
            }

            int sectionRows = sec1End - ref1Row + 2;
            int insertAt    = sec1End + 2;

            for (int r = 0; r < sectionRows - 1; r++)
            {
                var srcRow = ws.Row(ref1Row + r);
                var dstRow = ws.Row(insertAt + r);
                bool isHeader = r < 3;
                foreach (var cell in srcRow.Cells(1, 10))
                {
                    var dst = dstRow.Cell(cell.Address.ColumnNumber);
                    dst.Style = cell.Style;
                    if (isHeader)
                    {
                        if (r == 0 && cell.Address.ColumnNumber == 2 && cell.GetString().Trim() == "첨부1")
                            dst.Value = $"첨부{attachNum}";
                        else if (cell.Address.ColumnNumber <= 5)
                            dst.Value = cell.Value;
                    }
                }
                dstRow.Height = srcRow.Height;
            }
            return insertAt;
        }

        // ── 헬퍼 ─────────────────────────────────────────────────────────────

        private static string? GetStr(JsonElement el, string key)
            => el.TryGetProperty(key, out var v) ? v.GetString() : null;

        private static IEnumerable<string?> EnumArray(JsonElement arr)
        {
            foreach (var e in arr.EnumerateArray())
                yield return e.GetString();
        }
    }
}
