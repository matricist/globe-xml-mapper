using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public enum SectionMode { Keyword, Fixed, SimpleRow }

    public record SectionDef(
        string DisplayName,
        string Group,           // "main" | "group" | "entity"
        string SheetName,
        SectionMode Mode,
        string Keyword    = "",
        int FixedStart    = 0,
        int FixedEnd      = 0,
        int FixedGap      = 0,
        int SimpleStartRow = 0
    );

    public static class BlockService
    {
        // 데이터 입력 열: O~R (15~18)
        private const int DATA_COL_START = 15;
        private const int DATA_COL_END   = 18;

        public static readonly SectionDef[] Sections =
        {
            new("최종모기업",      "main",   "최종모기업",    SectionMode.Keyword,   Keyword: "1.3.1"),
            new("기업구조 변동",   "main",   "그룹구조 변동", SectionMode.SimpleRow, SimpleStartRow: 6),
            new("정보 요약",       "main",   "요약",          SectionMode.SimpleRow, SimpleStartRow: 4),
            new("구성기업",        "group",  "그룹구조",      SectionMode.Keyword,   Keyword: "1.3.2.1"),
            new("제외기업",        "entity", "제외기업",      SectionMode.Fixed,     FixedStart: 3, FixedEnd: 6,  FixedGap: 2),
            new("국가별 적용면제", "entity", "적용면제",      SectionMode.Fixed,     FixedStart: 2, FixedEnd: 53, FixedGap: 2),
        };

        // ─────────────────────────────────────────────────────────────────────
        //  카운트
        // ─────────────────────────────────────────────────────────────────────

        public static int Count(string filePath, SectionDef def)
        {
            try
            {
                using var wb = new XLWorkbook(filePath);
                if (!wb.TryGetWorksheet(def.SheetName, out var ws)) return 1;
                return Count(ws, def);
            }
            catch { return 1; }
        }

        public static int Count(IXLWorksheet ws, SectionDef def) => def.Mode switch
        {
            SectionMode.Keyword   => CountKeyword(ws, def.Keyword),
            SectionMode.Fixed     => CountFixed(ws, def),
            SectionMode.SimpleRow => CountSimpleRows(ws, def.SimpleStartRow),
            _                     => 1,
        };

        private static int CountKeyword(IXLWorksheet ws, string keyword)
        {
            int count = 0;
            var last = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int r = 1; r <= last; r++)
            {
                var v = ws.Cell(r, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v) && v.Contains(keyword)) count++;
            }
            return Math.Max(count, 1);
        }

        private static int CountFixed(IXLWorksheet ws, SectionDef def)
        {
            var setSize  = def.FixedEnd - def.FixedStart + 1 + def.FixedGap;
            var lastRow  = ws.LastRowUsed()?.RowNumber() ?? 0;
            int count = 0;
            for (int n = 0; ; n++)
            {
                var blockStart = def.FixedStart + n * setSize;
                if (blockStart > lastRow + setSize) break;
                if (blockStart > lastRow && n > 0) break;
                count = n + 1;
            }
            return Math.Max(count, 1);
        }

        private static int CountSimpleRows(IXLWorksheet ws, int startRow)
        {
            int count = 0;
            var last = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int r = startRow; r <= last; r++)
            {
                for (int c = 2; c <= 20; c++)
                {
                    if (!string.IsNullOrEmpty(ws.Cell(r, c).GetString()))
                    { count++; break; }
                }
            }
            return Math.Max(count, 1);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  추가
        // ─────────────────────────────────────────────────────────────────────

        public static void Add(string filePath, SectionDef def)
        {
            ModifyWorkbook(filePath, wb =>
            {
                if (!wb.TryGetWorksheet(def.SheetName, out var ws))
                    throw new InvalidOperationException($"시트 '{def.SheetName}'를 찾을 수 없습니다.");

                switch (def.Mode)
                {
                    case SectionMode.Keyword:
                        AddKeywordBlock(ws, def.Keyword); break;
                    case SectionMode.Fixed:
                        AddFixedBlock(ws, def); break;
                    case SectionMode.SimpleRow:
                        AddSimpleRow(ws, def.SimpleStartRow); break;
                }
            });
        }

        private static void AddKeywordBlock(IXLWorksheet ws, string keyword)
        {
            var headers = FindHeaderRows(ws, keyword);
            if (headers.Count == 0) return;

            var lastHeader = headers[^1];
            var blockEnd   = FindBlockEnd(ws, lastHeader, keyword);
            var blockSize  = blockEnd - lastHeader + 1;
            var insertRow  = blockEnd + 2;   // +1 = 구분 행, +2 = 새 블록 시작

            ws.Row(insertRow).InsertRowsAbove(blockSize);
            ws.Range(lastHeader, 1, blockEnd, 20).CopyTo(ws.Cell(insertRow, 1));

            for (int r = insertRow; r <= insertRow + blockSize - 1; r++)
                for (int c = DATA_COL_START; c <= DATA_COL_END; c++)
                    ws.Cell(r, c).Clear(XLClearOptions.Contents);
        }

        private static void AddFixedBlock(IXLWorksheet ws, SectionDef def)
        {
            var blockSize = def.FixedEnd - def.FixedStart + 1;
            var setSize   = blockSize + def.FixedGap;
            var count     = CountFixed(ws, def);
            var gapStart  = def.FixedStart + count * setSize - def.FixedGap;
            var insertRow = gapStart + def.FixedGap;   // 새 블록 위치

            ws.Row(gapStart).InsertRowsAbove(def.FixedGap + blockSize);
            ws.Range(def.FixedStart, 1, def.FixedEnd, 20).CopyTo(ws.Cell(insertRow, 1));

            for (int r = insertRow; r <= insertRow + blockSize - 1; r++)
                for (int c = DATA_COL_START; c <= DATA_COL_END; c++)
                    ws.Cell(r, c).Clear(XLClearOptions.Contents);
        }

        private static void AddSimpleRow(IXLWorksheet ws, int startRow)
        {
            var last = ws.LastRowUsed()?.RowNumber() ?? startRow;
            int insertRow = startRow;
            for (int r = last; r >= startRow; r--)
            {
                for (int c = 2; c <= 20; c++)
                {
                    if (!string.IsNullOrEmpty(ws.Cell(r, c).GetString()))
                    { insertRow = r + 1; goto done; }
                }
            }
            done:
            ws.Row(insertRow).InsertRowsAbove(1);
            ws.Range(startRow, 1, startRow, 20).CopyTo(ws.Cell(insertRow, 1));
            for (int c = DATA_COL_START; c <= DATA_COL_END; c++)
                ws.Cell(insertRow, c).Clear(XLClearOptions.Contents);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  삭제
        // ─────────────────────────────────────────────────────────────────────

        public static bool Remove(string filePath, SectionDef def)
        {
            bool removed = false;
            ModifyWorkbook(filePath, wb =>
            {
                if (!wb.TryGetWorksheet(def.SheetName, out var ws))
                    throw new InvalidOperationException($"시트 '{def.SheetName}'를 찾을 수 없습니다.");

                if (Count(ws, def) <= 1) return;

                switch (def.Mode)
                {
                    case SectionMode.Keyword:   RemoveKeywordBlock(ws, def.Keyword); break;
                    case SectionMode.Fixed:     RemoveFixedBlock(ws, def); break;
                    case SectionMode.SimpleRow: RemoveSimpleRow(ws, def.SimpleStartRow); break;
                }
                removed = true;
            });
            return removed;
        }

        private static void RemoveKeywordBlock(IXLWorksheet ws, string keyword)
        {
            var headers = FindHeaderRows(ws, keyword);
            if (headers.Count <= 1) return;

            var lastHeader = headers[^1];
            var blockEnd   = FindBlockEnd(ws, lastHeader, keyword);
            ws.Rows(lastHeader - 1, blockEnd).Delete();  // 구분 행 포함 삭제
        }

        private static void RemoveFixedBlock(IXLWorksheet ws, SectionDef def)
        {
            var blockSize    = def.FixedEnd - def.FixedStart + 1;
            var setSize      = blockSize + def.FixedGap;
            var count        = CountFixed(ws, def);
            var lastGapStart = def.FixedStart + (count - 1) * setSize - def.FixedGap;
            var lastBlockEnd = lastGapStart + def.FixedGap + blockSize - 1;
            ws.Rows(lastGapStart, lastBlockEnd).Delete();
        }

        private static void RemoveSimpleRow(IXLWorksheet ws, int startRow)
        {
            var last = ws.LastRowUsed()?.RowNumber() ?? startRow;
            for (int r = last; r >= startRow; r--)
            {
                for (int c = 2; c <= 20; c++)
                {
                    if (!string.IsNullOrEmpty(ws.Cell(r, c).GetString()))
                    { ws.Row(r).Delete(); return; }
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  내부 유틸
        // ─────────────────────────────────────────────────────────────────────

        private static List<int> FindHeaderRows(IXLWorksheet ws, string keyword)
        {
            var result = new List<int>();
            var last = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int r = 1; r <= last; r++)
            {
                var v = ws.Cell(r, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v) && v.Contains(keyword)) result.Add(r);
            }
            return result;
        }

        private static int FindBlockEnd(IXLWorksheet ws, int startRow, string keyword)
        {
            var last = ws.LastRowUsed()?.RowNumber() ?? startRow;
            for (int r = startRow + 1; r <= last; r++)
            {
                var v = ws.Cell(r, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v) && v.Contains(keyword))
                    return r - 2;  // 구분 행(1) 제외
            }
            // 마지막 블록: 뒤쪽 빈 행 제거
            while (last > startRow)
            {
                bool empty = true;
                for (int c = 2; c <= 20; c++)
                {
                    if (!string.IsNullOrEmpty(ws.Cell(last, c).GetString()))
                    { empty = false; break; }
                }
                if (empty) last--; else break;
            }
            return last;
        }

        private static void ModifyWorkbook(string filePath, Action<XLWorkbook> action)
        {
            try
            {
                using var wb = new XLWorkbook(filePath);
                action(wb);
                wb.Save();
            }
            catch (System.IO.IOException ex) when (ex.Message.Contains("another process"))
            {
                throw new InvalidOperationException(
                    "파일이 Excel에서 열려 있습니다.\nExcel을 저장하고 닫은 후 다시 시도하세요.", ex);
            }
        }
    }
}
