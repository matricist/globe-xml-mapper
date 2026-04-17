using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace GlobeMapper.Services
{
    /// <summary>
    /// Excel COM late-binding 래퍼.
    /// 시트 복제(CE/제외기업) + 시트 내 행 블록 반복(UPE) 지원.
    /// </summary>
    public class ExcelController : IDisposable
    {
        private dynamic _app;
        private dynamic _workbook;
        private bool _disposed;
        private bool _ownsApp;   // Open()으로 직접 생성한 경우만 true → Quit() 호출 여부 결정

        public const string MetaSheetName = "_META";

        public bool IsOpen
        {
            get
            {
                // _app.Visible 대신 _workbook.Name 접근으로 실제 열림 여부 확인.
                // Excel이 다이얼로그/작업 중일 때 Visible은 false가 아니지만 COM이 throw할 수 있음.
                try
                {
                    if (_workbook == null || _app == null) return false;
                    var _ = (string)_workbook.Name;
                    return true;
                }
                catch { return false; }
            }
        }

        public string GetActiveSheetName()
        {
            try { return (string)_app?.ActiveSheet?.Name; }
            catch { return null; }
        }

        public void ActivateSheet(object sheetNameOrIndex)
        {
            try { _workbook.Sheets[sheetNameOrIndex].Activate(); }
            catch { }
        }

        public void Open(string path)
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Excel이 설치되어 있지 않습니다.");

            _app = Activator.CreateInstance(excelType);
            _app.Visible = true;
            _ownsApp = true;

            dynamic workbooks = _app.Workbooks;
            _workbook = workbooks.Open(path);
            Marshal.ReleaseComObject(workbooks);

            dynamic firstSheet = _workbook.Sheets[1];
            EnsureMetaSheet();
            firstSheet.Activate();
            Marshal.ReleaseComObject(firstSheet);
        }

        // oleaut32.dll의 GetActiveObject 직접 호출 (Marshal.GetActiveObject는 .NET Core에서 제거됨)
        [System.Runtime.InteropServices.DllImport("oleaut32.dll")]
        private static extern void GetActiveObjectNative(
            ref Guid rclsid,
            IntPtr pvReserved,
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)]
            out object ppunk);

        // Excel.Application CLSID (버전 무관 공통)
        private static readonly Guid ExcelClsid = new Guid("00024500-0000-0000-C000-000000000046");

        /// <summary>
        /// 이미 열려 있는 Excel 인스턴스의 활성 워크북에 연결.
        /// 사용자가 직접 Excel에서 파일을 열고 이 메서드를 호출.
        /// </summary>
        public void AttachToActive()
        {
            try
            {
                var clsid = ExcelClsid;
                GetActiveObjectNative(ref clsid, IntPtr.Zero, out object obj);
                _app = obj;
            }
            catch (Exception ex) when (!(ex is InvalidOperationException))
            {
                throw new InvalidOperationException(
                    "실행 중인 Excel 인스턴스를 찾을 수 없습니다.\nExcel에서 파일을 먼저 열어주세요.", ex);
            }

            _workbook = _app.ActiveWorkbook;
            if (_workbook == null)
                throw new InvalidOperationException(
                    "열려 있는 Excel 통합문서가 없습니다.\nExcel에서 파일을 열고 다시 시도하세요.");

            _ownsApp = false;  // 외부 인스턴스에 연결 → Quit() 하지 않음
            EnsureMetaSheet();
        }

        public void CreateNew(string templatePath, string savePath)
        {
            System.IO.File.Copy(templatePath, savePath, true);
            Open(savePath);
        }

        public void Save() => _workbook?.Save();

        public string GetFilePathForMapping()
        {
            Save();
            return (string)_workbook.FullName;
        }

        public void CloseWithSavePrompt()
        {
            if (_workbook == null) return;
            bool doQuit = false;
            try
            {
                bool saved = (bool)_workbook.Saved;
                if (!saved)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "변경사항이 있습니다. 저장하시겠습니까?",
                        "저장 확인",
                        System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Cancel) return; // Quit 하지 않음
                    _workbook.Close(SaveChanges: result == System.Windows.Forms.DialogResult.Yes);
                }
                else
                {
                    _workbook.Close(SaveChanges: false);
                }
                doQuit = true;
            }
            catch { doQuit = true; }
            finally { if (doQuit) QuitApp(); }
        }

        /// <summary>
        /// COM RPC_E_CALL_REJECTED(0x80010001) 발생 시 최대 maxRetries회 재시도.
        /// Excel이 일시적으로 바쁜 상태일 때 COM 호출이 거부되는 것을 처리.
        /// </summary>
        private static void ComRetry(Action action, int maxRetries = 5, int delayMs = 150)
        {
            for (int i = 0; ; i++)
            {
                try { action(); return; }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80010001) && i < maxRetries)
                {
                    System.Threading.Thread.Sleep(delayMs);
                }
            }
        }

        /// <summary>
        /// 벌크 작업 전용 성능 스코프.
        /// ScreenUpdating/EnableEvents/Calculation을 끄고 복구 보장.
        /// 병합셀 많은 시트에서 5~10배 빠름.
        /// </summary>
        private IDisposable PerfScope()
        {
            return new PerfScopeImpl(_app);
        }

        private sealed class PerfScopeImpl : IDisposable
        {
            private const int xlCalculationManual = -4135;
            private const int xlCalculationAutomatic = -4105;

            private readonly dynamic _app;
            private readonly bool _prevScreenUpdating;
            private readonly bool _prevEnableEvents;
            private readonly int _prevCalculation;

            public PerfScopeImpl(dynamic app)
            {
                _app = app;
                try { _prevScreenUpdating = (bool)_app.ScreenUpdating; } catch { _prevScreenUpdating = true; }
                try { _prevEnableEvents   = (bool)_app.EnableEvents;   } catch { _prevEnableEvents   = true; }
                try { _prevCalculation    = (int)_app.Calculation;     } catch { _prevCalculation    = xlCalculationAutomatic; }

                try { _app.ScreenUpdating = false; } catch { }
                try { _app.EnableEvents   = false; } catch { }
                try { _app.Calculation    = xlCalculationManual; } catch { }
            }

            public void Dispose()
            {
                try { _app.Calculation    = _prevCalculation;    } catch { }
                try { _app.EnableEvents   = _prevEnableEvents;   } catch { }
                try { _app.ScreenUpdating = _prevScreenUpdating; } catch { }
            }
        }

        #region 시트 내 행 블록 반복 (1.3.1 UPE)

        /// <summary>
        /// B열에서 header 텍스트를 포함하는 행 번호 목록 반환.
        /// blockHeader 기반 탐지에 사용.
        /// </summary>
        private List<int> FindBlockHeaderRows(dynamic ws, string header)
        {
            var result = new List<int>();
            int lastRow;
            try { lastRow = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count; }
            catch { lastRow = 500; }

            for (int r = 1; r <= lastRow; r++)
            {
                string val = ws.Cells[r, 2].Value?.ToString()?.Trim();
                if (val != null && val.Contains(header))
                    result.Add(r);
            }
            return result;
        }

        /// <summary>
        /// 시트 내 행 블록을 복제하여 아래에 추가.
        /// blockHeader 지정 시 B열 헤더 스캔으로 마지막 블록 위치를 동적 탐지.
        /// blockHeader 미지정 시 _META blockCount 기반.
        /// </summary>
        public void AddRowBlock(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            int dataColStart = 15, int dataColEnd = 18, string blockHeader = null)
            => AddRowBlocks(sheetName, sourceStartRow, sourceEndRow, gap, 1, dataColStart, dataColEnd, blockHeader);

        /// <summary>
        /// 벌크 블록 추가 — count개의 블록을 한 번에 복제.
        /// 성능 최적화:
        ///  - ScreenUpdating/Calculation 정지 (PerfScope)
        ///  - FindBlockHeaderRows 1회만 호출 (시작점 계산)
        ///  - Insert() 대신 시트 끝에 직접 Copy (밀어낼 행 없음, 병합 셀 안전)
        ///  - 행 높이는 원본 1회만 읽고 배열 재사용
        ///  - ClearDataCells를 Range 단위로 일괄 처리
        /// </summary>
        public void AddRowBlocks(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            int count, int dataColStart = 15, int dataColEnd = 18, string blockHeader = null)
        {
            if (count <= 0) return;

            ComRetry(() =>
            {
                using var _ = PerfScope();

                dynamic ws = _workbook.Sheets[sheetName];
                var blockSize = sourceEndRow - sourceStartRow + 1;
                var setSize   = blockSize + gap;

                // 원본 블록 범위 (모든 추가의 소스)
                dynamic sourceRange = ws.Range[
                    ws.Cells[sourceStartRow, 1],
                    ws.Cells[sourceEndRow, 18]
                ];

                // 원본 행 높이 1회만 읽음
                var heights = new double[blockSize];
                for (int i = 0; i < blockSize; i++)
                {
                    try { heights[i] = (double)ws.Rows[sourceStartRow + i].RowHeight; }
                    catch { heights[i] = -1; }
                }

                // 첫 번째 추가 위치 계산
                int existingCount;
                int firstInsertRow;
                if (blockHeader != null)
                {
                    var headerRows = FindBlockHeaderRows(ws, blockHeader);
                    existingCount = headerRows.Count;
                    var lastStart = existingCount > 0 ? headerRows[existingCount - 1] : sourceStartRow;
                    var lastEnd   = lastStart + blockSize - 1;
                    firstInsertRow = lastEnd + 1 + gap;
                }
                else
                {
                    existingCount = GetMetaInt(sheetName, "blockCount", 1);
                    firstInsertRow = sourceEndRow + 1 + (existingCount - 1) * setSize + gap;
                }

                // count개 블록을 시트 끝에 순차 복사 (Insert 불필요 — 밀어낼 행 없음)
                for (int k = 0; k < count; k++)
                {
                    int insertRow = firstInsertRow + k * setSize;

                    dynamic destRange = ws.Range[
                        ws.Cells[insertRow, 1],
                        ws.Cells[insertRow + blockSize - 1, 18]
                    ];
                    sourceRange.Copy(destRange);

                    // 행 높이 적용
                    for (int i = 0; i < blockSize; i++)
                    {
                        if (heights[i] > 0)
                        {
                            try { ws.Rows[insertRow + i].RowHeight = heights[i]; } catch { }
                        }
                    }
                }

                // 데이터 셀 일괄 초기화 — 추가된 전체 범위를 Range.ClearContents로 한 번에
                int clearStart = firstInsertRow;
                int clearEnd   = firstInsertRow + count * setSize - gap - 1;
                ClearDataCellsBulk(ws, clearStart, clearEnd, dataColStart, dataColEnd);

                try { _app.CutCopyMode = false; } catch { }

                if (blockHeader == null)
                    SetMetaInt(sheetName, "blockCount", existingCount + count);
            });
        }

        /// <summary>
        /// 마지막 행 블록 삭제.
        /// blockHeader 지정 시 B열 헤더 스캔으로 마지막 블록 위치를 동적 탐지.
        /// </summary>
        public bool RemoveRowBlock(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            string blockHeader = null)
        {
            bool result = false;
            ComRetry(() =>
            {
            using var _ = PerfScope();
            dynamic ws = _workbook.Sheets[sheetName];
            var blockSize = sourceEndRow - sourceStartRow + 1;

            int count;
            int lastBlockStart, lastBlockEnd;

            if (blockHeader != null)
            {
                var headerRows = FindBlockHeaderRows(ws, blockHeader);
                count = headerRows.Count;
                if (count <= 1) return;
                lastBlockStart = headerRows[count - 1];
                lastBlockEnd = lastBlockStart + (sourceEndRow - sourceStartRow);
            }
            else
            {
                count = GetMetaInt(sheetName, "blockCount", 1);
                if (count <= 1) return;
                lastBlockStart = sourceEndRow + 1 + (count - 2) * (blockSize + gap) + gap;
                lastBlockEnd = lastBlockStart + blockSize - 1;
            }

            _app.DisplayAlerts = false;
            try
            {
                dynamic deleteRange = ws.Rows[$"{lastBlockStart - gap}:{lastBlockEnd}"];
                deleteRange.Delete();
            }
            finally
            {
                _app.DisplayAlerts = true;
            }

            if (blockHeader == null)
                SetMetaInt(sheetName, "blockCount", count - 1);
            result = true;
            }); // ComRetry
            return result;
        }

        /// <summary>
        /// 시트를 원래 상태로 초기화 (추가된 블록 모두 제거 + 데이터 초기화).
        /// blockHeader 지정 시 B열 헤더 스캔 사용.
        /// </summary>
        public void ResetSheet(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            int dataColStart = 15, int dataColEnd = 18, string blockHeader = null)
        {
            if (blockHeader != null)
            {
                // 헤더 탐색 없이 sourceEndRow+1 ~ lastUsed 전체 삭제.
                // 어떤 상태(블록이 내부에 삽입됐든, 헤더가 없든)에서도 확실히 초기화.
                dynamic ws = _workbook.Sheets[sheetName];
                int deleteFrom = sourceEndRow + 1;
                int lastUsed;
                try { lastUsed = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count - 1; }
                catch { lastUsed = deleteFrom + 300; }

                if (deleteFrom <= lastUsed)
                {
                    _app.DisplayAlerts = false;
                    try
                    {
                        dynamic deleteRange = ws.Rows[$"{deleteFrom}:{lastUsed}"];
                        deleteRange.Delete();
                    }
                    finally { _app.DisplayAlerts = true; }
                }
            }
            else
            {
                var count = GetMetaInt(sheetName, "blockCount", 1);
                if (count > 1)
                {
                    dynamic ws = _workbook.Sheets[sheetName];
                    var blockSize = sourceEndRow - sourceStartRow + 1;
                    var firstExtraRow = sourceEndRow + 1 + gap;
                    var lastRow = sourceEndRow + (count - 1) * (blockSize + gap);

                    _app.DisplayAlerts = false;
                    try
                    {
                        dynamic deleteRange = ws.Rows[$"{firstExtraRow}:{lastRow}"];
                        deleteRange.Delete();
                    }
                    finally { _app.DisplayAlerts = true; }
                }
                SetMetaInt(sheetName, "blockCount", 1);
            }

            // 첫 번째 블록 데이터 초기화
            dynamic sheet = _workbook.Sheets[sheetName];
            ClearDataCells(sheet, sourceStartRow, sourceEndRow, dataColStart, dataColEnd);
        }

        /// <summary>
        /// blockHeader 지정 시 B열 헤더 스캔으로 블록 수 반환, 미지정 시 _META 사용.
        /// </summary>
        public int GetRowBlockCount(string sheetName, int defaultCount = 1, string blockHeader = null)
        {
            if (blockHeader != null)
            {
                dynamic ws = _workbook.Sheets[sheetName];
                return FindBlockHeaderRows(ws, blockHeader).Count;
            }
            return GetMetaInt(sheetName, "blockCount", defaultCount);
        }

        private void ClearDataCells(dynamic ws, int startRow, int endRow,
            int dataColStart = 15, int dataColEnd = 18)
        {
            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = dataColStart; c <= dataColEnd; c++)
                {
                    dynamic cell = ws.Cells[r, c];
                    if (cell.MergeCells)
                    {
                        dynamic mergeArea = cell.MergeArea;
                        if ((int)mergeArea.Row == r && (int)mergeArea.Column == c)
                            mergeArea.ClearContents();
                    }
                    else
                    {
                        cell.ClearContents();
                    }
                }
            }
        }

        /// <summary>
        /// 데이터 영역 Range를 한 번에 ClearContents.
        /// 병합셀도 Range.ClearContents로 안전하게 처리됨 (각 병합 영역의 앵커 셀만 값 보유).
        /// 셀 단위 순회보다 수십~수백 배 빠름.
        /// </summary>
        private void ClearDataCellsBulk(dynamic ws, int startRow, int endRow,
            int dataColStart = 15, int dataColEnd = 18)
        {
            if (endRow < startRow) return;
            try
            {
                dynamic range = ws.Range[
                    ws.Cells[startRow, dataColStart],
                    ws.Cells[endRow, dataColEnd]
                ];
                range.ClearContents();
            }
            catch
            {
                // 폴백: 셀 단위 순회
                ClearDataCells(ws, startRow, endRow, dataColStart, dataColEnd);
            }
        }

        #endregion

        #region CE 블록

        private const int CE_BLOCK_START = 3;
        private const int CE_BLOCK_END = 21;  // row 3~21 = 19행 (기존 20은 오류)
        private const int CE_BLOCK_GAP = 2;

        /// <summary>
        /// CE 블록 추가: 행 블록 복제(헤더 기반).
        /// 소유지분은 블록 내 통합 셀(O12)에 인라인 입력.
        /// </summary>
        public void AddCeBlock(string ceSheetName) => AddCeBlocks(ceSheetName, 1);

        /// <summary>
        /// 벌크 CE 블록 추가.
        /// </summary>
        public void AddCeBlocks(string ceSheetName, int count)
        {
            AddRowBlocks(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP, count,
                blockHeader: "1.3.2.1");
        }

        /// <summary>
        /// 마지막 CE 블록 삭제.
        /// </summary>
        public bool RemoveCeBlock(string ceSheetName)
        {
            dynamic ws = _workbook.Sheets[ceSheetName];
            var count = FindBlockHeaderRows(ws, "1.3.2.1").Count;
            if (count <= 1) return false;

            RemoveRowBlock(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP,
                blockHeader: "1.3.2.1");
            return true;
        }

        /// <summary>
        /// CE 시트 초기화.
        /// </summary>
        public void ResetCeSheet(string ceSheetName)
        {
            ResetSheet(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP,
                blockHeader: "1.3.2.1");
        }

        public int GetCeBlockCount(string ceSheetName)
        {
            dynamic ws = _workbook.Sheets[ceSheetName];
            return FindBlockHeaderRows(ws, "1.3.2.1").Count;
        }

        #endregion

        #region 시트2 복합 블록 (3~23 + 26~54)

        // 시트2는 블록1(2~22) + 간격(23~24) + 블록2(25~53) = 총 52행이 한 세트
        private const int S2_BLOCK1_START = 2;
        private const int S2_BLOCK1_END = 22;
        private const int S2_GAP_ROWS = 2;  // 23~24행 (간격)
        private const int S2_BLOCK2_START = 25;
        private const int S2_BLOCK2_END = 53;
        private const int S2_TOTAL_SIZE = 52; // (23-3+1) + 2 + (54-26+1)
        private const int S2_INSERT_GAP = 2;  // 세트 간 간격

        public void AddSheet2Block(string sheetName) => AddSheet2Blocks(sheetName, 1);

        /// <summary>
        /// 벌크 Sheet2 블록 추가 (블록1 + gap + 블록2 = 52행 복합 세트).
        /// Insert 대신 시트 끝에 직접 Copy. 원본 행 높이 1회만 읽음.
        /// </summary>
        public void AddSheet2Blocks(string sheetName, int count)
        {
            if (count <= 0) return;

            ComRetry(() =>
            {
                using var _ = PerfScope();

                dynamic ws = _workbook.Sheets[sheetName];
                var existingCount = GetMetaInt(sheetName, "blockCount", 1);

                var block1Size = S2_BLOCK1_END - S2_BLOCK1_START + 1;
                var block2Size = S2_BLOCK2_END - S2_BLOCK2_START + 1;
                var setStride  = S2_TOTAL_SIZE + S2_INSERT_GAP;

                // 원본 범위 (모든 추가의 소스)
                dynamic src1 = ws.Range[ws.Cells[S2_BLOCK1_START, 1], ws.Cells[S2_BLOCK1_END, 18]];
                dynamic src2 = ws.Range[ws.Cells[S2_BLOCK2_START, 1], ws.Cells[S2_BLOCK2_END, 18]];

                // 원본 행 높이 1회만 읽음
                var heights1 = new double[block1Size];
                for (int i = 0; i < block1Size; i++)
                {
                    try { heights1[i] = (double)ws.Rows[S2_BLOCK1_START + i].RowHeight; }
                    catch { heights1[i] = -1; }
                }
                var heights2 = new double[block2Size];
                for (int i = 0; i < block2Size; i++)
                {
                    try { heights2[i] = (double)ws.Rows[S2_BLOCK2_START + i].RowHeight; }
                    catch { heights2[i] = -1; }
                }

                // 첫 번째 추가 위치
                int firstInsertRow = S2_BLOCK2_END + 1 + (existingCount - 1) * setStride + S2_INSERT_GAP;

                for (int k = 0; k < count; k++)
                {
                    int insertRow  = firstInsertRow + k * setStride;
                    int block2Start = insertRow + block1Size + S2_GAP_ROWS;

                    // 블록1 복사
                    dynamic dst1 = ws.Range[ws.Cells[insertRow, 1], ws.Cells[insertRow + block1Size - 1, 18]];
                    src1.Copy(dst1);
                    for (int i = 0; i < block1Size; i++)
                    {
                        if (heights1[i] > 0)
                        {
                            try { ws.Rows[insertRow + i].RowHeight = heights1[i]; } catch { }
                        }
                    }

                    // 블록2 복사
                    dynamic dst2 = ws.Range[ws.Cells[block2Start, 1], ws.Cells[block2Start + block2Size - 1, 18]];
                    src2.Copy(dst2);
                    for (int i = 0; i < block2Size; i++)
                    {
                        if (heights2[i] > 0)
                        {
                            try { ws.Rows[block2Start + i].RowHeight = heights2[i]; } catch { }
                        }
                    }

                    // 데이터 셀 초기화 (블록1, 블록2 각각)
                    ClearDataCellsBulk(ws, insertRow, insertRow + block1Size - 1);
                    ClearDataCellsBulk(ws, block2Start, block2Start + block2Size - 1);
                }

                try { _app.CutCopyMode = false; } catch { }
                SetMetaInt(sheetName, "blockCount", existingCount + count);
            });
        }

        public bool RemoveSheet2Block(string sheetName)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];

            // 마지막 세트의 시작 위치
            var lastSetStart = S2_BLOCK2_END + 1 + (count - 2) * (S2_TOTAL_SIZE + S2_INSERT_GAP) + S2_INSERT_GAP;
            var lastSetEnd = lastSetStart + S2_TOTAL_SIZE - 1;

            // 간격 포함 삭제
            _app.DisplayAlerts = false;
            try
            {
                dynamic deleteRange = ws.Rows[$"{lastSetStart - S2_INSERT_GAP}:{lastSetEnd}"];
                deleteRange.Delete();
            }
            finally { _app.DisplayAlerts = true; }

            SetMetaInt(sheetName, "blockCount", count - 1);
            return true;
        }

        public void ResetSheet2(string sheetName)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count > 1)
            {
                dynamic ws = _workbook.Sheets[sheetName];
                var firstExtraStart = S2_BLOCK2_END + 1 + S2_INSERT_GAP;
                var lastEnd = S2_BLOCK2_END + (count - 1) * (S2_TOTAL_SIZE + S2_INSERT_GAP) + S2_TOTAL_SIZE;

                _app.DisplayAlerts = false;
                try { ws.Rows[$"{firstExtraStart}:{lastEnd}"].Delete(); }
                finally { _app.DisplayAlerts = true; }
            }

            dynamic sheet = _workbook.Sheets[sheetName];
            ClearDataCells(sheet, S2_BLOCK1_START, S2_BLOCK1_END);
            ClearDataCells(sheet, S2_BLOCK2_START, S2_BLOCK2_END);
            SetMetaInt(sheetName, "blockCount", 1);
        }

        #endregion


        #region 메타 시트 관리

        // main_template 섹션→시트 매핑 (ControlPanel + _META용)
        internal static readonly (string section, string sheetName)[] SheetMap = new[]
        {
            ("1.1~1.2",     "다국적기업그룹 정보"),
            ("1.3.1",       "최종모기업"),
            ("1.3.2.1",     "그룹구조"),
            ("1.3.2.2",     "제외기업"),
            ("1.3.3",       "그룹구조 변동"),
            ("1.4",         "요약"),
            ("2",           "적용면제"),
            ("JurCal",      "국가별 계산"),
            ("EntityCe",    "구성기업 계산"),
            ("UTPR",        "UTPR 배분"),
        };

        private void EnsureMetaSheet()
        {
            dynamic meta = GetMetaSheet();

            if (meta == null)
            {
                dynamic lastSheet = _workbook.Sheets[_workbook.Sheets.Count];
                meta = _workbook.Sheets.Add(After: lastSheet);
                meta.Name = MetaSheetName;
                meta.Visible = -1; // xlSheetVeryHidden
                meta.Cells[1, 1] = "key";
                meta.Cells[1, 2] = "value";
            }

            // 기존 blockCount 값 보존
            var savedBlockCounts = new Dictionary<string, int>();
            var r = 2;
            while (true)
            {
                var k = (string)meta.Cells[r, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                if (k.StartsWith("blockCount:"))
                {
                    var name = k.Substring(11);
                    if (int.TryParse((string)meta.Cells[r, 2].Value?.ToString(), out var cnt))
                        savedBlockCounts[name] = cnt;
                }
                r++;
            }

            // 헤더 이후 전체 초기화
            for (var clearRow = 2; clearRow < r + 1; clearRow++)
            {
                meta.Cells[clearRow, 1] = "";
                meta.Cells[clearRow, 2] = "";
            }

            // blockCount: 항목 재작성 (보존된 값 또는 1)
            var row = 2;
            var blockSheets = new[] { "최종모기업", "그룹구조", "제외기업", "그룹구조 변동", "요약", "적용면제" };
            foreach (var name in blockSheets)
            {
                bool exists = false;
                try { var _ = _workbook.Sheets[name]; exists = true; } catch { }
                if (exists)
                {
                    meta.Cells[row, 1] = $"blockCount:{name}";
                    meta.Cells[row, 2] = savedBlockCounts.TryGetValue(name, out var cnt) ? cnt : 1;
                    row++;
                }
            }
        }

        private dynamic GetMetaSheet()
        {
            try { return _workbook.Sheets[MetaSheetName]; }
            catch { return null; }
        }

        private int GetMetaInt(string context, string key, int defaultValue)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return defaultValue;
            var fullKey = $"{key}:{context}";
            var row = 2;
            while (true)
            {
                string k = meta.Cells[row, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                if (k == fullKey)
                {
                    var val = meta.Cells[row, 2].Value;
                    return val != null ? Convert.ToInt32(val) : defaultValue;
                }
                row++;
            }
            return defaultValue;
        }

        private void SetMetaInt(string context, string key, int value)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return;
            var fullKey = $"{key}:{context}";
            var row = 2;
            while (true)
            {
                string k = meta.Cells[row, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                if (k == fullKey)
                {
                    meta.Cells[row, 2] = value;
                    return;
                }
                row++;
            }
            // 새 항목 추가
            row = FindMetaEmptyRow(meta);
            meta.Cells[row, 1] = fullKey;
            meta.Cells[row, 2] = value;
        }

        private int FindMetaEmptyRow(dynamic meta)
        {
            var row = 2;
            while (!string.IsNullOrEmpty(meta.Cells[row, 1].Value?.ToString()))
                row++;
            return row;
        }

        #endregion

        #region MappingOrchestrator용 메타 읽기

        public static int ReadBlockCount(ClosedXML.Excel.IXLWorksheet metaWs, string sheetName)
        {
            var key = $"blockCount:{sheetName}";
            var row = 2;
            while (true)
            {
                var k = metaWs.Cell(row, 1).GetString()?.Trim();
                if (string.IsNullOrEmpty(k)) break;
                if (k == key)
                {
                    var val = metaWs.Cell(row, 2).GetString()?.Trim();
                    return int.TryParse(val, out var n) ? n : 1;
                }
                row++;
            }
            return 1;
        }

        #endregion

        #region Dispose

        private void QuitApp()
        {
            // 순서 중요: 워크북 → (소유 시) Quit → 앱 → GC
            // GC.Collect 없이는 dynamic으로 생성된 중간 COM 객체(RCW)가 남아 Excel 프로세스가 살아있게 됨
            try
            {
                if (_workbook != null)
                {
                    Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }
                // AttachToActive()로 연결한 경우 Quit() 호출 금지 — 사용자 Excel을 닫지 않음
                if (_ownsApp)
                    _app?.Quit();
            }
            catch { }
            finally
            {
                if (_app != null)
                {
                    Marshal.ReleaseComObject(_app);
                    _app = null;
                }
                _ownsApp = false;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            CloseWithSavePrompt();
        }

        #endregion
    }
}
