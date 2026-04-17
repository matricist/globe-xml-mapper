using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class ControlPanelForm : Form
    {
        private readonly ExcelController _excel;
        private Label _lblCurrentSheet;
        private Panel _bodyPanel;
        private Panel _divider;
        private Panel _dynamicPanel;
        private Button btnToggle;
        private Timer _sheetTracker;
        private string _lastSheetName;
        private int _trackerFailCount;
        private Point _dragStart;
        private Point _formStart;
        private bool _dragging;
        private bool _collapsed;

        // ── 블록 설정: 페이지 추가가 필요한 시트만 ──────────────────────────
        // 벌크 상한(MaxBulkAdd)은 블록 크기 기준 보수적으로:
        //   가벼운 시트(9~21행) = 10, 적용면제(52행 복합) = 5,
        //   구성기업 계산(167행) = 5, 국가별 계산(259행) = 3
        private sealed record BlockConfig(
            string SheetName,
            string ItemName,
            int BlockStart,
            int BlockEnd,
            int BlockGap,
            string BlockHeader,
            int MaxBulkAdd,
            int DataColStart = 15,
            int DataColEnd = 18);

        private static readonly BlockConfig[] BlockConfigs =
        {
            new("최종모기업",     "최종모기업",     3,   11,  2, "1.3.1",   MaxBulkAdd: 10, DataColStart: 10, DataColEnd: 10),
            new("그룹구조",       "구성기업",       3,   21,  2, "1.3.2.1", MaxBulkAdd: 10),
            new("제외기업",       "제외기업",       2,   5,   2, null,      MaxBulkAdd: 10),
            new("적용면제",       "국가별 적용면제",2,   53,  2, null,      MaxBulkAdd: 5),   // 52행 복합 블록
            new("국가별 계산",    "합산단위",       2,   260, 2, "3.1 국가별 글로벌", MaxBulkAdd: 3),
            new("구성기업 계산",  "구성기업",       2,   167, 2, "3.2.4 구성기업", MaxBulkAdd: 5),
        };

        // 수량 TextBox (시트 전환 시 1로 리셋)
        private TextBox _bulkCountBox;

        private const int PAD     = 14;
        private const int PANEL_W = 500;
        private const int TITLE_H = 36;
        private const int ROW_H   = 44;

        public ControlPanelForm(ExcelController excel)
        {
            _excel = excel;
            InitializeComponent();
            StartSheetTracker();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            LayoutBodyPanel();
            if (_lastSheetName != null)
                UpdateDynamicPanel(_lastSheetName);
        }

        // DPI 스케일링 완료 후 body 요소 좌표 확정
        private void LayoutBodyPanel()
        {
            var w = ClientSize.Width - PAD * 2;
            _lblCurrentSheet.SetBounds(PAD, PAD, w, 32);
            _divider.SetBounds(PAD, PAD + 36, w, 1);
            _dynamicPanel.SetBounds(PAD, PAD + 36 + 1 + 6, w, 0);
        }

        private void InitializeComponent()
        {
            Text = "GIR 2 XML Mapper";
            FormBorderStyle = FormBorderStyle.None;
            TopMost = true;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = Color.FromArgb(37, 37, 38);
            ForeColor = Color.White;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(Screen.PrimaryScreen.WorkingArea.Right - PANEL_W - 30, 50);
            Size = new Size(PANEL_W, 200);

            // ── 타이틀바: TableLayoutPanel으로 글씨·버튼 잘림 방지 ────────
            var titleBar = new Panel
            {
                Dock = DockStyle.Top, Height = TITLE_H,
                BackColor = Color.FromArgb(28, 28, 28), Cursor = Cursors.SizeAll
            };
            titleBar.MouseDown += TitleDrag_Down;
            titleBar.MouseMove += TitleDrag_Move;
            titleBar.MouseUp   += (s, ev) => _dragging = false;

            var titleLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1,
                BackColor = Color.Transparent,
                Margin = Padding.Empty, Padding = new Padding(10, 0, 0, 0)
            };
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 44));
            titleLayout.MouseDown += TitleDrag_Down;
            titleLayout.MouseMove += TitleDrag_Move;
            titleLayout.MouseUp   += (s, ev) => _dragging = false;

            var lblTitle = new Label
            {
                Text = "GIR 2 XML Mapper", Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(220, 220, 220),
                Font = new Font("Segoe UI Semibold", 10),
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };
            lblTitle.MouseDown += TitleDrag_Down;
            lblTitle.MouseMove += TitleDrag_Move;
            lblTitle.MouseUp   += (s, ev) => _dragging = false;

            btnToggle = new Button
            {
                Text = "─", Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat, ForeColor = Color.Gray,
                BackColor = Color.FromArgb(50, 50, 50),
                Font = new Font("Segoe UI", 9),
                Margin = new Padding(2, 4, 4, 4)
            };
            btnToggle.FlatAppearance.BorderSize = 0;
            btnToggle.Click += (s, ev) =>
            {
                _collapsed = !_collapsed;
                _bodyPanel.Visible = !_collapsed;
                btnToggle.Text = _collapsed ? "□" : "─";
                ResizeToFit();
            };

            titleLayout.Controls.Add(lblTitle, 0, 0);
            titleLayout.Controls.Add(btnToggle, 1, 0);
            titleBar.Controls.Add(titleLayout);

            // ── 본문 패널: 절대좌표 (OnLoad에서 실제 크기로 조정) ─────────
            _bodyPanel = new Panel { Dock = DockStyle.Fill };

            _lblCurrentSheet = new Label
            {
                Text = "현재 시트: -",
                ForeColor = Color.FromArgb(86, 186, 240),
                Font = new Font("Segoe UI", 11f),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _divider = new Panel { BackColor = Color.FromArgb(60, 60, 60) };
            _dynamicPanel = new Panel { BackColor = Color.Transparent };

            _bodyPanel.Controls.Add(_lblCurrentSheet);
            _bodyPanel.Controls.Add(_divider);
            _bodyPanel.Controls.Add(_dynamicPanel);

            Controls.Add(_bodyPanel);
            Controls.Add(titleBar);
        }

        private void TitleDrag_Down(object s, MouseEventArgs e)
        {
            _dragging  = true;
            _dragStart = ((Control)s).PointToScreen(e.Location);
            _formStart = Location;
        }
        private void TitleDrag_Move(object s, MouseEventArgs e)
        {
            if (!_dragging) return;
            var cur = ((Control)s).PointToScreen(e.Location);
            Location = new Point(
                _formStart.X + cur.X - _dragStart.X,
                _formStart.Y + cur.Y - _dragStart.Y);
        }

        #region 시트별 동적 UI

        private void UpdateDynamicPanel(string sheetName)
        {
            _dynamicPanel.Controls.Clear();
            var y = 0;

            var cfg = BlockConfigs.FirstOrDefault(c => c.SheetName == sheetName);
            if (cfg != null)
            {
                y = RenderBlockPanel(cfg, y);
            }

            y += 10;
            y = AddActionButton("엑셀 종료하기", Color.FromArgb(48, 48, 52), y, BtnCloseExcel_Click);

            _dynamicPanel.Height = y + 4;
            ResizeToFit();
        }

        /// <summary>
        /// 페이지 추가가 필요한 6개 시트의 공통 UI 렌더링.
        /// [수량][+][-] + 상한 힌트 + 시트 초기화 버튼.
        /// </summary>
        private int RenderBlockPanel(BlockConfig cfg, int y)
        {
            var sheet = cfg.SheetName;
            var count = GetBlockCount(cfg);

            y = AddBulkSectionRow(cfg, $"{count}개", y,
                n => RunBulkAdd(cfg, n),
                n => RunBulkRemove(cfg, count, n));
            y += 4;
            y = AddActionButton("시트 초기화", Color.FromArgb(52, 52, 56), y, () =>
            {
                if (!Confirm("시트를 초기 상태로 되돌리시겠습니까?", true)) return;
                using var _ = new WaitCursorScope(this);
                ResetBlocks(cfg);
                UpdateDynamicPanel(sheet);
            });
            return y;
        }

        private int GetBulkCount(BlockConfig cfg)
        {
            if (_bulkCountBox == null) return 1;
            if (!int.TryParse(_bulkCountBox.Text.Trim(), out var n) || n < 1) n = 1;
            if (n > cfg.MaxBulkAdd) n = cfg.MaxBulkAdd;
            return n;
        }

        private void RunBulkAdd(BlockConfig cfg, int n)
        {
            if (n <= 0) return;
            using (new WaitCursorScope(this))
            {
                AddBlocks(cfg, n);
            }
            UpdateDynamicPanel(cfg.SheetName);
        }

        private void RunBulkRemove(BlockConfig cfg, int current, int n)
        {
            if (n <= 0) return;
            // 최소 1개 유지
            int maxRemove = Math.Max(0, current - 1);
            if (maxRemove == 0) { Warn("최소 1개는 유지해야 합니다."); return; }
            if (n > maxRemove)
            {
                Warn($"현재 {current}개 — 최대 {maxRemove}개까지만 삭제할 수 있습니다.");
                n = maxRemove;
            }
            if (n >= 2 && !Confirm($"{n}개를 삭제하시겠습니까?")) return;

            using (new WaitCursorScope(this))
            {
                for (int i = 0; i < n; i++) RemoveBlock(cfg);
            }
            UpdateDynamicPanel(cfg.SheetName);
        }

        private sealed class WaitCursorScope : IDisposable
        {
            private readonly Form _form;
            private readonly Cursor _prev;
            public WaitCursorScope(Form f) { _form = f; _prev = f.Cursor; f.Cursor = Cursors.WaitCursor; }
            public void Dispose() { try { _form.Cursor = _prev; } catch { } }
        }

        private int GetBlockCount(BlockConfig cfg) => cfg.SheetName switch
        {
            "그룹구조"   => _excel.GetCeBlockCount(cfg.SheetName),
            "적용면제"   => _excel.GetRowBlockCount(cfg.SheetName),
            _            => _excel.GetRowBlockCount(cfg.SheetName, blockHeader: cfg.BlockHeader),
        };

        private void AddBlocks(BlockConfig cfg, int count)
        {
            // 2차 방어: GetBulkCount에서 이미 클램프되지만, 혹시 다른 경로에서 호출돼도 상한 초과 불가
            if (count > cfg.MaxBulkAdd) count = cfg.MaxBulkAdd;
            if (count <= 0) return;

            switch (cfg.SheetName)
            {
                case "그룹구조":
                    _excel.AddCeBlocks(cfg.SheetName, count);
                    break;
                case "적용면제":
                    _excel.AddSheet2Blocks(cfg.SheetName, count);
                    break;
                default:
                    _excel.AddRowBlocks(cfg.SheetName, cfg.BlockStart, cfg.BlockEnd, cfg.BlockGap,
                        count, dataColStart: cfg.DataColStart, dataColEnd: cfg.DataColEnd,
                        blockHeader: cfg.BlockHeader);
                    break;
            }
        }

        private void RemoveBlock(BlockConfig cfg)
        {
            switch (cfg.SheetName)
            {
                case "그룹구조":
                    _excel.RemoveCeBlock(cfg.SheetName);
                    break;
                case "적용면제":
                    _excel.RemoveSheet2Block(cfg.SheetName);
                    break;
                default:
                    _excel.RemoveRowBlock(cfg.SheetName, cfg.BlockStart, cfg.BlockEnd, cfg.BlockGap,
                        blockHeader: cfg.BlockHeader);
                    break;
            }
        }

        private void ResetBlocks(BlockConfig cfg)
        {
            switch (cfg.SheetName)
            {
                case "그룹구조":
                    _excel.ResetCeSheet(cfg.SheetName);
                    break;
                case "적용면제":
                    _excel.ResetSheet2(cfg.SheetName);
                    break;
                default:
                    _excel.ResetSheet(cfg.SheetName, cfg.BlockStart, cfg.BlockEnd, cfg.BlockGap,
                        dataColStart: cfg.DataColStart, dataColEnd: cfg.DataColEnd,
                        blockHeader: cfg.BlockHeader);
                    break;
            }
        }

        #endregion

        #region UI 헬퍼

        // 벌크 섹션 행: [이름 ......] [N개] [수량][+][-]  (+ /max 힌트)
        // onAdd/onRemove는 수량(int)을 받음.
        private int AddBulkSectionRow(BlockConfig cfg, string countText, int y,
            Action<int> onAdd, Action<int> onRemove)
        {
            var w = _dynamicPanel.Width;
            const int BTN_W = 28;
            const int BTN_H = 30;
            const int BTN_GAP = 4;
            const int QTY_W = 40;
            const int HINT_W = 28;  // "/10" 힌트 폭
            const int CNT_W = 46;
            var btnY = y + (ROW_H - BTN_H) / 2;

            // 오른쪽부터 역순 배치: [hint] [-] [+] [qty] [count]
            int xHint = w - HINT_W;
            int xRem  = xHint - BTN_W;
            int xAdd  = xRem - BTN_W - BTN_GAP;
            int xQty  = xAdd - QTY_W - BTN_GAP;
            int xCnt  = xQty - CNT_W - 6;

            _dynamicPanel.Controls.Add(new Label
            {
                Text = cfg.ItemName,
                Bounds = new Rectangle(0, y, xCnt, ROW_H),
                ForeColor = Color.FromArgb(200, 200, 204),
                Font = new Font("Segoe UI", 10.5f),
                TextAlign = ContentAlignment.MiddleLeft
            });
            _dynamicPanel.Controls.Add(new Label
            {
                Text = countText,
                Bounds = new Rectangle(xCnt, y, CNT_W, ROW_H),
                ForeColor = Color.FromArgb(130, 130, 136),
                Font = new Font("Segoe UI", 10f),
                TextAlign = ContentAlignment.MiddleRight
            });

            // 수량 TextBox (항상 1로 시작, 상한 자동 클램프)
            // MaxLength를 상한 자릿수에 맞춰 제한해 상한 초과 입력 자체를 어렵게 함
            int maxLen = cfg.MaxBulkAdd.ToString().Length;
            _bulkCountBox = new TextBox
            {
                Text = "1",
                Bounds = new Rectangle(xQty, btnY + 1, QTY_W, BTN_H - 2),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.FromArgb(50, 50, 52),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 10f),
                MaxLength = maxLen,
            };
            _bulkCountBox.KeyPress += (s, e) =>
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                    e.Handled = true;
            };
            // 입력값이 상한 초과하면 즉시 TextBox 표시를 상한으로 교체 (사용자에게 시각적 피드백)
            _bulkCountBox.TextChanged += (s, e) =>
            {
                var t = _bulkCountBox.Text.Trim();
                if (int.TryParse(t, out var n) && n > cfg.MaxBulkAdd)
                {
                    _bulkCountBox.Text = cfg.MaxBulkAdd.ToString();
                    _bulkCountBox.SelectionStart = _bulkCountBox.Text.Length;
                }
            };
            _bulkCountBox.Leave += (s, e) =>
            {
                // 포커스 벗어날 때 0/빈값 → 1 정규화
                if (!int.TryParse(_bulkCountBox.Text.Trim(), out var n) || n < 1)
                    _bulkCountBox.Text = "1";
            };
            _bulkCountBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    onAdd(GetBulkCount(cfg));
                    e.SuppressKeyPress = true;
                }
            };

            var btnAdd = MakeIconButton("▲",
                Color.FromArgb(88, 190, 110),
                new Rectangle(xAdd, btnY, BTN_W, BTN_H),
                () => onAdd(GetBulkCount(cfg)));
            var btnRem = MakeIconButton("▼",
                Color.FromArgb(210, 80, 80),
                new Rectangle(xRem, btnY, BTN_W, BTN_H),
                () => onRemove(GetBulkCount(cfg)));

            _dynamicPanel.Controls.Add(new Label
            {
                Text = $"/{cfg.MaxBulkAdd}",
                Bounds = new Rectangle(xHint, y, HINT_W, ROW_H),
                ForeColor = Color.FromArgb(100, 100, 106),
                Font = new Font("Segoe UI", 9f),
                TextAlign = ContentAlignment.MiddleLeft,
            });

            _dynamicPanel.Controls.Add(_bulkCountBox);
            _dynamicPanel.Controls.Add(btnAdd);
            _dynamicPanel.Controls.Add(btnRem);
            return y + ROW_H + 2;
        }

        private Button MakeIconButton(string text, Color iconColor, Rectangle bounds, Action click)
        {
            var btn = new Button
            {
                Text = text, Bounds = bounds,
                FlatStyle = FlatStyle.Flat,
                ForeColor = iconColor,
                BackColor = Color.FromArgb(52, 52, 56),
                Font = new Font("Segoe UI", 7.5f),
                Padding = Padding.Empty
            };
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = Color.FromArgb(68, 68, 74);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(64, 64, 70);
            btn.Click += (s, e) => { try { click(); } catch (Exception ex) { MessageBox.Show(ex.Message); } };
            return btn;
        }

        private int AddActionButton(string text, Color bgColor, int y, Action click)
        {
            var btn = new Button
            {
                Text = text,
                Bounds = new Rectangle(0, y, _dynamicPanel.Width, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.FromArgb(190, 190, 194),
                Font = new Font("Segoe UI", 9f)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                Math.Min(bgColor.R + 12, 255),
                Math.Min(bgColor.G + 12, 255),
                Math.Min(bgColor.B + 12, 255));
            btn.Click += (s, e) => { try { click(); } catch (Exception ex) { MessageBox.Show(ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); } };
            _dynamicPanel.Controls.Add(btn);
            return y + 40;
        }

        private void ResizeToFit()
        {
            if (_collapsed) { Height = TITLE_H; return; }
            // TITLE_H + PAD + 32(라벨) + 36(간격+구분선) + _dynamicPanel.Height + PAD
            var h = TITLE_H + PAD + 32 + 6 + 1 + 6 + _dynamicPanel.Height + PAD;
            Height = Math.Max(h, 160);
        }

        private static void Warn(string msg) =>
            MessageBox.Show(msg, "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        private static bool Confirm(string msg, bool warning = false) =>
            MessageBox.Show(msg, "확인", MessageBoxButtons.YesNo,
                warning ? MessageBoxIcon.Warning : MessageBoxIcon.Question) == DialogResult.Yes;

        #endregion

        #region 시트 트래커 (2초 폴링)

        private void StartSheetTracker()
        {
            _sheetTracker = new Timer { Interval = 2000 };
            _sheetTracker.Tick += (s, e) =>
            {
                // IsOpen 실패만 종료 카운트에 반영.
                // 시트명 조회 등 이후 COM 실패는 Excel이 바쁜 상태일 뿐이므로 무시.
                if (!_excel.IsOpen)
                {
                    _trackerFailCount++;
                    // 2초 × 6 = 12초 연속 실패 시 종료
                    if (_trackerFailCount >= 6) { _sheetTracker.Stop(); Close(); }
                    return;
                }
                _trackerFailCount = 0;

                try
                {
                    var current = _excel.GetActiveSheetName();
                    if (current == ExcelController.MetaSheetName)
                    {
                        _excel.ActivateSheet(_lastSheetName != null ? (object)_lastSheetName : 1);
                        return;
                    }
                    if (current != null && current != _lastSheetName)
                    {
                        _lastSheetName = current;
                        _lblCurrentSheet.Text = $"현재 시트: {current}";
                        UpdateDynamicPanel(current);
                    }
                }
                catch { /* Excel 일시적으로 바쁜 상태 — 무시 */ }
            };
            _sheetTracker.Start();

            // 초기 시트명 기록 (UpdateDynamicPanel은 OnLoad에서 실행)
            var initial = _excel.GetActiveSheetName();
            if (initial != null)
            {
                _lastSheetName = initial;
                _lblCurrentSheet.Text = $"현재 시트: {initial}";
            }
        }

        #endregion

        #region 액션

        private void BtnConvert_Click()
        {
            using var terms = new TermsDialog();
            if (terms.ShowDialog(this) != DialogResult.OK) return;

            using var dlg = new SaveFileDialog
            {
                Filter = "XML 파일 (*.xml)|*.xml",
                Title = "XML 파일 저장",
                FileName = "GLOBE_OECD.xml"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            var filePath = _excel.GetFilePathForMapping();
            var globe = new Globe.GlobeOecd
            {
                Version = "2.0",
                MessageSpec = new Globe.MessageSpecType(),
                GlobeBody = new Globe.GlobeBodyType()
            };

            var orchestrator = new MappingOrchestrator();
            var mappingErrors = orchestrator.MapWorkbook(filePath, globe);
            var xml = XmlExportService.Serialize(globe);
            File.WriteAllText(dlg.FileName, xml, System.Text.Encoding.UTF8);

            var validationErrors = ValidationUtil.Validate(globe);
            var allErrors = new List<string>();
            if (mappingErrors.Count > 0) { allErrors.Add("── 매핑 오류 ──"); allErrors.AddRange(mappingErrors); allErrors.Add(""); }
            if (validationErrors.Count > 0) { allErrors.Add("── 검증 오류 (에러코드 기준) ──"); allErrors.AddRange(validationErrors); }

            var errorsPath = Path.ChangeExtension(dlg.FileName, ".errors.txt");
            if (allErrors.Count > 0)
            {
                File.WriteAllText(errorsPath,
                    $"[오류 목록] {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                    $"매핑 오류 {mappingErrors.Count}건 / 검증 오류 {validationErrors.Count}건{Environment.NewLine}{Environment.NewLine}" +
                    string.Join(Environment.NewLine, allErrors),
                    System.Text.Encoding.UTF8);
                MessageBox.Show($"XML 생성 완료.\n\n매핑 오류: {mappingErrors.Count}건\n검증 오류: {validationErrors.Count}건\n\n오류 목록: {errorsPath}",
                    "완료 (오류 있음)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (File.Exists(errorsPath)) File.Delete(errorsPath);
                MessageBox.Show("XML 생성이 완료되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnCloseExcel_Click()
        {
            _sheetTracker?.Stop();
            _excel.CloseWithSavePrompt();
            Close();
        }

        #endregion

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _sheetTracker?.Stop();
            _sheetTracker?.Dispose();
            base.OnFormClosing(e);
        }
    }
}
