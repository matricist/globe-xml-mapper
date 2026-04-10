using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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

        private const int UPE_BLOCK_START = 3;
        private const int UPE_BLOCK_END   = 11;
        private const int UPE_BLOCK_GAP   = 2;
        private const int EX_BLOCK_START  = 2;
        private const int EX_BLOCK_END    = 5;
        private const int EX_BLOCK_GAP    = 2;
        private const string ATTACH_SHEET_NAME = "그룹구조 첨부";

        private int _selectedAttachNum  = 1;
        private int _selectedSheet3Page = 1;

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

            if (sheetName == "최종모기업")
            {
                var count = _excel.GetRowBlockCount(sheetName, blockHeader: "1.3.1");
                y = AddSectionRow("최종모기업", $"{count}개", y,
                    () => { _excel.AddRowBlock(sheetName, UPE_BLOCK_START, UPE_BLOCK_END, UPE_BLOCK_GAP, dataColStart: 10, dataColEnd: 10, blockHeader: "1.3.1"); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (count <= 1) { Warn("최소 1개는 유지해야 합니다."); return; }
                        if (!Confirm("마지막 최종모기업을 삭제하시겠습니까?")) return;
                        _excel.RemoveRowBlock(sheetName, UPE_BLOCK_START, UPE_BLOCK_END, UPE_BLOCK_GAP, blockHeader: "1.3.1");
                        UpdateDynamicPanel(sheetName);
                    });
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(52, 52, 56), y, () =>
                {
                    if (!Confirm("시트를 초기 상태로 되돌리시겠습니까?", true)) return;
                    _excel.ResetSheet(sheetName, UPE_BLOCK_START, UPE_BLOCK_END, UPE_BLOCK_GAP, dataColStart: 10, dataColEnd: 10, blockHeader: "1.3.1");
                    UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "그룹구조 첨부")
            {
                var ceCount = _excel.GetCeBlockCount("그룹구조");
                if (_selectedAttachNum < 1 || _selectedAttachNum > ceCount) _selectedAttachNum = 1;

                y = AddNumberSelector(y, "첨부 번호:", ceCount,
                    () => _selectedAttachNum, n => _selectedAttachNum = n, sheetName);
                y += 4;

                var ownerCount = _excel.GetOwnerRowCount(sheetName, _selectedAttachNum);
                y = AddSectionRow($"첨부{_selectedAttachNum} 주주", $"{ownerCount}행", y,
                    () => { _excel.AddOwnerRow(sheetName, _selectedAttachNum); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (ownerCount <= 0) { Warn("삭제할 행이 없습니다."); return; }
                        _excel.RemoveOwnerRow(sheetName, _selectedAttachNum); UpdateDynamicPanel(sheetName);
                    });
            }
            else if (sheetName == "3.4.1 첨부")
            {
                var ownerCount = _excel.GetOwnerRowCount(sheetName, 1);
                y = AddSectionRow("주주 목록", $"{ownerCount}행", y,
                    () => { _excel.AddOwnerRow(sheetName, 1); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (ownerCount <= 0) { Warn("삭제할 행이 없습니다."); return; }
                        _excel.RemoveOwnerRow(sheetName, 1); UpdateDynamicPanel(sheetName);
                    });
            }
            else if (sheetName == "그룹구조")
            {
                var count = _excel.GetCeBlockCount(sheetName);
                y = AddSectionRow("구성기업", $"{count}개", y,
                    () => { _excel.AddCeBlock(sheetName, ATTACH_SHEET_NAME); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (count <= 1) { Warn("최소 1개는 유지해야 합니다."); return; }
                        if (!Confirm("마지막 구성기업을 삭제하시겠습니까?")) return;
                        _excel.RemoveCeBlock(sheetName, ATTACH_SHEET_NAME); UpdateDynamicPanel(sheetName);
                    });
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(52, 52, 56), y, () =>
                {
                    if (!Confirm("시트를 초기 상태로 되돌리시겠습니까?\n모든 구성기업 및 첨부 데이터가 삭제됩니다.", true)) return;
                    _excel.ResetCeSheet(sheetName, ATTACH_SHEET_NAME);
                    UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "제외기업")
            {
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionRow("제외기업", $"{count}개", y,
                    () => { _excel.AddRowBlock(sheetName, EX_BLOCK_START, EX_BLOCK_END, EX_BLOCK_GAP); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (count <= 1) { Warn("최소 1개는 유지해야 합니다."); return; }
                        if (!Confirm("마지막 제외기업을 삭제하시겠습니까?")) return;
                        _excel.RemoveRowBlock(sheetName, EX_BLOCK_START, EX_BLOCK_END, EX_BLOCK_GAP);
                        UpdateDynamicPanel(sheetName);
                    });
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(52, 52, 56), y, () =>
                {
                    if (!Confirm("시트를 초기 상태로 되돌리시겠습니까?", true)) return;
                    _excel.ResetSheet(sheetName, EX_BLOCK_START, EX_BLOCK_END, EX_BLOCK_GAP);
                    UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "적용면제")
            {
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionRow("국가별 적용면제", $"{count}개", y,
                    () =>
                    {
                        _excel.AddSheet2Block(sheetName);
                        _excel.AddSheet2AttachPage(count + 1); // 새 첨부N 섹션 추가
                        UpdateDynamicPanel(sheetName);
                    },
                    () => {
                        if (count <= 1) { Warn("최소 1개는 유지해야 합니다."); return; }
                        if (!Confirm("마지막 페이지를 삭제하시겠습니까?")) return;
                        _excel.RemoveSheet2Block(sheetName);
                        _excel.RemoveSheet2AttachPage(count); // 마지막 첨부 섹션 삭제
                        UpdateDynamicPanel(sheetName);
                    });
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(52, 52, 56), y, () =>
                {
                    if (!Confirm("시트를 초기 상태로 되돌리시겠습니까?", true)) return;
                    _excel.ResetSheet2(sheetName); UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "요약")
            {
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionRow("정보 요약", $"{count}행", y,
                    () => { _excel.AddSimpleRowByMeta(sheetName, 4); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (count <= 1) { Warn("최소 1행은 유지해야 합니다."); return; }
                        _excel.RemoveSimpleRowByMeta(sheetName, 4); UpdateDynamicPanel(sheetName);
                    });
            }
            else if (sheetName == "그룹구조 변동")
            {
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionRow("기업구조 변동", $"{count}행", y,
                    () => { _excel.AddSimpleRowByMeta(sheetName, 6); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (count <= 1) { Warn("최소 1행은 유지해야 합니다."); return; }
                        _excel.RemoveSimpleRowByMeta(sheetName, 6); UpdateDynamicPanel(sheetName);
                    });
            }
            else if (sheetName == "3.1~3.2.3.2")
            {
                var pageCount = _excel.GetRowBlockCount(sheetName);
                y = AddSectionRow("페이지", $"{pageCount}개", y,
                    () => { _excel.AddSheet3Page(sheetName); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (pageCount <= 1) { Warn("최소 1페이지는 유지해야 합니다."); return; }
                        if (!Confirm("마지막 페이지를 삭제하시겠습니까?")) return;
                        _excel.RemoveSheet3Page(sheetName); UpdateDynamicPanel(sheetName);
                    });
                y += 8;

                if (_selectedSheet3Page > pageCount) _selectedSheet3Page = 1;
                y = AddNumberSelector(y, "페이지 번호:", pageCount,
                    () => _selectedSheet3Page, n => _selectedSheet3Page = n, sheetName);
                y += 8;

                var pk = $"p{_selectedSheet3Page}";

                var cfcCount = _excel.GetSheet3RowCount(sheetName, $"{pk}:cfc");
                y = AddSectionRow("통합형피지배", $"{cfcCount}행", y,
                    () => { _excel.AddSheet3Row(sheetName, $"{pk}:cfc", 97); UpdateDynamicPanel(sheetName); },
                    () => { if (cfcCount <= 1) { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, $"{pk}:cfc", 97); UpdateDynamicPanel(sheetName); });

                var cbCount = _excel.GetSheet3RowCount(sheetName, $"{pk}:carryback");
                y = AddSectionRow("결손금 소급공제", $"{cbCount}행", y,
                    () => { _excel.AddSheet3Row(sheetName, $"{pk}:carryback", 140); UpdateDynamicPanel(sheetName); },
                    () => { if (cbCount <= 1) { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, $"{pk}:carryback", 140); UpdateDynamicPanel(sheetName); });

                var artCount = _excel.GetSheet3RowCount(sheetName, $"{pk}:art89");
                y = AddSectionRow("제89조", $"{artCount}행", y,
                    () => { _excel.AddSheet3Row(sheetName, $"{pk}:art89", 170); UpdateDynamicPanel(sheetName); },
                    () => { if (artCount <= 1) { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, $"{pk}:art89", 170); UpdateDynamicPanel(sheetName); });

                y += 8;
                y = AddActionButton("시트 초기화", Color.FromArgb(52, 52, 56), y, () =>
                {
                    if (!Confirm("시트를 초기 상태로 되돌리시겠습니까?", true)) return;
                    _excel.ResetSheet3(sheetName); UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "3.2.4~3.2.4.5")
            {
                // 각 섹션 행 수 조회
                var grpCount    = _excel.GetSheet3RowCount(sheetName, "grp");
                var branchCount = _excel.GetSheet3RowCount(sheetName, "branch");
                var crossCount  = _excel.GetSheet3RowCount(sheetName, "cross");
                var upeCount    = _excel.GetSheet3RowCount(sheetName, "upe");
                var taxCount    = _excel.GetSheet3RowCount(sheetName, "tax");
                var fairCount   = _excel.GetSheet3RowCount(sheetName, "fair");
                var distCount   = _excel.GetSheet3RowCount(sheetName, "dist");
                var otherCount  = _excel.GetSheet3RowCount(sheetName, "other");

                // 앞 섹션 추가행 누적 반영 → 실제 첫 행 계산
                int r0 = 6;
                int r1 = 48  + (grpCount - 1);
                int r2 = 55  + (grpCount - 1) + (branchCount - 1);
                int r3 = 62  + (grpCount - 1) + (branchCount - 1) + (crossCount - 1);
                int r4 = 94  + (grpCount - 1) + (branchCount - 1) + (crossCount - 1) + (upeCount - 1);
                int r5 = 137 + (grpCount - 1) + (branchCount - 1) + (crossCount - 1) + (upeCount - 1) + (taxCount - 1);
                int r6 = 161 + (grpCount - 1) + (branchCount - 1) + (crossCount - 1) + (upeCount - 1) + (taxCount - 1) + (fairCount - 1);
                int r7 = 168 + (grpCount - 1) + (branchCount - 1) + (crossCount - 1) + (upeCount - 1) + (taxCount - 1) + (fairCount - 1) + (distCount - 1);

                y = AddSectionRow("(b) 연결납세그룹 통합신고",   $"{grpCount}행",    y,
                    () => { _excel.AddSheet3Row(sheetName, "grp",    r0); UpdateDynamicPanel(sheetName); },
                    () => { if (grpCount <= 1)    { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "grp",    r0); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("(b) 손익 국가간 배분",         $"{branchCount}행", y,
                    () => { _excel.AddSheet3Row(sheetName, "branch", r1); UpdateDynamicPanel(sheetName); },
                    () => { if (branchCount <= 1) { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "branch", r1); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("(c) 국가간 손익 조정",         $"{crossCount}행",  y,
                    () => { _excel.AddSheet3Row(sheetName, "cross",  r2); UpdateDynamicPanel(sheetName); },
                    () => { if (crossCount <= 1)  { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "cross",  r2); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("(d) 최종모기업 소득 감액",     $"{upeCount}행",    y,
                    () => { _excel.AddSheet3Row(sheetName, "upe",    r3); UpdateDynamicPanel(sheetName); },
                    () => { if (upeCount <= 1)    { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "upe",    r3); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("(b) 대상조세 국가간 배분",     $"{taxCount}행",    y,
                    () => { _excel.AddSheet3Row(sheetName, "tax",    r4); UpdateDynamicPanel(sheetName); },
                    () => { if (taxCount <= 1)    { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "tax",    r4); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("k. 공정가액조정 선택",         $"{fairCount}행",   y,
                    () => { _excel.AddSheet3Row(sheetName, "fair",   r5); UpdateDynamicPanel(sheetName); },
                    () => { if (fairCount <= 1)   { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "fair",   r5); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("과세분배방법 적용 선택",       $"{distCount}행",   y,
                    () => { _excel.AddSheet3Row(sheetName, "dist",   r6); UpdateDynamicPanel(sheetName); },
                    () => { if (distCount <= 1)   { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "dist",   r6); UpdateDynamicPanel(sheetName); });
                y = AddSectionRow("그 밖의 회계기준",             $"{otherCount}행",  y,
                    () => { _excel.AddSheet3Row(sheetName, "other",  r7); UpdateDynamicPanel(sheetName); },
                    () => { if (otherCount <= 1)  { Warn("최소 1행은 유지해야 합니다."); return; }
                            _excel.RemoveSheet3Row(sheetName, "other",  r7); UpdateDynamicPanel(sheetName); });
            }
            else if (sheetName == "UTPR 배분")
            {
                var count = _excel.GetRowBlockCount(sheetName, defaultCount: 2);
                y = AddSectionRow("구성기업 배분내역", $"{count}행", y,
                    () => { _excel.AddSimpleRowByMeta(sheetName, 4, defaultCount: 2); UpdateDynamicPanel(sheetName); },
                    () => {
                        if (count <= 2) { Warn("최소 2행은 유지해야 합니다."); return; }
                        _excel.RemoveSimpleRowByMeta(sheetName, 4, defaultCount: 2); UpdateDynamicPanel(sheetName);
                    });
            }
            else
            {
                // 인식할 수 없는 시트명
                var lbl = new Label
                {
                    Text = $"인식할 수 없는 시트명입니다.\n({sheetName})",
                    ForeColor = Color.FromArgb(140, 140, 150),
                    Font = new Font("Segoe UI", 10f),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Width = _dynamicPanel.Width,
                    Height = 60,
                    Top = y,
                };
                _dynamicPanel.Controls.Add(lbl);
                y += 60;
            }

            y += 10;
            y = AddActionButton("엑셀 종료하기", Color.FromArgb(48, 48, 52), y, BtnCloseExcel_Click);

            _dynamicPanel.Height = y + 4;
            ResizeToFit();
        }

        #endregion

        #region UI 헬퍼

        // 섹션 행: [이름 ..........] [카운트] [+][-]
        private int AddSectionRow(string name, string countText, int y, Action onAdd, Action onRemove)
        {
            var w = _dynamicPanel.Width;
            const int BTN_W = 28;
            const int BTN_H = 28;
            const int BTN_GAP = 4;
            var btnArea = BTN_W * 2 + BTN_GAP + 2; // 오른쪽 끝 2px 여백
            const int CNT_W = 54;
            var btnY = y + (ROW_H - BTN_H) / 2;

            _dynamicPanel.Controls.Add(new Label
            {
                Text = name,
                Bounds = new Rectangle(0, y, w - CNT_W - btnArea - 6, ROW_H),
                ForeColor = Color.FromArgb(200, 200, 204),
                Font = new Font("Segoe UI", 10.5f),
                TextAlign = ContentAlignment.MiddleLeft
            });
            _dynamicPanel.Controls.Add(new Label
            {
                Text = countText,
                Bounds = new Rectangle(w - CNT_W - btnArea - 6, y, CNT_W, ROW_H),
                ForeColor = Color.FromArgb(130, 130, 136),
                Font = new Font("Segoe UI", 10f),
                TextAlign = ContentAlignment.MiddleRight
            });

            var btnAdd = MakeIconButton("+",
                Color.FromArgb(88, 190, 110),
                new Rectangle(w - btnArea, btnY, BTN_W, BTN_H), onAdd);
            var btnRem = MakeIconButton("−",
                Color.FromArgb(210, 80, 80),
                new Rectangle(w - BTN_W - 2, btnY, BTN_W, BTN_H), onRemove);

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
                Font = new Font("Segoe UI", 12, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = Color.FromArgb(68, 68, 74);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(64, 64, 70);
            btn.Click += (s, e) => { try { click(); } catch (Exception ex) { MessageBox.Show(ex.Message); } };
            return btn;
        }

        private int AddNumberSelector(int y, string label, int maxNum,
            Func<int> getCurrent, Action<int> setCurrent, string sheetName)
        {
            var cur = getCurrent();
            const int H = 30;

            _dynamicPanel.Controls.Add(new Label
            {
                Text = label, Location = new Point(2, y + 5), AutoSize = true,
                ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 9.5f)
            });

            var baseX = 110;
            var btnPrev = new Button
            {
                Text = "◀", Bounds = new Rectangle(baseX, y, 28, H),
                FlatStyle = FlatStyle.Flat, ForeColor = Color.FromArgb(150, 150, 155),
                BackColor = Color.FromArgb(52, 52, 56), Font = new Font("Segoe UI", 8)
            };
            btnPrev.FlatAppearance.BorderSize = 1;
            btnPrev.FlatAppearance.BorderColor = Color.FromArgb(68, 68, 74);
            btnPrev.FlatAppearance.MouseOverBackColor = Color.FromArgb(64, 64, 70);
            btnPrev.Click += (s, e) => { if (cur > 1) { setCurrent(cur - 1); UpdateDynamicPanel(sheetName); } };

            var txtNum = new TextBox
            {
                Text = cur.ToString(),
                Bounds = new Rectangle(baseX + 34, y + 1, 38, H - 4),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.FromArgb(50, 50, 52), ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10)
            };
            txtNum.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                if (int.TryParse(txtNum.Text, out var n) && n >= 1 && n <= maxNum)
                { setCurrent(n); UpdateDynamicPanel(sheetName); }
                else txtNum.Text = cur.ToString();
                e.SuppressKeyPress = true;
            };

            var btnNext = new Button
            {
                Text = "▶", Bounds = new Rectangle(baseX + 72, y, 28, H),
                FlatStyle = FlatStyle.Flat, ForeColor = Color.FromArgb(150, 150, 155),
                BackColor = Color.FromArgb(52, 52, 56), Font = new Font("Segoe UI", 8)
            };
            btnNext.FlatAppearance.BorderSize = 1;
            btnNext.FlatAppearance.BorderColor = Color.FromArgb(68, 68, 74);
            btnNext.FlatAppearance.MouseOverBackColor = Color.FromArgb(64, 64, 70);
            btnNext.Click += (s, e) => { if (cur < maxNum) { setCurrent(cur + 1); UpdateDynamicPanel(sheetName); } };

            _dynamicPanel.Controls.Add(new Label
            {
                Text = $"/ {maxNum}", Location = new Point(baseX + 104, y + 6),
                AutoSize = true, ForeColor = Color.FromArgb(100, 100, 106), Font = new Font("Segoe UI", 9f)
            });
            _dynamicPanel.Controls.AddRange(new Control[] { btnPrev, txtNum, btnNext });
            return y + H + 6;
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
