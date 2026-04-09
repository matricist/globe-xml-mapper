using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class MainForm : Form
    {
        private ExcelController _excel;
        private ControlPanelForm _controlPanel;

        private static readonly string TemplatePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Resources", "main_template.xlsx");

        // ── 색상 상수 ──────────────────────────────────────────────────────
        private static readonly Color BG        = Color.FromArgb(30, 30, 32);
        private static readonly Color BG2       = Color.FromArgb(36, 36, 40);
        private static readonly Color BG3       = Color.FromArgb(44, 44, 50);
        private static readonly Color BORDER    = Color.FromArgb(55, 55, 62);
        private static readonly Color FG        = Color.FromArgb(215, 215, 220);
        private static readonly Color FG_DIM    = Color.FromArgb(120, 120, 130);
        private static readonly Color FG_ACCENT = Color.FromArgb(86, 186, 240);
        private static readonly Color ACCENT    = Color.FromArgb(210, 160, 0);
        private static readonly Color PRIMARY   = Color.FromArgb(200, 90, 15);
        private static readonly Color GREEN     = Color.FromArgb(40, 160, 80);

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text            = "GIR 2 XML Mapper";
            AutoScaleMode   = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox     = false;
            StartPosition   = FormStartPosition.CenterScreen;
            ClientSize      = new Size(480, 380);
            BackColor       = BG;
            ForeColor       = FG;
            Font            = new Font("Segoe UI", 11f);

            var iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "app.ico");
            if (File.Exists(iconPath)) Icon = new Icon(iconPath);

            // ── 타이틀 ────────────────────────────────────────────────────
            var title = new Label
            {
                Text      = "GIR 2 XML Mapper",
                Dock      = DockStyle.Top, Height = 64,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = new Font("Segoe UI Semibold", 17f),
                ForeColor = Color.FromArgb(230, 230, 235),
                BackColor = Color.FromArgb(22, 22, 24),
            };
            var titleDiv = Divider(DockStyle.Top);

            // ── 버전 ──────────────────────────────────────────────────────
            var ver = new Label
            {
                Text      = "v1   만료일 2026.6.30   라이선스 DA",
                Dock      = DockStyle.Bottom, Height = 28,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(70, 70, 80),
                BackColor = Color.FromArgb(22, 22, 24),
            };

            // ── 하단 버튼 영역 ─────────────────────────────────────────────
            var bottomDiv = Divider(DockStyle.Bottom);
            var bottom = new Panel
            {
                Dock      = DockStyle.Bottom, Height = 136,
                BackColor = Color.FromArgb(24, 24, 26),
                Padding   = new Padding(28, 14, 28, 14),
            };

            var btnXml = MakeBtn("XML 변환하기", BtnConvert_Click, accent: true);
            btnXml.Dock   = DockStyle.Bottom;
            btnXml.Height = 48;
            btnXml.Margin = new Padding(0, 8, 0, 0);

            var btnPanel = MakeBtn("서식 작업 시작", BtnSwitchToPanel_Click, primary: true);
            btnPanel.Dock   = DockStyle.Top;
            btnPanel.Height = 48;

            bottom.Controls.Add(btnXml);
            bottom.Controls.Add(btnPanel);

            // ── 중앙 콘텐츠 영역 ──────────────────────────────────────────
            var content = new Panel
            {
                Dock       = DockStyle.Fill,
                BackColor  = Color.Transparent,
                Padding    = new Padding(28, 20, 28, 8),
            };

            // 섹션 라벨
            var lblSection = new Label
            {
                Text      = "새 서식 파일 만들기",
                Height    = 28,
                ForeColor = FG_ACCENT,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
            };

            // 3개 버튼 그리드
            var grid = new TableLayoutPanel
            {
                ColumnCount = 3, RowCount = 1,
                Height      = 60,
                BackColor   = Color.Transparent,
                Margin      = Padding.Empty, Padding = Padding.Empty,
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.4f));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var btnMne   = MakeBtn("MNE 생성",    BtnCreateMne_Click,  primary: true);
            var btnGroup = MakeBtn("합산단위 생성", BtnCreateGroup_Click, primary: true);
            var btnCe    = MakeBtn("구성기업 생성", BtnCreateCe_Click,   primary: true);
            btnMne.Dock   = DockStyle.Fill; btnMne.Font   = new Font("Segoe UI", 10.5f); btnMne.Margin   = new Padding(0, 0, 6, 0);
            btnGroup.Dock = DockStyle.Fill; btnGroup.Font = new Font("Segoe UI", 10.5f); btnGroup.Margin = new Padding(3, 0, 3, 0);
            btnCe.Dock    = DockStyle.Fill; btnCe.Font    = new Font("Segoe UI", 10.5f); btnCe.Margin    = new Padding(6, 0, 0, 0);

            grid.Controls.Add(btnMne,   0, 0);
            grid.Controls.Add(btnGroup, 1, 0);
            grid.Controls.Add(btnCe,    2, 0);

            // 절대 배치
            content.Controls.Add(lblSection);
            content.Controls.Add(grid);

            content.Resize += (s, e) =>
            {
                var w = content.ClientSize.Width - content.Padding.Left - content.Padding.Right;
                lblSection.SetBounds(0, 0, w, 28);
                grid.SetBounds(0, 36, w, 60);
            };

            Controls.Add(content);
            Controls.Add(bottom);
            Controls.Add(bottomDiv);
            Controls.Add(titleDiv);
            Controls.Add(title);
            Controls.Add(ver);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  버튼 핸들러
        // ─────────────────────────────────────────────────────────────────────

        private void BtnSwitchToPanel_Click(object sender, EventArgs e)
        {
            try
            {
                _excel = new ExcelController();
                _excel.AttachToActive();
                ShowControlPanel();
            }
            catch
            {
                // 활성 Excel 없으면 파일 열기로 폴백
                using var dlg = new OpenFileDialog
                {
                    Filter = "Excel 파일 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                    Title  = "서식 파일 열기",
                };
                if (dlg.ShowDialog() != DialogResult.OK) return;
                try
                {
                    _excel = new ExcelController();
                    _excel.Open(dlg.FileName);
                    ShowControlPanel();
                }
                catch (Exception ex2) { ShowError($"파일 열기 오류:\n{ex2.Message}"); }
            }
        }

        private void BtnCreateMne_Click(object sender, EventArgs e)
        {
            if (!File.Exists(TemplatePath))
            { ShowError($"템플릿 파일을 찾을 수 없습니다.\n{TemplatePath}"); return; }

            using var dlg = new SaveFileDialog
            {
                Filter   = "Excel 파일 (*.xlsx)|*.xlsx",
                Title    = "MNE 파일 저장",
                FileName = $"MNE_{DateTime.Now:yyyyMMdd}.xlsx",
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                File.Copy(TemplatePath, dlg.FileName, overwrite: true);
                MessageBox.Show($"MNE 파일이 생성되었습니다.\n{dlg.FileName}",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (MessageBox.Show("생성된 파일을 Excel로 여시겠습니까?", "열기",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _excel = new ExcelController();
                    _excel.Open(dlg.FileName);
                    ShowControlPanel();
                }
            }
            catch (Exception ex) { ShowError($"파일 생성 오류:\n{ex.Message}"); }
        }

        private void BtnCreateGroup_Click(object sender, EventArgs e)
        {
            if (!File.Exists(TemplatePath))
            { ShowError($"템플릿 파일을 찾을 수 없습니다.\n{TemplatePath}"); return; }

            // 저장 폴더 선택
            using var folderDlg = new FolderBrowserDialog
            {
                Description         = "합산단위 파일을 저장할 폴더를 선택하세요.",
                UseDescriptionForTitle = true,
            };
            if (folderDlg.ShowDialog() != DialogResult.OK) return;

            // 개수 입력
            int count = AskCount("합산단위를 몇 개 생성할까요?");
            if (count <= 0) return;

            try
            {
                for (int i = 1; i <= count; i++)
                {
                    var subDir = Path.Combine(folderDlg.SelectedPath, $"합산단위_{i}");
                    Directory.CreateDirectory(subDir);
                    var dest = Path.Combine(subDir, $"합산단위_{i}.xlsx");
                    File.Copy(TemplatePath, dest, overwrite: true);
                }
                MessageBox.Show($"합산단위 {count}개가 생성되었습니다.\n{folderDlg.SelectedPath}",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { ShowError($"파일 생성 오류:\n{ex.Message}"); }
        }

        private void BtnCreateCe_Click(object sender, EventArgs e)
        {
            if (!File.Exists(TemplatePath))
            { ShowError($"템플릿 파일을 찾을 수 없습니다.\n{TemplatePath}"); return; }

            // 저장 폴더 선택
            using var folderDlg = new FolderBrowserDialog
            {
                Description         = "구성기업 파일을 저장할 폴더를 선택하세요.",
                UseDescriptionForTitle = true,
            };
            if (folderDlg.ShowDialog() != DialogResult.OK) return;

            // 개수 입력
            int count = AskCount("구성기업을 몇 개 생성할까요?");
            if (count <= 0) return;

            try
            {
                for (int i = 1; i <= count; i++)
                {
                    var dest = Path.Combine(folderDlg.SelectedPath, $"구성기업_{i}.xlsx");
                    File.Copy(TemplatePath, dest, overwrite: true);
                }
                MessageBox.Show($"구성기업 {count}개가 생성되었습니다.\n{folderDlg.SelectedPath}",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { ShowError($"파일 생성 오류:\n{ex.Message}"); }
        }

        // 개수 입력 미니 다이얼로그 (0 이하 반환 = 취소)
        private static int AskCount(string prompt)
        {
            var dlg = new Form
            {
                Text            = "개수 입력",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                ClientSize      = new Size(300, 110),
                MaximizeBox     = false, MinimizeBox = false,
                BackColor       = Color.FromArgb(36, 36, 40),
                ForeColor       = Color.FromArgb(215, 215, 220),
            };

            var lbl = new Label
            {
                Text     = prompt,
                AutoSize = false,
                ForeColor = Color.FromArgb(215, 215, 220),
                Font     = new Font("Segoe UI", 10f),
            };
            lbl.SetBounds(16, 14, 268, 22);

            var txt = new TextBox
            {
                Font      = new Font("Segoe UI", 11f),
                BackColor = Color.FromArgb(44, 44, 50),
                ForeColor = Color.FromArgb(215, 215, 220),
                BorderStyle = BorderStyle.FixedSingle,
                Text      = "1",
            };
            txt.SetBounds(16, 42, 268, 28);

            var btnOk = new Button
            {
                Text      = "확인",
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(200, 90, 15),
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 10f),
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.SetBounds(148, 76, 136, 26);

            dlg.Controls.AddRange(new Control[] { lbl, txt, btnOk });
            dlg.AcceptButton = btnOk;

            if (dlg.ShowDialog() != DialogResult.OK) return 0;
            return int.TryParse(txt.Text.Trim(), out int n) && n > 0 ? n : 0;
        }

        private void BtnConvert_Click(object sender, EventArgs e)
        {
            using var terms = new TermsDialog();
            if (terms.ShowDialog(this) != DialogResult.OK) return;

            // 폴더 선택
            using var folderDlg = new FolderBrowserDialog
            {
                Description            = "변환할 서식 파일이 있는 폴더를 선택하세요.",
                UseDescriptionForTitle = true,
            };
            if (folderDlg.ShowDialog() != DialogResult.OK) return;
            var rootPath = folderDlg.SelectedPath;

            using var saveDlg = new SaveFileDialog
            {
                Filter           = "XML 파일 (*.xml)|*.xml",
                Title            = "XML 파일 저장",
                FileName         = "GLOBE_OECD.xml",
                InitialDirectory = rootPath,
            };
            if (saveDlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                var globe = new Globe.GlobeOecd
                {
                    Version     = "2.0",
                    MessageSpec = new Globe.MessageSpecType(),
                    GlobeBody   = new Globe.GlobeBodyType(),
                };

                var orchestrator  = new MappingOrchestrator();
                var mappingErrors = orchestrator.MapFolder(rootPath, globe);

                var xml = XmlExportService.Serialize(globe);
                File.WriteAllText(saveDlg.FileName, xml, System.Text.Encoding.UTF8);

                var validationErrors = ValidationUtil.Validate(globe);
                var errorsPath = Path.ChangeExtension(saveDlg.FileName, ".errors.txt");

                if (mappingErrors.Count > 0 || validationErrors.Count > 0)
                {
                    File.WriteAllText(errorsPath,
                        $"[오류 목록] {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}"
                        + $"매핑 오류 {mappingErrors.Count}건 / 검증 오류 {validationErrors.Count}건{Environment.NewLine}{Environment.NewLine}"
                        + string.Join(Environment.NewLine, mappingErrors)
                        + Environment.NewLine
                        + string.Join(Environment.NewLine, validationErrors),
                        System.Text.Encoding.UTF8);
                    MessageBox.Show(
                        $"XML 생성 완료.\n매핑 오류: {mappingErrors.Count}건\n검증 오류: {validationErrors.Count}건\n\n오류 목록: {errorsPath}",
                        "완료 (오류 있음)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (File.Exists(errorsPath)) File.Delete(errorsPath);
                    MessageBox.Show("XML 생성이 완료되었습니다.", "완료",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { ShowError($"XML 변환 오류:\n{ex.Message}"); }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Excel COM 헬퍼
        // ─────────────────────────────────────────────────────────────────────

        private void ShowControlPanel()
        {
            Hide();
            _controlPanel = new ControlPanelForm(_excel);
            _controlPanel.FormClosed += (s, e) =>
            {
                _excel?.Dispose(); _excel = null;
                _controlPanel = null;
                Show();
            };
            _controlPanel.Show();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _excel?.Dispose();
            _controlPanel?.Close();
            base.OnFormClosing(e);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  UI 헬퍼
        // ─────────────────────────────────────────────────────────────────────

        private static Panel Divider(DockStyle dock) => new Panel
        {
            Height    = 1,
            BackColor = Color.FromArgb(55, 55, 62),
            Dock      = dock,
        };

        private static Button MakeBtn(string text, EventHandler click,
            bool accent = false, bool primary = false)
        {
            var bg    = accent ? ACCENT : primary ? PRIMARY : BG3;
            var hover = accent  ? Color.FromArgb(225, 175, 10)
                      : primary ? Color.FromArgb(218, 105, 25)
                      : Color.FromArgb(54, 54, 60);
            var btn = new Button
            {
                Text      = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = (accent || primary) ? Color.White : FG,
                Font      = new Font("Segoe UI", 12f),
                Cursor    = Cursors.Hand,
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = hover;
            btn.Click += click;
            return btn;
        }

        private static void ShowError(string msg) =>
            MessageBox.Show(msg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
