using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class MainForm : Form
    {
        private TextBox txtFolderPath;
        private Button btnBrowse;
        private Button btnGenerate;
        private Button btnDownloadTemplates;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Globe XML Mapper";
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                ColumnCount = 3,
                RowCount = 3,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // Row 0: 폴더 선택
            var lblFile = new Label
            {
                Text = "서식 폴더:",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 6, 0)
            };

            txtFolderPath = new TextBox
            {
                ReadOnly = true,
                Width = 350,
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };

            btnBrowse = new Button
            {
                Text = "찾아보기",
                AutoSize = true,
                Margin = new Padding(6, 0, 0, 0)
            };
            btnBrowse.Click += BtnBrowse_Click;

            layout.Controls.Add(lblFile, 0, 0);
            layout.Controls.Add(txtFolderPath, 1, 0);
            layout.Controls.Add(btnBrowse, 2, 0);

            // Row 1: XML 생성 버튼
            btnGenerate = new Button
            {
                Text = "XML 생성하기",
                Enabled = false,
                Dock = DockStyle.Fill,
                Height = 35,
                Margin = new Padding(0, 8, 0, 0)
            };
            btnGenerate.Click += BtnGenerate_Click;

            layout.Controls.Add(btnGenerate, 0, 1);
            layout.SetColumnSpan(btnGenerate, 3);

            // Row 2: 템플릿 다운로드 버튼
            btnDownloadTemplates = new Button
            {
                Text = "템플릿 다운로드",
                Dock = DockStyle.Fill,
                Height = 30,
                Margin = new Padding(0, 4, 0, 0)
            };
            btnDownloadTemplates.Click += BtnDownloadTemplates_Click;

            layout.Controls.Add(btnDownloadTemplates, 0, 2);
            layout.SetColumnSpan(btnDownloadTemplates, 3);

            Controls.Add(layout);
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using var dlg = new FolderBrowserDialog
            {
                Description = "서식 xlsx 파일이 있는 폴더를 선택하세요.",
                UseDescriptionForTitle = true
            };

            if (dlg.ShowDialog() != DialogResult.OK) return;

            var xlsxFiles = Directory.GetFiles(dlg.SelectedPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .ToArray();

            if (xlsxFiles.Length == 0)
            {
                MessageBox.Show("선택한 폴더에 xlsx 파일이 없습니다.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            txtFolderPath.Text = dlg.SelectedPath;
            btnGenerate.Enabled = true;
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
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

            try
            {
                var globe = new Globe.GlobeOecd
                {
                    Version = "2.0",
                    MessageSpec = new Globe.MessageSpecType(),
                    GlobeBody = new Globe.GlobeBodyType()
                };

                var orchestrator = new MappingOrchestrator();
                var mappingErrors = orchestrator.MapFolder(txtFolderPath.Text, globe);

                var xml = XmlExportService.Serialize(globe);
                File.WriteAllText(dlg.FileName, xml, System.Text.Encoding.UTF8);

                var validationErrors = ValidationUtil.Validate(globe);

                var allErrors = new List<string>();
                if (mappingErrors.Count > 0)
                {
                    allErrors.Add("── 매핑 오류 ──");
                    allErrors.AddRange(mappingErrors);
                    allErrors.Add("");
                }
                if (validationErrors.Count > 0)
                {
                    allErrors.Add("── 검증 오류 (에러코드 기준) ──");
                    allErrors.AddRange(validationErrors);
                }

                var errorsPath = Path.ChangeExtension(dlg.FileName, ".errors.txt");
                if (allErrors.Count > 0)
                {
                    File.WriteAllText(errorsPath,
                        $"[오류 목록] {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                        $"매핑 오류 {mappingErrors.Count}건 / 검증 오류 {validationErrors.Count}건{Environment.NewLine}{Environment.NewLine}" +
                        string.Join(Environment.NewLine, allErrors),
                        System.Text.Encoding.UTF8);

                    MessageBox.Show(
                        $"XML 생성 완료.\n\n매핑 오류: {mappingErrors.Count}건\n검증 오류: {validationErrors.Count}건\n\n오류 목록: {errorsPath}",
                        "완료 (오류 있음)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (File.Exists(errorsPath)) File.Delete(errorsPath);
                    MessageBox.Show("XML 생성이 완료되었습니다.", "완료",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XML 생성 중 오류 발생:\n{ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDownloadTemplates_Click(object sender, EventArgs e)
        {
            using var dlg = new FolderBrowserDialog
            {
                Description = "템플릿을 저장할 폴더를 선택하세요.",
                UseDescriptionForTitle = true
            };

            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                var targetDir = Path.Combine(dlg.SelectedPath, "GlobeMapper_Templates");
                Directory.CreateDirectory(targetDir);

                var templatesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "templates");
                if (!Directory.Exists(templatesDir))
                {
                    MessageBox.Show("템플릿 파일을 찾을 수 없습니다.", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 디렉토리 구조로 복사
                var subDirs = new[] { "1.3.1", "1.3.2.1", "1.3.2.2" };
                var count = 0;

                // 루트용 템플릿 (1.1~1.2)
                foreach (var file in Directory.GetFiles(templatesDir, "template_1.1*"))
                {
                    File.Copy(file, Path.Combine(targetDir, Path.GetFileName(file)), true);
                    count++;
                }

                // 하위 디렉토리별 템플릿
                foreach (var sub in subDirs)
                {
                    var subTarget = Path.Combine(targetDir, sub);
                    Directory.CreateDirectory(subTarget);
                    foreach (var file in Directory.GetFiles(templatesDir, $"template_{sub}*"))
                    {
                        File.Copy(file, Path.Combine(subTarget, Path.GetFileName(file)), true);
                        count++;
                    }
                }

                MessageBox.Show($"템플릿 {count}개가 다운로드되었습니다.\n\n위치: {targetDir}\n\n구조:\n  루트/ - 기본정보(1.1~1.2)\n  1.3.1/ - UPE\n  1.3.2.1/ - 구성기업\n  1.3.2.2/ - 제외기업",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"템플릿 다운로드 오류:\n{ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
