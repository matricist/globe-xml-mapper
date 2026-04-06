using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace GlobeMapper
{
    public class TermsDialog : Form
    {
        private CheckBox chkAgree;
        private Button btnNext;
        private Button btnCancel;

        public TermsDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "약관 동의";
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(500, 350);

            // 약관 내용
            var txtTerms = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new Point(12, 12),
                Size = new Size(476, 240),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = SystemColors.Window
            };

            var termsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "terms.txt");
            if (File.Exists(termsPath))
                txtTerms.Text = File.ReadAllText(termsPath);
            else
                txtTerms.Text = "(약관 파일을 찾을 수 없습니다)";

            // 동의 체크박스
            chkAgree = new CheckBox
            {
                Text = "위 약관에 동의합니다.",
                AutoSize = true,
                Location = new Point(12, 262),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            chkAgree.CheckedChanged += (s, e) => btnNext.Enabled = chkAgree.Checked;

            // 버튼
            btnCancel = new Button
            {
                Text = "취소",
                Size = new Size(80, 30),
                Location = new Point(408, 308),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };

            btnNext = new Button
            {
                Text = "다음으로",
                Size = new Size(80, 30),
                Location = new Point(320, 308),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Enabled = false,
                DialogResult = DialogResult.OK
            };

            AcceptButton = btnNext;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[] { txtTerms, chkAgree, btnNext, btnCancel });
        }
    }
}
