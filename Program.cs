using System;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Windows.Forms;

namespace GlobeMapper
{
    internal static class Program
    {
        private static readonly DateTime ExpiryDate = new DateTime(2026, 6, 30, 23, 59, 59, DateTimeKind.Utc);

        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            // ── 1. 만료일 검사 ─────────────────────────────────────────────
            if (DateTime.UtcNow > ExpiryDate)
            {
                var msg = LoadExpiredMessage();
                MessageBox.Show(msg, "사용 기간 만료",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ── 2. 온라인 활성화 검사 (현재 비활성) ───────────────────────
            // if (!CheckActivation()) return;

            Application.Run(new MainForm());
        }

        private static string LoadExpiredMessage()
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Resources", "expired_message.txt");
            try { return File.ReadAllText(path, System.Text.Encoding.UTF8).Trim(); }
            catch { return "유효기간 만료입니다. DA의 승인을 얻어야 사용이 가능합니다."; }
        }

        /// <summary>
        /// activation_config.json의 URL로 GET 요청 → "true" 응답이면 허용.
        /// enabled: false 이면 항상 통과.
        /// </summary>
        private static bool CheckActivation()
        {
            try
            {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "Resources", "activation_config.json");
                if (!File.Exists(configPath)) return true;

                var json = File.ReadAllText(configPath, System.Text.Encoding.UTF8);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // enabled: false면 검사 건너뜀
                if (root.TryGetProperty("enabled", out var enabledProp) &&
                    enabledProp.ValueKind == JsonValueKind.False)
                    return true;

                var url = root.GetProperty("activation_url").GetString();
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                var response = client.GetStringAsync(url).GetAwaiter().GetResult();

                if (response.Trim().ToLower() == "true") return true;

                MessageBox.Show("프로그램 활성화에 실패했습니다.\nDA에 문의하세요.",
                    "활성화 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch
            {
                // 네트워크 오류 등 → 일단 통과 (필요 시 정책 변경)
                return true;
            }
        }
    }
}
