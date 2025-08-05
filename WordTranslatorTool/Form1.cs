using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace WordTranslatorTool
{

    public partial class Form1 : Form
    {
        private string currentFilePath;
        public Form1()
        {
            InitializeComponent();
            this.Load += new System.EventHandler(this.Form1_Load);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            cmbSourceLang.Items.AddRange(new string[] { "繁體中文", "英文", "日文", "韓文" });
            cmbSourceLang.SelectedIndex = 0;

            cmbTargetLang.Items.AddRange(new string[] { "英文", "繁體中文", "日文", "韓文" });
            cmbTargetLang.SelectedIndex = 0;

            cmbModel.Items.AddRange(new string[] { "gpt-3.5-turbo", "gpt-4o" });
            cmbModel.SelectedIndex = 0;

            cmbTranslateMode.Items.AddRange(new string[] { "保留原文 + 翻譯", "取代原文（全文翻譯）" });
            cmbTranslateMode.SelectedIndex = 0;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word files (*.docx)|*.docx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    currentFilePath = openFileDialog.FileName;
                    lblInputFile.Text = $"已選擇：{Path.GetFileName(currentFilePath)}";
                    Log($"載入檔案：{currentFilePath}");
                }
            }
        }

        private async void btnSaveFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentFilePath))
            {
                MessageBox.Show("請先選擇一個 Word 檔案");
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word files (*.docx)|*.docx";
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(currentFilePath) + "_translated.docx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    cancelRequested = false;
                    btnCancel.Enabled = true;
                    progressBar.Value = 0;

                    string outputPath = saveFileDialog.FileName;
                    try
                    {
                        File.Copy(currentFilePath, outputPath, true);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("⚠️ 複製檔案失敗，請確認原始或目標檔案是否已開啟。\n\n詳細錯誤：" + ex.Message,
                                        "檔案使用中", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Log("❌ 複製失敗：" + ex.Message);
                        return;
                    }

                    using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
                    {
                        var body = doc.MainDocumentPart.Document.Body;
                        // 改成抓所有段落（含表格中段落）
                        var paragraphs = GetAllParagraphs(body);

                        // 顯示偵測到多少段落
                        Log($"📄 偵測段落數量：{paragraphs.Count}");

                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            if (cancelRequested)
                            {
                                Log("⛔ 翻譯已中止");
                                break;
                            }

                            var paragraph = paragraphs[i];
                            string originalText = paragraph.InnerText?.Trim();



                            // ✅ 加入更細緻的條件，只翻譯含中文或英文的實際描述性段落
                            if (string.IsNullOrWhiteSpace(originalText) ||
                                originalText.Contains("PAGEREF") ||
                                originalText.Contains("TOC") ||
                                originalText.Contains("\\h") ||
                                (!System.Text.RegularExpressions.Regex.IsMatch(originalText, @"[\u4e00-\u9fffA-Za-z]") && originalText.Length < 4)) // 有中文字或英文字
                            {
                                Log($"⏭ 略過段落 {i}：{originalText}");
                                continue;
                            }

                            try
                            {
                                Log($"🔤 翻譯中（段落 {i}）：{originalText}");
                                if (!System.Text.RegularExpressions.Regex.IsMatch(originalText, @"[\u4e00-\u9fff]"))
                                {
                                    Log($"⏭ 已略過非中文段落：{originalText}");
                                    continue;
                                }
                                string translatedText = await TranslateWithChatGptAsync(originalText);

                                // 根據模式做不同處理
                                string mode = cmbTranslateMode.SelectedItem.ToString();

                                if (mode == "保留原文 + 翻譯")
                                {
                                    // 原本的：保留原文，加翻譯
                                    var translationRun = new Run(
                                        new Break() { Type = BreakValues.TextWrapping },
                                        new Text(translatedText)
                                    );
                                    paragraph.AppendChild(translationRun);
                                }
                                else if (mode == "取代原文（全文翻譯）")
                                {
                                    // 移除原有文字內容的 Run，但保留段落樣式與結構
                                    paragraph.RemoveAllChildren<Run>();

                                    // 插入新的翻譯後內容
                                    var translatedRun = new Run(new Text(translatedText));
                                    paragraph.AppendChild(translatedRun);
                                }

                            }
                            catch (Exception ex)
                            {
                                Log($"⚠️ 翻譯失敗（段落 {i}）：{ex.Message}");
                            }

                            progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
                        }


                        doc.MainDocumentPart.Document.Save();
                        Log("✅ 已儲存翻譯後檔案：" + outputPath);

                    }
                    MessageBox.Show("已儲存翻譯版本");

                }
            }
        }


        private void Log(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:T}] {message}{Environment.NewLine}");
        }

        private async Task<string> TranslateWithChatGptAsync(string originalText)
        {
            string apiKey = txtApiKey.Text.Trim();
            string model = cmbModel.SelectedItem.ToString();
            string sourceLang = cmbSourceLang.SelectedItem.ToString();
            string targetLang = cmbTargetLang.SelectedItem.ToString();

            if (string.IsNullOrEmpty(apiKey))
            {
                throw new Exception("API Key 為空");
            }

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                var requestBody = new
                {
                    model = model,
                    messages = new[]
                    {
                new { role = "system", content = $"你是一位翻譯專家，請將使用者提供的句子從「{sourceLang}」翻譯為「{targetLang}」。整句都是英文的不需要翻譯留著原文即可" },
                new { role = "user", content = originalText }
            }
                };

                var json = System.Text.Json.JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("https://api.openai.com/v1/chat/completions", content);

                if (!response.IsSuccessStatusCode)
                {
                    string errorMsg = await response.Content.ReadAsStringAsync();
                    throw new Exception($"API 錯誤：{response.StatusCode}\n{errorMsg}");
                }

                using var responseStream = await response.Content.ReadAsStreamAsync();
                using var doc = await System.Text.Json.JsonDocument.ParseAsync(responseStream);
                var translated = doc.RootElement
                    .GetProperty("choices")[0]
                    .GetProperty("message")
                    .GetProperty("content")
                    .GetString();

                return translated.Trim();
            }
        }

        private bool cancelRequested = false;

        private void btnCancel_Click(object sender, EventArgs e)
        {
            cancelRequested = true;
            btnCancel.Enabled = false;
            Log("⚠️ 使用者請求中止翻譯");
        }
        private List<Paragraph> GetAllParagraphs(OpenXmlElement element)
        {
            var result = new List<Paragraph>();

            foreach (var child in element.Elements())
            {
                if (child is Paragraph p)
                {
                    result.Add(p);
                }

                // 遞迴找所有子節點
                result.AddRange(GetAllParagraphs(child));
            }

            return result;
        }

        private void tableProgress_Paint(object sender, PaintEventArgs e)
        {

        }

        private List<int> matchIndexes = new List<int>();
        private int currentMatchIndex = -1;

        private void HighlightAllMatches()
        {
            string keyword = txtSearchLog.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword)) return;

            // 解鎖以改變樣式
            txtLog.ReadOnly = false;

            // 清除樣式
            txtLog.SelectAll();
            txtLog.SelectionBackColor = txtLog.BackColor;

            matchIndexes.Clear();
            int startIndex = 0;
            while (startIndex < txtLog.TextLength)
            {
                int index = txtLog.Text.IndexOf(keyword, startIndex, StringComparison.OrdinalIgnoreCase);
                if (index == -1) break;

                matchIndexes.Add(index);

                txtLog.Select(index, keyword.Length);
                txtLog.SelectionBackColor = System.Drawing.Color.Yellow;

                startIndex = index + keyword.Length;
            }

            txtLog.ReadOnly = true;
            txtLog.Select(0, 0); // 避免保留黃色選取
        }

        private void JumpToNextMatch()
        {
            if (matchIndexes.Count == 0) return;

            currentMatchIndex++;
            if (currentMatchIndex >= matchIndexes.Count)
            {
                currentMatchIndex = 0;
            }

            int matchPos = matchIndexes[currentMatchIndex];
            string keyword = txtSearchLog.Text.Trim();

            txtLog.ReadOnly = false;

            // 先把全部標回黃色
            foreach (int index in matchIndexes)
            {
                txtLog.Select(index, keyword.Length);
                txtLog.SelectionBackColor = System.Drawing.Color.Yellow;
            }

            // 將目前這筆改為綠色
            txtLog.Select(matchPos, keyword.Length);
            txtLog.SelectionBackColor = System.Drawing.Color.LimeGreen;
            txtLog.ScrollToCaret();

            txtLog.ReadOnly = true;
            txtLog.Invalidate(); // 保留選取但不搶焦點
        }


        private void btnNextMatch_Click(object sender, EventArgs e)
        {
            JumpToNextMatch();
        }

        private void txtSearchLog_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;

                if (matchIndexes.Count == 0) // 第一次搜尋
                    HighlightAllMatches();

                JumpToNextMatch(); // 跳下一筆
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
