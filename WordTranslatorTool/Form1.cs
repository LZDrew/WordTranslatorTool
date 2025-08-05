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
                        var paragraphs = GetAllParagraphs(body);
                        Log($"📄 偵測段落數量：{paragraphs.Count}");

                        int batchSize = 10; // 可調整大小，例如 5 或 20，視文件和 API 率限
                        List<Paragraph> batchParagraphs = new List<Paragraph>();
                        List<string> batchTexts = new List<string>();

                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            if (cancelRequested)
                            {
                                Log("⛔ 翻譯已中止");
                                break;
                            }
                            var paragraph = paragraphs[i];
                            string originalText = paragraph.InnerText?.Trim();

                            if (string.IsNullOrWhiteSpace(originalText) ||
                                originalText.Contains("PAGEREF") ||
                                originalText.Contains("TOC") ||
                                originalText.Contains("\\h") ||
                                (!System.Text.RegularExpressions.Regex.IsMatch(originalText, @"[\u4e00-\u9fffA-Za-z]") && originalText.Length < 4))
                            {
                                Log($"⏭ 略過段落 {i}：{originalText}");
                                progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
                                continue;
                            }

                            Log($"🔤 收集段落 {i} 至批次：{originalText}");

                            if (!System.Text.RegularExpressions.Regex.IsMatch(originalText, @"[\u4e00-\u9fff]"))
                            {
                                Log($"⏭ 已略過非中文段落：{originalText}");
                                progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
                                continue;
                            }

                            // 收集到批次
                            batchParagraphs.Add(paragraph);
                            batchTexts.Add(originalText);

                            // 若批滿或到最後，處理批次
                            if (batchTexts.Count == batchSize || i == paragraphs.Count - 1)
                            {
                                if (batchTexts.Count > 0) // 確保有內容才翻譯
                                {
                                    try
                                    {
                                        var translatedTexts = await TranslateBatchAsync(batchTexts);
                                        for (int j = 0; j < batchParagraphs.Count; j++)
                                        {
                                            string translatedText = translatedTexts[j];
                                            string mode = cmbTranslateMode.SelectedItem.ToString();
                                            ReplaceParagraphTextWithTranslation(batchParagraphs[j], translatedText, mode);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"⚠️ 批次翻譯失敗（影響段落範圍 {i - batchTexts.Count + 1} 到 {i}）：{ex.Message}");
                                    }
                                }
                                // 清空批次
                                batchParagraphs.Clear();
                                batchTexts.Clear();
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

        private void ReplaceParagraphTextWithTranslation(Paragraph paragraph, string translatedText, string mode)
        {
            if (string.IsNullOrWhiteSpace(translatedText)) return;

            if (mode == "保留原文 + 翻譯")
            {
                var lastRun = paragraph.Elements<Run>().LastOrDefault();
                var newRun = new Run(
                    new Break() { Type = BreakValues.TextWrapping },
                    new Text(translatedText) { Space = SpaceProcessingModeValues.Preserve }
                );

                if (lastRun?.RunProperties != null)
                {
                    newRun.RunProperties = (RunProperties)lastRun.RunProperties.CloneNode(true);
                }

                paragraph.AppendChild(newRun);
            }
            else if (mode == "取代原文（全文翻譯）")
            {
                var runs = paragraph.Elements<Run>().ToList();
                if (runs.Count == 0) return;

                var firstRun = runs[0];
                firstRun.RemoveAllChildren<Text>();
                firstRun.AppendChild(new Text(translatedText) { Space = SpaceProcessingModeValues.Preserve });

                foreach (var run in runs.Skip(1))
                {
                    run.RemoveAllChildren<Text>();
                }
            }
        }

        private void Log(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:T}] {message}{Environment.NewLine}");
        }

        private async Task<List<string>> TranslateBatchAsync(List<string> originalTexts)
        {
            string apiKey = txtApiKey.Text.Trim();
            string model = cmbModel.SelectedItem.ToString();
            string sourceLang = cmbSourceLang.SelectedItem.ToString();
            string targetLang = cmbTargetLang.SelectedItem.ToString();
            if (string.IsNullOrEmpty(apiKey))
            {
                throw new Exception("API Key 為空");
            }

            string combinedText = string.Join("\n---\n", originalTexts.Select((text, index) => $"段落 {index + 1}: {text}"));

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                var requestBody = new
                {
                    model = model,
                    messages = new[]
                    {
                new { role = "system", content = $"你是一位專業的技術文件翻譯專家，請將使用者提供的段落從「{sourceLang}」翻譯為「{targetLang}」，並嚴格遵守以下規則：\r\n1. 保留所有專有名詞，例如產品名稱、技術術語（如：Mail2000、Daemon OTP、OutlookSync），不進行翻譯，請保持原文。\r\n2. 保留格式不變：URL、Email 地址、程式碼片段、Log 格式（例如 [2024/08/30 10:14:26] [INF]）等請原樣保留，不得修改或翻譯。\r\n3. 維持段落格式與語意邏輯：若原文有列表、縮排、表格結構或粗體標示，翻譯時請盡可能使用相對應的格式（例如 1. 項目 → 1. Item）。\r\n4. 混合語言處理：若句子中含有純英文、數字、符號或已屬專有名詞，請勿翻譯，直接保留原文。\r\n5. 翻譯品質要求：語句需通順自然，符合技術文件的專業語調，避免直譯與語意錯誤。Keep translations concise to fit all items.\r\n6. 返回格式：嚴格返回 JSON 物件，如 {{\"translations\": [\"譯文1\", \"譯文2\", ...]}}。The translations array must have exactly {originalTexts.Count} items, one for each input paragraph. Do not add '段落 N:' or 'Paragraph N:' or any prefixes/labels to the translations - pure translation text only. Always return ONLY the JSON object, no additional text or explanations." },
                new { role = "user", content = combinedText }
            },
                    response_format = new { type = "json_object" }  // 強制 JSON mode
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
                var responseContent = doc.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString();

                // 加 log 幫助 debug
                Log("🔍 批次 API 回應內容：" + responseContent.Substring(0, Math.Min(200, responseContent.Length)));

                // 解析 JSON
                var jsonDoc = System.Text.Json.JsonDocument.Parse(responseContent);
                var translations = jsonDoc.RootElement.GetProperty("translations").EnumerateArray().Select(e => e.GetString()?.Replace("Paragraph ", "").Replace("paragraph ", "").Replace("段落 ", "").Trim() ?? "").ToList();  // 移除可能的 "Paragraph N:" 前綴

                if (translations.Count != originalTexts.Count)
                {
                    throw new Exception("翻譯結果數量不匹配");
                }

                return translations;
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