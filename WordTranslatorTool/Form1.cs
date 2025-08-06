using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace WordTranslatorTool
{

    public partial class Form1 : Form
    {
        private string currentFilePath;
        private SecureString _apiKeySecure = new SecureString();

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

            txtApiKey.PasswordChar = '*'; // 隱藏
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
                string translateMode = cmbTranslateMode.SelectedItem.ToString();
                string suffix = translateMode == "保留原文 + 翻譯" ? "_appended" : "_replaced"; // 或自訂名稱，如 "_with_original" / "_translated_only"
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(currentFilePath) + suffix + ".docx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    cancelRequested = false;

                    _apiKeySecure.Clear();
                    foreach (char c in txtApiKey.Text)
                    {
                        _apiKeySecure.AppendChar(c);
                    }
                    txtApiKey.Text = ""; // 明文
                    txtApiKey.PasswordChar = '*'; // 隱藏輸入

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

                        int batchSize = 5; // 可調整大小，例如 5 或 20，視文件和 API 率限
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

                            bool hasMeaningfulText = System.Text.RegularExpressions.Regex.IsMatch(originalText, @"[\u4e00-\u9fffA-Za-z]");
                            if (string.IsNullOrWhiteSpace(originalText) ||
                                originalText.Contains("PAGEREF") ||
                                originalText.Contains("TOC") ||
                                originalText.Contains("\\h") ||
                                (!hasMeaningfulText && originalText.Length < 4) ||
                                !System.Text.RegularExpressions.Regex.IsMatch(originalText, @"[\u4e00-\u9fff]")) // 合併非中文
                            {
                                Log($"⏭ 略過段落 {i}：{originalText}");
                                progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
                                continue;
                            }

                            // 真正收集
                            Log($"🔤 收集段落 {i} 至批次：{originalText}");

                            // 收集到批次
                            batchParagraphs.Add(paragraph);
                            batchTexts.Add(originalText);

                            // 若批滿或到最後，處理批次
                            if (batchTexts.Count == batchSize)
                            {
                                if (batchTexts.Count > 0) // 確保有內容才翻譯
                                {
                                    Log($"🔤 執行剩餘批次翻譯，段落數: {batchTexts.Count}");
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
                                    // 清空批次
                                    batchParagraphs.Clear();
                                    batchTexts.Clear();
                                }
                                
                            }

                            progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
                        }
                        // 迴圈結束後，處理剩餘批次
                        if (batchTexts.Count > 0)
                        {
                            Log($"🔤 執行剩餘批次翻譯，段落數: {batchTexts.Count}");
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
                                Log($"⚠️ 剩餘批次翻譯失敗：{ex.Message}");
                            }
                            batchParagraphs.Clear();
                            batchTexts.Clear();
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

            var runs = paragraph.Descendants<Run>().ToList();
            if (runs.Count == 0) return;

            if (mode == "保留原文 + 翻譯")
            {
                var newRun = new Run(
                    new Break() { Type = BreakValues.TextWrapping },
                    new Text(translatedText) { Space = SpaceProcessingModeValues.Preserve }
                );
                var lastRun = paragraph.Elements<Run>().LastOrDefault();
                if (lastRun?.RunProperties != null)
                {
                    newRun.RunProperties = (RunProperties)lastRun.RunProperties.CloneNode(true);
                }

                paragraph.AppendChild(newRun);
            }
            else if (mode == "取代原文（全文翻譯）")
            {

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

        private static readonly HttpClient sharedClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) }; // 全局 client 重複使用 + timeout

        private async Task<List<string>> TranslateBatchAsync(List<string> originalTexts)
        {
            if (originalTexts.Count == 0) return new List<string>();

            Log($"🔤 批次輸入段落數: {originalTexts.Count}");

            // 使用乾淨段落合併，不加入段落 N: 前綴
            string combinedText = string.Join("\n---\n", originalTexts);

            IntPtr ptr = Marshal.SecureStringToBSTR(_apiKeySecure);
            string apiKey = Marshal.PtrToStringBSTR(ptr);
            Marshal.FreeBSTR(ptr);

            if (string.IsNullOrEmpty(apiKey))
            {
                throw new Exception("API Key 為空");
            }

            string model = cmbModel.SelectedItem.ToString();
            string sourceLang = cmbSourceLang.SelectedItem.ToString();
            string targetLang = cmbTargetLang.SelectedItem.ToString();

            var prompt = $"你是一位專業的技術文件翻譯專家，請將使用者提供的段落從「{sourceLang}」翻譯為「{targetLang}」，並嚴格遵守以下規則：\\r\\n1. 保留所有專有名詞，例如產品名稱、技術術語（如：Mail2000、Daemon OTP、OutlookSync），不進行翻譯，請保持原文。\\r\\n2. 保留格式不變：URL、Email 地址、程式碼片段、Log 格式（例如 [2024/08/30 10:14:26] [INF]）等請原樣保留，不得修改或翻譯。\\r\\n3. 維持段落格式與語意邏輯：若原文有列表、縮排、表格結構或粗體標示，翻譯時請盡可能使用相對應的格式（例如 1. 項目 → 1. Item）。\\r\\n4. 混合語言處理：若句子中含有純英文、數字、符號或已屬專有名詞，請勿翻譯，直接保留原文。\\r\\n5. 翻譯品質要求：語句需通順自然，符合技術文件的專業語調，避免直譯與語意錯誤。Keep translations concise to fit all items.\\r\\n6. 返回格式：嚴格返回 JSON 物件，如 {{\\\"translations\\\": [\\\"譯文1\\\", \\\"譯文2\\\", ...]}}。The translations array must have exactly {originalTexts.Count} items. Translate all paragraphs completely, even if long. Do not add any prefixes/labels like 'Paragraph N:' to the translations - pure translation text only. Always return ONLY the JSON object, no additional text or explanations.";

            var requestBody = new
            {
                model = model,
                messages = new[]
                    {
                        new { role = "system", content = prompt },
                        new { role = "user", content = combinedText }
                    },
                response_format = new { type = "json_object" },  // 強制 JSON mode
                max_tokens = 4096
            };
            string requestJson = System.Text.Json.JsonSerializer.Serialize(requestBody);
            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    sharedClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                    var response = await sharedClient.PostAsync("https://api.openai.com/v1/chat/completions", content);
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorBody = await response.Content.ReadAsStringAsync();
                        throw new Exception($"API 錯誤：{response.StatusCode}\n{errorBody}");
                    }
                    using var rawResponse = await response.Content.ReadAsStreamAsync();

                    // 抽取 JSON 格式的字串回應
                    using var outerJson = System.Text.Json.JsonDocument.Parse(rawResponse);
                    if (!outerJson.RootElement.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0)
                        throw new Exception("回應缺少 choices");
                    string messageContent = choices[0].GetProperty("message").GetProperty("content").GetString() ?? throw new Exception("回應內容空");

                    // 加 log 幫助 debug
                    Log("🔍 批次 API 回應內容：" + messageContent.Substring(0, Math.Min(200, messageContent.Length)));

                    // 解析 JSON
                    using var innerJson = System.Text.Json.JsonDocument.Parse(messageContent);
                    if (innerJson.RootElement.TryGetProperty("translations", out var translationsElement) && translationsElement.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        var translations = translationsElement.EnumerateArray().Select(t => {
                            string text = t.GetString() ?? "";
                            text = System.Text.RegularExpressions.Regex.Replace(text, @"(?i)^(paragraph|段落)\s*\d+:\s*", "").Trim(); // 後處理移除前綴
                            return text;
                        }).ToList();
                        if (translations.Count != originalTexts.Count)
                        {
                            Log("⚠️ 批次翻譯結果不匹配，fallback 到逐段翻譯");
                            translations.Clear();
                            for (int index = 0; index < originalTexts.Count; index++)
                            {
                                string text = originalTexts[index];
                                try
                                {
                                    // fallback 用單段 (簡化 prompt 減 token)
                                    var singleTexts = new List<string> { text };
                                    var result = await TranslateBatchAsync(singleTexts); // 復用批量，但傳單段
                                    translations.Add(result.FirstOrDefault() ?? "");
                                }
                                catch (Exception ex)
                                {
                                    Log($"⚠️ 單段 fallback 失敗 (index={index})：{ex.Message}");
                                    translations.Add(""); // fallback 空，保留原文
                                }
                            }
                        }
                        return translations;
                    }
                    else
                    {
                        throw new Exception("JSON 回傳格式不符，找不到 translations 陣列");
                    }
                }
                catch (Exception ex)
                {
                    Log($"⚠️ API 嘗試第 {attempt} 次失敗：{ex.Message}");
                    if (attempt == 3) throw;
                    await Task.Delay(1000 * attempt); // exponential backoff
                }
            }
            throw new Exception("翻譯失敗：超過最大重試次數");
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