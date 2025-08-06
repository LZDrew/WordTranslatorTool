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
        private OpenAIClient _openAIClient;
        private TranslatorService _translatorService;

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
            txtFileName.Text = ""; // 初始化空
            txtFileName.ReadOnly = true; // 確保不可編輯
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
                    txtFileName.Text = currentFilePath; // 加這行更新 TextBox
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
                    txtApiKey.Text = ""; // 清明文
                    txtApiKey.PasswordChar = '*'; // 隱藏輸入

                    btnCancel.Enabled = true;
                    progressBar.Value = 0;

                    _openAIClient = new OpenAIClient(_apiKeySecure, Log);
                    _translatorService = new TranslatorService(_openAIClient, Log, cmbModel.SelectedItem.ToString(), cmbSourceLang.SelectedItem.ToString(), cmbTargetLang.SelectedItem.ToString());

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

                        await _translatorService.TranslateAndReplaceParagraphs(paragraphs, translateMode, progressBar);

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

        private void cmbSourceLang_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}