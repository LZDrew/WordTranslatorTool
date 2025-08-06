using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;

public class TranslatorService
{
    private readonly OpenAIClient _openAIClient;
    private readonly Action<string> _logAction;
    private readonly string _model;
    private readonly string _sourceLang;
    private readonly string _targetLang;

    public TranslatorService(OpenAIClient openAIClient, Action<string> logAction, string model, string sourceLang, string targetLang)
    {
        _openAIClient = openAIClient;
        _logAction = logAction;
        _model = model;
        _sourceLang = sourceLang;
        _targetLang = targetLang;
    }

    public async Task TranslateAndReplaceParagraphs(List<Paragraph> paragraphs, string mode, ProgressBar progressBar)
    {
        int batchSize = 5;
        List<Paragraph> batchParagraphs = new List<Paragraph>();
        List<string> batchTexts = new List<string>();

        for (int i = 0; i < paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];
            string originalText = paragraph.InnerText?.Trim();

            if (string.IsNullOrWhiteSpace(originalText) ||
                originalText.Contains("PAGEREF") ||
                originalText.Contains("TOC") ||
                originalText.Contains("\\h") ||
                (!Regex.IsMatch(originalText, @"[\u4e00-\u9fffA-Za-z]") && originalText.Length < 4) ||
                !Regex.IsMatch(originalText, @"[\u4e00-\u9fff]"))
            {
                _logAction($"⏭ 略過段落 {i}：{originalText}");
                progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
                continue;
            }

            _logAction($"🔤 收集段落 {i} 至批次：{originalText}");
            batchParagraphs.Add(paragraph);
            batchTexts.Add(originalText);

            if (batchTexts.Count == batchSize)
            {
                await ProcessBatch(batchParagraphs, batchTexts, mode);
                batchParagraphs.Clear();
                batchTexts.Clear();
            }
            progressBar.Value = (int)((i + 1) * 100.0 / paragraphs.Count);
        }

        // 處理剩餘批次
        if (batchTexts.Count > 0)
        {
            await ProcessBatch(batchParagraphs, batchTexts, mode);
        }
    }

    private async Task ProcessBatch(List<Paragraph> batchParagraphs, List<string> batchTexts, string mode)
    {
        _logAction($"🔤 執行批次翻譯，段落數: {batchTexts.Count}");
        var translatedTexts = await _openAIClient.TranslateBatchAsync(batchTexts, _model, _sourceLang, _targetLang);
        for (int j = 0; j < batchParagraphs.Count; j++)
        {
            ReplaceParagraphTextWithTranslation(batchParagraphs[j], translatedTexts[j], mode);
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
}