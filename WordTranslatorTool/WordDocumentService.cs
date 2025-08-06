using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

public class WordDocumentService
{
    private readonly Action<string> _logAction;

    public WordDocumentService(Action<string> logAction)
    {
        _logAction = logAction;
    }

    public List<Paragraph> GetAllParagraphs(OpenXmlElement element)
    {
        var result = new List<Paragraph>();
        var queue = new Queue<OpenXmlElement>();
        queue.Enqueue(element);

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            foreach (var child in current.Elements())
            {
                if (child is Paragraph p)
                {
                    result.Add(p);
                }
                queue.Enqueue(child); // 佇列取代遞迴，防堆疊溢位
            }
        }
        return result;
    }

    public void SaveDocument(WordprocessingDocument doc, string outputPath)
    {
        try
        {
            doc.MainDocumentPart.Document.Save();
            doc.Save();
            _logAction("✅ 已儲存翻譯後檔案：" + outputPath);
        }
        catch (IOException ex)
        {
            _logAction("❌ 儲存失敗：" + ex.Message);
            MessageBox.Show("⚠️ 儲存檔案失敗，請確認檔案是否開啟。\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public WordprocessingDocument OpenDocument(string outputPath)
    {
        return WordprocessingDocument.Open(outputPath, true);
    }
}