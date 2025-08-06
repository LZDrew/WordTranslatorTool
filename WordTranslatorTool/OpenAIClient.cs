using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

public class OpenAIClient
{
    private readonly SecureString _apiKeySecure;
    private readonly Action<string> _logAction;
    private readonly HttpClient _client = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };

    public OpenAIClient(SecureString apiKeySecure, Action<string> logAction)
    {
        _apiKeySecure = apiKeySecure;
        _logAction = logAction;
    }
    public async Task<List<string>> TranslateBatchAsync(List<string> originalTexts, string model, string sourceLang, string targetLang)
    {
        if (originalTexts.Count == 0) return new List<string>();

        _logAction($"🔤 批次輸入段落數: {originalTexts.Count}");

        // 使用乾淨段落合併，不加入段落 N: 前綴
        string combinedText = string.Join("\n---\n", originalTexts);

        IntPtr ptr = Marshal.SecureStringToBSTR(_apiKeySecure);
        string apiKey = Marshal.PtrToStringBSTR(ptr);
        Marshal.FreeBSTR(ptr);

        if (string.IsNullOrEmpty(apiKey))
        {
            throw new Exception("API Key 為空");
        }

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
                _client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                var response = await _client.PostAsync("https://api.openai.com/v1/chat/completions", content);
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
                _logAction("🔍 批次 API 回應內容：" + messageContent.Substring(0, Math.Min(200, messageContent.Length)));

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
                        _logAction("⚠️ 批次翻譯結果不匹配，fallback 到逐段翻譯");
                        translations.Clear();
                        for (int index = 0; index < originalTexts.Count; index++)
                        {
                            string text = originalTexts[index];
                            try
                            {
                                // fallback 用單段 (簡化 prompt 減 token)
                                var singleTexts = new List<string> { text };
                                var result = await TranslateBatchAsync(singleTexts, model, sourceLang, targetLang); // 復用批量，但傳單段，加 model, sourceLang, targetLang
                                translations.Add(result.FirstOrDefault() ?? "");
                            }
                            catch (Exception ex)
                            {
                                _logAction($"⚠️ 單段 fallback 失敗 (index={index})：{ex.Message}");
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
                _logAction($"⚠️ API 嘗試第 {attempt} 次失敗：{ex.Message}");
                if (attempt == 3) throw;
                await Task.Delay(1000 * attempt); // exponential backoff
            }
        }
        throw new Exception("翻譯失敗：超過最大重試次數");
    }
    
}