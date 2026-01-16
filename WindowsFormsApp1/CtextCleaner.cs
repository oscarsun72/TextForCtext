//https://gemini.google.com/share/b05c32b4427b
//https://gemini.google.com/share/d922b92fd363

using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WebSocketSharp;
using static WindowsFormsApp1.Form1;

namespace TextForCtext
{
    public static class CtextCleaner
    {

        /// <summary>
        /// 即WordVBA「清除頁前的分段符號」之核心程式
        /// </summary>
        /// <param name="xmlContent"></param>
        /// <returns>失敗傳回false</returns>
        public static bool FixXMLParagraphMarkPosition_SetPage1Content(ref string xmlContent)//public static void ProcessClipboard()
        {
            try
            {
                //if (!Clipboard.ContainsText())
                //{
                //    MessageBox.Show("剪貼簿是空的。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    //MessageBoxShowOKExclamationDefaultDesktopOnly("剪貼簿是空的。");
                //    return;
                //}

                //string content = Clipboard.GetText();
                if (xmlContent.IsNullOrEmpty()) return false;

                // 1. 統一換行符號 (標準化)
                xmlContent = xmlContent.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");

                // 2. 處理第一頁代碼 (setPage1Code)
                xmlContent = SetPage1Code(ref xmlContent);

                // 3. 清除多餘代碼 (clearRedundantCode)
                // 確保 scanend 與 scanbegin 緊緊相連，中間無雜訊
                xmlContent = ClearRedundantCode(ref xmlContent);

                // 4. 【修正重點】將星號前的分段符號移置前段之末
                // 包含了 "跨頁雙標籤跳躍" 與 "單標籤跳躍" 的邏輯
                xmlContent = MoveAsteriskUp(ref xmlContent);

                // 5. 清除頁前的分段符號 (處理 scanbegin 後的空行)
                // 原 VBA: rng.Delete
                xmlContent = CleanScanBegin(ref xmlContent);

                // 6. 處理 scanend 後的分段符號
                // 原 VBA: Cut -> SetRange s,s -> Paste (即移到 scanend 之前)
                xmlContent = CleanScanEnd(ref xmlContent);

                //Clipboard.SetText(xmlContent);
                System.Media.SystemSounds.Beep.Play(); // 播放完成音效
                return true;
            }
            catch (Exception ex)
            {
                MessageBoxShowOKExclamationDefaultDesktopOnly("發生錯誤: " + ex.Message);
                return false;

            }
        }

        /// <summary>
        /// 處理第一頁的 XML 標籤 (包含自動清理與詢問機制)
        /// </summary>
        private static string SetPage1Code(ref string text)
        {
            // 1. 如果完全沒有 page="1"，嘗試補上 (這部分邏輯不變)
            if (!text.Contains("page=\"1\""))
            {
                var match = Regex.Match(text, "page=\"(\\d+)\"");
                if (match.Success && int.Parse(match.Groups[1].Value) < 10)
                {
                    var fileMatch = Regex.Match(text, "file=\"(\\d+)\"");
                    if (fileMatch.Success)
                    {
                        string bID = fileMatch.Groups[1].Value;
                        string header = $"<scanbegin file=\"{bID}\" page=\"1\" />●<scanend file=\"{bID}\" page=\"1\" />";
                        return header + text;
                    }
                }
                return text;
            }

            // 2. 如果有 page="1"，檢查內容是否需要清理
            // 抓取第一頁的完整標籤與內容
            string pattern = "(<scanbegin[^>]*page=\"1\"[^>]*>)([\\s\\S]*?)(<scanend[^>]*page=\"1\"[^>]*>)";

            return Regex.Replace(text, pattern, (m) =>
            {
                string beginTag = m.Groups[1].Value;
                string innerContent = m.Groups[2].Value; // 頁面中間的實際內容
                string endTag = m.Groups[3].Value;

                // --- 邏輯 A: 自動清除的強條件 (對應 VBA: InStr 與 Page1Exam_ContainsRegex) ---
                // 如果包含 "}}<scanskip" 或 "}}\r\n\r\n<scanbreak"，直接清除，不問使用者
                if (innerContent.Contains("}}<scanskip") ||
                    Regex.IsMatch(innerContent, @"\}\}\r\n\r\n<scanbreak"))
                {
                    return beginTag + "●" + endTag;
                }

                // --- 邏輯 B: 自動保留的條件 (對應 VBA: Page1Exam_NotContainsRegex) ---
                // 如果包含書名號《》，通常表示是正常的目錄或書名，不應清除
                if (Regex.IsMatch(innerContent, "[《》]"))
                {
                    return m.Value; // 保持原狀
                }

                // --- 邏輯 C: 模糊地帶，詢問使用者 (對應 VBA: MsgBox) ---
                // 既不符合自動刪除，也不符合自動保留，這時跳出視窗讓您決定
                DialogResult result = MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否清除第一頁的內容，以利連續自動輸入之進程？\n\n" +
                    "目前頁一的xml內容是：\n\n" +
                    innerContent.Trim(), // 顯示內容給使用者看 (Trim 去掉前後空白較整潔)
                    "TextForCtext - 第一頁清理確認");
                //DialogResult result = MessageBox.Show(
                //    "是否清除第一頁的內容，以利連續自動輸入之進程？\n\n" +
                //    "目前頁一的xml內容是：\n\n" +
                //    innerContent.Trim(), // 顯示內容給使用者看 (Trim 去掉前後空白較整潔)
                //    "TextForCtext - 第一頁清理確認",
                //    MessageBoxButtons.OKCancel,
                //    MessageBoxIcon.Question);

                if (result == DialogResult.OK)
                {
                    // 使用者按了「確定」，執行取代
                    return beginTag + "●" + endTag;
                }
                else
                {
                    // 使用者按了「取消」，保持原狀
                    return m.Value;
                }
            });
        }
        private static string SetPage1Code_Old(string text)
        {
            if (!text.Contains("page=\"1\""))
            {
                var match = Regex.Match(text, "page=\"(\\d+)\"");
                if (match.Success && int.Parse(match.Groups[1].Value) < 10)
                {
                    var fileMatch = Regex.Match(text, "file=\"(\\d+)\"");
                    if (fileMatch.Success)
                    {
                        string bID = fileMatch.Groups[1].Value;
                        string header = $"<scanbegin file=\"{bID}\" page=\"1\" />●<scanend file=\"{bID}\" page=\"1\" />";
                        return header + text;
                    }
                }
            }
            else
            {
                string pattern = "(<scanbegin[^>]*page=\"1\"[^>]*>)([\\s\\S]*?)(<scanend[^>]*page=\"1\"[^>]*>)";
                text = Regex.Replace(text, pattern, (m) =>
                {
                    string innerContent = m.Groups[2].Value;
                    if (Regex.IsMatch(innerContent, @"\}\}\r\n\r\n<scanbreak") || innerContent.Contains("}}<scanskip"))
                    {
                        return m.Groups[1].Value + "●" + m.Groups[3].Value;
                    }
                    return m.Value;
                });
            }
            return text;
        }

        private static string ClearRedundantCode(ref string text)
        {
            // 將 <scanend ...> ...雜訊... <scanbegin ...> 替換為 <scanend ...><scanbegin ...>
            // 讓它們緊貼，符合「上一頁的 end 要和下一頁的 begin 角括弧接在一起」
            return Regex.Replace(text,
                @"(<scanend[^>]+>)([\s\S]*?)(<scanbegin[^>]+>)",
                "$1$3",
                RegexOptions.IgnoreCase);
        }

        private static string MoveAsteriskUp(ref string text)
        {
            // 這裡是原本出錯的地方，現在分為兩步處理：

            // 【步驟 A：跨頁星號處理】
            // 如果結構是：<scanend...><scanbegin...>\r\n\r\n*
            // 則將 \r\n\r\n 移到 <scanend> 的前面。
            // $1 = scanend標籤, $2 = scanbegin標籤
            // 替換成：\r\n\r\n$1$2*
            text = Regex.Replace(text,
                @"(<scanend[^>]+>)(<scanbegin[^>]+>)\r\n\r\n\*",
                "\r\n\r\n$1$2*");

            // 【步驟 B：普通星號處理】 (如 scanbreak 或其他標籤)
            // 如果結構是：<標籤...>\r\n\r\n*
            // 則將 \r\n\r\n 移到 <標籤> 的前面。
            // $1 = 標籤本體
            // 替換成：\r\n\r\n$1*
            // 注意：這裡只會匹配到步驟 A 沒處理到的，因為步驟 A 處理後星號前已經沒有換行了。
            text = Regex.Replace(text,
                @"(<[^>]+>)\r\n\r\n\*",
                "\r\n\r\n$1*");

            return text;
        }

        private static string CleanScanBegin(ref string text)
        {
            // 對應 VBA: Do While rng.Find.Execute("<scanbegin ")... rng.Delete
            // 如果 scanbegin 後面有空行，直接刪除 (因為如果是星號的情況，上面 MoveAsteriskUp 已經先移走了，這裡處理的是純文字的情況)
            return Regex.Replace(text,
                @"(<scanbegin[^>]+>)\r\n\r\n",
                "$1");
        }

        private static string CleanScanEnd(ref string text)
        {
            // 對應 VBA: Do While ... <scanend ...> ... rng.Cut ... rng.Paste (移到前面)
            // 如果 scanend 後面有空行，移到 scanend 前面
            return Regex.Replace(text,
                @"(<scanend[^>]+>)\r\n\r\n",
                "\r\n\r\n$1");
        }
    }


    //https://gemini.google.com/share/dc92cc76c402

    /// <summary>
    /// 提取人名資料
    /// </summary>
    public static class NameExtractor
    {
        /// <summary>
        /// 即 WordVBA Sub 提取人名_二字人名中有空白者()
        /// 選取剪貼簿以操作，結果亦傳回剪貼簿中
        /// </summary>
        public static void ExtractNamesWithSpaces()
        {
            try
            {
                if (!Clipboard.ContainsText()) return;
                string input = Clipboard.GetText();

                // 1. 定義 Ctext 空白字元 w (U+10FFFD)
                string w = "\uDBFF\uDFFD";

                // 2. 定義「真正的漢字」範圍 (排除掉私人使用區的 w)
                // [ \u4E00-\u9FFF\u3400-\u4DBF\uF900-\uFAFF ] : 基本區 + 擴展A + 兼容區
                // [ \uD840-\uD888 ][ \uDC00-\uDFFF ] : 擴展B、C、D、E、F、G、H (U+20000 ~ U+323AF)
                // 這樣定義會跳過 Plane 16 的私人使用區 (也就是 w 所在的區域)
                string realCjk = @"(?:[\u4E00-\u9FFF\u3400-\u4DBF\uF900-\uFAFF]|[\uD840-\uD888][\uDC00-\uDFFF])";

                // 模式：(真漢字) + (w) + (真漢字)
                string pattern = "(" + realCjk + ")" + Regex.Escape(w) + "(" + realCjk + ")";

                StringBuilder sb = new StringBuilder();
                MatchCollection matches = Regex.Matches(input, pattern);

                foreach (Match m in matches)
                {
                    string char1 = m.Groups[1].Value;
                    string char2 = m.Groups[2].Value;

                    // 原始版：字􏿽字
                    string original = $"{char1}{w}{char2}";
                    // 排版版：字　字 (全形空格)
                    string formatted = $"{char1}　{char2}";

                    // 輸出格式：原始 [Tab] 修正
                    sb.AppendLine($"{original}\t{formatted}");
                }

                if (sb.Length > 0)
                {
                    Clipboard.SetText(sb.ToString());
                    MessageBox.Show($"提取完成！共找到 {matches.Count} 筆二字人名。\n已排除連串空白之誤判。", "成功");
                }
                else
                {
                    MessageBox.Show("未找到符合「漢字 + 􏿽 + 漢字」結構的內容。", "提示");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤: " + ex.Message);
            }
        }
    }
}