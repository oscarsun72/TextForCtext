using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using WindowsFormsApp1;

namespace TextForCtext
{
    /// <summary>
    /// 源文本全形空格修正器
    /// creedit_with_Copilot大菩薩　20260121
    /// </summary>
    public class FullWidthSpaceFixer : TextFixerBase
    {
        private const string _checkMark = "❌";

        public List<ExclusionZone> ExclusionZones { get; } = new List<ExclusionZone>
        {
            new ExclusionZone { StartMarker = "*", EndMarker = "<p>" },
            new ExclusionZone { StartMarker = "{", EndMarker = "}" }
        };

        public int RequiredCount { get; set; } = 3;
        public InsertionMode Mode { get; set; } = InsertionMode.HalfSpace;
        public char TargetSpaceChar { get; set; } = '　'; // U+3000
        public bool AllowHalfWidthSpace { get; set; } = false;

        /// <summary>
        /// Ctrl + Shift + Alt + x : 檢查textBox1文本中指定連續全形空格位置不當者，起於對《四庫全書》源文本標題錯置者之檢查，如《曝書亭集》卷六) x:Exam
        /// </summary>
        /// <param name="originalRaw"></param>
        /// <param name="textBox1"></param>
        /// <param name="richTextBox1"></param>
        /// <returns>若通過檢查則傳回false（不需檢查）</returns>
        public static bool ExamFullWidthSpaceSequencesFormattingError_inserttoCheck_HighlighttoWatch(
            string originalRaw, System.Windows.Forms.TextBox textBox1, System.Windows.Forms.RichTextBox richTextBox1)
        {
            // 保持 richTextBox1 與 textBox1 的外觀一致（字型、大小）
            richTextBox1.Size = textBox1.Size;
            richTextBox1.Location = textBox1.Location;
            richTextBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, richTextBox1.Font.Size);

            var fixer = new FullWidthSpaceFixer
            {
                RequiredCount = 3,
                Mode = InsertionMode.HalfSpace,
                AllowHalfWidthSpace = false,
                //EnableDebugReport = true,// ← 平常註解掉，需要時再打開
                HighlightColor = System.Drawing.Color.Yellow,
                AutoScrollToFirstHighlight = true
            };

            string normalized = originalRaw.Replace("\r\n", "\n");

            var rangesNormalized = fixer.CollectValidSequences(normalized, out int totalSpaces);
            var processedNormalized = fixer.ApplyInsertions(normalized, rangesNormalized, out var shiftedRangesNormalized, out string debugReport);

            textBox1.Text = processedNormalized.Replace("\n", Environment.NewLine);
            richTextBox1.Text = originalRaw;

            // 清除既有高亮
            richTextBox1.SelectAll();
            richTextBox1.SelectionBackColor = richTextBox1.BackColor;
            richTextBox1.SelectionLength = 0;

            int firstHighlightPos = -1;
            foreach (var r in rangesNormalized)
            {
                string before = r.Start > 0 ? normalized.Substring(r.Start - 1, 1) : "";
                string spaces = normalized.Substring(r.Start, r.Length);
                string after = (r.Start + r.Length < normalized.Length) ? normalized.Substring(r.Start + r.Length, 1) : "";
                string snippet = before + spaces + after;

                int pos = richTextBox1.Text.IndexOf(snippet, StringComparison.Ordinal);
                if (pos >= 0)
                {
                    if (firstHighlightPos < 0) firstHighlightPos = pos;
                    richTextBox1.Select(pos, snippet.Length);
                    richTextBox1.SelectionBackColor = fixer.HighlightColor;
                }
            }

            if (fixer.AutoScrollToFirstHighlight && firstHighlightPos >= 0)
            {
                richTextBox1.Select(firstHighlightPos, 0);
                richTextBox1.ScrollToCaret();
            }

            if (rangesNormalized.Count > 0)
            {
                richTextBox1.Show();
                return true;
            }
            else
            {
                //richTextBox1.Hide();
                return false;
            }

            #region 偵錯用

            //var sb = new System.Text.StringBuilder();
            //sb.AppendFormat("找到合格序列數: {0}\r\n", rangesNormalized.Count);
            //sb.AppendFormat("合格空格總數: {0}\r\n", totalSpaces);
            //sb.AppendLine();
            //sb.AppendLine("每段原始位置 (start, length) [normalized]:");
            //foreach (var r in rangesNormalized) sb.AppendFormat("  ({0}, {1})\r\n", r.Start, r.Length);
            //sb.AppendLine();
            //sb.AppendLine("位移後 shiftedRanges (start, length):");
            //foreach (var r in shiftedRangesNormalized) sb.AppendFormat("  ({0}, {1})\r\n", r.Start, r.Length);
            //sb.AppendLine();
            //sb.AppendLine("DebugReport:");
            //sb.AppendLine(debugReport);

            //System.Windows.Forms.Clipboard.SetText(sb.ToString());
            //System.Windows.Forms.MessageBox.Show(sb.ToString(), "FullWidthSpaceFixer 偵測結果", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            #endregion
        }

        private bool IsTargetSpace(char c)
        {
            if (AllowHalfWidthSpace)
                return char.IsWhiteSpace(c) || c == TargetSpaceChar;
            // 若要更寬鬆可改為 char.IsWhiteSpace(c)
            return c == TargetSpaceChar;
        }

        public override List<SpaceExamRange> CollectValidSequences(string text, out int totalSpaces)
        {
            ClearDebugReport();
            var ranges = new List<SpaceExamRange>();
            totalSpaces = 0;
            if (string.IsNullOrEmpty(text)) return ranges;

            int i = 0;
            while (i < text.Length)
            {
                if (IsTargetSpace(text[i]))
                {
                    int start = i;
                    int runLen = 0;
                    while (i < text.Length && IsTargetSpace(text[i]))
                    {
                        runLen++;
                        i++;
                    }

                    if (runLen >= RequiredCount)
                    {
                        bool valid = IsValid(text, start, runLen);
                        if (EnableDebugReport)
                        {
                            DebugBuilder.AppendFormat("Candidate start={0}, len={1}, IsValid={2}\r\n", start, runLen, valid);
                            int ctxStart = Math.Max(0, start - 6);
                            int ctxLen = Math.Min(runLen + 12, text.Length - ctxStart);
                            DebugBuilder.AppendFormat("  Context: [{0}]\r\n", text.Substring(ctxStart, ctxLen).Replace("\n", "\\n"));
                        }
                        if (valid)
                        {
                            ranges.Add(new SpaceExamRange { Start = start, Length = runLen });
                            totalSpaces += runLen;
                        }
                    }
                }
                else
                {
                    i++;
                }
            }

            return ranges;
        }

        public override string ApplyInsertions(string original, List<SpaceExamRange> ranges, out List<SpaceExamRange> shiftedRanges, out string debugReport)
        {
            var sb = new StringBuilder(original ?? string.Empty);
            shiftedRanges = new List<SpaceExamRange>();
            int offset = 0;

            foreach (var r in ranges.OrderBy(r => r.Start))
            {
                int insertPos = r.Start + offset;
                DebugBuilder.AppendFormat("原始序列: start={0}, len={1}, insertPos(before)={2}\r\n", r.Start, r.Length, insertPos);

                switch (Mode)
                {
                    case InsertionMode.HalfSpace:
                        sb.Insert(insertPos, ' ');
                        DebugBuilder.AppendFormat("  插入半形空格 at {0}\r\n", insertPos);
                        offset += 1;
                        shiftedRanges.Add(new SpaceExamRange { Start = insertPos + 1, Length = r.Length });
                        break;

                    case InsertionMode.CheckMark:
                        sb.Insert(insertPos, _checkMark);
                        DebugBuilder.AppendFormat("  插入檢查標記(❌) at {0}\r\n", insertPos);
                        offset += _checkMark.Length;
                        shiftedRanges.Add(new SpaceExamRange { Start = insertPos + _checkMark.Length, Length = r.Length });
                        break;

                    case InsertionMode.None:
                    default:
                        DebugBuilder.AppendFormat("  不插入（保留原始位置） at {0}\r\n", insertPos);
                        shiftedRanges.Add(new SpaceExamRange { Start = insertPos, Length = r.Length });
                        break;
                }
            }

            debugReport = GetDebugReport();
            return sb.ToString();
        }

        private bool IsValid(string text, int start, int runLen)
        {
            if (string.IsNullOrEmpty(text)) return false;

            // 排除區段 判斷與您原先版本一致（使用 start-1）
            foreach (var zone in ExclusionZones)
            {
                int lastStart = -1;
                int lastEnd = -1;
                if (start - 1 >= 0)
                    lastStart = text.LastIndexOf(zone.StartMarker, start - 1, StringComparison.Ordinal);
                if (start - 1 >= 0)
                    lastEnd = text.LastIndexOf(zone.EndMarker, start - 1, StringComparison.Ordinal);

                if (lastStart != -1 && (lastEnd == -1 || lastStart > lastEnd))
                {
                    if (EnableDebugReport) DebugBuilder.AppendFormat("  排除: 在排除區段 {0}\r\n", zone.StartMarker);
                    return false;
                }
            }

            // 行內含星號則排除（與原版一致）
            int lastNewline = text.LastIndexOf('\n', Math.Max(0, start - 1));
            int lineStart = (lastNewline == -1) ? 0 : lastNewline + 1;
            int nextNewline = text.IndexOf('\n', start);
            int lineEnd = (nextNewline == -1) ? text.Length : nextNewline;
            if (lineEnd < lineStart) lineEnd = lineStart;
            string line = text.Substring(lineStart, lineEnd - lineStart);
            if (line.Contains("*"))
            {
                if (EnableDebugReport) DebugBuilder.AppendFormat("  排除: 行含星號 -> [{0}]\r\n", line.Replace("\n", "\\n"));
                return false;
            }

            // 前一字元不可為換行或目標空格
            if (start == 0)
            {
                if (EnableDebugReport) DebugBuilder.AppendLine("  排除: start == 0");
                return false;
            }
            char prev = text[start - 1];
            if (prev == '\n' || prev == '\r' || IsTargetSpace(prev))
            {
                if (EnableDebugReport) DebugBuilder.AppendFormat("  排除: 前一字元為換行或空格 U+{0:X4}\r\n", (int)prev);
                return false;
            }

            // 新增條件：字數差異 ≤ 4
            int l = TextLengthHelper.CountWordsLenPerLinePara(line);
            int normalLength = Form1.InstanceForm1.NormalLineParaLength;//l - runLen;
            if (l - normalLength > 4)//誤差4，由 CheckAbnormalLinePara 方法來決定
            {
                if (EnableDebugReport)
                    DebugBuilder.AppendFormat("  排除: 行字數差異過大 (l={0}, normal={1})\r\n", l, normalLength);
                return false;
            }

            return true;
        }

        public override List<HighlightRange> GetHighlightRangesForOriginal(string originalText, List<SpaceExamRange> originalRanges)
        {
            var highlights = new List<HighlightRange>();
            if (string.IsNullOrEmpty(originalText) || originalRanges == null || originalRanges.Count == 0)
                return highlights;

            var si = new StringInfo(originalText);

            foreach (var r in originalRanges)
            {
                int elementIndex = 0;
                int charPos = 0;
                while (charPos < r.Start && elementIndex < si.LengthInTextElements)
                {
                    charPos += si.SubstringByTextElements(elementIndex, 1).Length;
                    elementIndex++;
                }

                int prevLen = elementIndex > 0 ? si.SubstringByTextElements(elementIndex - 1, 1).Length : 0;

                int endElementIndex = elementIndex;
                int endCharPos = charPos;
                int targetEnd = r.Start + r.Length;
                while (endCharPos < targetEnd && endElementIndex < si.LengthInTextElements)
                {
                    endCharPos += si.SubstringByTextElements(endElementIndex, 1).Length;
                    endElementIndex++;
                }

                int nextLen = endElementIndex < si.LengthInTextElements ? si.SubstringByTextElements(endElementIndex, 1).Length : 0;

                int highlightStart = Math.Max(0, r.Start - prevLen);
                int highlightLength = prevLen + r.Length + nextLen;
                if (highlightStart + highlightLength > originalText.Length)
                    highlightLength = originalText.Length - highlightStart;
                if (highlightLength <= 0) continue;

                highlights.Add(new HighlightRange { Start = highlightStart, Length = highlightLength });
            }

            return highlights;
        }
    }

    public class ExclusionZone { public string StartMarker; public string EndMarker; }
    public enum InsertionMode { None, HalfSpace, CheckMark }
}

//https://copilot.microsoft.com/shares/gWdFMKkRVTazmVXsHpLzt
//https://copilot.microsoft.com/shares/cJ3LdPep6x4Djp2mUDyMt
//https://copilot.microsoft.com/shares/j68B9rn64K8agmxW62XCX




//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.Linq;
//using System.Text;

//namespace TextForCtext
//{
//    public class FullWidthSpaceFixer
//    {
//        private const string _checkMark = "❌"; // 註解保留

//        public List<ExclusionZone> ExclusionZones { get; } = new List<ExclusionZone>
//        {
//            new ExclusionZone { StartMarker = "*", EndMarker = "<p>" },
//            new ExclusionZone { StartMarker = "{", EndMarker = "}" }
//        };

//        public int RequiredCount { get; set; } = 3;
//        public InsertionMode Mode { get; set; } = InsertionMode.HalfSpace;
//        public char TargetSpaceChar { get; set; } = '　'; // 全形空格
//        public bool AllowHalfWidthSpace { get; set; } = false;
//        public bool EnableDebugReport { get; set; } = false;

//        private bool IsTargetSpace(char c)
//        {
//            if (AllowHalfWidthSpace)
//                return c == TargetSpaceChar || c == ' ';
//            return c == TargetSpaceChar;
//        }

//        public List<SpaceExamRange> CollectValidSequences(string text, out int totalSpaces)
//        {
//            var ranges = new List<SpaceExamRange>();
//            totalSpaces = 0;
//            if (string.IsNullOrEmpty(text)) return ranges;

//            int i = 0;
//            while (i < text.Length)
//            {
//                if (IsTargetSpace(text[i]))
//                {
//                    int start = i;
//                    int runLen = 0;
//                    while (i < text.Length && IsTargetSpace(text[i]))
//                    {
//                        runLen++;
//                        i++;
//                    }

//                    if (runLen >= RequiredCount && IsValid(text, start, runLen))
//                    {
//                        ranges.Add(new SpaceExamRange { Start = start, Length = runLen });
//                        totalSpaces += runLen;
//                    }
//                }
//                else
//                {
//                    i++;
//                }
//            }

//            return ranges;
//        }

//        // ApplyInsertions 回傳 debugReport（供測試用）
//        public string ApplyInsertions(string original, List<SpaceExamRange> ranges, out List<SpaceExamRange> shiftedRanges, out string debugReport)
//        {
//            var sb = new StringBuilder(original ?? string.Empty);
//            shiftedRanges = new List<SpaceExamRange>();
//            int offset = 0;
//            var dbg = new StringBuilder();

//            foreach (var r in ranges.OrderBy(r => r.Start))
//            {
//                int insertPos = r.Start + offset;
//                if (EnableDebugReport)
//                    dbg.AppendFormat("原始序列: start={0}, len={1}, insertPos(before)={2}\r\n", r.Start, r.Length, insertPos);

//                switch (Mode)
//                {
//                    case InsertionMode.HalfSpace:
//                        sb.Insert(insertPos, ' ');
//                        offset += 1;
//                        shiftedRanges.Add(new SpaceExamRange { Start = insertPos + 1, Length = r.Length });
//                        if (EnableDebugReport) dbg.AppendFormat("  插入半形空格 at {0}\r\n", insertPos);
//                        break;

//                    case InsertionMode.CheckMark:
//                        sb.Insert(insertPos, _checkMark);
//                        offset += _checkMark.Length;
//                        shiftedRanges.Add(new SpaceExamRange { Start = insertPos + _checkMark.Length, Length = r.Length });
//                        if (EnableDebugReport) dbg.AppendFormat("  插入檢查標記 at {0}\r\n", insertPos);
//                        break;

//                    case InsertionMode.None:
//                    default:
//                        shiftedRanges.Add(new SpaceExamRange { Start = insertPos, Length = r.Length });
//                        if (EnableDebugReport) dbg.AppendFormat("  未插入 (None)\r\n");
//                        break;
//                }
//            }

//            debugReport = dbg.ToString();
//            return sb.ToString();
//        }

//        private bool IsValid(string text, int start, int runLen)
//        {
//            if (string.IsNullOrEmpty(text)) return false;

//            // 排除區段（簡潔判斷：若最近的 startMarker 在最近的 endMarker 之後，視為在區段內）
//            foreach (var zone in ExclusionZones)
//            {
//                int lastStart = -1;
//                int lastEnd = -1;
//                if (start - 1 >= 0)
//                    lastStart = text.LastIndexOf(zone.StartMarker, start - 1, StringComparison.Ordinal);
//                if (start - 1 >= 0)
//                    lastEnd = text.LastIndexOf(zone.EndMarker, start - 1, StringComparison.Ordinal);

//                if (lastStart != -1 && (lastEnd == -1 || lastStart > lastEnd))
//                    return false;
//            }

//            // 行內含星號則排除
//            int lastNewline = text.LastIndexOf('\n', Math.Max(0, start - 1));
//            int lineStart = (lastNewline == -1) ? 0 : lastNewline + 1;
//            int nextNewline = text.IndexOf('\n', start);
//            int lineEnd = (nextNewline == -1) ? text.Length : nextNewline;
//            if (lineEnd < lineStart) lineEnd = lineStart;
//            string line = text.Substring(lineStart, lineEnd - lineStart);
//            if (line.Contains("*")) return false;

//            // 前一字元不可為換行或目標空格
//            if (start == 0) return false;
//            char prev = text[start - 1];
//            if (prev == '\n' || prev == '\r' || IsTargetSpace(prev)) return false;

//            return true;
//        }

//        public List<HighlightRange> GetHighlightRangesForOriginal(string originalText, List<SpaceExamRange> originalRanges)
//        {
//            var highlights = new List<HighlightRange>();
//            if (string.IsNullOrEmpty(originalText) || originalRanges == null || originalRanges.Count == 0)
//                return highlights;

//            var si = new StringInfo(originalText);

//            foreach (var r in originalRanges)
//            {
//                int elementIndex = 0;
//                int charPos = 0;
//                while (charPos < r.Start && elementIndex < si.LengthInTextElements)
//                {
//                    charPos += si.SubstringByTextElements(elementIndex, 1).Length;
//                    elementIndex++;
//                }

//                int prevLen = elementIndex > 0 ? si.SubstringByTextElements(elementIndex - 1, 1).Length : 0;

//                int endElementIndex = elementIndex;
//                int endCharPos = charPos;
//                int targetEnd = r.Start + r.Length;
//                while (endCharPos < targetEnd && endElementIndex < si.LengthInTextElements)
//                {
//                    endCharPos += si.SubstringByTextElements(endElementIndex, 1).Length;
//                    endElementIndex++;
//                }

//                int nextLen = endElementIndex < si.LengthInTextElements ? si.SubstringByTextElements(endElementIndex, 1).Length : 0;

//                int highlightStart = Math.Max(0, r.Start - prevLen);
//                int highlightLength = prevLen + r.Length + nextLen;

//                if (highlightStart + highlightLength > originalText.Length)
//                    highlightLength = originalText.Length - highlightStart;
//                if (highlightLength <= 0) continue;

//                highlights.Add(new HighlightRange { Start = highlightStart, Length = highlightLength });
//            }

//            return highlights;
//        }

//        public List<HighlightRange> GetHighlightRangesForProcessed(string processedText, List<SpaceExamRange> shiftedRanges)
//        {
//            var highlights = new List<HighlightRange>();
//            if (string.IsNullOrEmpty(processedText) || shiftedRanges == null || shiftedRanges.Count == 0)
//                return highlights;

//            var si = new StringInfo(processedText);

//            foreach (var r in shiftedRanges)
//            {
//                int elementIndex = 0;
//                int charPos = 0;
//                while (charPos < r.Start && elementIndex < si.LengthInTextElements)
//                {
//                    charPos += si.SubstringByTextElements(elementIndex, 1).Length;
//                    elementIndex++;
//                }

//                int prevLen = elementIndex > 0 ? si.SubstringByTextElements(elementIndex - 1, 1).Length : 0;

//                int endElementIndex = elementIndex;
//                int endCharPos = charPos;
//                int targetEnd = r.Start + r.Length;
//                while (endCharPos < targetEnd && endElementIndex < si.LengthInTextElements)
//                {
//                    endCharPos += si.SubstringByTextElements(endElementIndex, 1).Length;
//                    endElementIndex++;
//                }

//                int nextLen = endElementIndex < si.LengthInTextElements ? si.SubstringByTextElements(endElementIndex, 1).Length : 0;

//                int highlightStart = Math.Max(0, r.Start - prevLen);
//                int highlightLength = prevLen + r.Length + nextLen;
//                if (highlightStart + highlightLength > processedText.Length)
//                    highlightLength = processedText.Length - highlightStart;
//                if (highlightLength <= 0) continue;

//                highlights.Add(new HighlightRange { Start = highlightStart, Length = highlightLength });
//            }

//            return highlights;
//        }
//    }

//    public struct SpaceExamRange { public int Start; public int Length; }
//    public struct HighlightRange { public int Start; public int Length; }
//    public class ExclusionZone { public string StartMarker; public string EndMarker; }
//    public enum InsertionMode { None, HalfSpace, CheckMark }
//}


// https://copilot.microsoft.com/shares/GgCFwg7wLx2HaKXAespby
