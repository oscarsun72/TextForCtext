using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace TextForCtext
{
    public class LineFormatter
    {
        public class LineFormatRule
        {
            public string MarkersAtStart { get; set; }  // 新增：行首標記
            public Dictionary<int, string> MarkersAfterHanCount { get; } = new Dictionary<int, string>();
            public string MarkersAtEnd { get; set; }
        }

        public static LineFormatRule InferRuleFromFirstLine(string raw, string formatted)
        {
            var rule = new LineFormatRule();

            var rawElems = EnumerateTextElements(raw).ToList();
            var fmtElems = EnumerateTextElements(formatted).ToList();

            int iRaw = 0;
            int iFmt = 0;
            int hanCount = 0;

            while (iRaw < rawElems.Count && iFmt < fmtElems.Count)
            {
                string eRaw = rawElems[iRaw];
                string eFmt = fmtElems[iFmt];

                if (TextElementEquals(eRaw, eFmt))
                {
                    if (IsHanTextElement(eRaw))
                        hanCount++;

                    iRaw++;
                    iFmt++;
                }
                else
                {
                    int markerStartHanCount = hanCount;
                    var markerSb = new StringBuilder();

                    while (iFmt < fmtElems.Count)
                    {
                        var currentFmt = fmtElems[iFmt];
                        if (iRaw < rawElems.Count && TextElementEquals(currentFmt, rawElems[iRaw]))
                            break;

                        markerSb.Append(currentFmt);
                        iFmt++;
                    }

                    string marker = markerSb.ToString();
                    if (!string.IsNullOrEmpty(marker))
                    {
                        if (markerStartHanCount == 0)
                        {
                            // 在第一個漢字之前出現的標記，記錄到行首
                            rule.MarkersAtStart = marker;
                        }
                        else
                        {
                            if (rule.MarkersAfterHanCount.ContainsKey(markerStartHanCount))
                                rule.MarkersAfterHanCount[markerStartHanCount] += marker;
                            else
                                rule.MarkersAfterHanCount[markerStartHanCount] = marker;
                        }
                    }
                }
            }

            if (iFmt < fmtElems.Count)
            {
                var tailSb = new StringBuilder();
                for (; iFmt < fmtElems.Count; iFmt++)
                    tailSb.Append(fmtElems[iFmt]);

                var tail = tailSb.ToString();
                if (!string.IsNullOrEmpty(tail))
                    rule.MarkersAtEnd = tail;
            }

            return rule;
        }

        public static string ApplyRuleToLine(string line, LineFormatRule rule)
        {
            var sb = new StringBuilder();
            int hanCount = 0;

            // 先加行首標記
            if (!string.IsNullOrEmpty(rule.MarkersAtStart))
                sb.Append(rule.MarkersAtStart);

            foreach (var elem in EnumerateTextElements(line))
            {
                sb.Append(elem);

                if (IsHanTextElement(elem))
                {
                    hanCount++;
                    string marker;
                    if (rule.MarkersAfterHanCount.TryGetValue(hanCount, out marker))
                        sb.Append(marker);
                }
            }

            if (!string.IsNullOrEmpty(rule.MarkersAtEnd))
                sb.Append(rule.MarkersAtEnd);

            return sb.ToString();
        }

        private static IEnumerable<string> EnumerateTextElements(string s)
        {
            var enumerator = StringInfo.GetTextElementEnumerator(s);
            while (enumerator.MoveNext())
                yield return enumerator.GetTextElement();
        }

        private static bool TextElementEquals(string a, string b)
        {
            return string.Equals(a, b, StringComparison.Ordinal);
        }

        private static bool IsHanTextElement(string elem)
        {
            if (string.IsNullOrEmpty(elem))
                return false;

            int codePoint = GetFirstCodePoint(elem);

            if (codePoint >= 0x4E00 && codePoint <= 0x9FFF) return true;
            if (codePoint >= 0x3400 && codePoint <= 0x4DBF) return true;
            if (codePoint >= 0x20000 && codePoint <= 0x3134F) return true;
            if (codePoint >= 0xF900 && codePoint <= 0xFAFF) return true;

            return false;
        }

        private static int GetFirstCodePoint(string elem)
        {
            if (elem.Length >= 2 && char.IsHighSurrogate(elem[0]) && char.IsLowSurrogate(elem[1]))
                return char.ConvertToUtf32(elem[0], elem[1]);
            return (int)elem[0];
        }
    }
}


//https://copilot.microsoft.com/shares/Hbe6F7c4AJ4ZsiALqrkBm
//Ctrl + Shift + Alt + f：根據第一行/段的排版格式套用到後面的內容各行/段中 (f: lineFormatter 的 f formatter) 20260117 https://copilot.microsoft.com/shares/Kr7ZFUZ2aaKQuHofeUxaR