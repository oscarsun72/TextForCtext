/* 
 * 將XML標記文本每頁最後一行/段搬移到下一頁的第一行/段。感恩感恩　讚歎讚歎　Copilot大菩薩　南無阿彌陀佛　讚美主
 * 如此2頁：https://ctext.org/library.pl?if=gb&file=195799&page=69 
 * https://ctext.org/library.pl?if=gb&file=30127&page=66 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TextForCtext
{
    public enum MoveDirection
    {
        Forward,
        Backward
    }

    public struct PagePair
    {
        public string FromPage;
        public string ToPage;
        public PagePair(string from, string to)
        {
            FromPage = from;
            ToPage = to;
        }
    }
    /// <summary>
    /// 處置XML文本中與頁碼相關的內容    
    /// </summary>
    public static class PageRearranger
    {
        /// <summary>
        /// 將XML文本前後頁相鄰行移動之功能
        /// 將指定行/段之XML標記及其內容搬到後一頁之首或前一頁之末
        /// </summary>
        /// <param name="xmlContent">要處理的XML文本</param>
        /// <param name="count">指定的行/段數</param>
        /// <param name="direction">要向後或向前搬移</param>
        /// <returns></returns>
        public static string Rearrange(string xmlContent, int count = 1, MoveDirection direction = MoveDirection.Forward)
        {
            if (count <= 0) return xmlContent;

            var pagePattern = new Regex(@"<scanbegin\b.*?>.*?<scanend\b.*?>", RegexOptions.Singleline);
            var matches = pagePattern.Matches(xmlContent);
            var pages = matches.Cast<Match>().Select(m => m.Value).ToList();
            if (pages.Count == 0) return xmlContent;

            for (int i = 0; i < pages.Count; i++)
            {
                if (direction == MoveDirection.Forward && i < pages.Count - 1)
                {
                    var pair = MoveSegmentsForward(pages[i], pages[i + 1], count);
                    pages[i] = pair.FromPage;
                    pages[i + 1] = pair.ToPage;
                }
                else if (direction == MoveDirection.Backward && i > 0)
                {
                    var pair = MoveSegmentsBackward(pages[i], pages[i - 1], count);
                    pages[i] = pair.FromPage;
                    pages[i - 1] = pair.ToPage;
                }
            }

            return string.Join("", pages);
        }

        // Forward: 從 fromPage 取最後 count 段，插入到 toPage 的 <scanbegin ...> '>' 之後
        private static PagePair MoveSegmentsForward(string fromPage, string toPage, int count)
        {
            var breakPattern = new Regex(@"<scanbreak\b", RegexOptions.Compiled);
            var breakMatches = breakPattern.Matches(fromPage).Cast<Match>().Select(m => m.Index).ToList();
            if (breakMatches.Count == 0) return new PagePair(fromPage, toPage);

            int take = Math.Min(count, breakMatches.Count);
            var selectedStarts = breakMatches.Skip(breakMatches.Count - take).ToList();

            int scanendIndex = IndexOfTag(fromPage, "<scanend");
            if (scanendIndex < 0) return new PagePair(fromPage, toPage);

            // 收集 segments (start, length) 並同時建立要插入的字串（已對調 tag 與內容）
            var segments = new List<Tuple<int, int>>();
            var movingStrings = new List<string>();
            foreach (int start in selectedStarts)
            {
                int tagEnd = fromPage.IndexOf('>', start);
                if (tagEnd < 0) tagEnd = start; // 保險
                // 找下一個 <scanbreak 在 start 之後且在 scanend 之前
                int nextBreak = breakMatches.FirstOrDefault(b => b > start);
                int endPos = (nextBreak > 0 && nextBreak < scanendIndex) ? nextBreak : scanendIndex;
                int len = endPos - start;
                if (len < 0) len = 0;
                segments.Add(Tuple.Create(start, len));

                // 分割 tag 與內容，並對調順序：內容 + tag
                int contentStart = tagEnd + 1;
                int contentLen = endPos - contentStart;
                string tagStr = fromPage.Substring(start, tagEnd - start + 1);
                string contentStr = (contentLen > 0 && contentStart < fromPage.Length) ? fromPage.Substring(contentStart, Math.Max(0, contentLen)) : "";
                movingStrings.Add(contentStr + tagStr);
            }

            // 刪除原段（從後往前）
            for (int r = segments.Count - 1; r >= 0; r--)
            {
                var seg = segments[r];
                if (seg.Item2 > 0 && seg.Item1 >= 0 && seg.Item1 + seg.Item2 <= fromPage.Length)
                    fromPage = fromPage.Remove(seg.Item1, seg.Item2);
            }

            // 插入到 toPage 的 <scanbegin ...> 的 '>' 之後（緊接位置）
            int beginIdx = IndexOfTag(toPage, "<scanbegin");
            if (beginIdx >= 0)
            {
                int closePos = toPage.IndexOf('>', beginIdx);
                int insertPos = (closePos >= 0) ? closePos + 1 : 0;
                string toInsert = string.Join("", movingStrings);
                toPage = toPage.Insert(insertPos, toInsert);
            }
            else
            {
                // 若找不到 scanbegin，則插到開頭
                toPage = string.Join("", movingStrings) + toPage;
            }

            return new PagePair(fromPage, toPage);
        }

        // Backward: 從 fromPage 取最前 count 段，插入到 toPage 的 <scanend ...> 之前
        private static PagePair MoveSegmentsBackward(string fromPage, string toPage, int count)
        {
            var breakPattern = new Regex(@"<scanbreak\b", RegexOptions.Compiled);
            var breakMatches = breakPattern.Matches(fromPage).Cast<Match>().Select(m => m.Index).ToList();
            if (breakMatches.Count == 0) return new PagePair(fromPage, toPage);

            int take = Math.Min(count, breakMatches.Count);
            var selectedStarts = breakMatches.Take(take).ToList();

            var segments = new List<Tuple<int, int>>();
            var movingStrings = new List<string>();
            for (int k = 0; k < selectedStarts.Count; k++)
            {
                int start = selectedStarts[k];
                int tagEnd = fromPage.IndexOf('>', start);
                if (tagEnd < 0) tagEnd = start;
                var nextBreak = breakMatches.FirstOrDefault(b => b > start);
                int endPos = (nextBreak > 0) ? nextBreak : fromPage.Length;
                int len = endPos - start;
                if (len < 0) len = 0;
                segments.Add(Tuple.Create(start, len));

                int contentStart = tagEnd + 1;
                int contentLen = endPos - contentStart;
                string tagStr = fromPage.Substring(start, tagEnd - start + 1);
                string contentStr = (contentLen > 0 && contentStart < fromPage.Length) ? fromPage.Substring(contentStart, Math.Max(0, contentLen)) : "";
                movingStrings.Add(contentStr + tagStr);
            }

            // 刪除原段（從後往前）
            for (int r = segments.Count - 1; r >= 0; r--)
            {
                var seg = segments[r];
                if (seg.Item2 > 0 && seg.Item1 >= 0 && seg.Item1 + seg.Item2 <= fromPage.Length)
                    fromPage = fromPage.Remove(seg.Item1, seg.Item2);
            }

            // 插入到 toPage 的 <scanend ...> 之前（即在 scanend 的起始位置插入）
            int insertPos = IndexOfTag(toPage, "<scanend");
            if (insertPos < 0) insertPos = toPage.Length;
            toPage = toPage.Insert(insertPos, string.Join("", movingStrings));

            return new PagePair(fromPage, toPage);
        }

        private static int IndexOfTag(string text, string tagStart)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(tagStart)) return -1;
            return text.IndexOf(tagStart, StringComparison.Ordinal);
        }
    }
}
//XMLPageRearranger: https://copilot.microsoft.com/shares/5AvwuBZfsbcTjwjbro3mo  20260124 https://copilot.microsoft.com/shares/5LfBiJYoQaheC1PERttWF