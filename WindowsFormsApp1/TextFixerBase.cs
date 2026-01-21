using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Drawing;

namespace TextForCtext
{
    public abstract class TextFixerBase
    {
        public bool EnableDebugReport { get; set; } = false;
        public Color HighlightColor { get; set; } = Color.Yellow;
        public bool AutoScrollToFirstHighlight { get; set; } = true;
        protected StringBuilder DebugBuilder { get; } = new StringBuilder();

        public string GetDebugReport() => DebugBuilder.ToString();
        public void ClearDebugReport() => DebugBuilder.Clear();

        public void WriteDebugReportToFile(string path)
        {
            try { File.WriteAllText(path, GetDebugReport(), Encoding.UTF8); }
            catch (Exception ex) { DebugBuilder.AppendLine("[WriteDebugReportToFile] 失敗：" + ex.Message); }
        }

        public abstract List<SpaceExamRange> CollectValidSequences(string text, out int totalSpaces);
        public abstract string ApplyInsertions(string original, List<SpaceExamRange> ranges, out List<SpaceExamRange> shiftedRanges, out string debugReport);

        public virtual List<HighlightRange> GetHighlightRangesForOriginal(string original, List<SpaceExamRange> ranges)
        {
            return new List<HighlightRange>();
        }
    }

    public struct SpaceExamRange { public int Start; public int Length; }
    public struct HighlightRange { public int Start; public int Length; }
}




/*
 
 TextForCtext 修復機制總覽
========================

目的：
- 修復《四庫全書電子版》等匯出文本中的排版錯置（如連續全形空格）
- 提供可擴充的修復架構，支援多種修復器（空格、標點、標記等）
- 保持 UI 友善：插入提示、原文高亮、可選顏色、自動滾動

架構：
- TextFixerBase：基底類，提供共用設定（EnableDebugReport、HighlightColor、AutoScrollToFirstHighlight）、Debug 輸出、檔案寫入。
- FullWidthSpaceFixer：空格修復器，支援：
  - 排除區段（ExclusionZones：如 *...<p>、{...}）
  - 插入策略互斥（InsertionMode：HalfSpace / CheckMark / None）
  - 允許半形空格檢測（AllowHalfWidthSpace）
  - Debug 報告（插入位置、位移後索引）
- 呼叫端（JustForTest_）：採用「文字定位式高亮」，避免 \r\n、代理對、全形字造成偏移。

工作流程：
1. 取得原始文本 originalRaw（textBox1.Text）
2. 正規化換行 normalized = originalRaw.Replace("\r\n", "\n")
3. CollectValidSequences(normalized, out totalSpaces)
4. ApplyInsertions(normalized, ranges, out shiftedRanges, out debugReport)
5. textBox1 顯示插入後 processedNormalized.Replace("\n", Environment.NewLine)
6. richTextBox1 顯示原始 originalRaw
7. 文字定位式高亮：對每段序列取「前一字元 + 空格序列 + 下一字元」片段，在 richTextBox1.Text 中搜尋並高亮
8. 若啟用 AutoScrollToFirstHighlight，滾動到第一處
9. 輸出 Debug 報告（MessageBox、Clipboard、或檔案）

擴充建議：
- 新增 PunctuationFixer：修復標點錯置（如全形/半形混用）
- 新增 MarkerFixer：修復標記遺漏（如 <p>、{{...}}）
- 將 ExclusionZones 改為可配置（讀取 JSON 或設定檔）
- 將高亮顏色與策略暴露到 UI（使用者可選）

注意：
- RichTextBox 的 Select(start, length) 以 UTF-16 索引為準；避免索引式高亮，改用文字定位式。
- 代理對（surrogate pairs）與全形字可能造成索引偏移；文字定位式可避免此問題。
- Debug 報告可協助比對長篇文本的插入位置與位移後索引。
https://copilot.microsoft.com/shares/TXSjwCTHoPy4PWHhxYvYb https://copilot.microsoft.com/shares/F3bX1ZnGjTZkn8xgywqxa 20260121
 */

