using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using WindowsFormsApp1;
using static WindowsFormsApp1.Form1;
using static TextForCtext.AncientTextExamine;

namespace TextForCtext
{
    /// <summary>
    /// 提供計算文本長度的輔助方法（正文、注文、異常行長度檢查）
    /// </summary>
    public static class TextLengthHelper
    {
        /// <summary>
        /// 表示非常行（段）的資訊：起點、選取長度、正常長度、異常行長
        /// </summary>
        public readonly struct AbnormalLineInfo
        {
            /// <summary>起點（lineSeprtStart）</summary>
            public int StartIndex { get; }

            /// <summary>選取長度（lineSeprtEnd - lineSeprtStart）</summary>
            public int Length { get; }

            /// <summary>通常長度（NormalLineParaLength）</summary>
            public int NormalLength { get; }

            /// <summary>異常行長（len）</summary>
            public int AbnormalLength { get; }

            public AbnormalLineInfo(int startIndex, int length, int normalLength, int abnormalLength)
            {
                StartIndex = startIndex;
                Length = length;
                NormalLength = normalLength;
                AbnormalLength = abnormalLength;
            }

            public override string ToString() =>
                $"Start={StartIndex}, Length={Length}, Normal={NormalLength}, Abnormal={AbnormalLength}";
        }

        /// <summary>
        /// 計算單行/段的字數
        /// </summary>
        internal static int CountWordsLenPerLinePara(string xLinePara)
        {
            if (Regex.IsMatch(xLinePara, "{{{.*?}}}"))
                xLinePara = Regex.Replace(xLinePara, "{{{.*?}}}", string.Empty);

            xLinePara = Regex.Replace(xLinePara, "＝.*?＝", string.Empty);
            xLinePara = Regex.Replace(xLinePara, "[*<p>|]", string.Empty);

            foreach (var item in Form1.PunctuationsNum)
                xLinePara = xLinePara.Replace(item.ToString(), "");

            int openCurly = xLinePara.IndexOf("{{");
            int closeCurly = xLinePara.IndexOf("}}");

            if (openCurly == -1 && closeCurly == -1)
                return new StringInfo(xLinePara).LengthInTextElements;

            if (openCurly == 0 && closeCurly == xLinePara.Length - 2)
                return new StringInfo(xLinePara.Replace("{{", "").Replace("}}", "")).LengthInTextElements;

            if (openCurly > 0 && closeCurly == -1)
                return new StringInfo(xLinePara.Substring(0, openCurly)).LengthInTextElements +
                       CountNoteLen(xLinePara.Substring(openCurly + 2));

            if (openCurly == -1 && closeCurly < xLinePara.Length - 2)
                return CountNoteLen(xLinePara.Substring(0, closeCurly)) +
                       new StringInfo(xLinePara.Substring(closeCurly + 2)).LengthInTextElements;

            int countResult = 0;
            int cursor = 0;

            while (openCurly > -1 && closeCurly > -1)
            {
                string textSegment = xLinePara.Substring(cursor, openCurly - cursor);
                countResult += new StringInfo(textSegment).LengthInTextElements;

                string noteSegment = xLinePara.Substring(openCurly + 2, closeCurly - (openCurly + 2));
                countResult += CountNoteLen(noteSegment);

                cursor = closeCurly + 2;
                openCurly = xLinePara.IndexOf("{{", cursor);
                closeCurly = xLinePara.IndexOf("}}", cursor);
            }

            if (cursor < xLinePara.Length)
                countResult += new StringInfo(xLinePara.Substring(cursor)).LengthInTextElements;

            return countResult;
        }

        /// <summary>
        /// 計算注文長度（每兩字算一單位，若有餘數則進位）
        /// </summary>
        internal static int CountNoteLen(string notePure)
        {
            int l = new StringInfo(notePure).LengthInTextElements;
            int quotient = l / 2;
            int remainder = l % 2;
            return remainder == 0 ? quotient : quotient + 1;
        }

        /// <summary>
        /// 簡化版：檢查非常長度的行（段）
        /// </summary>
        public static AbnormalLineInfo? CheckAbnormalLinePara(string xChk, int normalLength)
        {
            string[] xLineParas = Regex.Split(xChk, @"\r?\n");
            foreach (string line in xLineParas)
            {
                int l = CountWordsLenPerLinePara(line);
                if (l - normalLength > 4)
                {
                    int startIndex = xChk.IndexOf(line);
                    return new AbnormalLineInfo(startIndex, l, normalLength, l);
                }
            }
            return null;
        }

        /// <summary>
        /// 複雜版：包含手動輸入模式、OCR、跨頁注文等特殊情況
        /// </summary>
        public static AbnormalLineInfo? CheckAbnormalLinePara(string xChk)
        {
            if (!InstanceForm1.FastMode) InstanceForm1.SaveText();

            string[] xLineParas = xChk.Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);

            if (InstanceForm1.KeyinTextMode)
            {
                InstanceForm1.Lines_perPage = CountLinesPerPage(xChk);
                InstanceForm1.LinesParasPerPage = InstanceForm1.Lines_perPage;
                xLineParas = xLineParas.Where(x => x.Trim('　') != "").ToArray();
            }
            else
            {
                InstanceForm1.Lines_perPage = (InstanceForm1.LinesParasPerPage > 0) ? InstanceForm1.LinesParasPerPage : CountLinesPerPage(xChk);
                if (InstanceForm1.LinesParasPerPage == -1) InstanceForm1.LinesParasPerPage = InstanceForm1.Lines_perPage;
            }

            if (InstanceForm1.NormalLineParaLength == 0)
            {
                if (InstanceForm1.WordsPerLinePara != -1) InstanceForm1.NormalLineParaLength = InstanceForm1.WordsPerLinePara;
                else if (xLineParas.Length > 0)
                    InstanceForm1.NormalLineParaLength = CountWordsLenPerLinePara(xLineParas.Last());
            }

            if (InstanceForm1.NormalLineParaLength < 4)
            {
                if (InstanceForm1.KeyinTextMode) { InstanceForm1.NormalLineParaLength = 0; InstanceForm1.WordsPerLinePara = -1; }
                return null;
            }

            for (int i = 0; i < xLineParas.Length; i++)
            {
                string line = xLineParas[i];
                int len = CountWordsLenPerLinePara(line);
                int gap = Math.Abs(len - InstanceForm1.NormalLineParaLength);

                const int gapRef = 0;

                if (gap > gapRef && IsAbnormalLine(line, len))
                {
                    int startIndex = xChk.IndexOf(line);
                    return new AbnormalLineInfo(startIndex, len, InstanceForm1.NormalLineParaLength, len);
                }
            }

            return null;
        }

        private static bool IsAbnormalLine(string line, int len)
        {
            if (line == "　") return false;
            if (line.Contains("*") || line.Contains("|")) return false;
            if (len < InstanceForm1.NormalLineParaLength && line.Contains("<p>")) return false;
            return true;
        }
    }
}

//https://copilot.microsoft.com/shares/kHPV5udSdMjrBTHiYtthk
//https://copilot.microsoft.com/shares/KeGExaXHDraKCNfwUxkfJ

//https://copilot.microsoft.com/shares/HSuFN8WURrN3XKSuugheL
//https://copilot.microsoft.com/shares/QvQMxUuaqDzHpDENNZhqo