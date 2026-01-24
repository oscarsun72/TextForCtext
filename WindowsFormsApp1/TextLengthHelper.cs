using System;
using System.Globalization;
using System.Linq;
using System.Media;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1;
using static TextForCtext.AncientTextExamine;
using static WindowsFormsApp1.Form1;

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
            public int StartIndex { get; }
            public int Length { get; }
            public int NormalLength { get; }
            public int AbnormalLength { get; }

            public AbnormalLineInfo(int startIndex, int length, int normalLength, int abnormalLength)
            {
                StartIndex = startIndex;
                Length = length;
                NormalLength = normalLength;
                AbnormalLength = abnormalLength;
            }
        }

        /// <summary>
        /// 計算單行/段的字數
        /// 含空格空白；而標點符號、校按、標題標記、段落標記、分行標記均不計
        /// （即只計算原書內文之長度，以供比較行/段長度）
        /// </summary>
        internal static int CountWordsLenPerLinePara(string xLinePara)
        {
            string pattern;
            if (xLinePara.IndexOf("{{{") > -1 || xLinePara.IndexOf("}}}") > -1)
            {
                pattern = "{{{.*?}}}";
                xLinePara = Regex.Replace(xLinePara, pattern, string.Empty);
            }

            pattern = "＝.*?＝";
            xLinePara = Regex.Replace(xLinePara, pattern, string.Empty);

            xLinePara = Regex.Replace(xLinePara, "[*<p>|]", string.Empty);

            foreach (var item in PunctuationsNum)
            {
                xLinePara = xLinePara.Replace(item.ToString(), "");
            }

            int openCurly = xLinePara.IndexOf("{{");
            int closeCurly = xLinePara.IndexOf("}}");
            int s = 0, countResult = 0;
            string txt = "", note = "";

            if (openCurly == -1 && closeCurly == -1)
                return new StringInfo(xLinePara).LengthInTextElements;
            else if (!(openCurly == 0 && closeCurly == xLinePara.Length - 2)
                && !(openCurly == 0 && closeCurly == -1)
                && !(openCurly == -1 && closeCurly == xLinePara.Length - 2))
            {
                while (true)
                {
                    if (closeCurly > -1 &&
                        (openCurly > closeCurly) || openCurly == -1)
                    {
                        countResult += CountNoteLen(xLinePara.Substring(0, closeCurly));
                    }
                    else if (closeCurly > -1)
                    {
                        txt = xLinePara.Substring(s, openCurly - s);
                        countResult += new StringInfo(txt).LengthInTextElements;
                        note = xLinePara.Substring(openCurly + 2, closeCurly - (openCurly + 2));
                        countResult += CountNoteLen(note);
                    }
                    else if (closeCurly == -1 && openCurly > -1)
                    {
                        txt = xLinePara.Substring(s, openCurly);
                        countResult += new StringInfo(txt).LengthInTextElements;
                        note = xLinePara.Substring(openCurly + 2);
                        countResult += CountNoteLen(note);
                        break;
                    }
                    s = closeCurly + 2;
                    openCurly = xLinePara.IndexOf("{{", closeCurly);
                    if (openCurly == -1)
                    {
                        txt = (closeCurly + 2 > xLinePara.Length) ? "" : xLinePara.Substring(closeCurly + 2);
                        return countResult += new StringInfo(txt).LengthInTextElements;
                    }
                    closeCurly = xLinePara.IndexOf("}}", closeCurly + 2);
                    if (closeCurly == -1)
                    {
                        txt = xLinePara.Substring(s, openCurly - s);
                        countResult += new StringInfo(txt).LengthInTextElements;
                        note = xLinePara.Substring(openCurly + 2);
                        return countResult += CountNoteLen(note);
                    }
                    if (openCurly < 0) break;
                }
                return countResult;
            }
            else if (openCurly > 0 && closeCurly == -1)
            {
                return new StringInfo(xLinePara.Substring(0, openCurly)).LengthInTextElements +
                       CountNoteLen(xLinePara.Substring(openCurly + 2));
            }
            else if (closeCurly < xLinePara.Length - 2 && openCurly == -1)
            {
                return CountNoteLen(xLinePara.Substring(0, closeCurly)) +
                       new StringInfo(xLinePara.Substring(closeCurly + 2)).LengthInTextElements;
            }
            else
            {
                return new StringInfo(xLinePara.Replace("{{", "").Replace("}}", "")).LengthInTextElements;
            }
        }

        /// <summary>
        /// 注文長度：每兩字算一單位，餘數進位
        /// </summary>
        internal static int CountNoteLen(string notePure)
        {
            int l = new StringInfo(notePure).LengthInTextElements;
            int x = l / 2;   // 商數
            int y = l - (x * 2); // 餘數
            return y == 0 ? x : ++x;
        }

        /// <summary>
        /// 簡單版：檢查非常長度的行（段）
        /// </summary>
        public static AbnormalLineInfo? CheckAbnormalLinePara(string xChk, int normalLength)
        {
            string[] xLineParas = Regex.Split(xChk, @"\r?\n");
            foreach (string line in xLineParas)
            {
                int l = CountWordsLenPerLinePara(line);
                if (l - normalLength > 4)
                {
                    return new AbnormalLineInfo(xChk.IndexOf(line), l, normalLength, l);
                }
            }
            return null;
        }



        /// <summary>
        /// 檢查非常長度的行（段） 複雜版：包含手動輸入模式、OCR、跨頁注文等特殊情況
        /// </summary>
        /// <param name="xChk">這引數是指定要傳入檢查的文本</param>
        /// <returns>若發現非常長度的行，則傳回一個 AbnormalLineInfo 類別，其中4屬性分別表示非常行諸特徵：
        /// 1.lineSeprtStart（起點；Seprt=Separate） 2. lineSeprtEnd - lineSeprtStart（非常長度-位置長度） 此2者作為選取方法Select之參數，以選取異常行便於檢視
        /// 3. normalLineParaLength（通常長度） 4. len（長度-漢字長度）}此二者表正常行之漢字數與所發現之異常行之漢字數
        /// </returns>
        public static AbnormalLineInfo? CheckAbnormalLinePara(string xChk)
        {
            if (!InstanceForm1.FastMode) InstanceForm1.SaveText();//備份以防萬一
            string[] xLineParas = xChk.Split(
                Environment.NewLine.ToArray(),
                StringSplitOptions.RemoveEmptyEntries);

            #region get lines_perPage//取得該頁的每行（段）文字
            //lines_perPage = 0;
            if (InstanceForm1.KeyinTextMode)
            {//如果是手動輸入模式：
                InstanceForm1.Lines_perPage = CountLinesPerPage(xChk);
                InstanceForm1.LinesParasPerPage = InstanceForm1.Lines_perPage;
                //略過只含有「　」的元素
                xLineParas = xLineParas.Where(x => x.Trim('　') != "").ToArray();
                /* 20230907 Bing大菩薩：
                 太好了，我很高興能幫助您！LINQ 是 Language Integrated Query 的簡稱，它是一種用於查詢和操作數據的語言。在 C# 中，LINQ 可以用來查詢各種數據源，包括數組、列表、XML 文件等。
                    在上面的例子中，我們使用了 LINQ 的 `Where` 方法來過濾出 `xLineParas` 中不符合條件的元素。`Where` 方法接受一個委託作為參數，該委託定義了過濾條件。在這個例子中，我們定義了一個匿名函數 `x => x.Trim('　') != ""` 作為過濾條件。這個函數會對每個元素進行判斷，如果元素去掉前後的「　」後不為空字符串，則返回 `true`，否則返回 `false`。
                    `Where` 方法會根據過濾條件返回一個新的序列，其中只包含符合條件的元素。最後，我們使用 `ToArray` 方法將序列轉換為數組。
                    希望這對您有所幫助！如果您還有其他問題，請隨時告訴我。感恩感恩　讚歎讚歎　南無阿彌陀佛。
                */
            }
            else//不是手動輸入模式：
                InstanceForm1.Lines_perPage = (InstanceForm1.LinesParasPerPage != -1 && InstanceForm1.LinesParasPerPage != 0) ? InstanceForm1.LinesParasPerPage : CountLinesPerPage(xChk);
            if (InstanceForm1.LinesParasPerPage == -1) InstanceForm1.LinesParasPerPage = InstanceForm1.Lines_perPage;
            //lines_perPage = xLineParas.Length;
            /*
            foreach (string item in xLineParas)
            {
                if (lines_perPage == 0 & xChk.IndexOf("}}") < xChk.IndexOf("{{") && xChk.IndexOf("}}") > item.Length)
                    lines_perPage++;
                else if (item.Length > 4 && item.Substring(0, 2) == "{{" && item.Substring(item.Length - 2, 2) == "}}"
                        && item.Substring(2, item.Length - 4).IndexOf("{{") == -1 && item.Substring(2, item.Length - 4).IndexOf("}}") == -1)
                    lines_perPage++;
                else if (item.Length > 2 && (item.Substring(0, 2) == "{{" && item.IndexOf("}}") == -1 || item.Substring(item.Length - 2, 2) == "}}" && item.IndexOf("{{") == -1))
                    lines_perPage++;
                else
                    lines_perPage += 2;
            }
            */
            #endregion
            if (InstanceForm1.NormalLineParaLength == 0)
            {
                if (InstanceForm1.WordsPerLinePara != -1) InstanceForm1.NormalLineParaLength = InstanceForm1.WordsPerLinePara;
                else
                {
                    if (xLineParas.Length > 0)//通常第一行會有卷首篇題，故不準；最末行又可能收尾，故取其次末行
                        InstanceForm1.NormalLineParaLength = CountWordsLenPerLinePara(xLineParas[xLineParas.Length - 1]);// new StringInfo(xLineParas[0]).LengthInTextElements;
                }
            }

            /////暫時取消此條件，7改成4（即每行3字內，自行目測檢查。）20230822
            //if (normalLineParaLength < 7)
            if (InstanceForm1.NormalLineParaLength < 4)
            {//如果正常漢字數小於7則不執行
             //normalLineParaLength歸零、wordsPerLinePara歸零
                if (InstanceForm1.KeyinTextMode) { InstanceForm1.NormalLineParaLength = 0; InstanceForm1.WordsPerLinePara = -1; }
                return null;//new int[0];
            }

            int i = -1, len = 0;
            foreach (string lineParaText in xLineParas)
            {
                //if (lineParaText.IndexOf("竝當與") > -1) //just for check 
                //    Debugger.Break();


                i++;
                if (lineParaText.IndexOf("{{{") > -1 || lineParaText.IndexOf("孫守真") > -1 || lineParaText.IndexOf("＝") > -1)//{{{孫守真按：}}}、缺字說明等略去，以人工校對
                {
                    continue;
                }
                int noteTextBlendStart = lineParaText.IndexOf("{"),
                    noteTextBlendEnd = lineParaText.IndexOf("}");
                int gap;
                if (noteTextBlendStart != -1 || noteTextBlendEnd != -1)
                {//blend text and note                     
                    string text = "", note = "";
                    if (noteTextBlendStart != -1 && noteTextBlendEnd == -1)
                    {// {{ only
                        text = ClearOmitChar(lineParaText.Substring(0, noteTextBlendStart));
                        note = ClearOmitChar(lineParaText.Substring(noteTextBlendStart + 2));
                        if (text == "")
                            len = new StringInfo(note).LengthInTextElements;
                        else
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements / 2);
                    }
                    if (noteTextBlendStart == -1 && noteTextBlendEnd != -1)
                    {// }} only
                        note = ClearOmitChar(lineParaText.Substring(0, noteTextBlendEnd));
                        text = ClearOmitChar(lineParaText.Substring(noteTextBlendEnd + 2));
                        if (text == "")
                            len = new StringInfo(note).LengthInTextElements;
                        else
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements / 2);
                    }
                    if (noteTextBlendStart != -1 && noteTextBlendEnd != -1)
                    {// {{ and }} both
                        if (noteTextBlendStart == 0 && noteTextBlendEnd + "}}".Length == lineParaText.Length)
                        {
                            note += lineParaText.Substring
                                (noteTextBlendStart + 2,
                                noteTextBlendEnd == -1 ?
                                lineParaText.Length - (noteTextBlendStart + 2)
                                : noteTextBlendEnd - (noteTextBlendStart + 2));
                            note += new String('　', CountWordsLenPerLinePara(note));//單行注文則補上空格以方便計算字數
                            len = CountWordsLenPerLinePara(note) / 2;
                        }
                        else if (noteTextBlendStart < noteTextBlendEnd)
                        {
                            int st = 0, lText = noteTextBlendStart;
                            while (noteTextBlendStart != -1)
                            {
                                text += lineParaText.Substring(st,
                                    lText);
                                note += lineParaText.Substring
                                    (noteTextBlendStart + 2,
                                    noteTextBlendEnd == -1 ?
                                    lineParaText.Length - (noteTextBlendStart + 2)
                                    : noteTextBlendEnd - (noteTextBlendStart + 2));
                                if (CountWordsLenPerLinePara(note) % 2 == 1)
                                    note += "　";//如果一行中有兩處注文以上，可能剛好都缺1字（即均為單數長，又剛好有2的倍數量），造成字數統計上的失誤，如此例：　斗字作斤{{詳前《急就篇》}}與什形近{{《説文·敘》云：人持十為斗。}}此什卽斗字趙
                                                //故補上空白以供計算
                                noteTextBlendStart = lineParaText.IndexOf
                                    ("{", noteTextBlendStart + 2);
                                if (noteTextBlendStart == -1)
                                {
                                    if (noteTextBlendEnd != -1)
                                        text += lineParaText.Substring
                                            (noteTextBlendEnd + 2);
                                    break;
                                }
                                st = noteTextBlendEnd + 2;
                                lText = noteTextBlendStart - st;//(noteTextBlendEnd + 2);
                                if (lText < 0)
                                {
                                    MessageBox.Show("somethins must be wrong,plx check it out !", "", MessageBoxButtons.OK, MessageBoxIcon.Error
                                        , MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                    //return new int[0];
                                    return null;
                                }
                                text += lineParaText.Substring
                                    (st, lText);
                                lText = noteTextBlendEnd;
                                noteTextBlendEnd = lineParaText.IndexOf("}",
                                   noteTextBlendStart);
                                if (noteTextBlendEnd == -1)
                                {
                                    note += lineParaText.Substring(
                                        noteTextBlendStart,
                                        lineParaText.Length - noteTextBlendStart);
                                    break;
                                }
                                st = noteTextBlendEnd + 2;
                                lText = noteTextBlendStart;
                                noteTextBlendStart = lineParaText.IndexOf("{", st);
                                if (noteTextBlendStart == -1)
                                {
                                    note += lineParaText.Substring(lText + 2,
                                        noteTextBlendEnd - (lText + 2));
                                    lText = lineParaText.Length - st;
                                    text += lineParaText.Substring(st,
                                        lineParaText.Length - st);
                                    break;
                                }
                                note += lineParaText.Substring(lText + 2, noteTextBlendEnd - (lText + 2));
                                noteTextBlendEnd = lineParaText.IndexOf("}", noteTextBlendStart);
                                lText = noteTextBlendStart - st;
                            }
                            text = ClearOmitChar(text); note = ClearOmitChar(note);
                            //len = new StringInfo(text).LengthInTextElements + (int)Math.Ceiling((decimal)(new StringInfo(note).LengthInTextElements
                            //    / ((lineParaText.Length - note.Length == 4 && lineParaText.StartsWith("{{") && lineParaText.EndsWith("}}") &&
                            //    new StringInfo(note).LengthInTextElements == normalLineParaLength) ? 1 : 2)));
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements
                                / ((lineParaText.StartsWith("{{") && lineParaText.EndsWith("}}") &&
                                new StringInfo(lineParaText.Substring(2, lineParaText.Length - "{{}}".Length)).LengthInTextElements == InstanceForm1.NormalLineParaLength) ? 1 : 2));
                            //單行小注而字數與正文大字同時，則不折半
                        }
                        else
                        {// noteTextBlendEnd < noteTextBlendStart  
                            int stNote = 0, lNote = noteTextBlendEnd;
                            while (noteTextBlendStart != -1)
                            {
                                note += lineParaText.Substring(stNote, lNote);
                                text += lineParaText.Substring(noteTextBlendEnd + 2,
                                    noteTextBlendStart - (noteTextBlendEnd + 2));
                                noteTextBlendEnd = lineParaText.IndexOf("}",
                                    noteTextBlendStart + 2);
                                stNote = noteTextBlendStart + 2;
                                if (noteTextBlendEnd == -1)
                                {
                                    note += lineParaText.Substring(stNote); break;
                                }
                                else
                                {
                                    note += lineParaText.Substring(stNote, noteTextBlendEnd -
                                        (noteTextBlendStart + 2));
                                    stNote = noteTextBlendStart;//暫記下備用
                                    noteTextBlendStart = lineParaText.IndexOf("{",
                                        noteTextBlendStart + 2);
                                    if (noteTextBlendStart == -1)
                                    {
                                        text += lineParaText.Substring(noteTextBlendEnd + 2);
                                        break;
                                    }
                                    text += lineParaText.Substring(noteTextBlendEnd + 2,
                                        noteTextBlendStart - (noteTextBlendEnd + 2));
                                    stNote = noteTextBlendStart + 2;
                                    lNote = noteTextBlendEnd;
                                    noteTextBlendEnd = lineParaText.IndexOf("}", stNote);
                                    if (noteTextBlendEnd == -1)
                                    {
                                        text += lineParaText.Substring(lNote + 2,
                                            noteTextBlendStart - (lNote + 2));
                                        lNote = lineParaText.Length - stNote;
                                        note += lineParaText.Substring(stNote, lNote);
                                        break;
                                    }
                                    lNote = noteTextBlendEnd - stNote;
                                    noteTextBlendStart = lineParaText.IndexOf("{",
                                        noteTextBlendEnd);
                                    if (noteTextBlendStart == -1)
                                    {
                                        text += lineParaText.Substring(
                                            noteTextBlendEnd + 2);
                                    }
                                }
                            }
                            text = ClearOmitChar(text); note = ClearOmitChar(note);
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)
                                new StringInfo(note).LengthInTextElements / 2);
                        }
                    }
                    gap = Math.Abs(len - InstanceForm1.NormalLineParaLength);
                }
                else//only text or note
                {
                    len = CountWordsLenPerLinePara(lineParaText.EndsWith("<p>") ? lineParaText.Substring(0, lineParaText.Length - "<p>".Length) : lineParaText);
                    //len = new StringInfo(clearOmitChar(lineParaText)).LengthInTextElements;
                    if ((xChk.IndexOf(lineParaText) + lineParaText.Length + lineParaText.Length <= xChk.Length
                        && xChk.Substring(xChk.IndexOf(lineParaText) + lineParaText.Length, "<p>".Length) == "<p>") ||
                        lineParaText.EndsWith("<p>"))
                        gap = 0;
                    else
                        gap = Math.Abs(len - InstanceForm1.NormalLineParaLength);
                }

                //誤差容錯值
                const int gapRef = 0;//3;//9;

                //the normal rule
                if (gap > gapRef && !(len < InstanceForm1.NormalLineParaLength
                    && lineParaText.IndexOf("<p>") > -1)
                    && lineParaText != "　" && lineParaText.IndexOf("*") == -1 &&
                        lineParaText.IndexOf("|") == -1) //&& gap < 8)
                {//select the abnormal one
                    bool alarm = true;
                    if (i + 1 < xLineParas.Length)
                    {
                        if (gap > gapRef && len < InstanceForm1.NormalLineParaLength
                            && xLineParas[i + 1].IndexOf("}}") > -1
                            && CountWordsLenPerLinePara(xLineParas[i + 1]) < InstanceForm1.NormalLineParaLength)
                        //&& xChk.IndexOf(lineParaText) + lineParaText.Length - 1 > 0
                        //&& xChk.Substring(xChk.IndexOf(lineParaText) + lineParaText.Length , "<p>".Length) == "<p>")
                        {
                            alarm = false;
                        }
                    }
                    if (alarm)
                    {
                        //string x = textBox1.Text;
                        string x = InstanceForm1.TextBox1_Text;
                        int j = -1;
                        //一行的起點與終點;Seprt=Separate
                        int lineSeprtEnd = 0, lineSeprtStart = lineSeprtEnd;
                        lineSeprtEnd = x.IndexOf(Environment.NewLine, lineSeprtEnd);
                        while (lineSeprtEnd > -1)
                        {

                            if (++j == i) break;
                            lineSeprtStart = lineSeprtEnd;
                            lineSeprtEnd = x.IndexOf(Environment.NewLine, ++lineSeprtEnd);
                        }
                        if (gap > 10)
                        {
                            SystemSounds.Hand.Play();
                        }
                        //return new int[] { lineSeprtStart, (lineSeprtEnd == -1 ? x.Length : lineSeprtEnd) - lineSeprtStart ,
                        //                  InstanceForm1.NormalLineParaLength,len};
                        return new AbnormalLineInfo(lineSeprtStart, (lineSeprtEnd == -1 ? x.Length : lineSeprtEnd) - lineSeprtStart
                                    , InstanceForm1.NormalLineParaLength, len);
                    }
                }
            }
            return null;
        }

    }
}

//https://copilot.microsoft.com/shares/ZxwV3BsbHPqm8YzSPiAVG TextLengthHelper類別與程式優化

//https://copilot.microsoft.com/shares/kHPV5udSdMjrBTHiYtthk
//https://copilot.microsoft.com/shares/KeGExaXHDraKCNfwUxkfJ

//https://copilot.microsoft.com/shares/HSuFN8WURrN3XKSuugheL
//https://copilot.microsoft.com/shares/QvQMxUuaqDzHpDENNZhqo