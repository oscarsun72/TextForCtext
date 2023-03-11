using ado = ADODB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WindowsFormsApp1;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Net.Mime.MediaTypeNames;

namespace TextForCtext
{
    /// <summary>
    /// Cntext=context=ChineseText，處理一切上下文、前後文文字的運算
    /// </summary>
    //public static class CnText
    public struct CnText
    {/*
      * 那像這樣的運算適合寫到我的 CnText struct 裡嗎 ？我打算在這個 struct 裡集中管理這樣類似的中文運算
          還是用 class 才適合呢？ 請教 感恩感恩　南無阿彌陀佛
        是否使用 struct 或 class 取決於您所管理的數據的性質。如果數據是不可變的，且大部分操作都是讀取而非修改，則使用 struct 比較合適。如果數據是可以被修改的，則使用 class 比較合適。

        而您所提到的分段中文字符串的運算是比較適合封裝在類別裡面的。所以，如果您有考慮將此類運算分析中文字符串放入類別裡面，那麼使用類別是比較合適的。
        孫守真
        所以還是用 public static class 比較合適麼？

        這其實是取決於你希望這個運算組成的類別如何使用，以及你希望他是如何實現的。
        如果這些運算是你需要在多個類別或多個對象之間共享且不需要對實例進行操作或維護狀態，那麼使用靜態類別 (Public Static Class) 可能是一個較好的選擇。
        如果你希望這些運算是對特定的實例進行的，那麼使用常規類別可能是一個較好的選擇。
        孫守真
        確實，這其中的運算應該是很多類別所共用的，且希望能方便呼叫，故用 public static class 應該是比較可取的吧
        是的，如果這些運算是整個程式中許多地方都需要呼叫的，且不需要建立類別的實體，那麼使用 public static class 可以使程式更加統一簡潔。這種情況下，靜態類別是一個很好的選擇。
      */
        //20230204癸卯立春元宵前夕 creedit with chatGPT大菩薩：
        /// <summary>
        /// 將傳入的中文字串分成二行；以中文字數判斷，忽略其中任何符號及數字
        /// </summary>
        /// <param name="text">要分成2行的含中文字串</param>
        /// <returns></returns>
        public static string SplitStringIntoTwoLines(string text)
        {


            ////20230204 YouChat菩薩：StringBuilder 類別是 C# 中一個可以建立可變長度字串的類別，它可以把多個字串串接起來，並且可以進行字串的插入、修改等操作。它的使用時機是當你要建立一個可變長度的字串且需要進行字串的操作時，使用 StringBuilder 類別可以提高字串的處理效率，因為它的執行時間較短。
            //StringBuilder sb = new StringBuilder();
            //sb.

            ////忽略標點符號與數字
            //string ch = Regex.Replace(XBefrTitleLine, "[{{}}<p>]" + Form1.punctuationsNum, "");
            ////不含標點符號與數字
            //StringInfo CH = new StringInfo(ch);
            //int llen = CH.LengthInTextElements;
            //if (llen > 1)//中文長度多於1個字才處理
            //{
            //    llen = llen % 2 == 0 ? llen / 2 : (llen + 1) / 2;
            //    //漢文部分
            //    //ch.Substring(0, llen) + newLine + ch.Substring(llen) +                                            
            //    int iS = 0;
            //    for (int i = 0; i < CH.LengthInTextElements; i++)
            //    {
            //        if (i > llen) { iS = i; break; }

            //    }
            //}

            //string text = "{{後《漢·志》注多引之}}<p>";
            int charCount = GetCharCount(text);

            if (charCount % 2 == 0)
            {
                int line1CharCount = charCount / 2;
                int line2CharCount = charCount / 2;
                //Console.WriteLine(SplitString(text, line1CharCount, line2CharCount));
                text = SplitString(text, line1CharCount, line2CharCount);
            }
            else
            {
                int line1CharCount = charCount / 2 + 1;
                int line2CharCount = charCount / 2;
                //Console.WriteLine(SplitString(text, line1CharCount, line2CharCount));
                text = SplitString(text, line1CharCount, line2CharCount);
            }
            //Console.ReadLine();
            return text;
        }

        private static int GetCharCount(string text)
        {
            int charCount = 0;
            StringInfo si = new StringInfo(text);

            for (int i = 0; i < si.LengthInTextElements; i++)
            {
                string subString = si.SubstringByTextElements(i, 1);
                if (!IsPunctuationNumTagSymbol(subString))
                {
                    charCount++;
                }
            }

            return charCount;
        }
        private static string SplitString(string text, int line1CharCount, int line2CharCount)
        {
            StringBuilder sb = new StringBuilder();
            StringInfo si = new StringInfo(text);
            int charCount = 0;

            for (int i = 0; i < si.LengthInTextElements; i++)
            {
                string subString = si.SubstringByTextElements(i, 1);
                if (!IsPunctuationNumTagSymbol(subString))
                {
                    charCount++;
                }

                sb.Append(subString);

                if (charCount == line1CharCount)
                {
                    sb.AppendLine();
                }
                //else if (charCount == line1CharCount + line2CharCount)
                //{
                //    break;
                //}
            }

            return sb.ToString();
        }

        private static bool IsPunctuationNumTagSymbol(string text)
        {
            //char[] punctuations = new char[] { '。', '，', '、', '；', '：', '？', '！' };
            //return Array.IndexOf(punctuations, text[0]) >= 0;
            return ("{{}}<p>" + Form1.punctuationsNum).IndexOf(text) > -1;
        }

        /// <summary>
        /// 自動加上書名號篇名號
        /// </summary>
        /// <param name="clpTxt">剪貼簿中的文字--需要加上書名號篇名號的文本。預防大文本，故以傳址（pass by reference）方式</param>
        /// <returns>傳址回傳clpTxt被標點後的結果</returns>
        internal static ref string booksPunctuation(ref string clpTxt)
        {
            //提示音
            //new SoundPlayer(@"C:\Windows\Media\Windows Balloon.wav").Play();
            System.Media.SystemSounds.Asterisk.Play();
            ado.Connection cnt = new ado.Connection();
            ado.Recordset rst = new ado.Recordset();
            Mdb.openDatabase("查字.mdb", ref cnt);
            rst.Open("select * from 標點符號_書名號_自動加上用 order by 排序", cnt, ado.CursorTypeEnum.adOpenForwardOnly);
            string w, rw;
            while (!rst.EOF)
            {
                w = rst.Fields["書名"].Value.ToString();
                rw = rst.Fields["取代為"].Value.ToString();
                rw = rw == "" ? "《" + w + "》" : rw;
                if (clpTxt.IndexOf(w) > -1)
                {
                    //clpTxt = clpTxt.Replace(w, rw);
                    booksPunctuationExamReplace(ref clpTxt, w, rw);
                }
                rst.MoveNext();
            }
            rst.Close();
            rst.Open("select * from 標點符號_篇名號_自動加上用 order by 排序", cnt, ado.CursorTypeEnum.adOpenForwardOnly);
            while (!rst.EOF)
            {
                w = rst.Fields["篇名"].Value.ToString();
                rw = rst.Fields["取代為"].Value.ToString();
                rw = rw == "" ? "〈" + w + "〉" : rw;
                if (clpTxt.IndexOf(w) > -1)
                {
                    //clpTxt = clpTxt.Replace(w, rw);
                    booksPunctuationExamReplace(ref clpTxt, w, rw);
                }
                rst.MoveNext();
            }

            //textBox1.Text = clpTxt;
            rst.Close(); cnt.Close();
            clpTxt = clpTxt.Replace("《《", "《").Replace("》》", "》").Replace("〈〈", "〈").Replace("〉〉", "〉");
            return ref clpTxt;
        }

        /// <summary>
        /// 檢查要標點上的書名號或篇名號詞彙，是否已經標過
        /// 20230309 creedit with chatGPT大菩薩：書名號標點與正則表達式ADO.NET、LINQ：
        /// 20230311 creedit with Bing大菩薩
        /// </summary>
        /// <param name="context">預防大文本，故以傳址（pass by reference）方式。呼叫端也當如此，如booksPunctuation()函式</param>
        /// <param name="term">須標上書名號或篇名號之詞彙</param>        
        static void booksPunctuationExamReplace(ref string context, string term, string termReplaced)
        {
            //string pattern = "(?<!《)(?<!〈)" + Regex.Escape(item[0]) + "(?!》)(?!〉)";
            //string pattern = "(?<!《)(?<!〈)" + Regex.Escape(term) + "(?!》)(?!〉)";
            //chatGPT大菩薩：其中的 (?<![\\p{P}&&[^》〉]]+) 表示前面沒有其他的標點符號（但是可以有其他非標點符號的字符）。這樣就可以避免誤標的情況了。
            //string pattern = "(?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉)";
            /*\\p{P}是什麼意思 \\ 前一個 \ 是逸出字元吧？
                是的，你說的沒錯。在正則表達式中，反斜杠符號 \ 是用來表示特殊字符的逸出字符，它可以使得某些字符在正則表達式中具有特殊含義，例如 .、* 等。如果要匹配 \ 本身，那麼需要使用 \\ 逸出一個反斜杠字符。
                而 \\p{P} 則是一個 Unicode 屬性表達式，表示匹配任何一個 Unicode 的標點符號字符。其中 \\p{} 是用來匹配具有某個 Unicode 屬性的字符，而 {P} 則是用來指定匹配的 Unicode 屬性是「標點符號」。
                孫守真
                為什麼要用 p 是內建的關鍵字還是需要自己宣告的變數？
                孫守真
                喔 所以 p 就是 punct 也是 punctiaton 的簡寫（縮寫）囉
                是的，您理解得很對。p是punctuation(標點符號)的簡寫，而\p{P}是正則表達式的一種語法，用於匹配任何標點符號。             
             */
            //string replacement = item[1];
            //string replacement = termReplaced;
            //text = Regex.Replace(text, pattern, replacement);
            //20230311 Bing大菩薩：
            //string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            //20230311 合併Bing大菩薩與之前chatGPT大菩薩的：
            string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            /* the result of this one : (?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉) will be 《續《資治通鑑》長編》 not 《續資治通鑑長編》 however it should not to be "《續《資治通鑑》長編》" ,because when ready to mark "資治通鑑"，the code should check out of the "資治通鑑" is already marked "《續資治通鑑長編》",so it should jump this replacement
             * Bing大菩薩：
             * I see. It seems that the regular expression pattern "(?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉)" is not working as intended and is resulting in nested replacements like “《續《資治通鑑》長編》” instead of “《續資治通鑑長編》”.
                One possible solution could be to modify the regular expression pattern to check if the term is already enclosed within another pair of 《》 or 〈〉 symbols before replacing it. For example, you could try using a pattern like "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)".
                This pattern uses negative lookbehind and negative lookahead assertions to check if the term is not enclosed within another pair of 《》 or 〈〉 symbols before replacing it with termReplaced.
                Would you like to try this solution?
                已接收訊息. I see. It seems that the regular expression pattern `"(?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉)"` is not working as intended and is resulting in nested replacements like "《續《資治通鑑》長編》" instead of "《續資治通鑑長編》". One possible solution could be to modify the regular expression pattern to check if the term is already enclosed within another pair of 《》 or 〈〉 symbols before replacing it. For example, you could try using a pattern like `"(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)"`. This pattern uses negative lookbehind and negative lookahead assertions to check if the term is not enclosed within another pair of 《》 or 〈〉 symbols before replacing it with `termReplaced`. Would you like to try this solution?
             *  已傳送訊息. Yes.
             *  Great! You can try modifying the booksPunctuationExamReplace method to use the new regular expression pattern like this:
                static void booksPunctuationExamReplace(ref string context, string term, string termReplaced)
                {
                    string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
                    context = Regex.Replace(context, pattern, termReplaced);
                }
                This should prevent nested replacements and ensure that terms that are already enclosed within another pair of 《》 or 〈〉 symbols are not replaced again.
                Let me know if this works for you or if you have any further questions.
             */
            context = Regex.Replace(context, pattern, termReplaced);
        }

    }
}
