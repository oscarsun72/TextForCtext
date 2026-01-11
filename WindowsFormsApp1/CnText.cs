using FuzzySharp;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1;
using ado = ADODB;
using static WindowsFormsApp1.Form1;
//using OpenQA.Selenium.DevTools.V125.Runtime;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;
//using static System.Net.Mime.MediaTypeNames;
//using System.Reflection;

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
            int charCount = getCharCount(text);

            if (charCount % 2 == 0)
            {
                int line1CharCount = charCount / 2;
                int line2CharCount = charCount / 2;
                //Console.WriteLine(SplitString(text, line1CharCount, line2CharCount));
                //text = splitString(text, line1CharCount, line2CharCount);
                text = splitString(text, line1CharCount);
            }
            else
            {
                int line1CharCount = charCount / 2 + 1;
                int line2CharCount = charCount / 2;
                //Console.WriteLine(SplitString(text, line1CharCount, line2CharCount));
                //text = splitString(text, line1CharCount, line2CharCount);
                text = splitString(text, line1CharCount);
            }
            //Console.ReadLine();
            return text;
        }

        private static int getCharCount(string text)
        {
            int charCount = 0;
            StringInfo si = new StringInfo(text);

            for (int i = 0; i < si.LengthInTextElements; i++)
            {
                string subString = si.SubstringByTextElements(i, 1);
                if (!isPunctuationNumTagSymbol(subString))
                {
                    charCount++;
                }
            }

            return charCount;
        }
        private static string splitString(string text, int line1CharCount)
        {
            StringBuilder sb = new StringBuilder();
            StringInfo si = new StringInfo(text);
            int charCount = 0;

            for (int i = 0; i < si.LengthInTextElements; i++)
            {
                string subString = si.SubstringByTextElements(i, 1);
                if (!isPunctuationNumTagSymbol(subString))
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

        private static bool isPunctuationNumTagSymbol(string text)
        {
            //char[] punctuations = new char[] { '。', '，', '、', '；', '：', '？', '！' };
            //return Array.IndexOf(punctuations, text[0]) >= 0;
            return ("{{}}<p>" + Form1.PunctuationsNum).IndexOf(text) > -1;
        }

        /// <summary>
        /// 自動加上書名號篇名號、及相關標點（如冒號、句號等）
        /// </summary>
        /// <param name="clpTxt">剪貼簿中的文字（不必是剪貼簿）--需要加上書名號篇名號的文本。預防大文本，故以傳址（pass by reference）方式</param>
        /// <param name="force2mark">強制執行標點，不管已有、或已做過了沒。</param>
        /// <returns>傳址回傳clpTxt被標點後的結果</returns>
        internal static ref string BooksPunctuation(ref string clpTxt, bool force2mark = false)
        {
            if (!force2mark) if (HasEditedWithPunctuationMarks(ref clpTxt)) { Form1.PlaySound(Form1.SoundLike.error); return ref clpTxt; }
            //提示音
            if (!Form1.MuteProcessing)
                new System.Media.SoundPlayer(@"C:\Windows\Media\Windows Balloon.wav").Play();
            string clpTxtOriginal = clpTxt;
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
            Form1.PlaySound(Form1.SoundLike.processing);
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
            clpTxt = clpTxt.Replace("《《", "《").Replace("》》", "》").Replace("〈〈", "〈").Replace("〉〉", "〉")
                .Replace("。。", "。").Replace("，，", "，").Replace("：：", "：").Replace("；；", "；")
                .Replace("、、", "、");
            if (clpTxt != clpTxtOriginal)
                System.Media.SystemSounds.Asterisk.Play();
            else
                Form1.PlaySound(Form1.SoundLike.warn);

            //取代有規則的標點：20240221大年十二
            clpTxt = clpTxt.Replace("》云", "》云：").Replace("〉云", "〉云：").Replace("》曰", "》曰：")
                .Replace("〉曰", "〉曰：");
            clpTxt = clpTxt.Replace("：：", "：");
            return ref clpTxt;
        }
        /// <summary>
        ///重新標識書名號篇名號
        /// </summary>
        /// <param name="clpTxt">要重新標識書名號篇名號的文本</param>        
        /// <returns>重新標好的文本</returns>
        internal static ref string RemarkBooksPunctuation(ref string clpTxt)
        {
            clpTxt = RemoveBooksPunctuation(ref clpTxt);
            BooksPunctuation(ref clpTxt, true);
            return ref clpTxt;
        }

        internal static ref string RemoveBooksPunctuation(ref string clpTxt)
        {
            //Regex rx = new Regex("[《·》〈〉]");
            Regex rx = new Regex("[《·》〈〉：]");
            clpTxt = rx.Replace(clpTxt, string.Empty);
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
            int pos_Term = context.IndexOf(term);//ABBREVIATION:1.position https://www.collinsdictionary.com/dictionary/english/pos
            if (pos_Term == -1) return;

            //202503314 GitHub　Copilot大菩薩
            //你可以考慮先用程式碼將星號* 和<p> 之間的文字單獨提取或保護起來，再進行其他替換操作。例如：

            //string input = "*欽定四庫全書<p>";
            //string term = "四庫全書";

            // 將星號和 <p> 之間的範圍保護
            string protectedPattern = @"\*.*?<p>";
            var matches = Regex.Matches(context, protectedPattern);

            // 暫時保護匹配到的範圍
            foreach (Match match in matches)
            {
                context = context.Replace(match.Value, $"{{PROTECTED:{match.Index}}}");
            }

            // 對其餘部分進行替換
            //context = Regex.Replace(context, Regex.Escape(term), $"《{term}》");

            //避開有星號「*」的段落：
            string pattern = "(?<!\\*)(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";

            context = Regex.Replace(context, pattern, termReplaced);

            // 還原保護範圍
            foreach (Match match in matches)
            {
                context = context.Replace($"{{PROTECTED:{match.Index}}}", match.Value);
            }

            //Console.WriteLine(context);



            //////我自己的，更簡短，更優化；運作邏輯忖度，應當還是用正則表達式好，因為這是一次取代，與WordVBA中的逐一檢查不同；以下這樣寫，則只瞻前、未顧後，誠掛一漏萬者也。感恩感恩　南無阿彌陀佛 20230312
            ////string patternCntext = context.Substring(0, pos_Term);
            ////if (patternCntext.LastIndexOf("《") <= patternCntext.LastIndexOf("》")
            ////    && patternCntext.LastIndexOf("〈") <= patternCntext.LastIndexOf("〉"))
            ////    context = context.Replace(term, termReplaced);


            ////if(patternCntext.LastIndexOf("")==-1&& patternCntext.LastIndexOf("")==-1)
            ////if (termReplaced.IndexOf("《") > -1)
            ////{
            ////    if (patternCntext.LastIndexOf("《") <= patternCntext.LastIndexOf("》")
            ////        && patternCntext.LastIndexOf("〈") <= patternCntext.LastIndexOf("〉"))
            ////        context = context.Replace(term, termReplaced);
            ////}
            ////else
            ////{
            ////    if (patternCntext.LastIndexOf("〈") <= patternCntext.LastIndexOf("〉")
            ////        && patternCntext.LastIndexOf("《") <= patternCntext.LastIndexOf("》"))
            ////        context = context.Replace(term, termReplaced);
            ////}



            ////string pattern = "(?<!《)(?<!〈)" + Regex.Escape(item[0]) + "(?!》)(?!〉)";
            ////string pattern = "(?<!《)(?<!〈)" + Regex.Escape(term) + "(?!》)(?!〉)";
            ////chatGPT大菩薩：其中的 (?<![\\p{P}&&[^》〉]]+) 表示前面沒有其他的標點符號（但是可以有其他非標點符號的字符）。這樣就可以避免誤標的情況了。
            ////string pattern = "(?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉)";
            ///*\\p{P}是什麼意思 \\ 前一個 \ 是逸出字元吧？
            //    是的，你說的沒錯。在正則表達式中，反斜杠符號 \ 是用來表示特殊字符的逸出字符，它可以使得某些字符在正則表達式中具有特殊含義，例如 .、* 等。如果要匹配 \ 本身，那麼需要使用 \\ 逸出一個反斜杠字符。
            //    而 \\p{P} 則是一個 Unicode 屬性表達式，表示匹配任何一個 Unicode 的標點符號字符。其中 \\p{} 是用來匹配具有某個 Unicode 屬性的字符，而 {P} 則是用來指定匹配的 Unicode 屬性是「標點符號」。
            //    孫守真
            //    為什麼要用 p 是內建的關鍵字還是需要自己宣告的變數？
            //    孫守真
            //    喔 所以 p 就是 punct 也是 punctiaton 的簡寫（縮寫）囉
            //    是的，您理解得很對。p是punctuation(標點符號)的簡寫，而\p{P}是正則表達式的一種語法，用於匹配任何標點符號。             
            // */
            //////string replacement = item[1];
            //////string replacement = termReplaced;
            //////text = Regex.Replace(text, pattern, replacement);
            //////20230311 Bing大菩薩：
            //////string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            ////20230311 合併Bing大菩薩與之前chatGPT大菩薩的：
            ////不計星號「*」：
            ////string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";

            ////避開有星號「*」的段落：
            //string pattern = "(?<!\\*)(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            //////GitHub　Copilot大菩薩：
            ////string pattern = "(?<!\\*)(?<!\\*[^<]*<p>)(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            //////Copilot大菩薩：
            ////string pattern = "(?<!\\*)(?<!\\*[^<]*<p>)(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)(?![^*]*\\*[^<]*<p>)";


            ///* Bing大菩薩： 20230709:
            // * 這是因為您的正則表達式中使用了否定的零寬度前瞻断言 (?<!\\*) 來排除前面有星號的情況。這個断言只會查看比對位置之前的字符，而不會跨越段落分隔符（例如換行符）去查看其他段落中的字符。
            // * 
            // * 如果您希望跨越段落進行比對，那麼您可以使用正則表達式中的「點」字符 . 來匹配任意字符（除了換行符之外）。例如，您可以使用 .* 來匹配任意數量的任意字符。
            //    但是，請注意，由於否定的零寬度前瞻断言 (?<!) 只能查看固定長度的字符，因此您不能直接在 (?<!) 中使用 .* 來跨越段落進行比對。如果您希望跨越段落進行比對，那麼您可能需要使用其他方法來實現。
            //    例如，您可以先將輸入字串中的所有段落分隔符（例如換行符）替換為某個特殊字符（例如 #），然後再使用正則表達式進行比對。這樣，您就可以使用 .* 來匹配任意數量的任意字符，並且能夠跨越原來的段落分隔符進行比對了。
            // */


            ///* the result of this one : (?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉) will be 《續《資治通鑑》長編》 not 《續資治通鑑長編》 however it should not to be "《續《資治通鑑》長編》" ,because when ready to mark "資治通鑑"，the code should check out of the "資治通鑑" is already marked "《續資治通鑑長編》",so it should jump this replacement
            // * Bing大菩薩：
            // * I see. It seems that the regular expression pattern "(?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉)" is not working as intended and is resulting in nested replacements like “《續《資治通鑑》長編》” instead of “《續資治通鑑長編》”.
            //    One possible solution could be to modify the regular expression pattern to check if the term is already enclosed within another pair of 《》 or 〈〉 symbols before replacing it. For example, you could try using a pattern like "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)".
            //    This pattern uses negative lookbehind and negative lookahead assertions to check if the term is not enclosed within another pair of 《》 or 〈〉 symbols before replacing it with termReplaced.
            //    Would you like to try this solution?
            //    已接收訊息. I see. It seems that the regular expression pattern `"(?<!《)(?<!〈)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?!》)(?!〉)"` is not working as intended and is resulting in nested replacements like "《續《資治通鑑》長編》" instead of "《續資治通鑑長編》". One possible solution could be to modify the regular expression pattern to check if the term is already enclosed within another pair of 《》 or 〈〉 symbols before replacing it. For example, you could try using a pattern like `"(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)"`. This pattern uses negative lookbehind and negative lookahead assertions to check if the term is not enclosed within another pair of 《》 or 〈〉 symbols before replacing it with `termReplaced`. Would you like to try this solution?
            // *  已傳送訊息. Yes.
            // *  Great! You can try modifying the booksPunctuationExamReplace method to use the new regular expression pattern like this:
            //    static void booksPunctuationExamReplace(ref string context, string term, string termReplaced)
            //    {
            //        string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            //        context = Regex.Replace(context, pattern, termReplaced);
            //    }
            //    This should prevent nested replacements and ensure that terms that are already enclosed within another pair of 《》 or 〈〉 symbols are not replaced again.
            //    Let me know if this works for you or if you have any further questions.
            // */

            //context = Regex.Replace(context, pattern, termReplaced);
        }

        /// <summary>
        /// NotionAI大菩薩所提的3、4種方式也是會清除中文，茲不錄。此即其所建構者。以或者可用到，暫先保留
        /// </summary>
        enum CharType { Letter, Digit, Other }
        /// <summary>
        /// 清除所有英文字與數字
        /// 20230313 creedit with YouChat菩薩
        /// 也會留下「-」
        /// </summary>
        /// <param name="text">要處理的文本。考慮到大文本，故用傳址（pass by reference）</param>
        internal static ref string ClearLettersAndDigits(ref string text)
        {
            //Adrenaline菩薩的這個會把中文也清除了：(NotionAI大菩薩所提的3、4種方式也是，茲不錄。enum CharType 即其所建構者)
            //text= new string(text.Where(c => !char.IsLetterOrDigit(c)).ToArray());
            //return ref text;

            text = Regex.Replace(text, "[a-zA-Z0-9]+", "");//消不掉「-」，蓋條件僅限在英數故
            return ref text;

            /* YouChat菩薩：
             * Hello! How can I help you?
                C# 清除所有英文字與數字 怎麼寫 感恩感恩　南無阿彌陀佛
                To remove all the English alphabets and numeric digits from a string in C#, you can use a regular expression. Here's an example:

                using System.Text.RegularExpressions;
                //...
                string str = "C# 清除所有英文字與數字怎麼寫 感恩感恩 南無阿彌陀佛";
                str = Regex.Replace(str, "[a-zA-Z0-9]+", "");
                Console.WriteLine(str);
                This code will output "C# 清除所有英文字與數字怎麼寫 感恩感恩 南無阿彌陀佛", which is the original string with all English alphabets and numeric digits removed. The regular expression "[a-zA-Z0-9]+" matches one or more English alphabets or numeric digits. The Regex.Replace() method is then used to replace all matches with an empty string.

                有現成的函式方法可以套用嗎？
                Yes, there are built-in methods in C# that you can use to achieve the same result. Here's an example using only built-in methods:

                string str = "C# 清除所有英文字與數字怎麼寫 感恩感恩 南無阿彌陀佛";
                str = new string(str.Where(c => !char.IsLetterOrDigit(c)).ToArray());
                Console.WriteLine(str);
                This code will also output "C# 清除所有英文字與數字怎麼寫 感恩感恩 南無阿彌陀佛", which is the original string with all English alphabets and numeric digits removed. The LINQ Where() method is used to filter out all the letters and digits, and the ToArray() method is then used to convert the resulting IEnumerable<char> to a char[]. Finally, the new string(char[]) constructor is used to create a new string from the resulting character array.
             *「"[a-zA-Z0-9]+"」 最後的加號（「+」）是什麼意思？ 謝謝
                The plus sign ("+") in the regular expression "[a-zA-Z0-9]+" means "match one or more occurrences of the preceding character or group". In this case, it means "match one or more occurrences of any uppercase or lowercase letter in the English alphabet or any numeric digit". So this regular expression will match any sequence of one or more consecutive English alphabets or numeric digits in a string.
             */
        }

        /// <summary>
        /// 清除所有英文字與數字
        ///  20230314 creedit with Bing大菩薩
        ///  會留下「-」
        /// </summary>
        /// <param name="text">要處理的文本。考慮到大文本，故用傳址（pass by reference）</param>
        /// <returns>傳址回傳結果</returns>
        internal static ref string ClearLettersAndDigits_UseUnicodeCategory(ref string text)
        {
            // 這是對的，測試成功：Bing大菩薩：您好!這是一個使用 C# 來清除英數字但保留中文的程式碼示例:
            string input = text;// "這是1個測試ABC";
            string output = string.Empty;
            foreach (char c in input)
            {
                if (!char.IsLetterOrDigit(c) || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.OtherLetter)
                {//會留下「-」清不乾淨
                    output += c;
                }
            }
            //Console.WriteLine(output);
            //ConsoIe.WriteLine(output);這段程式碼會輸出 這是個測試 。希望對您有所幫助!
            text = output;
            return ref text;

        }
        /// <summary>
        /// 只保留unicode文字（不限中文），不含其他符號
        /// </summary>
        /// <param name="text">要清理的文本，以傳址（pass by reference）方式傳遞</param>
        /// <returns></returns>
        internal static ref string ClearOthers_ExceptUnicodeCharacters(ref string text)
        {
            //20230314 creedit with Bing大菩薩（參見ClearLettersAndDigitsUseUnicodeCategory()）：
            StringBuilder sb = new StringBuilder();//bool isUnicodeCharacters=false;
            foreach (char c in text)//這行設中斷點暫停，可以明白各個字元究竟在哪個UnicodeCategory中  https://learn.microsoft.com/zh-tw/dotnet/api/system.globalization.unicodecategory?view=netframework-4.8 可配合Form1_Activated()事件程序中來測試
            {
                //if (c == "O".ToCharArray()[0]) sb.Append("◯");
                if (c == "ｏ".ToCharArray()[0]) Debugger.Break();
                if (c == "Ｏ".ToCharArray()[0]) Debugger.Break();
                if (c == "O".ToCharArray()[0]) Debugger.Break();
                if (c == "o".ToCharArray()[0]) Debugger.Break();
                if (c == "◯".ToCharArray()[0]) Debugger.Break();
                //if (c == "〇".ToCharArray()[0]) Debugger.Break();
                //if (c == "{".ToCharArray()[0]) Debugger.Break();
                //if (c == "}".ToCharArray()[0]) Debugger.Break();
                switch (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c))
                {
                    case UnicodeCategory.UppercaseLetter:
                        if (c == "O".ToCharArray()[0]) sb.Append("◯");
                        break;
                    case UnicodeCategory.LowercaseLetter:
                        break;
                    case UnicodeCategory.TitlecaseLetter:
                        break;
                    case UnicodeCategory.ModifierLetter:
                        break;
                    case UnicodeCategory.OtherLetter://cjk非surrogate的在這，如「積(7A4D=31309)」。
                        sb.Append(c);
                        break;
                    case UnicodeCategory.NonSpacingMark:
                        break;
                    case UnicodeCategory.SpacingCombiningMark:
                        break;
                    case UnicodeCategory.EnclosingMark:
                        break;
                    case UnicodeCategory.DecimalDigitNumber://[0-9]
                        if (c == "0".ToCharArray()[0]) sb.Append("◯");
                        break;
                    case UnicodeCategory.LetterNumber:
                        //if (c == "〇".ToCharArray()[0]) sb.Append(c);
                        if (c == "〇".ToCharArray()[0]) sb.Append("◯");
                        break;
                    case UnicodeCategory.OtherNumber:
                        break;
                    case UnicodeCategory.SpaceSeparator:
                        break;
                    case UnicodeCategory.LineSeparator:
                        break;
                    case UnicodeCategory.ParagraphSeparator:
                        break;
                    case UnicodeCategory.Control://\r\n在這
                        switch (c)
                        {
                            case '\r':
                                sb.Append(c);
                                break;
                            case '\n':
                                sb.Append(c);
                                break;
                            default:
                                break;
                        }
                        break;
                    case UnicodeCategory.Format:
                        break;
                    case UnicodeCategory.Surrogate:
                        sb.Append(c);
                        break;
                    case UnicodeCategory.PrivateUse:
                        break;
                    case UnicodeCategory.ConnectorPunctuation://_在這
                        break;
                    case UnicodeCategory.DashPunctuation://-在這
                        break;
                    case UnicodeCategory.OpenPunctuation:
                        if (c == "{".ToCharArray()[0]) sb.Append("{");
                        break;
                    case UnicodeCategory.ClosePunctuation:
                        if (c == "}".ToCharArray()[0]) sb.Append("}");
                        break;
                    case UnicodeCategory.InitialQuotePunctuation:
                        break;
                    case UnicodeCategory.FinalQuotePunctuation:
                        break;
                    case UnicodeCategory.OtherPunctuation://*!/在這
                        break;
                    case UnicodeCategory.MathSymbol://+在這
                        break;
                    case UnicodeCategory.CurrencySymbol:
                        break;
                    case UnicodeCategory.ModifierSymbol://^在這
                        break;
                    case UnicodeCategory.OtherSymbol:
                        if (c == "○".ToCharArray()[0]) sb.Append("◯");
                        break;
                    case UnicodeCategory.OtherNotAssigned:
                        break;
                    default:
                        break;
                }
            }
            text = sb.ToString();
            return ref text;
        }
        /// <summary>
        /// 清除首段誤空二格或誤空格的情形（《四庫全書》文本適用，會與篇標題空2格者混） 20250418
        /// </summary>
        /// <param name="text"></param>
        internal static void ClearFirstParaLeadingSpace(ref string text)
        {
            string[] paragraphs = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            int i = 0, PCount = paragraphs.Length;
            foreach (string p in paragraphs)
            {
                if (p.StartsWith("　"))
                {
                    if (i + 1 < PCount && !paragraphs[i + 1].StartsWith("　") &&
                        Form1.CountWordsLenPerLinePara(p.TrimStart('　')) == Form1.InstanceForm1.NormalLineParaLength)
                    {
                        //string pNew = p.TrimStart('　');
                        int pSt = text.IndexOf(p);
                        if (text.IndexOf(p, pSt + 1) == -1)
                            text = text.Substring(0, pSt) + p.TrimStart('　') + text.Substring(pSt + p.Length);
                    }
                }
                i++;
            }
        }

        /* 20230414 YouChat大菩薩：
         * 我想在textBox1中，將前後不是英數字的小數點（英文句號）「.」取代為中文句號「。」要怎麼寫呢 C# Windows.Forms應用程式
            * 要將textbox1中前後不是英數字的小數點取代為中文句號，你可以使用正則表達式和replace方法。下面是一個可能的實現方式：
         */
        /// <summary>
        /// 將中文文本中的小數點取代為句號
        /// </summary>
        /// <param name="text">要置換的文本，以傳址（pass　by　reference）方式傳遞</param>        
        internal static void PeriodsReplace_ChinesePunctuationMarks(ref string text)
        {
            //string pattern = @"(?<![a-zA-Z0-9])\.(?![a-zA-Z0-9])";
            string pattern = @"(?<![a-zA-Z0-9])\.";
            // 那如果只要前面是英數字，而後面未必是英數字才取代為「。」呢？
            //YouChat大菩薩：如果你只想將前面是英文或數字的小數點取代為中文句號，而後面的字符不需要是英文或數字，你可以將正則表達式修改為： 
            //string pattern = @"(?<=\w)\.(?!\w)";
            //這樣，正則表達式會匹配前面是任何英文字母或數字的小數點，並且後面不是任何英文字母或數字的位置。 
            string replacement = "。";
            text = Regex.Replace(text, pattern, replacement);
            /* YouChat大菩薩：
             *這樣會用中文句號"。"取代文本框中不在英文或數字前後出現的小數點。注意，我們使用正則表達式來定義模式，它的含義如下：
            (?<![a-zA-Z0-9]) 匹配前面不是英文字母或數字的位置（負向零寬度後願）
            \. 匹配小數點（需要用反斜杠對點進行轉義）
            (?![a-zA-Z0-9]) 匹配後面不是英文字母或數字的位置（正向零寬度先行斷言）
            該詞匹配任何前後包含字母或數字的字串中的小數點。replace方法在這種情況下會使用中文句號替換它們。
             */
        }
        /// <summary>
        /// 判斷是否已經人工手動編輯了
        /// </summary>
        /// <param name="text">要檢查的文本</param>
        /// <returns>若有人工手動編輯過的記號，則傳回true</returns>
        internal static bool HasEditedWithPunctuationMarks(ref string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            if (text.Length > 1000)
            {
                Regex regex = new Regex(@"\，|\。|\？|\！|\〈|\〉|\《|\》|\：|\『|\』|\〖|\〗|\【|\】|\「|\」|\􏿽|、|●|□|■|·|\*\*|\{\{\{|\}\}\}|\||〇|◯|　}}|\*　");
                Match match = regex.Match(text);
                return match.Success;
            }
            else
            {
                return (text.Contains("，") || text.Contains("。") || text.Contains("：") || text.Contains("􏿽")
                    || text.Contains("！") || text.Contains("？") || text.Contains("《") || text.Contains("〈")
                    || text.Contains("『") || text.Contains("』") || text.Contains("〖") || text.Contains("〗")
                    || text.Contains("「") || text.Contains("」") || text.Contains("【") || text.Contains("】")
                    || text.Contains("》") || text.Contains("〉")
                    || text.Contains("□") || text.Contains("■")
                    || text.Contains("●") || text.Contains("、")
                    || text.Contains("·") || text.Contains("**")
                    || text.Contains("|") || text.Contains("　}}")
                    || text.Contains("◯")
                    || text.Contains("〇") || text.Contains("*　")
                    || text.Contains(@"{{{") || text.Contains(@"}}}"));
            }

        }
        /// <summary>
        /// 清除所有人為/手動的標記（如標點符號等）
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        internal static bool ClearHasEditedWithPunctuationMarks(ref string text)
        {
            if (text.Length == 0) return false;
            if (!HasEditedWithPunctuationMarks(ref text)) return false;
            Regex regex = new Regex(@"\，|\。|\？|\！|\〈|\〉|\《|\》|\：|\『|\』|\「|\」|\􏿽|、|●|□|■|·|\*\*|\{\{\{|\}\}\}|\||〇|◯|　}}|\*　");
            text = regex.Replace(text, "");
            return true;
        }


        /// <summary>
        /// 文本規範化，正規化，以合於[簡單修改模式]接受的形式。如半形標點符號轉全形。置換為全形符號。
        /// 由narrow2WidePunctuationMarks擴展而來
        /// 從newTextBox1抽離出來的
        /// </summary>
        /// <param name="x">須轉換的文本。傳址（pass　by　reference）</param>        
        internal static void FormalizeText(ref string x)
        {
            if (x.Length == 0) return;
            //有語法內容者不執行（暫時先這樣，不做細緻的區別）20240803
            if (x.IndexOf("\" ") > -1 && x.IndexOf("=\"") > -1 && x.IndexOf(" />") > -1)
            {
                x = x.Replace("/><p>＝", "/>＝");
                return;
            }

            #region narrow2WidePunctuationMarks 半形轉全形。置換為全形符號。
            //20230806Bing大菩薩：
            //string pattern = "[\\u0021-\\u002F\\u003A-\\u0040\\u005B-\\u0060\\u007B-\\u007E]";
            ////string pattern = "[,.;]";
            //MatchEvaluator evaluator = match => ((char)(match.Value[0] + 65248)).ToString();
            ////x= Regex.Replace(x, pattern, evaluator);
            //x = Regex.Replace(x, pattern, evaluator).Replace("．", "。");
            #endregion

            //被取代者
            string[] replaceDChar = { "＠","○","〇","!","！！","'", ",", ";", ":", "．", "?", "：：","：\r\n：", "《《", "》》", "〈〈", "〉〉",
                "。}}<p>。}}","。}}。}}", "。}}}。<p>", "}}}。<p>", "。}}。<p>", "}}。<p>",".<p>","·<p>" ,"<p>。<p>","<p>。","􏿽。<p>","　。<p>"
                ,"。。", "，，", "@" 
                //,"}}<p>\r\n{{"//像《札迻》就有此種格式，不能取代掉！ https://ctext.org/library.pl?if=en&file=36575&page=12&editwiki=800245#editor
                ,"\r\n。<p>","\r\n〗","\r\n。}}","\r\n："
                ,"！。<p>","？。<p>","+<p>","<p>++","<p>+","：。<p>","。\r\n。"
                ,"《\r\n　　","《\r\n　","《\r\n"
                ,"：。","\r\n，","\r\n。","\r\n、","\r\n？","\r\n」","「\r\n" ,"{{\r\n" ,"\r\n}}"
                ,"􏿽？","􏿽。","，〉","。〉","〈、","！，","〈，","。〈。","〈：","：〉","〈。","：，","|。"//自動標點結果的訂正
                ,"，。","〉·","》·"
                ,"}}\r\n}"
                ,"*欽定《四庫全書》"

            };

            //取代為
            string[] replaceChar = { string.Empty,"◯","◯","！","！","、", "，", "；", "：", "·", "？", "：","：\r\n", "《", "》", "〈", "〉",
                "。}}","。}}", "。}}}<p>", "。}}}<p>", "。}}<p>", "。}}<p>","。<p>","。<p>","<p>","<p>","　","　"
                , "。", "，", "●" 
                //,"}}\r\n{{"//像《札迻》就有此種格式，不能取代掉！ https://ctext.org/library.pl?if=en&file=36575&page=12&editwiki=800245#editor
                ,"\r\n","〗\r\n","。}}\r\n","：\r\n"
                ,"！<p>","？<p>","<p>","<p>","<p>","：<p>","。\r\n"
                ,"\r\n　　《","\r\n　《","\r\n《"
                ,"。","，\r\n","。\r\n","、\r\n","？\r\n","」\r\n","\r\n「" ,"\r\n{{", "}}\r\n"
                ,"？􏿽","。􏿽","〉，","〉。","、〈","！","，〈","。〈","〈","〉","。〈","：","。|"//自動標點結果的訂正
                ,"。","·","·"
                ,"}}}\r\n"
                ,"*欽定四庫全書"

            };
            if (replaceDChar.Count() != replaceChar.Count()) Debugger.Break();//請檢查！！
            for (int i = 0; i < replaceChar.Count(); i++)
            {
                //if (replaceDChar[i] == "{{\r\n") Debugger.Break();
                //if (replaceDChar[i].Contains("?") ) Debugger.Break();
                x = x.Replace(replaceDChar[i], replaceChar[i]);
            }


            //以下舊式
            //foreach (var item in replaceDChar)
            //{
            //    if (x.IndexOf(item) > -1)
            //    {
            //        //if (MessageBox.Show("含半形標點，是否取代為全形？", "", MessageBoxButtons.OKCancel,
            //        //    MessageBoxIcon.Error) == DialogResult.OK)
            //        //{//直接將半形標點符號轉成全形
            //        for (int i = 0; i < replaceChar.Length; i++)
            //        {
            //            x = x.Replace(replaceDChar[i], replaceChar[i]);
            //        }
            //        //}
            //        break;
            //    }
            //}

            //置換中文文本中的英文句號（小數點）
            CnText.PeriodsReplace_ChinesePunctuationMarks(ref x);

            //清除\r 20231114Bing大菩薩：C# 字串取代：
            string pattern = "(?<!\\n)\\r(?!\\n)";
            string replacement = "";
            Regex rgx = new Regex(pattern);
            x = rgx.Replace(x, replacement);
            pattern = "(?<!\\r)\\n(?!\\r)";
            rgx = new Regex(pattern);
            x = rgx.Replace(x, replacement);
            /* 20231115 Bing大菩薩：C# 字串取代：
             * 您的程式碼中有一個小錯誤。當您使用 Regex rgx = new Regex(pattern); 創建了一個正則表達式物件 rgx 之後，您在後續的程式碼中改變了 pattern 變數的值，但並沒有更新 rgx 物件。因此，當您再次調用 rgx.Replace(x, replacement); 時，它仍然使用的是原來的模式，也就是 "(?<!\\n)\\r(?!\\n)"，而不是您新指定的 "(?<!\\r)\\n(?!\\r)"。
             * 您需要在改變 pattern 變數的值之後，再次創建一個新的 Regex 物件。             
             */

            //RemoveInnerBraces(ref x);            
        }
        /// <summary>
        /// 移除文本中的標點符號和阿拉伯數字
        /// 20240818 Copilot大菩薩：移除標點符號的C#函式方法：https://sl.bing.net/erpgod8B3LM
        /// </summary>
        /// <param name="input">要移除標點符號及數字的文本</param>
        /// <returns>回傳移除後的結果</returns>
        public static string RemovePunctuationsNum(string input)
        {/*要寫一個將文本移除標點符號的函式方法，可以使用正則表達式來匹配並移除標點符號。以下是一個簡單且有效率的C#函式範例：
          這個函式使用了Regex.Escape來處理標點符號字串中的特殊字符，並使用Regex.Replace來移除匹配的標點符號。*/
            // 使用正則表達式匹配並移除標點符號
            //string pattern = "[" + Regex.Escape(Form1.PunctuationsNum) + "]";
            //return Regex.Replace(input, pattern, "");
            // 手動處理特殊字符，確保正則表達式正確
            string pattern = "[－.,;?@'\"。，；！？、—…:：《·》〈‧〉「」『』〖〗【】（）()\\[\\]〔〕［］0123456789-]";
            //string pattern = @"[" + Form1.PunctuationsNum + "]";//一定要上面這樣才行
            return Regex.Replace(input, pattern, "");
        }
        /// <summary>
        /// 版心文字比對引入相似度方法。
        /// 檢查文本中是否包括書名（title）。如版心等內容。如果有則傳回所在位置。沒有或有錯誤則傳回-1
        /// 20240818：creedit with Copilot大菩薩：模糊比對與相似度比對的程式改寫：https://sl.bing.net/gnYNHR1sxRA
        /// 這段程式碼使用 FuzzySharp 函式庫來計算書名與文本行之間的相似度。如果相似度達到或超過 80%，則傳回找到的文本起始位置。
        /// </summary>
        /// <param name="xChecking">檢查的文本</param>
        /// <returns>傳回出現的位置。沒有或有錯誤則傳回-1</returns>
        internal static int HasPlatecenterTextIncluded(string xChecking)
        {
            string title = CTP.Title_Linkbox?.GetAttribute("textContent");
            if (title == null)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "HasPlatecenterTextIncluded");
                return -1;
            }

            int location = -1;
            double threshold = 0.36; // 相似度閾值            
            string xCheckingWithoutPunct = RemovePunctuationsNum(xChecking).Replace("　", string.Empty)
                .Replace("}", "").Replace("{", "");
            string[] lines = xCheckingWithoutPunct.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                //標題必非版心！//行數太少也不必檢查
                if (lines.Length < 3 || line.IndexOf("*") > -1
                    || line.IndexOf("孫守真按") > -1//有按語則不用汰除 
                    || (i > 1 && i < lines.Length - 2))//版心不可能在中間啦
                    continue;


                //double similarity = Fuzz.Ratio(title, lines[i].Replace("卷","")) / 100.0;
                string pattern = "[" + Regex.Escape("卷上下卄一二三四五六七八九十卅卌<p>") + "]";
                line = Regex.Replace(line, pattern, "");
                double similarity = Fuzz.Ratio(title, line) / 100.0;
                //Form1 frm = (Form1)Application.OpenForms[0] ;//依Copilot大菩薩建議，改用依賴注入（Dependency Injection, DI） 
                if (similarity >= threshold && line.Length < Form1.InstanceForm1.NormalLineParaLength)
                {
                    //前一段若為「|」通常是卷末題目
                    if (i > lines.Length - 2 &&
                        ("|" + Environment.NewLine).IndexOf(lines[i - 1]) > -1) //也分段/行符號可能還未自動轉換成「|」
                        continue;

                    location = xChecking.IndexOf(lines[i]);
                    if (location == -1)
                    {
                        // 計算第 i 行在 xChecking 中的行頭位置
                        location = 0;
                        for (int j = 0; j < i; j++)
                        {
                            location = xChecking.IndexOf(Environment.NewLine, location) + Environment.NewLine.Length;
                        }
                    }
                    break;
                }
            }

            return location;
        }        ///// <summary>
                 ///// 版心文字比對引入相似度方法。
                 ///// 檢查文本中是否包括書名（title）。如版心等內容。如果有則傳回所在位置。沒有或有錯誤則傳回-1
                 ///// 20240818：creedit with Copilot大菩薩：模糊比對與相似度比對的程式改寫：https://sl.bing.net/gnYNHR1sxRA
                 ///// 這段程式碼使用 FuzzySharp 函式庫來計算書名與文本行之間的相似度。如果相似度達到或超過 80%，則傳回找到的文本起始位置。
                 ///// </summary>
                 ///// <param name="xChecking">檢查的文本</param>
                 ///// <returns>傳回出現的位置。沒有或有錯誤則傳回-1</returns>
                 //internal static int HasPlatecenterTextIncluded(string xChecking)
                 //{/*要將您的程式碼增益為模糊比對與相似度比對，您可以使用一些字串相似度演算法，例如 Levenshtein 距離或 Cosine 相似度。以下是如何改寫您的程式碼以實現這一目標的範例：
                 //    引入模糊比對的函式庫，例如 FuzzySharp。
                 //    使用相似度演算法來計算文本與書名之間的相似度。
                 //    如果相似度達到 80%，則傳回找到的文本起始位置。
                 //    以下是改寫後的程式碼範例：*/
                 //    string title = CTP.Title_Linkbox?.GetAttribute("textContent");
                 //    if (title == null)
                 //    {
                 //        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "HasPlatecenterTextIncluded");
                 //        return -1;
                 //    }

        //    int location = -1;
        //    double threshold = 0.36; // 相似度閾值            
        //    string xCheckingWithoutPunct = RemovePunctuationsNum(xChecking).Replace("　",string.Empty)
        //        .Replace("}","").Replace("{","");
        //    string[] lines = xCheckingWithoutPunct.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

        //    for (int i = 0; i < lines.Length; i++)
        //    {
        //        double similarity = Fuzz.Ratio(title, lines[i]) / 100.0;
        //        if (similarity >= threshold)
        //        {
        //            location = xChecking.IndexOf(lines[i]);
        //            break;
        //        }
        //    }

        //    return location;
        //}
        /// <summary>
        /// 檢查文本中是否包括書名（title）。如版心等內容。如果有則傳回所在位置。沒有或有錯誤則傳回-1
        /// </summary>
        /// <param name="xChecking">檢查的文本</param>
        /// <returns>傳回出現的位置。沒有或有錯誤則傳回-1</returns>
        internal static int HasPlatecenterTextIncluded_exactly(string xChecking)
        {
            string title = CTP.Title_Linkbox?.GetAttribute("textContent");
            if (title == null)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "HasPlatecenterTextIncluded");
                return -1;
            }
            int location = xChecking.IndexOf(title),
                locationOriginal = location;
            ; string rn = Environment.NewLine;//\r\n
            while (location > -1)
            {
                //只有第一行/段或最後一行/段才符合所求
                int rnLoaction = xChecking.LastIndexOf(rn, location);
                if (rnLoaction == -1)//第一行/段
                    return location;
                rnLoaction = xChecking.IndexOf(rn, location);
                if (rnLoaction == -1)//最後一行/段
                    return location;
                //也有可能是最後2段/行，因為頁碼或版心下方的文字，往往以類似小注的方式呈現，OCR結果常會因此而換行/段
                else
                {
                    locationOriginal = location;
                    location = rnLoaction;
                    rnLoaction = xChecking.IndexOf(rn, location + 1);
                    if (rnLoaction == -1)//最後二行/段
                        return locationOriginal;
                }
                location = xChecking.IndexOf(title, location + 1);

            }

            return location;
        }

        /// <summary>
        /// 當取代模式輸入時，改變選取範圍的實際值（不會改變選取範圍）
        /// 此法可與Form1類別中的 overwriteModeSelectedTextSetting 互用。該法會改變文字方塊的選取範圍
        /// </summary>        
        /// <param name="insertMode">是否是取代模式；即Form1的insertMode欄位值</param>
        /// <param name="textBox1">要操作的文字方塊</param>
        /// <returns>回傳實際「選取」取得的值。</returns>
        internal static string ChangeSeltextWhenOvertypeMode(bool insertMode, TextBox textBox1)
        {
            string x = textBox1.SelectedText; int s = textBox1.SelectionStart;
            if (s + 1 <= textBox1.TextLength)
            {

                if (!insertMode)
                {
                    x = textBox1.Text;
                    x = char.IsHighSurrogate(x.Substring(s, 1).ToArray()[0]) ? textBox1.SelectedText + x.Substring(s + textBox1.SelectionLength, 2) : textBox1.SelectedText + x.Substring(s + textBox1.SelectionLength, 1);
                }
            }
            return x;
        }
        /// <summary>
        /// 這段程式碼會將4個連續的大括號替換為2個，但不會影響到5個或更多的連續大括號。
        /// 20231103 Bing大菩薩：正則表達式-大括弧的處理（和Bard大菩薩、chatGPT大菩薩一樣都不行； YouChat大菩薩先不問了）感恩感恩　南無阿彌陀佛
        /// 取代文字中有4個上下花括號（大括號）的為2個，但若是5個在一起的就不取代
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        internal static string CurlybracesFormalizer(ref string text)
        {
            string pattern = @"\{\{4\}\}";//這正則表達式應該都沒用，只是暫存而已
            string replacement = "{{}}";
            //text = text.Replace("{{{{", "{{").Replace("}}}}", "}}");
            //text = text.Replace("{{{", "{{{{{").Replace("}}}", "}}}}}");
            text = text.Replace("{{{{", "{{").Replace("}}}}", "}}").Replace("{{{", "{{{{{").Replace("}}}", "}}}}}");
            Regex rgx = new Regex(pattern);
            text = rgx.Replace(text, replacement);

            #region 全注文標記之處理
            text = text.Replace("}}" + Environment.NewLine + "{{", Environment.NewLine);

            #endregion

            return text;

            //string pattern = "(?<=^|[^{]){{4}(?=[^}]|$)|(?<=^|[^}])}}(?=[^}])(?<=[^}])(?<=[^}])(?<=[^}])(?=[^}]|$)";
            //string replacement = "{{}}";
            //Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Singleline);
            //text = rgx.Replace(text, replacement);
            //return text;


            //string pattern = "(?<!\\{)\\{\\{\\{(?!\\{)|(?<!\\})\\}\\}\\}(?!\\})";
            //string replacement = "{{}}";
            //Regex rgx = new Regex(pattern);            
            //text = rgx.Replace(text, replacement);
        }
        /// <summary>
        /// 將如《趙城金藏》3欄式的版面書圖《古籍酷》AI服務OCR結果重新排列
        /// 要先將各欄的文字區別開來，再執行
        /// </summary>
        /// <param name="text">要處理的文本。通常就是textBox1.Text</param>
        internal static void Rearrangement3ColumnLayout(ref string text)
        {//20240405 Copilot大菩薩：C# 重新排列字串
         //var paragraphs = text.Split(new[] { "\r\n\r\n" }, StringSplitOptions.None);
            var paragraphs = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var rearranged = paragraphs
                //.Select((p, i) => new { Paragraph = p, Index = i / 3 })
                .Select((p, i) => new { Paragraph = p, Index = i % 3 })
                .OrderBy(x => x.Index)
                .Select(x => x.Paragraph);
            //text = string.Join("\r\n\r\n", rearranged);
            text = string.Join(Environment.NewLine, rearranged);


            ////有群組的
            //var paragraphs = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            //var grouped = paragraphs.Select((p, i) => new { Paragraph = p, Group = i % 3 });
            //var rearranged = grouped.OrderBy(x => x.Group).ThenBy(x => x.Paragraph).Select(x => x.Paragraph);
            //text = string.Join(Environment.NewLine, rearranged);

            ////沒群組的（應該與前面的相同。只是原誤作求商，而當求餘數才是。）
            //var paragraphs = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            //var rearranged = paragraphs.Select((p, i) => new { Paragraph = p, Index = i % 3 }).OrderBy(x => x.Index).Select(x => x.Paragraph);
            //text = string.Join(Environment.NewLine, rearranged);
        }


        /// <summary>
        /// 檢查文本中的成對「【」「】」符號，並清除其中的嵌套符號（【、】）。
        /// 20240911 Copilot大菩薩不行了 還是我自己寫：https://sl.bing.net/coDbtujzqjk 改良後的「【」和「】」符號處理程式碼
        /// </summary>
        /// <param name="input">要清除的字串</param>
        /// <returns>傳回清除後的值</returns>
        public static string RemoveNestedBrackets(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            string result = @"[【】]";
            Regex regex = new Regex(result);
            if (!regex.IsMatch(input)) return input;
            result = input;
            int posOpenBoldSquareBracket = result.IndexOf("【", 0);
            int posCloseBoldSquareBracket = result.IndexOf("】", posOpenBoldSquareBracket);
            int nextOpenBoldSquareBracket = result.IndexOf("【", posOpenBoldSquareBracket + 1);
            int nextCloseBoldSquareBracket = result.IndexOf("】", posCloseBoldSquareBracket + 1);
            while (nextCloseBoldSquareBracket > -1 || nextOpenBoldSquareBracket > -1)
            {
                //如果下括號在括號之間
                if ((nextCloseBoldSquareBracket > -1 && nextOpenBoldSquareBracket > nextCloseBoldSquareBracket)
                    || (nextOpenBoldSquareBracket == -1 && nextCloseBoldSquareBracket > posCloseBoldSquareBracket))
                {
                    string str = result.Substring(posOpenBoldSquareBracket + 1, nextCloseBoldSquareBracket - posOpenBoldSquareBracket - 1);//1="【".Length
                    if (str.IndexOf("】") + posOpenBoldSquareBracket + 1 == posCloseBoldSquareBracket)
                        result = result.Substring(0, posCloseBoldSquareBracket) + result.Substring(posCloseBoldSquareBracket + 1);
                }
                //如果上括號在括號之間
                else if ((nextOpenBoldSquareBracket > -1 && nextOpenBoldSquareBracket < posCloseBoldSquareBracket)
                    || (nextCloseBoldSquareBracket == -1 && nextOpenBoldSquareBracket < posCloseBoldSquareBracket))
                {
                    string str = result.Substring(posOpenBoldSquareBracket + 1, posCloseBoldSquareBracket - posOpenBoldSquareBracket - 1);//1="【".Length
                    if (str.IndexOf("【") + posOpenBoldSquareBracket + 1 == nextOpenBoldSquareBracket)
                        result = result.Substring(0, nextOpenBoldSquareBracket) + result.Substring(nextOpenBoldSquareBracket + 1);
                }
                posOpenBoldSquareBracket = result.IndexOf("【", posOpenBoldSquareBracket + 1);
                if (posOpenBoldSquareBracket == -1) break;
                posCloseBoldSquareBracket = result.IndexOf("】", posOpenBoldSquareBracket);
                if (posCloseBoldSquareBracket == -1) break;
                nextOpenBoldSquareBracket = result.IndexOf("【", posOpenBoldSquareBracket + 1);
                nextCloseBoldSquareBracket = result.IndexOf("】", posCloseBoldSquareBracket + 1);

            }

            if (input != result)
            {
                //if (!IsBalanced(result, "【".ToCharArray()[0], "】".ToCharArray()[0]))
                if (!IsBalanced(result, '【', '】'))
                {
                    Form1.PlaySound(Form1.SoundLike.waiting, true);
                    Debugger.Break();
                }
            }
            return result;
        }
        /// <summary>
        /// 檢查 string result 中是否有不對稱的符號
        /// creedit_with_Copilot大菩薩 20241029
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool IsBalanced(string input, char openSymbol, char closeSymbol)
        {
            int countLeft = input.Count(c => c == openSymbol);
            int countRight = input.Count(c => c == closeSymbol);
            return countLeft == countRight;
        }
        /// <summary>
        /// 沒作用！檢查文本中的成對「【」「】」符號，並清除其中的嵌套符號（【、】）。
        /// 20240901 Copilot大菩薩：C# Windows.Forms中檢查並清除成對的「【」「】」符號：https://sl.bing.net/jZwKT5iL7gO
        /// </summary>
        /// <param name="input">要清除的字串</param>
        /// <returns>傳回清除的值</returns>
        public static string RemoveNestedBrackets_BAD(string input)
        {
            string pattern = @"[【】]";
            Regex regex = new Regex(pattern);
            if (!regex.IsMatch(input)) return input;
            pattern = @"【[^【】]*】";
            return Regex.Replace(input, pattern, match =>
            {
                string content = match.Value;
                content = Regex.Replace(content, @"[【】]", "");
                return $"【{content}】";
            });
        }
        /// <summary>
        /// 這段程式碼會將 textBox1 中的文字中的內部花括號移除。例如，它會將 {{Copilot{{大菩薩}} 或 {{Copilot}}大菩薩}} 轉換為 {{Copilot大菩薩}}。
        /// 請注意，這個程式碼只會移除最內層的花括號。如果有多層花括號，您可能需要多次執行這個程式碼。
        /// 20240515 Copilot大菩薩： C# 正則表達式移除內部花括號
        /// </summary>
        public static void RemoveInnerBraces(ref string text)
        {
            //string pattern = "{{([^{}]*?)}}([^}]*)}}";
            //string replacement = "{{$1$2}}";
            if (text.IndexOf("{") == -1 && text.IndexOf("}") == -1) return;
            string pattern = "{{([^{}]*?)}}([^}]*?)}}(?=\\W|$)";
            string replacement = "{{$1$2}}";
            string result = Regex.Replace(text, pattern, replacement);
            //MessageBox.Show(result);
            if (result != text)
            {
                int openBracesCount = result.Count(c => c == '{');
                int closeBracesCount = result.Count(c => c == '}');
                if (openBracesCount == closeBracesCount)
                {
                    //Debugger.Break();
                    text = result;
                }
            }
            //text = Regex.Replace(text, pattern, replacement);


            //string input = text;
            //while (input.Contains("{{") && input.Contains("}}"))
            //{
            //    int firstIndex = input.IndexOf("{{");
            //    int lastIndex = input.IndexOf("}}");
            //    if (firstIndex < lastIndex)
            //    {
            //        string start = input.Substring(0, firstIndex);
            //        string middle = input.Substring(firstIndex + 2, lastIndex - firstIndex - 2);
            //        string end = input.Substring(lastIndex + 2);
            //        input = start + "{{" + middle + "}}" + end;
            //    }
            //    else
            //    {
            //        break;
            //    }
            //}
            //text = input;

            //string pattern = @"\{\{([^{}]*\{[^{}]*\}[^{}]*)\}\}";
            //string replacement = "{{$1}}";
            //Regex rgx = new Regex(pattern);
            //string input = text;
            //string result = rgx.Replace(input, replacement);
            //text= result;
        }
        /// <summary>
        /// 將文本中成對一組的半形空格「 」轉換成成對一組的雙大括弧「{{}}」
        /// 20240529 Copilot大菩薩：C# 文本轉換
        /// </summary>
        /// <param name="text"></param>
        public static void Spaces2Braces(ref string text)
        {///這段程式碼使用了正則表達式（Regular Expression）來找出文本中成對一組的半形空格「 」，並將其替換為成對一組的雙大括弧「{{}}」。請將 “您的文本” 替換為您要處理的文本。
            //20240908 Copilot大菩薩：要在 C# 中快速且簡單地計算字串中半形空格的數量，可以使用 Count 方法。……這段程式碼使用了 System.Linq 命名空間中的 Count 擴充方法來計算字串中所有半形空格的數量。
            if (text.Count(c => c == ' ') < 2) return;
            string pattern = " ([^ ]*) ";
            string replacement = "{{$1}}";
            string result = Regex.Replace(text, pattern, replacement);
            if (result != text) { text = result; }
        }
        /// <summary>
        /// 縮排字級計算：計算分段符號後的全形空格數量
        /// 20240920 creedit_with_Copilot大菩薩：計算段落符號後的全形空格數量：https://sl.bing.net/f3ufxPqngjc
        /// </summary>
        /// <param name="strToCount">要計算的文本</param>
        /// <returns></returns>
        public static int IndentCounter(string strToCount)
        {
            int count = 0;
            //string pattern = @"\r\n　"; // 使用正則表達式匹配 "\r\n" 後面接的全形空格
            string pattern = Environment.NewLine + @"　"; // 使用正則表達式匹配 "\r\n" 後面接的全形空格
            foreach (Match match in Regex.Matches(strToCount, pattern))
            {
                count++;
            }
            return count;
        }
        /// <summary>
        /// 縮排字級計算：計算分段符號後的全形空格數量
        /// 20240920 creedit_with_Copilot大菩薩：計算段落符號後的全形空格數量：https://sl.bing.net/f3ufxPqngjc
        /// </summary>
        /// <param name="strToCount">要計算的文本</param>
        /// <param name="count"></param>
        /// <returns>傳回與縮排等量的全形空格字串</returns>
        public static string IndentCounter(string strToCount, out int count)
        {
            count = 0;
            int s = strToCount.IndexOf(Environment.NewLine);
            if (s == -1) return string.Empty;
            s += Environment.NewLine.Length;
            if (s + 1 >= strToCount.Length) return string.Empty;
            while (s + 1 < strToCount.Length && strToCount.Substring(s + count, 1) == "　")
                count++;
            return new string('　', count); // 根據計算出的全形空格數量生成等長的空格字串
                                           //int spaceCount = 0;
                                           //string pattern = @"\r\n　"; // 使用正則表達式匹配 "\r\n" 後面接的全形空格
                                           //string pattern = Environment.NewLine + @"　"; // 使用正則表達式匹配 "\r\n" 後面接的全形空格
                                           //foreach (Match match in Regex.Matches(strToCount, pattern))
                                           //{
                                           //    count++;
                                           //    spaceCount += match.Value.Length - 2; // 減去 "\r\n" 的長度，只計算全形空格的數量
                                           //}            
                                           //return new string('　', spaceCount); // 根據計算出的全形空格數量生成等長的空格字串

        }


        /// <summary>
        /// 20240808（臺灣父親節）creedit with Copilot大菩薩：《古籍酷》自動標點完成的文本重新插入分段符號
        /// </summary>
        /// <param name="originalText"></param>
        /// <param name="punctuatedText"></param>        
        /// <returns>傳回復原段落的 punctuatedText</returns>
        public static string RestoreParagraphs(string originalText, ref string punctuatedText)
        //public static string RestoreParagraphs(ref string originalText, ref string punctuatedText)
        {
            //記下縮排的字數
            string indentStr = IndentCounter(originalText, out int indentCount);

            // Define a set of punctuation marks to ignore
            HashSet<char> punctuationMarks = new HashSet<char> { '。', '，', '；', '：', '、', '？', '！', '《', '》', '「', '」', '『', '』' };
            /* `HashSet` 是 .NET 中的一種集合類別，它有一些特點使其在某些情況下非常有用。以下是 `HashSet` 的一些主要優點：
                    1. **快速查找**：`HashSet` 使用哈希表來存儲元素，因此查找元素的時間複雜度為 O(1)，這意味著無論集合中有多少元素，查找速度都非常快。這在需要頻繁查找元素的情況下特別有用。
                    2. **唯一性**：`HashSet` 保證集合中的每個元素都是唯一的。如果嘗試添加一個已經存在的元素，`HashSet` 不會添加重複的元素。
                    3. **靈活性**：`HashSet` 支持標準的集合操作，如聯集、交集和差集，這使得它在處理集合操作時非常靈活。
                    在您的情況下，使用 `HashSet` 來存儲標點符號集合的好處是可以快速查找和檢查某個字符是否是標點符號，從而提高程式的效率。
                    如果您對 `HashSet` 有更多的興趣或有其他問題，請隨時告訴我。南無阿彌陀佛 🙏
             */
            // Function to remove punctuation marks from a string

            DateTime dt = DateTime.Now;
            string RemovePunctuation(string text)
            {
                var result = new List<char>();
                foreach (var ch in text)
                {
                    if (!punctuationMarks.Contains(ch))
                    {
                        result.Add(ch);
                    }

                }
                return new string(result.ToArray());
            }

            bool error = false;
            // Function to find the adjusted position in punctuatedText
            //int FindAdjustedPosition(string text, string original, int pos, string before, string after)
            int FindAdjustedPosition(string text, int pos, string before, string after)
            {
                int offset1 = 0;
                int adjustedPos = pos;//原來分段符號所在位置
                while (adjustedPos < text.Length)
                //while (adjustedPos + offset1 < text.Length)
                //while ((adjustedPos + (before.Length + offset1)) < text.Length)
                {
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 5)
                    {
                        if (!error)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("復原段落有誤，請注意！！");
                            error = true;
                        }
                        return -1;
                    }
                    // Process the 'before' part
                    string subText = text.Substring(adjustedPos - (before.Length + offset1), before.Length + offset1);
                    string subTextWithoutPunctuation = RemovePunctuation(subText);
                    int cntr = 0;//20251118
                    while (subTextWithoutPunctuation.Length < before.Length)
                    {
                        cntr++; if (cntr > 2000) break;
                        //if ((adjustedPos + (before.Length + offset1)) < text.Length)
                        if (adjustedPos + after.Length + offset1 < text.Length)
                        {
                            offset1++;
                            int startIndex = adjustedPos - (before.Length + offset1);
                            if (startIndex < 0) break;////////●●●●●●●●● 20251118 待觀察！！！！
                                                      //{
                            subText = text.Substring(startIndex, before.Length + offset1);
                            subTextWithoutPunctuation = RemovePunctuation(subText);
                            //}
                            //else//●●●●●●●●●●●
                            //    if (offset1 > 2000) break;
                        }
                        else
                        {
                            Debugger.Break();
                            //text=RemovePunctuation(text);
                            adjustedPos = (before.Length + offset1 + 1);
                            Form1.PlaySound(Form1.SoundLike.error);
                            //return -1;
                            //break;
                        }
                    }
                    if (subTextWithoutPunctuation.Contains(before))
                    {
                        ////異常檢查（分段符號前文字）：
                        //if (subTextWithoutPunctuation != before &&
                        //    new StringInfo(subTextWithoutPunctuation).LengthInTextElements - 2 > new StringInfo(before).LengthInTextElements) Debugger.Break();
                        //adjustedPos += subText.Length;
                        //offset1 += subText.Length - before.Length;

                        // Process the 'after' part
                        int offset2 = 0;
                        int afterAdjustedPos = adjustedPos;
                        while (afterAdjustedPos + (after.Length + offset2) < text.Length)
                        {
                            string afterSubText = text.Substring(afterAdjustedPos, after.Length + offset2);
                            string afterSubTextWithoutPunctuation = RemovePunctuation(afterSubText);
                            while (afterSubTextWithoutPunctuation.Length < after.Length)
                            //while (afterSubTextWithoutPunctuation.Length <= after.Length)
                            {
                                if (afterAdjustedPos + (after.Length + offset2) < text.Length)
                                {
                                    offset2++;
                                    afterSubText = text.Substring(afterAdjustedPos, after.Length + offset2);
                                    afterSubTextWithoutPunctuation = RemovePunctuation(afterSubText);
                                }
                            }
                            if (afterSubTextWithoutPunctuation.Contains(after))
                            {
                                //異常檢查：
                                if (afterSubTextWithoutPunctuation != after)
                                    if (!afterSubTextWithoutPunctuation.EndsWith(after)) Debugger.Break();
                                return adjustedPos;
                            }
                            else
                            {
                                afterAdjustedPos++;
                            }
                        }
                    }
                    else
                    {
                        adjustedPos++;
                    }
                }
                return -1;
            }

            #region 先規範要操作的文本
            //先清除標點完成的文本中可能含有的分段符號，以利後續的比對
            punctuatedText = punctuatedText.Replace(Environment.NewLine, string.Empty);
            //清除標點符號以利分段符號之比對搜尋
            originalText = RemovePunctuation(originalText);
            ////清除縮排即凸排格式標記，即將分段符號前後的空格「　」均予清除//當寫在送去自動標點前！！20240918//發現問題出在使用了 .Text屬性值 故先還原再觀察
            //originalText = Regex.Replace(originalText, $@"\s*{Environment.NewLine}+\s*", Environment.NewLine);
            //因為自動標點為清除全形空格，故亦予清除以好比對
            originalText = originalText.Replace("　", string.Empty);
            #endregion

            // Step 1: Find the positions of the paragraph breaks in the original text
            //List<(int, string, string)> paragraphPositions = new List<(int, string, string)>();
            List<Tuple<int, string, string>> paragraphPositions = new List<Tuple<int, string, string>>();
            string newLine = Environment.NewLine;
            int index = 0;
            while ((index = originalText.IndexOf(newLine, index)) != -1)
            {
                // Store the position and the surrounding text for comparison
                int start = Math.Max(0, index - 5);
                int end = Math.Min(originalText.Length, index + 5);
                string before = originalText.Substring(start, index - start);
                string after = originalText.Substring(index + newLine.Length, end - index - newLine.Length);

                #region  surrogate判斷調適：會干擾後面的判斷，須再審測（已經測試，不可用！）20240920


                ////if (char.IsHighSurrogate(before.LastOrDefault()))
                ////    before = originalText.Substring(start, index - start + 1);
                //if (char.IsLowSurrogate(before.FirstOrDefault()))
                //{
                //    Debugger.Break();
                //    before = originalText.Substring(start - 1, index - start);
                //}
                //if (char.IsHighSurrogate(after.LastOrDefault()))
                //{
                //    Debugger.Break();
                //    after = originalText.Substring(index + newLine.Length, end - index - newLine.Length + 1);
                //}
                ////if (char.IsLowSurrogate(after.FirstOrDefault()))
                ////    after = originalText.Substring(index + newLine.Length, end - index - newLine.Length);

                #endregion

                // Ensure 'before' and 'after' do not include newline characters
                while (before.Contains('\r') || before.Contains('\n'))
                {
                    start++;
                    before = originalText.Substring(start, index - start);
                }
                while (after.Contains('\r') || after.Contains('\n'))
                {
                    end--;
                    after = originalText.Substring(index + newLine.Length, end - index - newLine.Length);
                }

                //paragraphPositions.Add((index, before, after));
                paragraphPositions.Add(new Tuple<int, string, string>(index, before, after));
                index += newLine.Length;
            }


            // Step 2: Insert paragraph breaks into the punctuated text
            int offset = 0;
            //foreach (var (pos, before, after) in paragraphPositions)
            foreach (var tp in paragraphPositions)
            {
                //int adjustedPos = FindAdjustedPosition(punctuatedText, originalText, pos + offset, before, after);
                //int adjustedPos = FindAdjustedPosition(punctuatedText, originalText, tp.Item1 + offset, tp.Item2, tp.Item3);
                int adjustedPos = FindAdjustedPosition(punctuatedText, tp.Item1 + offset, tp.Item2, tp.Item3);
                //int adjustedPos = FindAdjustedPosition(punctuatedText, originalText, pos + offset, before, after);
                //因為在子函式方法中，若沒有找到時會將標點符號清除再與原未標點之文本作比對，若原文本已略有標點，則會干擾比對結果，不如兩造一律均清除，則簡單有效 20240808
                if (adjustedPos != -1)
                {

                    //punctuatedText = punctuatedText.Insert(adjustedPos, newLine);
                    //offset += newLine.Length;
                    //以下改在迴圈後再處理--仍在這裡試看看：成功了！20240920                    
                    punctuatedText = punctuatedText.Insert(
                        punctuationMarks.Contains(punctuatedText[adjustedPos]) ?
                        ++adjustedPos : adjustedPos
                                            , newLine + //若分段符號後起首是「􏿽」，如 ： 􏿽人之多事，私欲使然也。無欲則無事矣。<p>
                                                        // 􏿽欲者，無涯之物也。原其端則一𫝹，要其極則無
                                                        // 則非縮排
                                            ((indentCount > 0 && punctuatedText[adjustedPos] == "􏿽"[0]) ? string.Empty : indentStr));
                    offset += newLine.Length + indentCount;

                }
            }

            if (Form1.CountWordsinDomain("\r", originalText)
                != Form1.CountWordsinDomain("\n", punctuatedText))
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("還原段落時出錯，請注意！！");
                //Debugger.Break();
            }

            //if (indentCount > 0)
            //{
            //    FormalizeText(ref punctuatedText);
            //    punctuatedText = punctuatedText.Replace(Environment.NewLine, Environment.NewLine + indentStr);
            //}

            if (punctuatedText.IndexOf("􏿽：") > -1) punctuatedText = punctuatedText.Replace("􏿽：", "􏿽");//《古籍酷》自動標點的問題 20240920
            if (originalText.StartsWith("　") && !punctuatedText.StartsWith("　"))
            {
                int s = 0;
                while (s < originalText.Length && originalText.Substring(s, 1) == "　")
                {
                    s++;
                }
                punctuatedText = originalText.Substring(0, s) + punctuatedText;//亦《古籍酷》自動標點的問題 20240920
            }

            return punctuatedText;
        }



        /// <summary>
        /// 將非縮排用的全形空格置換成引數 symbolToReplace 值
        /// 送去《古籍酷》自動標點前先將非縮排用的全形空格置換成引數 symbolToReplace 值
        /// 因為自動標點會清除文本中的全形空格，故先置換非縮排的全形空格為特殊符號以便後來還原
        /// </summary>
        /// <param name="x">要置換的文本</param>        
        /// <param name="symbolToReplace">要置換成的字符</param>
        /// <returns></returns>
        public static string ReplaceFullWidthSpace(ref string x, string symbolToReplace)
        {
            int s = x.IndexOf("　");
            if (s == -1) return x;
            if (s == 0) x = symbolToReplace + x.Substring(s + 1);
            while (s > -1)
            {
                if (s > 2 && x.Substring(s - 2, 2) != Environment.NewLine)
                {
                    x = x.Substring(0, s) + symbolToReplace + x.Substring(s + 1);
                }
                s = x.IndexOf("　", s + 1);
            }
            return x;
        }

        /* 20250307 creedit with Gemini大菩薩： https://g.co/gemini/share/6b0eda0a373f */
        /// <summary>
        /// 將全形空格（　）取代為指定的字符
        /// 進階版的 string sTxtChk = sTxt.Replace("　", "􏿽")
        /// </summary>
        /// <param name="input">要處理的字串</param>
        /// <param name="symbolToReplace">要將全形空格取代為的對象，預設為「􏿽」（空白）</param>
        /// <returns>傳回處理的字串</returns>
        public static string ReplaceFullWidthSpace_withBlank(string input, string symbolToReplace = "􏿽")
        {
            ////string pattern = @"(?<!(\*|^|\G(?!^)(?!<p>))(\s*))　(?!(?<=\<p\>)\s?　|(\s*\}\}))";
            ////string pattern = @"(?<!(\*|^|\G(?!^)(?!<p>))(\s*))　(?!(?<=\<p\>\s)\s　|(\s*\}\}))";
            //string pattern = @"(?<!(\*|^|\G(?!^)(?!<p>))(\s*))　(?!(?<=\<p\>\s?)　|(\s*\}\}))";

            //string sTxtChk = Regex.Replace(sTxt, pattern, "􏿽");

            //string sTxtChk = ReplaceSpacesCompletely(sTxt);
            //string ReplaceSpacesCompletely(string input)
            //{
            // 第一次處理
            string pattern = @"(?<!(\*|^|\G(?!^)(?!<p>))(\s*))　(?!(?<=\<p\>\s?)　|(\s*\}\}))";
            //input = Regex.Replace(input, pattern, "􏿽");
            input = Regex.Replace(input, pattern, symbolToReplace);

            // 第二次處理
            //string[] correct = { "<p>　􏿽", "}}􏿽" }; int st = 0;
            string[] correct = { "<p>　" + symbolToReplace, "}}" + symbolToReplace }; int st = 0;
            foreach (var item in correct)
            {
                st = input.IndexOf(item, st);
                while (st > -1 && st + item.Length < input.Length)
                {
                    int spsSt = st + item.Length; string inputOld = input;
                    while (spsSt < input.Length && input.Substring(spsSt, 1) == "　")
                    {
                        input = input.Substring(0, spsSt) + "􏿽" + input.Substring(spsSt + 1);
                        //spsSt += 2;//2="􏿽".Length;
                        spsSt += symbolToReplace.Length;
                    }
                    st = input.IndexOf(item, st + 1);
                }
                st = 0;
            }

            //第3次處理
            if (input.Contains("<p>" + Environment.NewLine + symbolToReplace))
            {
                //input = input.Replace("<p>" + Environment.NewLine + "􏿽", "<p>" + Environment.NewLine + "　");
                input = input.Replace("<p>" + Environment.NewLine + symbolToReplace, "<p>" + Environment.NewLine + "　");
            }

            //第4次處理
            while (input.Contains(symbolToReplace + "　"))
                input = input.Replace(symbolToReplace + "　", symbolToReplace + symbolToReplace);

            Document document = new Document(input);
            foreach (var item in document.GetParagraphs())
            {
                if (item.Text.IndexOf("*") == -1)
                {
                    int i = 0; StringBuilder sb = new StringBuilder();
                    while (i < item.Text.Length && item.Text.Substring(i, 1) == "　")
                    {
                        sb.Append("􏿽");
                        i++;
                    }
                    if (i > 0)
                        item.Text = sb.ToString() + item.Text.Substring(i);
                }
            }
            //input = document.Text;
            return document.Text;

        }

        /// <summary>
        /// 20250117
        /// 在數行/段（有\r\n作間隔符）文字裡將n行/段的文字選取起來、或取得這個n行/段範圍內的字串值
        /// Copilot大菩薩：在C# Windows Forms中，要選取特定行數或段落內的文字或取得這些範圍內的字串值，可以利用TextBox或RichTextBox控制項和一些文字處理邏輯來完成。
        /// 這段程式碼會將TextBox控制項中的文字以\r\n分割成行，然後根據指定的行數範圍選取並取得這些行的文字。
        /// </summary>
        /// <param name="textBox"></param>
        /// <param name="lineParaCount"></param>
        /// <returns></returns>
        internal static string GetSelectionTextByLineParaCount(ref TextBox textBox, int lineParaCount)
        {
            //// 假設textBox是您的TextBox控制項
            //// 將\r\n作為行分隔符將文字分割成行
            string[] lines = textBox.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

            // 要選取的行數範圍
            int startLine = 0;//2; // 開始行，行數從0開始
            int endLine = lineParaCount - 1;//4; // 結束行

            // 確保範圍在有效的行數內
            if (startLine >= 0 && endLine < lines.Length && startLine <= endLine)
            {
                // 取出指定範圍的行
                var selectedLines = lines.Skip(startLine).Take(endLine - startLine + 1);
                // 合併行為一個字串
                string selectedText = string.Join(Environment.NewLine, selectedLines);

                // 顯示或使用選取的文字
                return selectedText;
            }
            else
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("選取範圍無效");
                return textBox.Text;
            }

        }

        /// <summary>
        /// 將斜線「/」前後的文本倒置過來（即《Kanripo漢籍リポジトリ》或《國學大師》所藏《四庫全書》或《四部叢刊》本小注夾注文前後行倒置的情形）
        /// 20250118 Copilot大菩薩
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string SwapTextAroundSlash(string input)
        {
            int slashIndex = input.IndexOf('/');
            if (slashIndex >= 0)
            {
                string part1 = input.Substring(0, slashIndex);
                string part2 = input.Substring(slashIndex + 1);
                return part2 + "/" + part1;
            }
            return input;
        }
        //將誤填滿為空白的改為空格
        //private Document _document; _document = new Document(textBox1);
        internal static void ReplaceBlanksWithSpaces(TextBox textBox1)
        {/*
          * 1.	取得 textBox1 內容的所有段落。
            2.	遍歷各個段落，找出符合條件的段落。
            3.	更改符合條件的段落的下一個段落的第一個字元。
            4.	確保更改反映在 textBox1 的內容中。
            這樣，當……符合條件的段落的下一個段落的第一個字元將被更改，並且 textBox1 中的內容也會相應更新。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。

          */
            Document _document;
            _document = new Document(ref textBox1);


            var paragraphs = _document.GetParagraphs();

            for (int i = 0; i < paragraphs.Count - 1; i++)
            {
                var currentParagraph = paragraphs[i];
                var nextParagraph = paragraphs[i + 1];

                if (currentParagraph.Text.Length > 0 && nextParagraph.Text.Length > 0)
                {
                    string firstCharCurrent = currentParagraph.Text.Substring(0, char.IsHighSurrogate(currentParagraph.Text[0]) ? 2 : 1);
                    string firstCharNext = nextParagraph.Text.Substring(0, char.IsHighSurrogate(nextParagraph.Text[0]) ? 2 : 1);

                    if (firstCharCurrent == "􏿽" && firstCharNext == "􏿽" &&
                        !currentParagraph.Text.EndsWith("<p>") && nextParagraph.Text.EndsWith("<p>"))
                    {
                        _document.CurrentParagraphIndex = i + 1;

                        _document.UpdateParagraphFirstCharacter(i + 1, "　");
                    }
                }
            }
        }


        public static int IndexOf_StringInfo(string searchStr, string contextStr, int start = 0)
        {
            //StringInfo sInfo = new StringInfo(s);
            //StringInfo xInfo = new StringInfo(x);
            TextElementEnumerator xTE = StringInfo.GetTextElementEnumerator(contextStr);
            int i = 0;
            while (xTE.MoveNext())
            {
                if (i > start)
                {
                    string sCompare = xTE.Current.ToString(); bool found = true;
                    if (searchStr == sCompare)
                    {
                        return i;
                    }
                    else if (sCompare == "\r")
                    {

                        for (int j = 0; j < searchStr.Length; j++)
                        {
                            if (searchStr[j] != contextStr[xTE.ElementIndex + j])
                            {
                                found = false; break;
                            }
                        }
                        if (found)
                            return i;

                    }
                }
                i++;
            }
            return -1;
        }


        /// <summary>
        /// 取得給定字串「/」後方的數字
        /// 20240727 Copilot大菩薩：C# 字串處理：取出「/」後的數字
        /// </summary>
        /// <param name="input">所給定的字串</param>
        /// <returns></returns>
        /// <exception cref="Exception">出錯時顯示</exception>
        public static int ExtractNumberAfterSlash(string input)
        {
            // 使用正則表達式找到「/」後的數字
            Match match = Regex.Match(input, @"/(\d+)");
            if (match.Success)
            {
                // 將數字部分轉換為 int 型別
                return int.Parse(match.Groups[1].Value);
            }
            else
            {
                throw new Exception("未找到數字");
            }
        }

        /// <summary>
        /// 取得一個含有t個字串的字串
        /// </summary>
        /// <param name="t">想要的字串的數量</param>
        /// <param name="str">想要重複的字串</param>
        /// <returns>含有t個字串的字串</returns>
        public static string GetStrings(int t, string str)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < t; i++)
            {
                sb.Append(str);
            }
            return sb.ToString();
        }
        /// <summary>
        /// 將缺失字元後的空格轉換為空白
        /// </summary>
        /// <param name="x">要檢查與轉換的文本</param>
        /// <returns>傳回轉換後的文本</returns>
        public static string ConvertSpace2BlankAfterMissingCharacter(string x)
        {
            List<string> missingcharacters = new List<string>() { "{{缺}}", "{{闕}}" };
            foreach (var item in missingcharacters)
            {
                int missIndex = x.IndexOf(item);
                if (missIndex > -1)
                {
                    int spasceIndex = missIndex + item.Length;
                    while (spasceIndex + 1 < x.Length && x.Substring(spasceIndex, 1) == "　") //1="　".Length
                    {
                        x = x.Substring(0, spasceIndex) + "􏿽" + x.Substring(spasceIndex + 1);
                        spasceIndex += 2; //2 = "􏿽".Length;
                    }
                }
            }
            return x;
        }
        /// <summary>
        /// 在文本更動後，取得更動前原來位置的對應值。
        /// 在文本經過如去除標點符號的處理後，在新文本找到的定位，如何回溯至原文本的對應位置時用
        /// </summary>
        /// <param name="position_afterupdate">更動文本後的位置</param>
        /// <param name="textOrigin">更動前的文本</param>
        /// <param name="text">更動後的文本</param>
        /// <returns></returns>
        public static int ResetPositionAfterTextUpdateed(int position_afterupdate, string textOrigin, string text)
        {
            if (text == textOrigin) return position_afterupdate;
            else
                return position_afterupdate += (textOrigin.Substring(0, textOrigin.IndexOf(" ", textOrigin.IndexOf(" ") + 1)).Length - text.Substring(0, position_afterupdate).Length);
        }

        /* 20250219 creedit_with_Copilot大菩薩 https://copilot.microsoft.com/shares/WUwdpzQFHY57cyPUyLE89
                   https://ctext.org/library.pl?if=gb&file=62381&page=7#%E4%BB%A5%E4%B8%8B%E5%92%8C%E5%8F%A5%E4%BA%A1 */
        /// <summary>
        /// 訂正註文中空白錯亂的文本
        /// 如「{{帝和霍王以下 句亡}}」訂正為「{{帝 霍王以下和句亡}}」，將半形空格與其前半對應的漢字對調。
        /// </summary>
        /// <param name="text">要訂正的文本</param>        
        /// <param name="spacePosition">傳回半形空格所在位置；若範圍內多於1個空格，則將其選位置指出</param>
        /// <returns>若失敗或無須校正（如沒半形空格）則傳回null</returns>
        public static string CorrectNoteBlankContent(string text, out int spacePosition)
        {
            int splitIndex = text.IndexOf(Environment.NewLine); spacePosition = -1;
            #region 規範文本
            string textOrigin = text;

            if (splitIndex > -1)
            {
                //選取整個行/段
                Form1.InstanceForm1.TextBox1_SelectionStart =
                    Form1.InstanceForm1._document.Range(Form1.InstanceForm1.TextBox1_SelectionStart, Form1.InstanceForm1.TextBox1_SelectionStart).GetCurrentParagraph().Start;
                Form1.InstanceForm1.TextBox1_SelectionLength =
                    Form1.InstanceForm1._document.Range(Form1.InstanceForm1.TextBox1_SelectionStart + Form1.InstanceForm1.TextBox1_SelectionLength, Form1.InstanceForm1.TextBox1_SelectionStart + Form1.InstanceForm1.TextBox1_SelectionLength).GetCurrentParagraph().End
                     - Form1.InstanceForm1._document.Range(Form1.InstanceForm1.TextBox1_SelectionStart, Form1.InstanceForm1.TextBox1_SelectionStart).GetCurrentParagraph().Start;

                text = Form1.InstanceForm1._document.Text.Substring(Form1.InstanceForm1.TextBox1_SelectionStart, Form1.InstanceForm1.TextBox1_SelectionLength);
                text = text.Replace("}}" + Environment.NewLine + "{{", Environment.NewLine);
            }
            CnText.RemoveBooksPunctuation(ref text);
            #endregion

            if (text.IndexOf(" ") == -1) return null;
            int startIndex = text.IndexOf("{{");
            if (startIndex == -1) return null;
            int endIndex = text.IndexOf("}}", startIndex);
            if (endIndex == -1) return null;

            splitIndex = text.IndexOf(Environment.NewLine);
            //用以下的式子取值則不必再有下面的：splitIndex = new StringInfo(text.Substring(0, splitIndex)).LengthInTextElements; 
            //int splitIndex = IndexOf_StringInfo(Environment.NewLine, text);
            if (splitIndex > -1)
            {
                text = text.Replace(Environment.NewLine, string.Empty);
                //記下分行/段符號位置
                splitIndex = new StringInfo(text.Substring(0, splitIndex)).LengthInTextElements;
                startIndex = text.IndexOf("{{");
                endIndex = text.IndexOf("}}", startIndex);
            }


            if (startIndex != -1 && endIndex != -1)
            {

                string pattern = "{{.*?}}";//Copilot大菩薩： 如果要使用貪婪匹配，只需要把正則表達式中的「?」去掉即可。原本的模式 ＝.*?＝ 中，「?」使得匹配變成非貪婪模式，而去掉它後，模式 ＝.*＝ 就變成了貪婪匹配。 20250224
                MatchCollection matchCollection = Regex.Matches(text, pattern);
                foreach (Match item in matchCollection)
                {
                    //傳回半形空格所在位置；若範圍內多於1個空格，則將其選位置指出
                    string x = item.Value;
                    spacePosition = x.IndexOf(" ");
                    spacePosition = x.IndexOf(" ", spacePosition + 1);
                    if (spacePosition > -1)
                    {
                        spacePosition += (item.Index);
                        //spacePosition += (textOrigin.Substring(0, textOrigin.IndexOf(" ", textOrigin.IndexOf(" ") + 1)).Length - text.Substring(0, spacePosition).Length);
                        spacePosition = ResetPositionAfterTextUpdateed(spacePosition, textOrigin, text);
                        return null;
                    }

                    startIndex = item.Index;
                    endIndex = item.Index + item.Length - "}}".Length;


                    StringInfo segment = new StringInfo(text.Substring(startIndex + 2, endIndex - startIndex - 2));
                    string segmentStr = segment.String;
                    //int spaceIndex = segmentStr.IndexOf(' ');
                    int spaceIndex = IndexOf_StringInfo(" ", segmentStr);

                    //if (spaceIndex != -1 && spaceIndex > 0)
                    if (spaceIndex > 0)
                    {
                        //if (splitIndex > spaceIndex)//這該是原書原文的空格，在分段後，不是電子文本。未見原書是無法判斷的！20250227
                        //{//若原書原文的空格是在後半，其電子文本是不會出錯的！唯有在前半，才會出錯。
                        //    return text.Replace(" ", "􏿽");
                        //}

                        //int segmentLength = segmentStr.Length;
                        int segmentLength = segment.LengthInTextElements;
                        //int midIndex = splitIndex > -1 ? splitIndex : segmentLength / 2;
                        int midIndex = segmentLength / 2;

                        int correspondingIndex = spaceIndex < midIndex ? spaceIndex : spaceIndex - midIndex;

                        if (segmentLength % 2 == 0)  // Even length
                        {
                            correspondingIndex = spaceIndex - midIndex;
                        }
                        else  // Odd length
                        {
                            correspondingIndex = spaceIndex - (midIndex + 1);
                        }

                        if (correspondingIndex < 0) return null;
                        if (correspondingIndex > spaceIndex)
                        {
                            return text.Replace(" ", "􏿽");
                        }

                        //StringInfo precedingChar = new StringInfo(segment.SubstringByTextElements(correspondingIndex, 1));
                        //StringBuilder sb = new StringBuilder(segmentStr);
                        // Swap the space and the corresponding character in the first half
                        //sb[spaceIndex] = precedingChar.String[0];
                        //sb[correspondingIndex] = ' ';
                        //text = (text.Substring(0, startIndex) + "{{" + sb.ToString() + "}}" + text.Substring(endIndex + 2)).Replace(" ", "􏿽");

                        text = text.Substring(0, startIndex)
                                + "{{" +
                                    (segment.SubstringByTextElements(0, correspondingIndex) + "􏿽" + segment.SubstringByTextElements(correspondingIndex + 1, spaceIndex - (correspondingIndex + 1))
                                    + segment.SubstringByTextElements(correspondingIndex, 1)
                                    + segment.SubstringByTextElements(spaceIndex + 1, segmentLength - (spaceIndex + 1)))
                                + "}}"
                                + text.Substring(endIndex + 2);
                        if (splitIndex > -1)
                        {

                            text = new StringInfo(text).SubstringByTextElements(0, splitIndex)
                                + Environment.NewLine
                                + new StringInfo(text).SubstringByTextElements(splitIndex);
                        }

                        //string result = text.Substring(0, startIndex) + "{{" + sb.ToString() + "}}" + text.Substring(endIndex + 2);
                        //return result.Replace(" ", "􏿽");
                        //return (text.Substring(0, startIndex) + "{{" + sb.ToString() + "}}" + text.Substring(endIndex + 2)).Replace(" ", "􏿽");

                        CnText.RemarkBooksPunctuation(ref text);

                        return text;
                    }
                }
            }
            return null;
        }

        /* 20250212元宵節creedit_with_Copilot大菩薩： 
         */
        /// <summary>
        /// 將夾注文本倒置者重整，如
        /// 「{{雪電}}　{{雨霧}}　{{霽虹}}　{{雷}}」這樣的文本，改成「{{雪􏿽雨􏿽霽􏿽雷、電􏿽霧􏿽虹　　}}」
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string FixNoteDisorder_TransformText(string input)
        {/*
          {{元正月晦}}　　{{人日寒食}}　　{{正月十五日三月三日}}
          {{元正月晦}}　　{{人日寒食}}　　{{正月十五日三月三日}}
          */

            string spaces = "、";
            if (input.IndexOf("}}") > -1 && input.IndexOf("{{", input.IndexOf("{{") + 1) > input.IndexOf("}}"))
            {
                spaces = input.Substring(input.IndexOf("}}") + "}}".Length, input.IndexOf("{{", input.IndexOf("}}")) - (input.IndexOf("}}") + "}}".Length)).Replace("　", "􏿽");
            }

            var matches = Regex.Matches(input, @"\{\{(.*?)\}\}");
            List<string> firstItem = new List<string>();
            List<string> secondItem = new List<string>();

            foreach (Match match in matches)
            {
                string content = match.Groups[1].Value;
                StringInfo si = new StringInfo(content);

                if (si.LengthInTextElements >= 1)
                    firstItem.Add(si.SubstringByTextElements(0, si.LengthInTextElements % 2 == 1 ? si.LengthInTextElements / 2 + 1 : si.LengthInTextElements / 2));

                if (si.LengthInTextElements > 1)
                    secondItem.Add(si.SubstringByTextElements(si.LengthInTextElements % 2 == 1 ? si.LengthInTextElements / 2 + 1 : si.LengthInTextElements / 2));
            }

            string result = ("{{" + string.Join(spaces, firstItem) + "、" + string.Join(spaces, secondItem) + "}}").Replace(" ", "􏿽");
            return result;
            #region 單字且單空白間隔
            //var matches = Regex.Matches(input, @"\{\{(.*?)\}\}");

            //List<string> firstChars = new List<string>();
            //List<string> secondChars = new List<string>();

            //foreach (Match match in matches)
            //{
            //    string content = match.Groups[1].Value;

            //    if (content.Length >= 1)
            //        firstChars.Add(content.Substring(0, 1));

            //    if (content.Length > 1)
            //        secondChars.Add(content.Substring(1));
            //}

            //string result = "{{" + string.Join("􏿽", firstChars) + "、" + string.Join("􏿽", secondChars) + "　　}}";
            //return result;
            #endregion

        }




        /* creedit_with_Copilot大菩薩 20250214：……
         * 您已經使用了高效的方法，只是將其封裝成了函式。這是個好方法，也很專業。您可以考慮使用正則表達式來實現同樣的功能，這樣會使代碼更簡潔，但效能差不多：
         * 這種方法通過正則表達式來查找所有匹配的出現次數，如果出現次數等於1，則說明該字符串只出現一次。……
         */
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pageEndText10">要尋找的字串</param>
        /// <param name="textBox1">要比對的對象字串</param>
        /// <returns>若某字串在要比對的字串中只出現一之則傳回true</returns>
        public static bool IsPageEndTextUnique(string pageEndText10, TextBox textBox1)
        {
            string pattern = Regex.Escape(pageEndText10);
            MatchCollection matches = Regex.Matches(textBox1.Text, pattern);
            return matches.Count == 1;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pageEndText10">要尋找的字串</param>
        /// <param name="textBox1Text">要比對的對象字串</param>
        /// <returns>若某字串在要比對的字串中只出現一之則傳回true</returns>
        public static bool IsPageEndTextUnique(string pageEndText10, string textBox1Text)
        {
            int firstOccurrence = textBox1Text.IndexOf(pageEndText10);
            if (firstOccurrence == -1)
            {
                // Not found
                return false;
            }

            int secondOccurrence = textBox1Text.IndexOf(pageEndText10, firstOccurrence + pageEndText10.Length);
            return secondOccurrence == -1;
        }


        /* 20250217 GitHub　Copilot大菩薩：如果字符已經被截斷，那麼標準的 IndexOf 方法將無法正確地找出來。因為 IndexOf 方法是基於完整字符來匹配的，而截斷的字符會被視為兩個獨立的字符。
            為了處理這種情況，我們需要一個能夠在字符被截斷的情況下進行匹配的自定義函數。這可以通過比較每個字符（包括 surrogates）來實現。
            以下是一種方法，通過逐個字符比較來找到被截斷的字符串： …… https://copilot.microsoft.com/shares/Fexk2qkBqC4QuWdKo8xZQ
         這個方法確保即使字符串被截斷，也可以通過逐個字符（包括 surrogates）比較來找到匹配的字符串         */
        /// <summary>
        /// 自定義函數來查找部分匹配的字符串（當surrogate字符被截斷時）creedit_with_Copilot大菩薩。
        /// </summary>
        /// <param name="text"></param>
        /// <param name="pattern"></param>
        /// <param name="matchPosition"></param>
        /// <returns></returns>
        public static bool PartialMatch(String text, String pattern, out int matchPosition)
        {
            int textIndex = 0;
            int patternIndex = 0;
            matchPosition = -1;

            while (textIndex < text.Length && patternIndex < pattern.Length)
            {
                // 如果是代理對字符
                if (Char.IsHighSurrogate(text[textIndex]) && textIndex + 1 < text.Length &&
                    Char.IsLowSurrogate(text[textIndex + 1]))
                {
                    if (text.Substring(textIndex, 2) == pattern.Substring(patternIndex, 2))
                    {
                        if (matchPosition == -1)
                            matchPosition = textIndex;
                        patternIndex += 2;
                    }
                    else
                    {
                        patternIndex = 0;
                        matchPosition = -1;
                    }
                    textIndex += 2;
                }
                else
                {
                    if (text[textIndex] == pattern[patternIndex])
                    {
                        if (matchPosition == -1)
                            matchPosition = textIndex;
                        patternIndex++;
                    }
                    else
                    {
                        patternIndex = 0;
                        matchPosition = -1;
                    }
                    textIndex++;
                }
            }
            return patternIndex == pattern.Length;
        }
        /*
         // 使用自定義函數查找部分匹配的字符串位置
            int matchPosition;
            bool found = PartialMatch(text, pageEndText10, out matchPosition);

            if (found)
            {
                int end = matchPosition + pageEndText10.Length;
                if (end > pageTextEndPosition)
                    pageTextEndPosition = end;
                else if (CnText.IsPageEndTextUnique(pageEndText10, textBox1) && pageTextEndPosition != end)
                    pageTextEndPosition = end;
            }

            // end 變量現在將包含找到的匹配字符串位置
         */

        //Gemini大菩薩：202512223 https://gemini.google.com/share/23e368f6f258
        public static void JSON提取中文內容(ref string input)
        {
            //string input = @"顧氏"",""objectRef"":null...（您的文本內容）";

            // 定義匹配範圍：
            // 1. \u4e00-\u9fa5 : 基本漢字
            // 2. \u3000-\u303f : 中式標點符號 (如：，。、)
            // 3. \uff01-\uff0f, \uff1a-\uff20, \uff3b-\uff40, \uff5b-\uff5e : 其他全形標點
            string pattern = @"[^\u4e00-\u9fa5\u3000-\u303f\uff01-\uff0f\uff1a-\uff20\uff3b-\uff40\uff5b-\uff5e]";

            // 將不符合上述範圍的字元替換為空字串
            //string result = Regex.Replace(input, pattern, "");
            input = Regex.Replace(input, pattern, "");

            //Console.WriteLine(result);
        }

        /// <summary>
        /// 《識典古籍》檢視原始碼內容轉換成ctext.org簡單編輯輸入格式
        /// Gemini大菩薩：202512223 https://gemini.google.com/share/23e368f6f258
        /// </summary>
        /// <param name="rawJson"></param>
        /// <returns></returns>
        public static string JSON提取中文內容_ProcessTextWithLineBreaks(string rawJson)
        {
            // 補齊為 JSON 陣列格式
            string formattedJson = rawJson.Trim().StartsWith("[") ? rawJson : "[" + rawJson + "]";
            JArray lines = JArray.Parse(formattedJson);

            StringBuilder result = new StringBuilder();

            foreach (var line in lines)
            {
                int lineType = line["lineType"]?.Value<int>() ?? 0;
                string content = line["content"]?.ToString() ?? "";

                // 1. 處理內容：僅保留漢字與中式標點
                string cleanContent = Regex.Replace(content, @"[^\u4e00-\u9fa5\u3000-\u303f\uff01-\uff0f\uff1a-\uff20\uff3b-\uff40\uff5b-\uff5e]", "");

                // 2. 根據類型決定輸出格式
                if (lineType == 1) // 正文
                {
                    result.AppendLine(cleanContent); // 結尾加換行
                }
                else if (lineType == 2) // 注文
                {
                    if (!string.IsNullOrEmpty(cleanContent))
                    {
                        result.AppendLine($"{{{{{cleanContent}}}}}"); // 注文也換行，以對應原書行數
                    }
                }
                else if (lineType == 3) // 分頁標記
                {
                    // 如果 lineType 為 3，代表原書在此分頁，可插入分頁線
                    result.AppendLine("\n---【原書分頁點】---\n");
                }
            }
            return result.ToString();
        }

        public static bool IsChineseChar(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            char c = s[0];
            // 判定是否為中文字元範圍（包含擴展區）
            return (c >= 0x4E00 && c <= 0x9FFF) || (c >= 0x3400 && c <= 0x4DBF) || char.IsHighSurrogate(c);
        }
    }


    /// <summary>
    /// 《識典古籍》檢視網頁原始碼的內容轉換成ctext.org簡單修改模式的內容
    /// Gemini大菩薩：C# 文本清除與標點符號保留 202512223 https://gemini.google.com/share/23e368f6f258
    /// </summary>
    public class AncientTextConverter
    {
        /// <summary>
        /// 《識典古籍》檢視網頁原始碼的內容轉換成ctext.org簡單修改模式的內容
        /// </summary>
        /// <param name="inputSource"></param>
        /// <returns></returns>
        public static string ConvertHtmlToAncientFormat(string inputSource)
        {
            if (string.IsNullOrWhiteSpace(inputSource)) return "輸入內容為空";

            try
            {
                // 1. 基礎解碼與 Unicode 轉義還原 (確保 □、■、◯ 能正確顯示)
                string decoded = WebUtility.HtmlDecode(inputSource);
                decoded = UnescapeUnicode(decoded);

                // 2. 識典古籍專屬：鎖定包含 lineType 與 content 的數據區塊
                // 我們依然需要定位 content 欄位，否則會抓到 lineId 等數字
                string pattern = @"\\""lineType\\"":(?<type>\d+),\\""content\\"":\\""(?<text>.*?)(?<!\\)\\""";

                MatchCollection matches = Regex.Matches(decoded, pattern, RegexOptions.Singleline);

                if (matches.Count == 0) return "無法定位數據內容";

                StringBuilder finalResult = new StringBuilder();

                foreach (Match m in matches)
                {
                    string typeStr = m.Groups["type"].Value;
                    string rawContent = m.Groups["text"].Value;

                    // 3. 淨化內容 (只處理 JSON 轉義，不處理字元過濾)
                    // 將內容中的 \" 轉回 "，\\ 轉回 \
                    string cleanContent = rawContent.Replace("\\\"", "\"").Replace("\\\\", "\\");

                    // 移除內容中可能殘留的控制符 (如 \n, \t)
                    cleanContent = Regex.Replace(cleanContent, @"\\[rntf]", "");

                    if (string.IsNullOrWhiteSpace(cleanContent)) continue;

                    // 4. 根據類型排版，並保留所有字元（含 □ ■ ◯ 〇 等）
                    if (typeStr == "1") // 正文
                    {
                        finalResult.AppendLine(cleanContent);
                    }
                    else if (typeStr == "2") // 注文
                    {
                        finalResult.AppendLine($"{{{{{cleanContent}}}}}");
                    }
                }

                return finalResult.ToString().Trim();
            }
            catch (Exception ex)
            {
                return $"【處理出錯】：{ex.Message}";
            }
        }

        /// <summary>
        /// 將 \uXXXX 格式的字串轉回實際字元，這是保留缺字符號的關鍵
        /// </summary>
        private static string UnescapeUnicode(string input)
        {
            try
            {
                return Regex.Replace(input, @"\\u(?<val>[a-fA-F0-9]{4})", m =>
                    ((char)Convert.ToInt32(m.Groups["val"].Value, 16)).ToString());
            }
            catch { return input; }
        }
    }

    /// <summary>
    /// 注解排版校正
    /// Gemini大菩薩 20251224平安夜 https://gemini.google.com/share/fb51391730eb
    /// </summary>
    public class AncientTextRestorer
    {
        public static string UniversalSwap(string input)
        {

            if (string.IsNullOrEmpty(input)) return input;

            string specialBlank = "􏿽";
            string[] lines = input.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            StringBuilder finalResult = new StringBuilder();

            foreach (string line in lines)
            {
                //if (!line.Contains("/")) { finalResult.AppendLine(line); continue; }
                if (!line.Contains(@"\")) { finalResult.AppendLine(line); continue; }//現在在WordVBA處理源文本時，改用「\」來表示夾注小注換行之標識符，以與造字缺字者作區別

                //string[] parts = line.Split('/');
                string[] parts = line.Split('\\');//現在在WordVBA處理源文本時，改用「\」來表示夾注小注換行之標識符，以與造字缺字者作區別
                string partA = parts[0];
                //string partB = "/" + parts[1]; // 保留斜槓作為基準點
                string partB = @"\" + parts[1]; // 保留斜槓作為基準點//現在在WordVBA處理源文本時，改用「\」來表示夾注小注換行之標識符，以與造字缺字者作區別

                // 建立視覺寬度地圖
                var mapA = BuildVisualWidthMap(partA);
                var mapB = BuildVisualWidthMap(partB);

                // 找出 B 段中的坑位並從 A 段對應座標抓取文字
                // 注意：我們遍歷 B 段的每一個座標點
                var sortedBKeys = mapB.Keys.OrderBy(k => k).ToList();
                foreach (var pos in sortedBKeys)
                {
                    // 如果 B 段這個位置是半形空格，且下一個位置也是半形空格 (構成一個漢字坑)
                    if (mapB.ContainsKey(pos) && mapB[pos].Content == " " &&
                        mapB.ContainsKey(pos + 1) && mapB[pos + 1].Content == " ")
                    {
                        // 檢查 A 段在相同座標是否有「非空格」的文字
                        if (mapA.ContainsKey(pos) && mapA[pos].Content != " ")
                        {
                            var nodeA = mapA[pos];
                            var nodeB1 = mapB[pos];
                            var nodeB2 = mapB[pos + 1];

                            // 執行物理對調
                            string textToMove = nodeA.Content;
                            nodeB1.Content = textToMove; // 填入文字
                            nodeB2.Content = "REMOVE";   // 標記待移除的第二個半形位
                            nodeA.Content = "  ";        // A 段原位留空 (稍後轉特殊符號)
                        }
                    }
                }

                // 重組並處理特殊空白符號
                string resultA = Reconstruct(mapA).Replace("  ", specialBlank);
                //string resultB = Reconstruct(mapB).Replace("REMOVE", "");
                //string resultB = Reconstruct(mapB).Replace("/", "");
                string resultB = Reconstruct(mapB).Replace(@"\", "");//現在在WordVBA處理源文本時，改用「\」來表示夾注小注換行之標識符，以與造字缺字者作區別

                finalResult.AppendLine(resultA + resultB);
            }

            return finalResult.ToString().TrimEnd();
        }

        class WidthNode { public string Content; }

        private static Dictionary<int, WidthNode> BuildVisualWidthMap(string s)
        {
            var map = new Dictionary<int, WidthNode>();
            TextElementEnumerator charEnum = StringInfo.GetTextElementEnumerator(s);
            int currentWidth = 0;

            while (charEnum.MoveNext())
            {
                string element = charEnum.GetTextElement();
                // 判斷該 Text Element 的物理寬度：半形字元寬 1，其餘(漢字等)寬 2
                int w = (element.Length == 1 && element[0] < 128) ? 1 : 2;

                var node = new WidthNode { Content = element };
                map[currentWidth] = node;

                // 如果寬度是 2，我們讓地圖的兩個座標都指向同一個節點，方便偵測
                if (w == 2)
                {
                    map[currentWidth + 1] = node;
                }

                currentWidth += w;
            }
            return map;
        }

        private static string Reconstruct(Dictionary<int, WidthNode> map)
        {
            StringBuilder sb = new StringBuilder();
            int lastPos = -1;
            foreach (var kvp in map.OrderBy(k => k.Key))
            {
                // 避免重複添加寬度為 2 的同一個節點
                if (kvp.Value.Content == "REMOVE") continue;

                // 這裡用對象引用檢查，確保寬度為 2 的漢字只被添加一次
                if (lastPos != -1 && map[lastPos] == kvp.Value) continue;

                sb.Append(kvp.Value.Content);
                lastPos = kvp.Key;
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// 夾注文自動校正 Gemini大菩薩 20251224平安夜 ：https://gemini.google.com/share/639690183e06
    /// Provides utilities for processing and correcting ancient or complex text formats, including visual width
    /// adjustments and line formatting within delimited regions.
    /// </summary>
    /// <remarks>The AncientTextTool class is designed for scenarios where text contains special formatting,
    /// such as visually aligned columns or embedded markup (e.g., text enclosed in double curly braces). It offers
    /// static methods to process such text, making it suitable for applications involving historical documents, digital
    /// humanities, or advanced text layout correction. All members are static and thread-safe.</remarks>
    public class AncientTextTool
    {
        /// <summary>
        /// 處理入口：解析 {{ }} 區塊並執行校正與排版 https://gemini.google.com/share/8ed9c9c823dd
        /// </summary>
        /// <param name="input">原始文本</param>
        /// <param name="lineLengthLimit">單行長度閾值（字數）</param>
        public static string ProcessText(string input, int lineLengthLimit)
        {
            if (string.IsNullOrEmpty(input)) return input;

            return Regex.Replace(input, @"\{\{([\s\S]*?)\}\}", m =>
            {
                string innerText = m.Groups[1].Value;

                // --- 智慧判斷門檻開始 ---
                // 只有符合條件的注文才執行校正，否則直接進入排版判斷
                string processedContent = innerText;

                if (ShouldApplyCorrection(innerText))
                {
                    processedContent = CorrectVisualWidth(innerText);
                }
                // --- 智慧判斷門檻結束 ---

                // 處理斜槓邏輯
                StringInfo si = new StringInfo(processedContent.Replace("/", ""));
                if (si.LengthInTextElements < lineLengthLimit * 2)
                    processedContent = processedContent.Replace("/", "");
                else
                    processedContent = processedContent.Replace("/", Environment.NewLine);

                return "{{" + processedContent + "}}";
            });
        }

        /// <summary>
        /// 智慧判斷門檻：決定這段注文是否需要執行「空格/文字反轉」校正 https://gemini.google.com/share/022b7575825b
        /// </summary>
        private static bool ShouldApplyCorrection(string text)
        {
            // 條件 1：目前的已知條件，含有「字缺」二字
            if (text.Contains("字缺")) return true;

            // 條件 2：日後您若發現其他關鍵字，可以加在這裡
            // 例如：if (text.Contains("稱為")) return true;

            // 條件 3：敏感字詞庫（範例）
            /*
            string[] sensitiveWords = { "撰", "註", "補" };
            foreach (var word in sensitiveWords) {
                if (text.Contains(word)) return true;
            }
            */

            return false; // 預設不執行校正，避免誤傷正確的排版
        }

        private static string CorrectVisualWidth(string input)
        {
            string specialBlank = "􏿽";
            string[] lines = input.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            List<string> processedLines = new List<string>();

            foreach (var line in lines)
            {
                if (!line.Contains("/")) { processedLines.Add(line); continue; }

                int slashIdx = line.IndexOf('/');
                string partA = line.Substring(0, slashIdx);
                string partB = line.Substring(slashIdx);

                var mapA = BuildWidthMap(partA);
                var mapB = BuildWidthMap(partB);

                foreach (var pos in mapB.Keys.OrderBy(k => k))
                {
                    if (mapB.ContainsKey(pos) && mapB[pos].Content == " " &&
                        mapB.ContainsKey(pos + 1) && mapB[pos + 1].Content == " ")
                    {
                        if (mapA.ContainsKey(pos) && mapA[pos].Content != " ")
                        {
                            var nodeA = mapA[pos];
                            mapB[pos].Content = nodeA.Content;
                            mapB[pos + 1].Content = "REMOVE";
                            nodeA.Content = "  ";
                        }
                    }
                }
                processedLines.Add(Reconstruct(mapA).Replace("  ", specialBlank) + Reconstruct(mapB).Replace("REMOVE", ""));
            }
            return string.Join(Environment.NewLine, processedLines);
        }

        // --- 輔助工具 ---
        class WidthNode { public string Content; }

        private static Dictionary<int, WidthNode> BuildWidthMap(string s)
        {
            var map = new Dictionary<int, WidthNode>();
            TextElementEnumerator charEnum = StringInfo.GetTextElementEnumerator(s);
            int currentWidth = 0;
            while (charEnum.MoveNext())
            {
                string element = charEnum.GetTextElement();
                int w = (element.Length == 1 && element[0] < 128) ? 1 : 2;
                var node = new WidthNode { Content = element };
                map[currentWidth] = node;
                if (w == 2) map[currentWidth + 1] = node;
                currentWidth += w;
            }
            return map;
        }

        private static string Reconstruct(Dictionary<int, WidthNode> map)
        {
            StringBuilder sb = new StringBuilder();
            HashSet<WidthNode> seen = new HashSet<WidthNode>();
            foreach (var k in map.Keys.OrderBy(x => x))
            {
                if (map[k].Content == "REMOVE" || seen.Contains(map[k])) continue;
                sb.Append(map[k].Content);
                seen.Add(map[k]);
            }
            return sb.ToString();
        }
    }

    public static class TextBoxPagingHelper
    {
        /// <summary>
        /// 取得指定頁面的文字內容（單純依行數分頁）
        /// </summary>
        public static string GetPageText(TextBox textBox, int pageNumber, int linesPerPage)
        {
            if (textBox == null) throw new ArgumentNullException(nameof(textBox));
            if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber));
            if (linesPerPage < 1) throw new ArgumentOutOfRangeException(nameof(linesPerPage));

            string[] allLines = textBox.Lines;

            int startIndex = (pageNumber - 1) * linesPerPage;
            if (startIndex >= allLines.Length) return string.Empty;

            int endIndex = Math.Min(startIndex + linesPerPage, allLines.Length);

            string[] pageLines = allLines.Skip(startIndex).Take(endIndex - startIndex).ToArray();

            return string.Join(Environment.NewLine, pageLines);
        }

        /// <summary>
        /// 計算總頁數（單純依行數分頁）
        /// </summary>
        public static int GetTotalPages(TextBox textBox, int linesPerPage)
        {
            if (textBox == null) throw new ArgumentNullException(nameof(textBox));
            if (linesPerPage < 1) throw new ArgumentOutOfRangeException(nameof(linesPerPage));

            int totalLines = textBox.Lines.Length;
            return (int)Math.Ceiling((double)totalLines / linesPerPage);
        }

        /// <summary>
        /// 計算一頁中的「實際行/段數」
        /// （正文算 2，小注算 1，依 {{ }} 標記判斷）
        /// </summary>
        public static int CountLinesPerPage(string xPage)
        {
            // 這裡直接放你提供的邏輯，
            // 或者抽取成獨立的 private 方法以保持整潔。
            // 為了簡潔，這裡只示範呼叫你的原始方法。
            return AncientTextExamine.CountLinesPerPage(xPage);
        }
    }

    /// <summary>
    /// 專門處理古籍文本的行數字數頁面排版訊息等計算邏輯
    /// </summary>
    public static class AncientTextExamine
    {//https://copilot.microsoft.com/shares/xuC8fA9dA7oyxp9XGKP93 GitHub　Copilot大菩薩：C# 分頁顯示 TextBox 行內容 20260107
        const string omitStr = "．‧.…【】〖〗＝{}<p>（）《》〈〉：；、，。「」『』？！0123456789-‧·\r\n";//"　"
        /// <summary>
        /// 將全文依古籍實際行數分頁 依古籍實際行數（正文=2，注文=1，夾注=1）分頁。 
        /// 以「累積文字段」呼叫 CountLinesPerPage，避免逐行計算造成頁尾誤差。
        /// </summary>
        /// <param name="text">全文</param>
        /// <param name="linesPerPage">每頁的實際行數上限</param>
        /// <returns>分頁後的文字集合，每頁一個字串</returns>
        /// <summary>
        /// 將全文依古籍實際行數分頁
        /// </summary>
        public static List<string> SplitIntoPages(string text, int linesPerPage)
        {
            if (string.IsNullOrEmpty(text)) return new List<string>();
            if (linesPerPage < 1) throw new ArgumentOutOfRangeException(nameof(linesPerPage));

            string[] rawLines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var pages = new List<string>();
            var bufferLines = new List<string>();//https://copilot.microsoft.com/shares/SAeUq1UcVuawWFqKv4ZNh

            foreach (string line in rawLines)
            {
                bufferLines.Add(line);
                string bufferText = string.Join(Environment.NewLine, bufferLines);

                int counted = CountLinesPerPage(bufferText);

                if (counted >= linesPerPage)
                {
                    pages.Add(bufferText);
                    bufferLines.Clear();
                }
            }

            if (bufferLines.Count > 0)
            {
                pages.Add(string.Join(Environment.NewLine, bufferLines));
            }

            return pages;
        }

        /// <summary>
        /// 取得指定頁面的文字內容（傳入 List<string>）
        /// </summary>
        public static string GetPageText(List<string> pages, int pageNumber)
        {
            if (pages == null) throw new ArgumentNullException(nameof(pages));
            if (pageNumber < 1 || pageNumber > pages.Count) return string.Empty;
            return pages[pageNumber - 1];
        }

        /// <summary>
        /// 取得指定頁面的文字內容（傳入全文 string，自動分頁）
        /// </summary>
        public static string GetPageText(string text, int pageNumber, int linesPerPage)
        {
            var pages = SplitIntoPages(text, linesPerPage);
            return GetPageText(pages, pageNumber);
        }

        /// <summary>
        /// 計算總頁數
        /// </summary>
        public static int GetTotalPages(string text, int linesPerPage)
        {
            return SplitIntoPages(text, linesPerPage).Count;
        }


        ///// <summary>
        ///// 調用您現成的 countLinesPerPage 方法
        ///// </summary>
        //private static int CountLinesPerPage(string xPage)
        //{
        //    // 這裡直接放您原本的 countLinesPerPage 方法完整內容
        //    // （您已經提供了，我建議直接複製貼上）
        //    throw new NotImplementedException("請貼上您原本的 countLinesPerPage 方法內容");
        //}

        /// <summary>
        /// 計算xPage中的實際行/段數（正文算1，注文算2，故8行正文之文本傳回的值即16）
        /// </summary>
        /// <param name="xPage">要計算的文本</param>
        /// <returns>xPage中的實際行/段數</returns>
        public static int CountLinesPerPage(string xPage)
        {
            #region 規範文本

            while (xPage.Contains("􏿽<p>"))
            {
                xPage = xPage.Replace("􏿽<p>", string.Empty);
            }
            while (xPage.Contains("<p>+"))
            {
                xPage = xPage.Replace("<p>+", "<p>");
            }
            while (xPage.Contains("+<p>"))
            {
                xPage = xPage.Replace("+<p>", "<p>");
            }
            while (xPage.Contains("<p>+<p>"))
            {
                xPage = xPage.Replace("<p>+<p>", "<p>");
            }
            while (xPage.Contains(Environment.NewLine + Environment.NewLine))
                xPage = xPage.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine + "|" + Environment.NewLine);
            if (xPage.Length > Environment.NewLine.Length && xPage.Substring(xPage.Length - Environment.NewLine.Length, Environment.NewLine.Length) == Environment.NewLine)
                xPage = xPage + "|";

            #endregion
            int i = 0, openBracketS, closeBracketS, e = xPage.IndexOf(Environment.NewLine); bool openNote = false;
            string[] linesParasPage = xPage.Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            string item;
            if (linesParasPage.Length == 1)
            {
                item = linesParasPage[0];
                if (item == "|") return 2;
                else
                {
                    if (item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)
                        return 2;
                    else
                    {
                        if ((item.StartsWith("{{") && item.IndexOf("}}") == -1)
                            || (item.EndsWith("}}") && item.IndexOf("{{") == -1)
                            || (item.StartsWith("{{") && item.EndsWith("}}")))
                            return 1;
                        else
                            return 2;

                    }
                }
            }
            for (int index = 0; index < linesParasPage.Length; index++)
            //foreach (string item in linesParasPage)
            {
                #region 舊式
                item = linesParasPage[index];
                //if (item == "") return;
                openBracketS = item.IndexOf("{{"); closeBracketS = item.IndexOf("}}");

                if (item == "}}<p>" || item == "}}|" || (closeBracketS == -1 && openBracketS == 0 && item.Length < 5))//《維基文庫》純注文空及其前一行
                {
                    i++;
                    if (item == "}}<p>" || item == "}}|") openNote = false; else openNote = true;
                }

                else if (i == 0 && xPage.IndexOf("}}") > -1 && xPage.IndexOf("}}") < (xPage.IndexOf("{{") == -1 ? xPage.Length : xPage.IndexOf("{{")) && xPage.IndexOf("}}") > e)
                { i++; openNote = true; }//第一段/行是純注文        
                else if (i == 0 && item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)
                {
                    string xx = linesParasPage[i + 1];
                    if (xx.IndexOf("}}") > -1 && xx.IndexOf("{{") == -1)//&& x.IndexOf("}}") > e)
                    { i++; openNote = true; }//第一段/行是純注文
                    else { i += 2; openNote = false; }//第一段/行是純正文
                }

                else if (i == 0 && ((closeBracketS > -1 && openBracketS > closeBracketS) ||
                    (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2))) //第一行正、注夾雜
                {
                    if (openBracketS > 2)
                    {
                        i += 2;
                    }
                    else
                    {
                        if (openBracketS == -1) i += 2;
                        else if (openBracketS == 1)
                        {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1)
                            {
                                i += 2;
                            }
                            else i++;
                        }
                        else if (openBracketS == 2)
                        {
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1) i += 2;
                            else if (omitStr.IndexOf(item.Substring(1, 1)) == -1)
                            {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                                i += 2;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }

                    if (item.LastIndexOf("}}") > item.LastIndexOf("{{"))
                        openNote = false;
                    else
                        openNote = true;
                }

                else if (openBracketS == 0 && closeBracketS == -1)//獨立注文（純注文）（開始）
                { i++; openNote = true; }
                else if (openBracketS == -1 && openNote)
                {//純注文（末截）
                    if (closeBracketS == item.Length - 2)//第2行/段純注文、獨立注文
                    { i++; openNote = false; }
                    else if (item.Length > 4 || item.EndsWith("}}|"))
                    {
                        if (item.Substring(item.Length - 3) == "}}|" ||
                            item.Substring(item.Length - 5) == "}}<p>") { i++; openNote = false; }
                        else
                        {
                            if (closeBracketS == -1)
                            {
                                if (openNote)
                                    i++;
                                else
                                    i += 2;
                            }
                            else
                            {
                                i += 2;
                                openNote = false;
                            }

                        }
                    }
                }
                else if (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2)
                {//正注夾雜注文結束
                    { i += 2; openNote = false; }
                }
                else if (openBracketS > -1 && item.IndexOf("{{", openBracketS + 2) > -1
                    && item.IndexOf("{{{", openBracketS + 2) != item.IndexOf("{{", openBracketS + 2))//正注夾雜
                {
                    i += 2;
                    if (item.LastIndexOf("}}") < item.LastIndexOf("{{")) openNote = true;
                    else openNote = false;
                }
                else if (openBracketS > -1 && closeBracketS > -1 && closeBracketS < item.Length - 2)
                {
                    if (openBracketS == 0 && (item.EndsWith("}}") || item.EndsWith("}}<p>") || item.EndsWith("}}|") || item.EndsWith("}}。<p>"))
                        && ((item.IndexOf("{{", openBracketS + 2) == -1 && item.LastIndexOf("}}", closeBracketS) == -1)
                        || (item.IndexOf("{{", openBracketS + 2) == item.IndexOf("{{{", openBracketS + 2) &&
                                item.LastIndexOf("}}", closeBracketS) == item.LastIndexOf("}}}", closeBracketS))
                        ))//獨立注文
                    {
                        string preItem, preItem1;
                        if (index == 0) i += 2;
                        else if (index > 0)
                        {
                            preItem = linesParasPage[index - 1];
                            if (index > 1)
                                preItem1 = linesParasPage[index - 2];
                            else
                                preItem1 = null;
                            if (((preItem.IndexOf("{{") == -1 || (preItem.StartsWith("{{") && preItem.IndexOf("{{", 2) == -1))
                                        || (preItem.IndexOf("{{") > 0 && preItem.IndexOf("}}") > (preItem.IndexOf("{{"))))
                                && (preItem.EndsWith("}}") || preItem.EndsWith("}}<p>") || preItem.EndsWith("}}|") || preItem.EndsWith("}}。<p>"))
                                && preItem.LastIndexOf("}}", preItem.LastIndexOf("}}")) == -1
                                //前一行是獨立注文
                                && ((item.EndsWith("}}<p>") || item.EndsWith("}}|") || item.EndsWith("}}。<p>")) &&
                                    ((preItem.StartsWith("{{") && preItem.EndsWith("}}") && preItem.IndexOf("{{", 2) == -1)
                                    || preItem.Contains("{{") == false && preItem.EndsWith("}}")))
                                && (//本段是末有<p>的獨立注文，而前2段不是獨立注文
                                    !(((item.EndsWith("}}<p>") || item.EndsWith("}}|") || item.EndsWith("}}。<p>")) && item.LastIndexOf("}}", item.LastIndexOf("}}")) == -1)
                                        && ((preItem1 != null && (preItem1.StartsWith("{{") && preItem1.IndexOf("{{", 2) == -1)
                                        && preItem.EndsWith("}}") && preItem.LastIndexOf("}}", preItem.Length - 2) == -1)
                                        ||
                                        (preItem1 != null && (!preItem1.Contains("{{") && !preItem1.Contains("}}"))
                                            && !preItem.Contains("{{") && preItem.EndsWith("}}") && preItem.LastIndexOf("}}", preItem.Length - 2) == -1)
                                        ))
                                    //或者本/行段與前行/段字數差不超過1
                                    || Math.Abs(CountWordsLenPerLinePara(item) - CountWordsLenPerLinePara(preItem)) <= 1)
                                )
                                i++;
                            else
                                i += 2;
                        }
                        openNote = false;
                    }
                    else//●●●●●●●●●●●●●●●
                    {//正注夾雜
                        i += 2;
                        if (item.LastIndexOf("}}") < item.LastIndexOf("{{")) openNote = true;
                        else openNote = false;
                    }

                }

                //無{{}}標記：
                else if (openBracketS == -1 && closeBracketS == -1)
                {
                    if (openNote == false)//《維基文庫》純正文
                        i += 2;
                    else //《維基文庫》純注文
                        i++;
                }

                //《維基文庫》正注文夾雜
                else if (openBracketS > 0)//正注夾雜
                {
                    if (openBracketS > 2)
                    {
                        i += 2;
                    }
                    else
                    {
                        if (openBracketS == 1)
                        {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1)
                            {
                                i += 2;
                            }
                            else i++;
                        }
                        if (openBracketS == 2)
                        {
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1) i += 2;
                            else if (omitStr.IndexOf(item.Substring(1, 1)) == -1)
                            {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                                i += 2;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }
                    if (closeBracketS == -1) openNote = true;
                    else
                    {
                        if (item.LastIndexOf("}}") > item.LastIndexOf("{{"))
                            openNote = false;
                        else
                            openNote = true;
                    }
                }
                //else if (openBracketS > 0 && closeBracketS == -1) { i += 2; openNote = true; }
                else if (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2) { i += 2; openNote = false; }
                /*
                if ((item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)//純正文
                    || item.IndexOf("{{") > 0 || item.IndexOf("}}") + 2 < item.Length)////正注文夾雜
                {
                    count += 2;
                }
                else if ((item.IndexOf("{{") > -1 && item.IndexOf("}}") > -1)//純注文
                    || (item.IndexOf("{{") > -1 && item.IndexOf("}}") == -1)
                    || (item.IndexOf("{{") == -1 && item.IndexOf("}}") > -1))
                {
                    count++;
                }
                */

                //else if 一行裡面不能有2個「{{」或「}}」，請完成這樣的判斷式 20250304 GitHub Copilot大菩薩：
                else if (item.Count(c => c == '{') > 2 || item.Count(c => c == '}') > 2)//只要不止1個大括號，不論是上或下大括弧，就必非獨立注文
                {
                    // 處理包含兩個或更多「{{」或「}}」的行
                    Debugger.Break();
                    i += 2;//●●●●●●●●●●●●●●●●●●●●●
                }
                else if (openBracketS == 0 && closeBracketS == item.Length - 2
                    && item.IndexOf("{{", openBracketS + 2) == -1 && item.LastIndexOf("}}", closeBracketS) == -1) { i++; openNote = false; }
                //●●●●●●●●●●●●●●●●●●●●●
                else if (openBracketS == -1 && closeBracketS == item.Length - 2
                    && item.IndexOf("{{", openBracketS + 2) == -1 && item.LastIndexOf("}}", closeBracketS) == -1) { i++; openNote = false; }
            }
            return i;
            #endregion

        }
        /// <summary>
        /// 作為 countLinesPerPage 之改良，但未成功！202050301
        /// </summary>
        /// <param name="xPage"></param>
        /// <returns></returns>
        static public int CountLinesPerPage_NewTry(string xPage)
        {

            string[] linesParasPage = xPage.Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);

            int count = 0, index = -1;
            foreach (string item in linesParasPage)
            {
                index++;
                if (item == "|") count += 2;
                else
                {
                    if (item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)
                        count += 2;
                    else if (item.Contains("*"))
                        count += 2;
                    else if (item.IndexOf("{{") > 0)
                        count += 2;
                    else if (item.IndexOf("<p>") == -1 && item.LastIndexOf("}}") > -1 && item.LastIndexOf("}}") < item.Length - "}}".Length)
                        count += 2;
                    else if (item.LastIndexOf("}}") + 2 < item.IndexOf("<p>") || item.LastIndexOf("}}") + 2 < item.IndexOf("。<p>"))
                        count += 2;
                    else if (item.StartsWith("{{") && item.LastIndexOf("}}") == -1)
                        count++;
                    else
                    {
                        //落單的獨立注文，或夾注尾字數不足者
                        if ((item.StartsWith("{{") && item.EndsWith("}}"))
                            || (item.StartsWith("{{") && item.EndsWith("}}<p>"))
                            || (item.StartsWith("{{") && item.EndsWith("}}。<p>")))
                        {
                            string itemShort = item.Substring(2, item.IndexOf("}}") - 2);
                            if (!itemShort.Contains("{{") && !itemShort.Contains("}}"))
                            {
                                bool singleNoteLine = false;
                                if (index > 0)
                                {
                                    string preLine = linesParasPage[index - 1];
                                    string nextLine = string.Empty;

                                    if (index + 1 < linesParasPage.Length)
                                        nextLine = linesParasPage[index + 1];
                                    if (!preLine.Contains("{{") && !preLine.Contains("}}"))
                                        singleNoteLine = true;
                                    else if (preLine.EndsWith("}}"))
                                    {
                                        if (nextLine != string.Empty && !nextLine.StartsWith("{{"))
                                        {
                                            string preLineShort = preLine.Substring(2, preLine.IndexOf("}}") - 2);
                                            if (!(preLine.StartsWith("{{") && preLine.EndsWith("}}")
                                                && preLineShort.IndexOf("{{") == -1 && preLineShort.IndexOf("}}") == -1))
                                                singleNoteLine = true;
                                        }
                                        else if (new StringInfo(itemShort).LengthInTextElements == 1)
                                            singleNoteLine = true;
                                    }

                                    if (singleNoteLine)
                                        count += 2;
                                    else
                                        count++;
                                }
                                else
                                    count += 2;
                            }//以上落單獨立注文的處理
                             //獨立注文
                            else if ((item.StartsWith("{{") && item.IndexOf("}}") == -1)
                                || (item.EndsWith("}}") && item.IndexOf("{{") == -1)
                                || (item.EndsWith("}}<p>") && item.IndexOf("{{") == -1)
                                || (item.EndsWith("}}。<p>") && item.IndexOf("{{") == -1))
                                count++;
                            else
                                count += 2;

                        }
                    }
                }
            }
            return count;
        }


        //public static bool IsInBraces(string text, int start, int end)
        //{
        //    text = "　　　*念奴嬌{{錄一　}}　　　　　{{宋}}蘇　軾<p>";

        //    start = text.IndexOf('*') + 1;
        //    end = text.IndexOf("<p>");

        //    StringBuilder sb = new StringBuilder(text);

        //    bool inBraces = false; // 是否在 {{}} 内

        //    for (int i = start; i < end; i++)
        //    {
        //        // 检查是否进入或退出 {{}} 块
        //        if (text.Substring(i).StartsWith("{{"))
        //        {
        //            inBraces = true;
        //            i += 1; // 跳过 '{'，因为下一次循环会再加 1
        //            continue;
        //        }
        //        if (text.Substring(i).StartsWith("}}"))
        //        {
        //            inBraces = false;
        //            i += 1;
        //            continue;
        //        }
        //        if (!inBraces && text[i] == '　')
        //        {
        //            sb[i] = '􏿽';
        //        }
        //    }

        //    string result = sb.ToString();
        //    Console.WriteLine(result);
        //}

    }




}


