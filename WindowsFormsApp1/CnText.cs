using FuzzySharp;
using ado = ADODB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WindowsFormsApp1;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;
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
                text = splitString(text, line1CharCount, line2CharCount);
            }
            else
            {
                int line1CharCount = charCount / 2 + 1;
                int line2CharCount = charCount / 2;
                //Console.WriteLine(SplitString(text, line1CharCount, line2CharCount));
                text = splitString(text, line1CharCount, line2CharCount);
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
        private static string splitString(string text, int line1CharCount, int line2CharCount)
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
            if (!force2mark) if (HasEditedWithPunctuationMarks(ref clpTxt)) { Form1.playSound(Form1.soundLike.error); return ref clpTxt; }
            //提示音
            //new SoundPlayer(@"C:\Windows\Media\Windows Balloon.wav").Play();
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
                Form1.playSound(Form1.soundLike.warn);

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
            ////我自己的，更簡短，更優化；運作邏輯忖度，應當還是用正則表達式好，因為這是一次取代，與WordVBA中的逐一檢查不同；以下這樣寫，則只瞻前、未顧後，誠掛一漏萬者也。感恩感恩　南無阿彌陀佛 20230312
            //string patternCntext = context.Substring(0, pos_Term);
            //if (patternCntext.LastIndexOf("《") <= patternCntext.LastIndexOf("》")
            //    && patternCntext.LastIndexOf("〈") <= patternCntext.LastIndexOf("〉"))
            //    context = context.Replace(term, termReplaced);


            //if(patternCntext.LastIndexOf("")==-1&& patternCntext.LastIndexOf("")==-1)
            //if (termReplaced.IndexOf("《") > -1)
            //{
            //    if (patternCntext.LastIndexOf("《") <= patternCntext.LastIndexOf("》")
            //        && patternCntext.LastIndexOf("〈") <= patternCntext.LastIndexOf("〉"))
            //        context = context.Replace(term, termReplaced);
            //}
            //else
            //{
            //    if (patternCntext.LastIndexOf("〈") <= patternCntext.LastIndexOf("〉")
            //        && patternCntext.LastIndexOf("《") <= patternCntext.LastIndexOf("》"))
            //        context = context.Replace(term, termReplaced);
            //}



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
            ////string replacement = item[1];
            ////string replacement = termReplaced;
            ////text = Regex.Replace(text, pattern, replacement);
            ////20230311 Bing大菩薩：
            ////string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            //20230311 合併Bing大菩薩與之前chatGPT大菩薩的：
            //不計星號「*」：
            //string pattern = "(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            //避開有星號「*」的段落：
            string pattern = "(?<!\\*)(?<!《[^》]*)(?<!〈[^〉]*)(?<![\\p{P}&&[^》〉]]+)" + Regex.Escape(term) + "(?![^《]*》)(?![^〈]*〉)";
            /* Bing大菩薩： 20230709:
             * 這是因為您的正則表達式中使用了否定的零寬度前瞻断言 (?<!\\*) 來排除前面有星號的情況。這個断言只會查看比對位置之前的字符，而不會跨越段落分隔符（例如換行符）去查看其他段落中的字符。
             * 
             * 如果您希望跨越段落進行比對，那麼您可以使用正則表達式中的「點」字符 . 來匹配任意字符（除了換行符之外）。例如，您可以使用 .* 來匹配任意數量的任意字符。
                但是，請注意，由於否定的零寬度前瞻断言 (?<!) 只能查看固定長度的字符，因此您不能直接在 (?<!) 中使用 .* 來跨越段落進行比對。如果您希望跨越段落進行比對，那麼您可能需要使用其他方法來實現。
                例如，您可以先將輸入字串中的所有段落分隔符（例如換行符）替換為某個特殊字符（例如 #），然後再使用正則表達式進行比對。這樣，您就可以使用 .* 來匹配任意數量的任意字符，並且能夠跨越原來的段落分隔符進行比對了。
             */


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

        internal static bool HasEditedWithPunctuationMarks(ref string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            if (text.Length > 1000)
            {
                Regex regex = new Regex(@"\，|\。|\？|\！|\〈|\〉|\《|\》|\：|\『|\』|\「|\」|\􏿽|、|●|□|■|·|\*\*|\{\{\{|\}\}\}|\||〇|◯|　}}|\*　");
                Match match = regex.Match(text);
                return match.Success;
            }
            else
            {
                return (text.Contains("，") || text.Contains("。") || text.Contains("：") || text.Contains("􏿽")
                    || text.Contains("！") || text.Contains("？") || text.Contains("《") || text.Contains("〈")
                    || text.Contains("『") || text.Contains("』")
                    || text.Contains("「") || text.Contains("」")
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
            if (x.IndexOf("\" ") > -1 && x.IndexOf("=\"") > -1 && x.IndexOf(" />") > -1) return;

            #region narrow2WidePunctuationMarks 半形轉全形。置換為全形符號。
            //20230806Bing大菩薩：
            //string pattern = "[\\u0021-\\u002F\\u003A-\\u0040\\u005B-\\u0060\\u007B-\\u007E]";
            ////string pattern = "[,.;]";
            //MatchEvaluator evaluator = match => ((char)(match.Value[0] + 65248)).ToString();
            ////x= Regex.Replace(x, pattern, evaluator);
            //x = Regex.Replace(x, pattern, evaluator).Replace("．", "。");
            #endregion
            
            string[] replaceDChar = { "〇","!","！！","'", ",", ";", ":", "．", "?", "：：","：\r\n：", "《《", "》》", "〈〈", "〉〉",
                "。}}<p>。}}","。}}。}}", "。}}}。<p>", "}}}。<p>", "。}}。<p>", "}}。<p>",".<p>","·<p>" ,"<p>。<p>","<p>。","􏿽。<p>","　。<p>"
                ,"。。", "，，", "@" 
                //,"}}<p>\r\n{{"//像《札迻》就有此種格式，不能取代掉！ https://ctext.org/library.pl?if=en&file=36575&page=12&editwiki=800245#editor
                ,"\r\n。<p>","\r\n〗","\r\n。}}","\r\n："
                ,"！。<p>","？。<p>","+<p>","<p>+","：。<p>","。\r\n。"
                ,"：。","\r\n，","\r\n。","\r\n、","\r\n？","\r\n」","「\r\n" ,"{{\r\n" ,"\r\n}}"
                ,"􏿽？","􏿽。","，〉","。〉","〈、","！，","〈，"//自動標點結果的訂正
                ,"，。"
            };

            string[] replaceChar = { "◯","！","！","、", "，", "；", "：", "·", "？", "：","：\r\n", "《", "》", "〈", "〉",
                "。}}","。}}", "。}}}<p>", "。}}}<p>", "。}}<p>", "。}}<p>","。<p>","。<p>","<p>","<p>","　","　"
                , "。", "，", "●" 
                //,"}}\r\n{{"//像《札迻》就有此種格式，不能取代掉！ https://ctext.org/library.pl?if=en&file=36575&page=12&editwiki=800245#editor
                ,"\r\n","〗\r\n","。}}\r\n","：\r\n"
                ,"！<p>","？<p>","<p>","<p>","：<p>","。\r\n"
                ,"。","，\r\n","。\r\n","、\r\n","？\r\n","」\r\n","\r\n「" ,"\r\n{{", "}}\r\n"
                ,"？􏿽","。􏿽","〉，","〉。","、〈","！","，〈"//自動標點結果的訂正
                ,"。"
            };
            if (replaceDChar.Count() != replaceChar.Count()) Debugger.Break();//請檢查！！
            for (int i = 0; i < replaceChar.Count(); i++)
            {
                //if (replaceDChar[i] == "{{\r\n") Debugger.Break();
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
            string title = Browser.Title_Linkbox?.GetAttribute("textContent");
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
                //標題必非版心！
                if (line.IndexOf("*") > -1
                    || line.IndexOf("孫守真按") > -1//有按語則不用汰除 
                    || (i > 1 && i < lines.Length - 2))//版心不可能在中間啦
                    continue;


                //double similarity = Fuzz.Ratio(title, lines[i].Replace("卷","")) / 100.0;
                string pattern = "[" + Regex.Escape("卷上下卄一二三四五六七八九十卅卌<p>") + "]";
                line = Regex.Replace(line, pattern, "");
                double similarity = Fuzz.Ratio(title, line) / 100.0;
                if (similarity >= threshold && lines[i - 1] != "|")//前一段若為「|」通常是卷末題目
                {
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
        //    string title = Browser.Title_Linkbox?.GetAttribute("textContent");
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
            string title = Browser.Title_Linkbox?.GetAttribute("textContent");
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
            string pattern = " ([^ ]*) ";
            string replacement = "{{$1}}";
            string result = Regex.Replace(text, pattern, replacement);
            if (result != text) { text = result; }
        }

        /// <summary>
        /// 20240808（臺灣父親節）creedit with Copilot大菩薩：《古籍酷》自動標點完成的文本重新插入分段符號
        /// </summary>
        /// <param name="originalText"></param>
        /// <param name="punctuatedText"></param>
        /// <returns></returns>
        public static string RestoreParagraphs(ref string originalText, ref string punctuatedText)
        {

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

            // Function to find the adjusted position in punctuatedText
            int FindAdjustedPosition(string text, string original, int pos, string before, string after)
            {
                int offset1 = 0;
                int adjustedPos = pos;
                //while (adjustedPos < text.Length)
                while ((adjustedPos + (before.Length + offset1)) < text.Length)
                {
                    // Process the 'before' part
                    string subText = text.Substring(adjustedPos - (before.Length + offset1), before.Length + offset1);
                    string subTextWithoutPunctuation = RemovePunctuation(subText);
                    while (subTextWithoutPunctuation.Length < before.Length)
                    {
                        if ((adjustedPos + (before.Length + offset1)) < text.Length)
                        {
                            offset1++;
                            subText = text.Substring(adjustedPos - (before.Length + offset1), before.Length + offset1);
                            subTextWithoutPunctuation = RemovePunctuation(subText);
                        }
                        else
                        {
                            Debugger.Break();
                            break;
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
                                if (afterSubTextWithoutPunctuation != after) Debugger.Break();
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
            //清除標題符號以利分段符號之比對搜尋
            originalText = RemovePunctuation(originalText);
            //清除縮排即凸排格式標記，即將分段符號前後的空格「　」均予清除
            originalText = Regex.Replace(originalText, $@"\s*{Environment.NewLine}+\s*", Environment.NewLine);
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
                int adjustedPos = FindAdjustedPosition(punctuatedText, originalText, tp.Item1 + offset, tp.Item2, tp.Item3);
                //int adjustedPos = FindAdjustedPosition(punctuatedText, originalText, pos + offset, before, after);
                //因為在子函式方法中，若沒有找到時會將標點符號清除再與原未標點之文本作比對，若原文本已略有標點，則會干擾比對結果，不如兩造一律均清除，則簡單有效 20240808
                if (adjustedPos != -1)
                {
                    punctuatedText = punctuatedText.Insert(adjustedPos, newLine);
                    offset += newLine.Length;
                }
            }

            return punctuatedText;
        }
    }
}
