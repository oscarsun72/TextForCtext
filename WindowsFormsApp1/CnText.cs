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
using System.Reflection;
using System.Web.UI.WebControls;

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
            return ("{{}}<p>" + Form1.punctuationsNum).IndexOf(text) > -1;
        }

        /// <summary>
        /// 自動加上書名號篇名號
        /// </summary>
        /// <param name="clpTxt">剪貼簿中的文字--需要加上書名號篇名號的文本。預防大文本，故以傳址（pass by reference）方式</param>
        /// <returns>傳址回傳clpTxt被標點後的結果</returns>
        internal static ref string BooksPunctuation(ref string clpTxt)
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
                switch (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c))
                {
                    case UnicodeCategory.UppercaseLetter:
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
                        break;
                    case UnicodeCategory.LetterNumber:
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
                        break;
                    case UnicodeCategory.ClosePunctuation:
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
            if (text.Length == 0) return false;
            if (text.Length > 1000)
            {
                Regex regex = new Regex(@"\，|\。|\〈|\〉|\《|\》|\：|\『|\』|\「|\」|\􏿽");
                Match match = regex.Match(text);
                return match.Success;
            }
            else
            {
                return (text.Contains("，") || text.Contains("。") || text.Contains("：") || text.Contains("􏿽")
                    || text.Contains("《") || text.Contains("〈")
                    || text.Contains("『") || text.Contains("』")
                    || text.Contains("「") || text.Contains("」")
                    || text.Contains("》") || text.Contains("〉"));
            }

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
            //20230806Bing大菩薩：
            //string pattern = "[\\u0021-\\u002F\\u003A-\\u0040\\u005B-\\u0060\\u007B-\\u007E]";
            ////string pattern = "[,.;]";
            //MatchEvaluator evaluator = match => ((char)(match.Value[0] + 65248)).ToString();
            ////x= Regex.Replace(x, pattern, evaluator);
            //x = Regex.Replace(x, pattern, evaluator).Replace("．", "。");

            string[] replaceDChar = { "'", ",", ";", ":", "．", "?", "：：", "《《", "》》", "〈〈", "〉〉", "。}}。}}", "。。", "，，", "@" };
            string[] replaceChar = { "、", "，", "；", "：", "·", "？", "：", "《《", "》", "〈", "〉", "。}}", "。", "，", "●" };
            foreach (var item in replaceDChar)
            {
                if (x.IndexOf(item) > -1)
                {
                    //if (MessageBox.Show("含半形標點，是否取代為全形？", "", MessageBoxButtons.OKCancel,
                    //    MessageBoxIcon.Error) == DialogResult.OK)
                    //{//直接將半形標點符號轉成全形
                    for (int i = 0; i < replaceChar.Length; i++)
                    {
                        x = x.Replace(replaceDChar[i], replaceChar[i]);
                    }
                    //}
                    break;
                }
            }
            //置換中文文本中的英文句號（小數點）
            CnText.PeriodsReplace_ChinesePunctuationMarks(ref x);
        }

    }
}
