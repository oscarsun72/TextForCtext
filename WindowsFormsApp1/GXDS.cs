﻿using OpenQA.Selenium.DevTools.V85.ApplicationCache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WindowsFormsApp1;
using System.Windows.Forms;

namespace TextForCtext
{
    /// <summary>
    /// 《國學大師》網站的操作
    /// </summary>    
    public class GXDS : IDisposable
    {

        /// <summary>
        /// = Environment.NewLine;
        /// </summary>
        string newLine = Environment.NewLine;
        /// <summary>
        /// = Environment.NewLine.Length;
        /// </summary>
        int newLineLen = Environment.NewLine.Length;
        /// <summary>
        /// =Form1.cs（主表單）；
        /// </summary>
        Form1 frm, frmRef;
        /// <summary>

        /// <summary>
        /// =Form1.cs（主表單）； 不能被Dispose
        /// </summary>
        Form1 frm1 = Application.OpenForms[0] as Form1;
        /// <summary>

        /// 建構器（constructor）
        /// </summary>
        /// <param name="frm"></param>
        internal GXDS(Form1 form)
        {
            frm = form;
        }
        /// 建構器（constructor）
        /// </summary>
        /// <param name="frm"></param>
        internal GXDS(ref Form1 form)
        {
            frmRef = form;
        }

        /*20230125 Implement Dispose Method： chatGPT大菩薩新年吉祥 我想實作我以下類別的 Dispose 方法 請問這樣對嗎？
         * 您實作的 Dispose 方法顯示將 frm1 設為 null，這樣是可以的。如果您在使用 GXDS 類別時，需要手動釋放 frm1 物件的資源，可以在 Dispose 方法中使用 frm1.Dispose() 釋放資源。
         * 不過,在使用上,您還需要特別注意呼叫 Dispose 的時機, 確保呼叫在正確的生命週期中.
         * 孫守真
         * 所以我在 Form1.cs 中用這樣呼叫這個GSDS類別的方法 就可以了嗎
         * using (GXDS gxds = new GXDS(this)) { gxds.StandardizeSKQSContext(ref xClpBd); }
         * 對，您在 Form1.cs 中使用 using 語句呼叫 GXDS 類別的方法，可以在使用完 GXDS 類別後自動調用 Dispose 方法，釋放 GXDS 類別物件所佔用的資源。
         * 值得注意的是,這樣呼叫使用了 GXDS 的類別,在區塊結束後就會自動釋放GXDS物件,對應的資源也會被釋放,而不需要程式碼手動設置.
         * 這是一種簡化程式碼,避免程式碼錯誤的好寫法.
         * 孫守真
         * 那請問我在GXDS建構器（constructor）傳入了 Form1 引數 在GXDS 物件 Dispose時也會讓我原來主表單的Form1內的一切資源失效嗎？還是只有傳給 GXDS建構器（constructor）作為其物件欄位frm1值的引數所用到資源才會失效。因為我想那引數是傳值（pass by value）型的 應該是複製一個 Form1配置的資源過去，而不是影響我原來主表單Form1本身的資源與操作吧？
         * 在 GXDS 類別的建構子中，傳入的 Form1 引數會被複製給 GXDS 類別物件的 frm1 欄位，並不會影響到原來主表單 Form1 本身的資源與操作。
         * 當 GXDS 類別物件的 Dispose 方法被呼叫時，只會釋放 GXDS 類別物件所佔用的資源，不會影響到傳入的 Form1 引數。這是因為傳入的引數是傳值（pass by value）型的，複製了一個 Form1 配置的資源過去，而不是影響原來主表單 Form1 本身的資源與操作。
         * 另外，在您的程式碼中，Dispose 方法並沒有釋放任何資源(只有將frm1 設為 null)，因此不會有影響。
         * 孫守真
         * 那我GSDS 的 Dispose 改成這樣就完美了嗎？
         * public void Dispose()
         * {            frm1.Dispose();
         * 
         * frm1 = null;            
         * //throw new NotImplementedException();
         *}
         *這樣在 GXDS 類別物件的 Dispose 方法中呼叫 Form1 類別物件的 Dispose 方法，將會釋放傳入的 Form1 引數所佔用的資源。
         *但是,在這種情況下，您需要特別注意，當 Form1 引數物件被釋放後，對象及其相關資源都不能再使用。
         *如果 Form1 是主要的物件，並且還有其他物件還在使用它的資源，那麼釋放 Form1 的資源可能會發生錯誤。
         *因此,在您的情況下,這樣的寫法不適用，因為您未知是否還有其他物件正在使用 Form1 類別物件。
         *建議您可以在 Form1 中的 Dispose 方法中,自行管理相關的資源釋放,避免重複釋放或是未被釋放的問題.
         *孫守真
         *可是您之前不是說我這裡的 frm1 Form1類別的物件只是複製副本過去的，如果我在這個複製過來 的 frm1上呼叫它類別的這個物件方法 Dispose 應該只會釋放這個 GXDS 類別物件所用到的 frm1 物件的資源吧？ 並不會影響其他的 Form1類別物件才是啊
         *您說得對，在 GXDS 類別物件的 Dispose 方法中呼叫 frm1.Dispose() 將只會釋放 GXDS 類別物件所用到的 frm1 物件的資源，不會影響到其他的 Form1 類別物件。
         *因為這個 frm1 物件是一個副本，它只被這個 GXDS 類別物件所使用，不會影響到原來主表單 Form1 本身的資源與操作。
         *但是,基於使用上的考量,在這種情況下,仍建議您不要在 GXDS 類別物件的 Dispose 方法中呼叫 Form1 類別物件的 Dispose 方法，因為如果您的程式碼存在錯誤，可能會導致 Form1 類別物件被釋放多次，造成錯誤。
         *而且,在這種情況下,這樣呼叫是沒有必要的,因為在 GXDS 類別物件的 Dispose 方法中只有將 frm1 設為 null, 不會有影響.
         *孫守真
         *那我這個 GXDS 的 Dispose 方法 到底該怎麼寫才比較好呢？我原來也只有設 frm1= null 而已（詳前面所貼上的程式碼）
         *在 GXDS 類別的 Dispose 方法中，您可以僅將 frm1 設為 null，而不呼叫 frm1.Dispose()。
         *因為 frm1 是一個複製過去的副本，它只被 GXDS 類別物件所使用，不會影響到原來主表單 Form1 本身的資源與操作，所以不需要呼叫它的 Dispose 方法。
         *這樣的寫法符合 Dispose 方法的設計原則，並且也能避免因為重複釋放或是未被釋放的問題。
         *當然,如果您希望在 GXDS 類別物件的 Dispose 方法中釋放 Form1 類別物件的資源，那麼您需要保證 GXDS 類別物件是唯一使用 Form1 類別物件的物件,並且在程式碼中能夠確保正確的控制好呼叫的順序.
         *通常來說,在這種情況下,最好是在 Form1 物件的生命週期結束時,由 Form1 物件自己管理相關的資源釋放.
         *孫守真
         *那這樣寫呢？
         *public void Dispose()
         *{
         *if (frm1.IsDisposed)
         *frm1.Dispose();
         *frm1 = null;
         * //throw new NotImplementedException();        }
         * 在 GXDS 類別物件的 Dispose 方法中，先判斷 frm1 是否已經被釋放，如果已經被釋放則不進行釋放，否則呼叫 frm1.Dispose()，這樣的寫法可以避免因為重複釋放的錯誤發生。
         * 然而這種寫法也有可能會有問題，因為您還是無法確保 GXDS 類別物件是唯一使用 Form1 類別物件的物件，其他物件可能也會對 Form1 類別物件
         * 孫守真
         * 瞭解。我剛才這行「 if (frm1.IsDisposed)」應當寫作「 if (!frm1.IsDisposed)」才是，少了個「!」我只是怕在不用到 GXDS 的物件時，它所用到的 Form1類別物件的資源沒有被釋放而已。如今我大概只要寫原來的樣子，只有 frm1=null 這樣就好？ 是嗎
         *  public void Dispose()
         *  {
         *  //if (!frm1.IsDisposed)
         *  //    frm1.Dispose();
         *  frm1 = null;
         *  //throw new NotImplementedException();
         *  }
         *  沒錯,在 GXDS 類別物件的 Dispose 方法中，只需要將 frm1 設為 null 即可。
         *  因為這個 frm1 是一個複製過去的副本，它只被 GXDS 類別物件所使用，不會影響到原來主表單 Form1 本身的資源與操作。
         *  當 GXDS 類別物件被釋放時，這個副本就不再被使用，所以只需要將它設為 null 即可。
         *  至於資源的釋放,在 Form1 類別物件的生命週期結束時,由 Form1 類別物件自行管理相關的資源釋放.
         *  如果您的程式碼控制的正確，並且您使用了 using 或 try-finally 等方法來管理資源，那麼您不必擔心資源未被釋放的問題
         *  孫守真
         *  那若我這個GXDS 類別以後再擴充，有了更多的欄位或變數，是否我在未來都應該在這Dispose方法中，加入像「frm1=null」這樣的陳述？此外還有什麼該注意的麼？比如容易忽略掉類別裡的什麼成員（member）被有沒處置（dispose）？是不是只要釋放欄位與變數值就好？
         *  ，「using」是有點會了，以後再和老師您學「 try-finally 等方法」
         *  是的，當您的 GXDS 類別有更多的欄位或變數時，您應該在 Dispose 方法中為每個欄位或變數設置 null 值，以釋放他們所使用的資源。
         *  還有一些其他需要注意的事項：
         *  1.檢查類別中的所有資源是否都已被釋放
         *  2.檢查類別中是否有任何非同步操作正在進行，並確保已經等待完成
         *  3.檢查類別中是否有任何其他物件正在使用這個類別的資源，並確保它們已經結束使用
         *  *  使用「using」或「try-finally」來管理資源是很好的寫法，因為它們可以確保資源在使用完後被正確釋放。
         *  「using」語句是 C# 中的語法結構，它可以自動釋放在其區塊中定義的資源。
         *  而「try-finally」語句則是程式碼中的模式，用於在程式碼區塊中確保在異常發生時，資源仍然被正確釋放。
         *  如果您對這兩種方法有疑問，歡迎再詢問我。
         *  孫守真
         *  try-finally 如我想的，果然是用在例外情形發生時能保證資源被釋放，而不是程式當掉了卻還沒被釋放而浪費了作業系統的資源，有如我之前在學C++時所謂的記憶體外洩或洩漏。有機會再好好向老師菩薩您請教 感恩感恩　讚歎讚歎　南無阿彌陀佛
         *  是的，您說得很準確。
         *  try-finally 是一種常用的程式碼模式，用於在程式碼區塊中確保在異常發生時，資源仍然被正確釋放。
         *  在 try 區塊中的程式碼可能會拋出異常，而 finally 區塊中的程式碼則總是會在 try 區塊中的程式碼執行完後執行，無論是否發生異常。
         *  這種方式可以確保資源在程式碼執行完後被正確釋放，避免浪費*/
        public void Dispose()
        {
            //if (!frm1.IsDisposed)
            //    frm1.Dispose();
            frm = null;
            newLineLen = 0;
            newLine = null;
            //throw new NotImplementedException();
        }

        /// <summary>
        /// 《四庫全書》文本標準化
        /// </summary>
        /// <param name="xForStandardize">要處理的文本</param>
        /// <returns></returns>
        internal void StandardizeSKQSContext(ref string xForStandardize)
        {
            if (string.IsNullOrEmpty(xForStandardize)) return;

            xForStandardize = xForStandardize.Replace(" /\v", newLine)//.Replace(" /\v", newLine + newLine)
                                    .Replace(" /\v\v", newLine).Replace("\v", newLine)
                                    .Replace(" /", "");
            //這要做標題判斷，不能取代掉.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine)
            skqsTitleMark_WithoutPreSpaceWrappedinBlankLines(ref xForStandardize);

            xForStandardize = "*欽定四庫全書<p>〖文淵|閣寶〗<p>" + xForStandardize.Substring(xForStandardize.IndexOf("欽定《四庫全書》") + "欽定《四庫全書》".Length);


        }
        /// <summary>
        /// 標題/篇名標識。只要一行/段之前後均只是空行的話（跨行以上的不處理，手動自行處理）
        /// </summary>
        /// <param name="xForMark">要處理的文本</param>
        void skqsTitleMark_WithoutPreSpaceWrappedinBlankLines(ref string xForMark)
        {
            int s = xForMark.IndexOf(Environment.NewLine);
            while (s > -1)
            {//只要一行/段之前後均只是空行的話（跨行以上的不處理，手動自行處理）
                string tx = Form1.getLineTxt(xForMark, s, out int sLineStart, out int sLineLen);
                //如果後一行/段是空行段
                if (xForMark.Length >= s + newLineLen * 2 && xForMark.Substring(s + newLineLen, newLineLen) == newLine)
                {
                    if (sLineStart - newLineLen > 0)
                    {
                        //string x = "如圖\r\n\r\n中宫\r\n\r\n漢〈天文志〉{{史《天官書";
                        //如果前一行/段亦是空行段
                        int sPre = xForMark.LastIndexOf(newLine, sLineStart - newLineLen);
                        if (sPre > -1)
                        {
                            if (xForMark.Substring(sPre + newLineLen, newLineLen) == newLine)
                            {
                                xForMark = xForMark.Substring(0, sLineStart - newLineLen * 2) + "<p>" + newLine
                                    + "*" +//標題文字，清除書名篇名號等 20230125 creedit with chatGPT大菩薩：Regular Expression Replace Multi：
                                    Regex.Replace(xForMark.Substring(sLineStart, sLineLen), "[《》〈〉]", "").Replace("　", "􏿽")
                                    + "<p>" +
                                    xForMark.Substring(sLineStart + sLineLen + newLineLen);//清除標題後的空行/段
                                //xForMark = xForMark.Substring(0, sLineStart - newLineLen * 2) + "<p>" + newLine +
                                //    "*" + xForMark.Substring(sLineStart, sLineLen) + "<p>" +
                                //    xForMark.Substring(sLineStart + sLineLen + newLineLen);//清除標題後的空行/段
                            }
                        }
                        s = xForMark.IndexOf(newLine, sLineLen + newLineLen);
                    }
                }
                //else
                //設定下一個處理點
                s = xForMark.IndexOf(newLine, sLineStart + sLineLen + newLineLen);
            }
            //清除跨行標題誤標的後面部分: xForMark.Replace("<p><p>\r\n*", "\r\n");
            xForMark = xForMark.Replace("<p><p>" + newLine + "*", newLine);
            //Regex.Replace(xForMark, "<p><p>" + newLine + "*", newLine);//此無效，應該是要pattern才行，不能指定字串
            //清除末尾冗餘
            if (xForMark.Substring(xForMark.Length - 1, 1) == "\r") xForMark = xForMark.Substring(0, xForMark.Length - 1);
        }


        /// <summary>
        /// 更正《國學大師》《四庫全書》本小註文標識錯誤：
        /// Alt + - （字母區的減號）: 如果被選取的是「􏿽」則與下一個「{{」對調，反之亦然。（針對《國學大師》《四庫全書》文本小注文誤標而開發）
        /// </summary>
        internal void correctBlankAndUppercurlybrackets(ref TextBox txb)
        {
            const string Uppercurlybrackets = "{{";
            const string Lowercurlybrackets = "}}";
            string x = txb.Text,  wrng = txb.SelectedText, blank = "􏿽"; int s = txb.SelectionStart, sN=-1;

            switch (wrng)
            {
                case "􏿽":
                    blank = Uppercurlybrackets;
                    sN = x.IndexOf(Uppercurlybrackets, s);
                    break;
                //case Uppercurlybrackets:
                //    blank = "􏿽";
                //    break;
                case Lowercurlybrackets:
                    blank = "􏿽";
                    sN = x.IndexOf(blank, s);
                    break;
            }
            if (("􏿽"+Lowercurlybrackets).IndexOf(wrng) > -1)
            {
                if (x.Substring(sN, blank.Length) == blank)
                {
                    txb.SelectedText = blank;
                    txb.Select(sN, blank.Length);
                    txb.SelectedText = wrng;
                }
            }
        }

    }

}

