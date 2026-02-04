using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V127.PWA;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using WebSocketSharp;
using WindowsFormsApp1;
using static System.Net.Mime.MediaTypeNames;
using static TextForCtext.Browser;
using static TextForCtext.CTP;
using static TextForCtext.XML.ScanPageAdjuster;
using static WindowsFormsApp1.Form1;
using selenium = OpenQA.Selenium;

namespace TextForCtext
{
    /// <summary>
    /// 關於 CTP（《中國哲學書電子化計劃》）頁面的操作元件集中在這裡
    /// </summary>
    internal static class CTP
    {



        #region 上傳新資料 Submit a new text 頁面元件（二者重複，與修改原典後設資料不同者才收錄在此區域）        

        /// <summary>
        /// 上傳新資料 Submit a new text 頁面中的「檢查」（Analyze）按鈕
        /// </summary>
        internal static IWebElement Analyze_button_submitNewText
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#content > form > input[type=submit]:nth-child(10)", 3);
        }
        /// <summary>
        /// 上傳新資料 Submit a new text 頁面中的「上傳資料」（Create resource）按鈕
        /// </summary>
        internal static IWebElement CreateResource_button_submitNewText
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#createresource", 3);
        }
        /// <summary>
        /// 上傳新資料 Submit a new text 頁面中的上傳內容、大文字方塊控制項
        /// </summary>
        internal static IWebElement XMLData_textarea_submitNewText
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#data", 3);
        }

        #endregion

        #region 修改原典後設資料頁面諸元件
        ///<summary> 
        ///修改原典後設資料(&amp;action=edit)、新增資源(&amp;action=new)、新增section(&amp;action=newchapter) 、編輯section(editchapter) 等頁面的「著作名稱:」或「標題／篇名:」文字方塊控制項
        /// </summary>
        internal static IWebElement Title_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#title"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「作者:	」文字方塊控制項
        /// </summary>
        internal static IWebElement Author_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#author"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「成書年代:」選項清單控制項
        /// </summary>
        internal static IWebElement Dynasty_selectListBox { get => WaitFindWebElementBySelector_ToBeClickable("#dynasty"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「版本:	」→「其他:」文字方塊控制項
        /// </summary>
        internal static IWebElement OtherEdition_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#otheredition"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「其它名稱:	」文字方塊控制項
        /// </summary>
        internal static IWebElement Alias_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#alias"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「標籤:」文字方塊控制項
        /// </summary>
        internal static IWebElement Tags_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#tags"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「編撰年份:」文字方塊控制項
        /// </summary>
        internal static IWebElement CompositionDate_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#compositiondate"); }
        ///<summary> 
        ///修改原典後設資料等頁面的「修改摘要:」文字方塊控制項
        /// </summary>
        internal static IWebElement Description_textBox { get => WaitFindWebElementBySelector_ToBeClickable("#description"); }


        ///<summary>
        ///修改原典後設資料等頁面的「保存」按鈕
        /// </summary>
        internal static IWebElement Submit_button_EditTextMetadata { get => WaitFindWebElementBySelector_ToBeClickable("#content > form:nth-child(8) > table > tbody > tr:nth-child(10) > td:nth-child(2) > input[type=submit]:nth-child(2)"); }

        #endregion

        #region DTO        
        /// <summary>
        /// 定義 DTO
        /// 使用 DTO 傳遞資料的範例
        /// 複製後設資料的實作方式建議: https://copilot.microsoft.com/shares/D2EEe7DpwTg7QMnmpZEfs 
        /// </summary>
        public class TextMetadataDto
        {
            public string Title { get; set; }
            public string Author { get; set; }
            public string Dynasty { get; set; }
            public string OtherEdition { get; set; }
            public string Alias { get; set; }
            public string Tags { get; set; }
            public string CompositionDate { get; set; }
            public string Description { get; set; }
        }

        // 封裝讀取方法
        public static TextMetadataDto ReadFromEditPage()
        {
            return new TextMetadataDto
            {
                Title = Title_textBox.GetDomProperty("value"),
                Author = Author_textBox.GetDomProperty("value"),
                Dynasty = Dynasty_selectListBox.GetDomProperty("value"),
                OtherEdition = OtherEdition_textBox.GetDomProperty("value"),
                Alias = Alias_textBox.GetDomProperty("value"),
                Tags = Tags_textBox.GetDomProperty("value"),
                CompositionDate = CompositionDate_textBox.GetDomProperty("value"),
                Description = Description_textBox.GetDomProperty("value")
            };
        }

        // 封裝寫入方法
        public static void WriteToNewTextPage(TextMetadataDto dto)
        {
            SetIWebElementValueProperty(Title_textBox, dto.Title);
            SetIWebElementValueProperty(Author_textBox, dto.Author);
            SetIWebElementValueProperty(Dynasty_selectListBox, dto.Dynasty);
            SetIWebElementValueProperty(OtherEdition_textBox, dto.OtherEdition);
            SetIWebElement_textContent_Property(Alias_textBox, dto.Alias);
            SetIWebElement_textContent_Property(Tags_textBox, dto.Tags);
            SetIWebElementValueProperty(CompositionDate_textBox, dto.CompositionDate);
            SetIWebElementValueProperty(Description_textBox, dto.Description);
            //NewTextPage.TitleTextBox.SendKeys(dto.Title);
            //NewTextPage.AuthorTextBox.SendKeys(dto.Author);
            //NewTextPage.DynastySelect.SendKeys(dto.Dynasty);
            //NewTextPage.OtherEditionTextBox.SendKeys(dto.OtherEdition);
            //NewTextPage.AliasTextBox.SendKeys(dto.Alias);
            //NewTextPage.TagsTextBox.SendKeys(dto.Tags);
            //NewTextPage.CompositionDateTextBox.SendKeys(dto.CompositionDate);
            //NewTextPage.DescriptionTextBox.SendKeys(dto.Description);
        }


        #endregion

        /// <summary>
        /// 書籍首頁各冊列表中的第一冊(file)超連結控制項
        /// </summary>
        internal static IWebElement FirstFileItem_td_linkbox_BookHomepage
        {
            get => WaitFindWebElementBySelector_ToBeClickable(FirstFileCSSSelector);
        }

        /// <summary>
        /// 到書籍首頁            
        /// </summary>
        /// <returns>失敗則為false</returns>
        internal static bool GotoBookHomepage(string url = "")
        {
            url = url.IsNullOrEmpty() ? driver.Url : url;
            //https://copilot.microsoft.com/shares/DDxkayShTPyJVicM7v8vv                
            var info = CtextPageClassifier.ParseUrl(url);
            //Console.WriteLine($"頁面類型：{info.PageType}");
            //Console.WriteLine($"ResId：{info.ResId}");
            //Console.WriteLine($"FileId：{info.FileId}");
            //Console.WriteLine($"ChapterId：{info.ChapterId}");
            //Console.WriteLine($"PageNumber：{info.PageNumber}");            
            switch (info.PageType)
            {
                case CtextPageType.Unknown:
                    return false;
                case CtextPageType.LibraryResource:
                    if (url != driver.Url) driver.Url = url;
                    return true;
                case CtextPageType.LibraryFile:
                    if (Title_BookName_Linkbox_ImageTextCorrespondencePage?.JsClick() == false) return false;
                    break;
                case CtextPageType.LibraryFileEditWiki:
                    if (Title_BookName_Linkbox_ImageTextCorrespondencePage?.JsClick() == false) return false;
                    break;
                case CtextPageType.WikiResource:
                    if (Img_divWikibox1?.JsClick() == false) return false;
                    break;
                case CtextPageType.WikiChapter:
                    if (Title_Chapter_BookName?.JsClick() == false) return false;
                    if (Img_divWikibox1?.JsClick() == false) return false;
                    break;
                case CtextPageType.WikiEditChapter:
                    if (Title_Chapter_BookName?.JsClick() == false) return false;
                    if (Img_divWikibox1?.JsClick() == false) return false; break;
                case CtextPageType.EditTextMetadata:
                    if (Title_linkbox_EditTextMetadata_BookName?.JsClick() == false) return false;
                    if (Img_divWikibox1?.JsClick() == false) return false; break;
                case CtextPageType.SubmitNewText:
                    string baseTextUrl = OtherEdition_textBox?.GetDomProperty("value");
                    if (!baseTextUrl.IsNullOrEmpty() && CtextPageClassifier.ParseUrl(baseTextUrl).PageType == CtextPageType.LibraryResource)
                    {
                        driver.Url = baseTextUrl;
                        return true;
                    }
                    else
                        return false;
                default:
                    return false;
            }
            return true;
        }
        /// <summary>
        /// 到書籍第一冊(file)之首頁
        /// </summary>
        /// <returns>失敗則為false</returns>
        internal static bool GotoFirstFile(string url = "")
        {
            url = url.IsNullOrEmpty() ? driver.Url : url;
            //https://copilot.microsoft.com/shares/DDxkayShTPyJVicM7v8vv
            var info = CtextPageClassifier.ParseUrl(url);
            if (info == null) return false;

            //先到書籍首頁
            if (!GotoBookHomepage(url)) return false;
            //再點擊第一冊，直接到其首頁
            return (bool)FirstFileItem_td_linkbox_BookHomepage?.JsClick();
        }


        /// <summary>
        /// 取得[查看歷史](History)顯示差異（Compare）結果頁面中的表格控制項（元件） 20260115
        /// </summary>        
        internal static IWebElement Table_action_diff_History_Compare
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#content > table:nth-child(6)");
            /*
             * copy selector
                #content > table:nth-child(6)
               copy Xpath
                //*[@id="content"]/table[2]
                /html/body/div[2]/table[2]
             */
        }
        /// <summary>
        /// 取得[簡單修改模式](quick edit)超連結控制項（元件）
        /// </summary>        
        internal static IWebElement QuickeditLinkIWebElement
        {
            get
            {
                if (Browser.driver == null) Browser.driver = Browser.DriverNew();
                IWebElement iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#quickedit > a", 5);
                if (iwe != null)
                {
                    string iweText = iwe.GetAttribute("text");
                    if (iweText != "簡單修改模式" && iweText != "Quick edit")
                    {
                        //#quickedit > a:nth-child(1)                        
                        //# quickedit > a:nth-child(2)
                        if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個超連結控制項？"
                            + Environment.NewLine + Environment.NewLine + iweText))
                            return iwe;
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("沒有找到正確的「簡單修改模式Quick edit」超連結控制項，請檢查！");
                    }
                }
                return iwe;
            }
        }

        /// <summary>
        /// 取得[簡單修改模式](quick edit)下的Save changes按鈕（元件）
        /// </summary>
        /// <returns>傳回[簡單修改模式](quick edit)下的Save changes按鈕控制項</returns>
        internal static IWebElement SavechangesButton
        {
            get
            {
                IWebElement iwe = null;
                //if (Browser.driver == null) Browser.driver = Browser.DriverNew();
                if (!IsDriverInvalid)
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#savechangesbutton");
                }
                return iwe;
            }
        }


        /// <summary>
        /// 取得CTP圖文對照網頁中的「書名」（title）超連結控制項，含 href 屬性者
        /// </summary>
        internal static IWebElement Title_Linkbox_Link_ImageTextCorrespondencePage
        {
            get
            {
                const string selector = "#content > div:nth-child(3) > span:nth-child(2) > a";//32位元免安裝版Chrome瀏覽器               
                IWebElement iwe;
                //if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                //{
                iwe = Browser.WaitFindWebElementBySelector_ToBeClickable(selector);
                return iwe;
                //}
                //else
                //    return null;                
            }
        }
        /// <summary>
        /// 取得圖文對照頁面中的「Add transcription」控制項（元件）
        /// </summary>
        internal static IWebElement AddTranscription_Linkbox
        {
            get
            {
                IWebElement iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > a");
                if (iwe == null) return null;
                if (iwe.GetAttribute("text") == "Add transcription")
                    return iwe;
                else
                    return null;
            }
        }



        //https://copilot.microsoft.com/shares/voPeEUL6rctZiefyeRkMg 20260116 Visual Studio XML 註解顯示 HTML 標籤

        /// <summary>
        /// 取得CTP圖文對照頁面中的「書名」（title）控制項
        /// 如 &lt;span itemprop="title"&gt;純常子枝語&lt;/span&gt;
        /// </summary>
        internal static IWebElement Title_BookName_Linkbox_ImageTextCorrespondencePage
        {//如「<span itemprop="title">純常子枝語</span>」
            get
            {
                const string selector = "#content > div:nth-child(3) > span:nth-child(2) > a > span";//（32、64位元）免安裝版Chrome瀏覽器
                const string selector1 = "#content > div:nth-child(5) > span:nth-child(2) > a > span"; //（64位元）安裝版Chrome瀏覽器
                IWebElement iwe;
                if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable(selector);
                reCheck:
                    if (iwe != null)
                    {
                        string tx = iwe.GetAttribute("outerHTML");
                        if (!tx.StartsWith("<span itemprop=\"title\">"))
                        {
                            iwe = Browser.WaitFindWebElementBySelector_ToBeClickable(selector1);//64位元安裝版Chrome瀏覽器
                            if (iwe != null)
                            {
                                tx = iwe.GetAttribute("outerHTML");
                                if (!tx.StartsWith("<span itemprop=\"title\">"))
                                {
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "Title_Linkbox div:nth-child(5) !tx.StartsWith(\"<span itemprop=\\\"title\\\">\"))");
                                    return null;
                                }
                            }
                            else
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "Title_Linkbox div:nth-child(5)=null");
                                return null;
                            }
                        }
                        else
                            return iwe;
                    }
                    else
                    {
                        iwe = Browser.WaitFindWebElementBySelector_ToBeClickable(selector1);
                        if (iwe != null)
                            goto reCheck;
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "Title_Linkbox");
                            return null;
                        }
                    }
                }
                else
                    return null;
                return iwe;
            }
        }
        //有多工（多個TextForCtext同時運行時，此技不通）
        //internal static int Please_confirm_that_you_are_human_Page_Occurrence_Counter = 0;
        /// <summary>
        /// 在碰到認證碼時要隨文附上的敬告訊息 20260204 今凌晨半夜天未明突然被「熱」醒而有之靈感 亦天意也。既反應無效（去信、討論區、臉書、x.com），且讓來者與站長德龍賢友菩薩明悉如是苛政之猛厲虐於校正者之瀝血也，感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主　哈利路亞　而今天天母家這又是大晴天，一掃連日之陰霾寒流，如是才一吐這一年多來被這樣惡整的窩囊氣，柳暗花明，真是天意啊！這樣做來才真是舒坦的。有人性多了。否則一個人兀自做著做著、辛勤地耕耘者像深宮怨婦一般，盡做一些吃力不討好的，何必？之前站長德龍賢友菩薩還會熱情回應，現在就完全地不理睬人了，真的不知何故何忍…… 人在做，天在看，末日審判時再見吧。看看上帝是怎麼判，又賢友此時到底在忙什麼，可以如此不理睬來無償奉獻者（當不止末學一人而已矣）。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主　哈利路亞
        /// </summary>
        internal static string Please_confirm_that_you_are_human_Page_Occurrence_Interrupt_Message = string.Empty;
        /// <summary>
        /// 取得「Please confirm that you are human! 敬請輸入認證圖案」元件
        /// </summary>
        internal static IWebElement Please_confirm_that_you_are_human_Page
        {
            get
            {
                IWebElement iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > font");
                if (iwe == null) return null;
                //if (iwe.GetAttribute("textContent") == "Please confirm that you are human! 敬請輸入認證圖案")
                if (iwe.GetDomProperty("textContent") == "Please confirm that you are human! 敬請輸入認證圖案")
                    return iwe;
                else
                    return null;
            }
        }
        /// <summary>
        /// 在文字版瀏覽頁面的圖文對照小圖標按鈕元件
        /// 即點擊後會進入圖文對照頁面的按鈕元件
        /// </summary>
        internal static IWebElement GraphicMatchingPagesLink_Button_TextVersionViewPage
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#p2 > td:nth-child(1) > div > a.sprite-photo > div", 3);
        }



        /// <summary>
        /// 取得CTP圖文對照頁面中的「編輯」（Edit）控制項
        /// 「編輯」連結元件
        /// </summary>
        internal static IWebElement Edit_Linkbox_ImageTextComparisonPage
        {
            get
            {
                IWebElement iwe;
                //if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                if (IsValidUrl_ImageTextComparisonPage(ActiveForm1.TextBox3Text))
                {
                    //會因位置而移動，如：Add to 學海蠡測 Add to 思舊錄 [文字版] [編輯] [簡單修改模式] [編輯指南] https://ctext.org/library.pl?if=gb&file=194081&page=75&editwiki=5083072#editor
                    //故得逐一比對，目前應該只會有2種情形，當然也可能會不止如此
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                reCheck:
                    if (iwe != null)
                    {
                        string tx = iwe.GetAttribute("text");
                        if (tx != "編輯" && tx != "Edit")
                        {
                            iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(4)");
                            if (iwe != null)
                                tx = iwe.GetAttribute("text");
                            if (tx != "編輯" && tx != "Edit")
                            {
                                iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(3)");
                                if (iwe != null)
                                    tx = iwe.GetAttribute("text");
                                if (tx != "編輯" && tx != "Edit")
                                {
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「編輯（Edit）」超連結控制項，請檢查！");
                                    return null;
                                }
                                else
                                    return iwe;
                            }
                            else
                                return iwe;
                        }
                        //Edit_Linkbox = waitFindWebElementByName_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)", WebDriverWaitTimeSpan);
                    }
                    else
                    {
                        iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(9) > div:nth-child(2) > a:nth-child(2)");
                        if (iwe != null) goto reCheck;
                    }
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「參考上下頁」（ check the adjacent pages）控制項
        /// </summary>
        internal static IWebElement CheckAdjacentPages_Linkbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#editor > a:nth-child(13)");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「上一頁」的編輯文字方塊
        /// </summary>
        internal static IWebElement CheckAdjacentPages_DataPrev
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#dataprev");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「下一頁」(Next page:)的編輯文字方塊
        /// </summary>
        internal static IWebElement CheckAdjacentPages_DataNext
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#datanext");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 圖文對照頁面中的下一頁（箭頭狀）控制項（元件）
        /// </summary>
        internal static IWebElement NextPageBtn_ArrowShapedButton_WikiVersionScannedEditionComparisonPages//https://ctext.org/wiki.pl?if=en
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(3) > div:nth-child(5) > a > div", 5);
        }
        /// <summary>
        /// 取得CTP網頁中的「顯示頁碼，可輸入頁碼的」（page）控制項
        /// </summary>
        internal static IWebElement PageNum_textbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                bool checkNamePorp()
                {
                    return iwe?.GetDomAttribute("name") == "page";
                    //return iwe?.GetAttribute("name") == "page";
                }
                if (IsValidUrl_ImageTextComparisonPage(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(3) > form > input[type=text]:nth-child(3)");
                    if (iwe == null)
                    {
                        iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(5) > form");
                    }
                    if (!checkNamePorp()) return null;
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得現前/現用的頁碼-CTP網頁中的「顯示頁碼，可輸入頁碼的」（page）控制項的值
        /// </summary>
        internal static string CurrentPageNum_textbox_Value
        {
            get
            {
                IWebElement iwe = PageNum_textbox;
                if (iwe == null) return string.Empty;
                return iwe.GetDomProperty("value");
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「顯示頁碼資訊的條幅」（page）控制項（以取得該書的末頁）
        /// </summary>
        internal static IWebElement Div_generic_IncludePathAndEndPageNum
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (IsValidUrl_ImageTextComparisonPage(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(3)");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得某書書頁碼的上限值
        /// 若出錯則傳回0
        /// </summary>
        /// <returns></returns>
        internal static int PageUBound
        {
            get
            {
                IWebElement iwe = Div_generic_IncludePathAndEndPageNum;
                if (iwe == null) return 0;
                string input = iwe.GetAttribute("textContent");//"線上圖書館 -> 松煙小錄 -> 松煙小錄三  /117 ";
                return CnText.ExtractNumberAfterSlash(input);
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「文本框」（文字框）（圖文對照的文框）控制項
        /// </summary>
        internal static IWebElement Div_generic_TextBoxFrame
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (IsValidUrl_ImageTextComparisonPage(ActiveForm1.TextBox3Text))
                {
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(1)");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「書圖框」（圖文對照的圖框.svg）控制項
        /// </summary>
        internal static IWebElement Svg_image_PageImageFrame
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                //if (IsValidUrl_ImageTextComparisonPage(ActiveForm1.TextBox3Text))
                //{
                iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#canvas > svg");
                //}
                //else
                //return null;
                if (iwe == null)
                    iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#previmg");
                if (iwe == null)
                {
                    if (!IsDriverInvalid)
                        iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#previmg");
                    //iwe = Browser.driver.FindElement(By.XPath("/html/body/div[2]/div[3]/img"));
                    else
                    {
                        try
                        {
                            if (!Browser.driver.WindowHandles.Contains(LastValidWindow))
                                LastValidWindow = Browser.driver.WindowHandles.Last();
                            Browser.driver.SwitchTo().Window(LastValidWindow);

                            iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#previmg");
                        }
                        catch (Exception ex)
                        {
                            if (IsDriverInvalid) RestartChromedriver();
                            else
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            //throw;
                        }
                    }
                }
                return iwe;
            }
        }
        /// <summary>
        /// 放大書圖以便檢視
        /// </summary>
        /// <returns>有按下來放大則傳回true</returns>
        internal static bool EnlargeSvgImageSize(bool chromeSetFocus = true)
        {
            IWebElement iwe = Svg_image_PageImageFrame;
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = Svg_image_PageImageFrame;
                if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break;
            }
            if (iwe?.Size.Width <= 500) // Leo AI 20260118 iwe?.GetAttribute("width");
            {
                if (chromeSetFocus)
                    ChromeSetFocus();
                iwe.Click();//這裡不能用 iwe.JsClick()，會出錯;
                return true;
            }
            return false;
        }
        /// <summary>
        /// 還原放大的書圖
        /// </summary>
        /// <returns>有按下以還原則為true</returns>
        internal static bool RestoreSvgImageSize(bool chromeSetFocus = true)
        {
            IWebElement iwe = Svg_image_PageImageFrame;
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = Svg_image_PageImageFrame;
                if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break;
            }
            if (iwe?.Size.Width > 500) // Leo AI 20260118 iwe?.GetAttribute("width");
            {
                if (chromeSetFocus)
                    ChromeSetFocus();
                iwe.Click();//這裡不能用 iwe.JsClick()，會出錯;
                return true;
            }
            return false;
        }

        /// <summary>
        /// 自動全選[Quick edit]的內容，方便有時候須用剪下貼上者
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool SelectAllQuickedit_data_textboxContent()
        {
            OpenQA.Selenium.IWebElement ie = Quickedit_data_textbox;//br.QuickeditIWebElement;
            if (ie != null)
            {
                ie.SendKeys(OpenQA.Selenium.Keys.Control + "a");
                return true;
            }
            return false;
        }


        /// <summary>
        /// 取得如欽定四庫全書的版本連結元件；若失敗則回傳null
        /// </summary>
        internal static IWebElement Version_LinkBox_ImageTextCorrespondencePage
        {
            get
            {
                IWebElement version_LinkBox;
                //if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                //{
                //  /html/body/div[2]/div[5]/div[3]/a[1]
                version_LinkBox = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(8) > div:nth-child(3) > a:nth-child(1)");
                if (version_LinkBox == null)
                    // /html/body/div[2]/div[5]/div/a[1] /html/body/div[2]/div[5]/div/a[1]
                    version_LinkBox = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(8) > div > a:nth-child(1)");

                //}
                //else
                //version_LinkBox = null;
                return version_LinkBox;
            }
        }
        /// <summary>
        /// 調整quick edit 文字方塊大小以便檢視 20260123
        /// </summary>
        internal static void AdjustQuickEditDataTextBoxSizetoWatch()
        {
            // 調整 UI 大小以便檢視
            //Quickedit_data_textbox.JsResize("1000px", "800px");
            selenium.IWebElement iwe = Quickedit_data_textbox;
            iwe.JsResize((iwe.Size.Width + 100).ToString() + "px", (iwe.Size.Height + 50).ToString() + "px");
        }

        /// <summary>
        /// 取得[簡單修改模式]的文字方塊（編輯區的文字方塊）；若失敗則回傳null        
        /// Get the textbox of [Quick edit] 
        /// </summary>
        internal static IWebElement Quickedit_data_textbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                {
                    _quickedit_data_textbox = WaitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan);
                }
                else
                    _quickedit_data_textbox = null;
                return _quickedit_data_textbox;
            }
            private set { _quickedit_data_textbox = value; }
        }
        /// <summary>
        /// 取得[編輯]的文字方塊（編輯區的文字方塊）；若失敗則回傳null 20240929 于52生日
        /// Get the textbox of [edit] 
        /// </summary>
        internal static IWebElement Textarea_data_Edit_textbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#data", WebDriverWaitTimeSpan);
            //{
            //if (Browser.driver.Url.IndexOf("&action=editchapter") > -1)
            //{
            //return Browser.WaitFindWebElementBySelector_ToBeClickable("#data", WebDriverWaitTimeSpan);
            //}
            //    else
            //        return null;
            //}
        }
        /// <summary>
        /// 完整編輯頁面下「修改摘要」欄位控件
        /// </summary>
        internal static IWebElement Description_Edit_textbox
        {
            get { return WaitFindWebElementBySelector_ToBeClickable("#description", WebDriverWaitTimeSpan); }
        }
        /// <summary>
        /// 取得[編輯]的文字；若失敗則回傳空字串 20260111
        /// </summary>
        internal static string Textarea_data_Edit_textboxTxt
        {
            get
            {
                IWebElement ie = Textarea_data_Edit_textbox;
                if (ie != null)
                    return ie.GetDomProperty("value");
                else
                    return string.Empty;
            }
        }
        /// <summary>
        /// 取得[編輯]的[保存編輯]按鈕；若失敗則回傳null 20260110
        /// Get the submit of [commit] 
        /// </summary>
        internal static IWebElement Commit
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#commit", WebDriverWaitTimeSpan);
            //get
            //{

            //    if (Browser.driver.Url.IndexOf("&action=editchapter") > -1)
            //    {
            //        return Browser.WaitFindWebElementBySelector_ToBeClickable("#commit", WebDriverWaitTimeSpan);
            //    }
            //    else
            //        return null;
            //}
        }
        /// <summary>
        /// 取得 修改原典後設資料(Edit text metadata)頁面的[標題]（書名）元件；若失敗則回傳null 20260116
        /// </summary>
        internal static IWebElement Title_linkbox_EditTextMetadata_BookName
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(4) > a:nth-child(2)", WebDriverWaitTimeSpan);
            //get
            //{
            //    if (Browser.driver.Url.IndexOf("chapter") > -1)
            //    {
            //        return Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(4) > a:nth-child(2)", WebDriverWaitTimeSpan);
            //    }
            //    else
            //        return null;
            //}
        }
        /// <summary>
        /// 取得文字版chapter章節瀏覽與編輯頁面的[標題]（書名）超連結元件；若失敗則回傳null 20260110
        /// </summary>
        internal static IWebElement Title_Chapter_BookName
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(4) > span:nth-child(2) > a > span", WebDriverWaitTimeSpan);
            //get
            //{
            //    if (Browser.driver.Url.IndexOf("chapter") > -1)
            //    {
            //        return Browser.WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(4) > span:nth-child(2) > a > span", WebDriverWaitTimeSpan);
            //    }
            //    else
            //        return null;
            //}
        }

        /// <summary>
        /// 儲存[簡單修改模式]的文字方塊
        /// </summary>
        private static IWebElement _quickedit_data_textbox = null;

        internal static string Quickedit_data_textbox_Txt = "";

        /// <summary>
        /// 取得[簡單修改模式]的文字；若失敗則回傳空字串
        /// 原來取該元件的「value」Property就可以了20240913
        /// </summary>
        internal static string Quickedit_data_textboxTxt
        {
            get
            {
                //if (!IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text)) return string.Empty;
                IWebElement ie = Quickedit_data_textbox;
                if (ie != null)
                {
                    //20240913 原來取該元件的「value、textContent……」等 Property 就可以了！
                    //.Text屬性會清除起首的全形空格！！20240313
                    //if (quickedit_data_textboxTxt != Quickedit_data_textbox.Text) quickedit_data_textboxTxt = quickedit_data_textbox.Text;                    
                    //string quickedit_data_textbox_Txt = CopyQuickedit_data_textboxText();                    
                    //if (quickedit_data_textboxTxt != quickedit_data_textbox_Txt) quickedit_data_textboxTxt = quickedit_data_textbox_Txt;
                    //return quickedit_data_textboxTxt;
                    Quickedit_data_textbox_Txt = ie.GetDomProperty("value");
                    return Quickedit_data_textbox_Txt;
                }
                else
                    return string.Empty;
            }
        }
        /// <summary>
        /// 設定Quickedit_data_textbox的value屬性值  20240913
        /// creedit_with_Copilot大菩薩：C# Selenium 屬性設定方法： https://sl.bing.net/jv1AQReen36
        /// </summary>
        /// <param name="txt">要設定的值</param>
        /// <returns>若失敗則傳回false</returns>
        internal static bool SetQuickedit_data_textboxTxt(string txt)
        {
            if (!IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text)) return false;
            IWebElement ie = Quickedit_data_textbox;
            if (ie != null)
            {
                if (SetIWebElementValueProperty(ie, txt))
                    return true;
                else
                    return false;
            }
            else
                return false;

        }

        /// <summary>
        /// 當Quickedit_data_textbox的內容是以全形空格開頭的會被清除，類似Trim的功能，故須用複製文本的方式取得正確的值
        /// 解決Selenium在[簡單修改模式]文字方塊內容若以全形空格為開頭的，會被截去的方案 20230829        /// 
        /// 原來取該元件的「value」Property就可以了 20240913
        /// </summary>
        /// <returns>回傳所複製的Quickedit_data_textbox文本</returns>
        internal static string CopyQuickedit_data_textboxText()
        {
            IWebElement ie = Quickedit_data_textbox;
            if (ie != null)
            {
                //[簡單修改模式]方塊若不存在
                if (Browser.WaitFindWebElementBySelector_ToBeClickable("#data") == null)
                {
                    //[簡單修改模式]超連結
                    if (Browser.WaitFindWebElementBySelector_ToBeClickable("#quickedit > a") != null)
                    {
                        //按下[簡單修改模式]超連結
                        Browser.WaitFindWebElementBySelector_ToBeClickable("#quickedit > a").Click();
                    }
                    else
                        return string.Empty;
                    _quickedit_data_textbox = Browser.WaitFindWebElementBySelector_ToBeClickable("#data");
                    ie = Quickedit_data_textbox;
                }
                if (ie.Text != string.Empty)
                {
                    //ie.SendKeys(OpenQA.Selenium.Keys.Control + "a");//會移動視窗焦點到文字方塊 ie（Quickedit_data_textbox）中
                    SelectAllQuickedit_data_textboxContent();
                    ie.SendKeys(OpenQA.Selenium.Keys.Control + "c");
                    WindowsScrolltoTop();
                    //Clipboard.SetText(ie.Text);//.Text屬性會清除前首的全形空格，不適用！！20240313
                    DateTime dt = DateTime.Now;
                    while (!Form1.IsClipBoardAvailable_Text())
                        if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break;
                }
                else
                    Clipboard.Clear();
                return Clipboard.GetText();
            }
            else
            {
                Clipboard.Clear();
                return string.Empty;
            }
        }

        internal static IWebElement Full_text_search_textbox_searchressingle
        {
            get
            {
                if (Browser.driver == null) return null;
                IWebElement full_text_search_textbox_searchressingle = null;
                try
                {
                    if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                    {
                        //< input type = "hidden" name = "searchressingle" id = "searchressingle" value = "wiki:728745" style = "width: 80px;" >
                        //full_text_search_textbox_searchressingle = Browser.driver.FindElement(By.Name("searchressingle"));
                        full_text_search_textbox_searchressingle = Browser.driver.FindElement(By.CssSelector("#searchressingle"));
                        if (full_text_search_textbox_searchressingle == null)
                        {
                            WebDriverWait wait = new WebDriverWait(Browser.driver, TimeSpan.FromSeconds(2));
                            full_text_search_textbox_searchressingle =
                                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Name("searchressingle")));
                        }
                    }
                    return full_text_search_textbox_searchressingle;
                }
                catch (Exception)
                {
                    return null;
                }

            }
        }

        /// <summary>
        /// 取出如以下這個字串中的「tr:nth-child(2)」這個個部分的「2」這個數值以供計算，如加1後變成3，而轉置回這個Selector的字串中
        /// #content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a
        /// 20250305 GitHub　Copilot大菩薩
        /// </summary>
        /// <param name="selector"></param>
        /// <returns></returns>
        internal static string IncrementNthChild(string selector)
        {
            var match = Regex.Match(selector, @"tr:nth-child\((\d+)\)");
            if (match.Success)
            {
                int number = int.Parse(match.Groups[1].Value);
                number++;
                return Regex.Replace(selector, @"tr:nth-child\(\d+\)", $"tr:nth-child({number})");
            }
            return selector;
        }
        /// <summary>
        /// 常數，作為圖書館(Library)書籍首頁各冊列表中第一冊超連結元件的Css selector值的存儲
        /// 書籍首頁如此頁：https://ctext.org/library.pl?if=gb&amp;res=414
        /// </summary>
        internal const string FirstFileCSSSelector = "#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a";
        /// <summary>
        /// 取得目前file(冊，原用chapter）的CSS Selector值，不存在則傳回null
        /// </summary>
        internal static string FileCSSSelector
        {
            set
            {
                if (!WindowHandles.TryGetValue("FileSelector", out _))
                    WindowHandles.Add("FileSelector", value);
                else
                    WindowHandles["FileSelector"] = value;
            }
            get
            {
                if (!WindowHandles.TryGetValue("FileSelector", out string fileSelector))
                    return null;
                else
                    return fileSelector;
            }
        }
        /// <summary>
        /// 本冊的 CSS selector 值
        /// </summary>
        internal static string CurrentFileSelector
        {
            get => FileCSSSelector;
        }
        /// <summary>
        /// 取得下一個file（冊，原作chapter）的Selector值，不存在則傳回null
        /// </summary>
        internal static string NextFileSelector
        {
            get
            {
                if (FileCSSSelector == null)
                    return null;

                string selector = FileCSSSelector;//"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a";
                                                  //if (!WindowHandles.TryGetValue("ChapterSelecto", out string chapterSelector))
                                                  //    WindowHandles.Add("ChapterSelector ",);
                                                  //else
                                                  //{
                string newSelector = IncrementNthChild(selector);
                //Console.WriteLine(newSelector); // 輸出: #content > div:nth-child(6) > table > tbody > tr:nth-child(3) > td:nth-child(1) > a
                FileCSSSelector = newSelector;
                return newSelector;
                //}
            }
        }
        /// <summary>
        /// 前往下一冊（file）
        /// </summary>
        /// <returns></returns>
        internal static bool GotoNextFile()
        {
            //先到書籍首頁
            ////若是在圖文對照頁面：如：https://ctext.org/library.pl?if=gb&file=76626&page=7 
            ////點擊圖文對照頁面中「書名(title)」連結控制項
            //if (Title_BookName_Linkbox_ImageTextCorrespondencePage != null) Title_BookName_Linkbox_ImageTextCorrespondencePage.JsClick();
            ////或者在文字版整部書頁面，如：https://ctext.org/wiki.pl?if=gb&res=53207
            //else if (Img_divWikibox != null) Img_divWikibox1.JsClick();

            ////如果仍找不到各冊連結元件
            //if (Files_Table == null) return false;
            if (!GotoBookHomepage()) return false;

            if (FileCSSSelector == null) { MessageBoxShowOKCancelExclamationDefaultDesktopOnly("請先在textBox2中以「fn」+n(數字)指定目前是第n冊。或直接貼入CssSelector值"); return false; }
            //點擊下一個file的連結
            string nextfileSelector = NextFileSelector;
            if (nextfileSelector.IsNullOrEmpty()) return false;
            //點擊本書首頁的冊(file,網址中有）連結
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable(nextfileSelector, 5);
            if (iwe == null)
            {
                //MessageBoxShowOKExclamationDefaultDesktopOnly("本書業畢，沒有下一冊了！");//20260201 改寫在呼叫端
                //if (Form1.InstanceForm1.FastMode) Form1.InstanceForm1.FastModeSwitcher();
                return false;
            }
            return iwe.JsClick();
        }
        /// <summary>
        /// 書籍首頁的分冊列表表格
        /// </summary>
        internal static IWebElement Files_Table
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(6) > table");
        }
        /// <summary>
        /// 文字版整部書頁面的書圖縮圖元件外框
        /// </summary>
        internal static IWebElement Img_divWikibox
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#content > div.wikibox > table > tbody > tr:nth-child(3) > td");
        }
        /// <summary>
        /// 文字版整部書頁面的書圖縮圖元件1（左側縮圖）
        /// </summary>
        internal static IWebElement Img_divWikibox1
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#content > div.wikibox > table > tbody > tr:nth-child(3) > td > a:nth-child(1) > img");
        }
        /// <summary>
        /// 文字版整部書頁面的書圖縮圖元件2（右側縮圖）
        /// </summary>
        internal static IWebElement Img_divWikibox2
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#content > div.wikibox > table > tbody > tr:nth-child(3) > td > a:nth-child(2) > img");
        }

        /// <summary>
        /// 取得目前章節chapter（篇）的序號控制項（元件） 20260111
        /// </summary>
        internal static IWebElement Sequence_Edit_Chapter
        {
            get => Browser.WaitFindWebElementBySelector_ToBeClickable("#sequence");
        }
        /// <summary>
        /// 取得目前章節chapter（篇）的序號值
        /// </summary>
        internal static string Sequence_Edit_Chapter_Value
        {
            get => Sequence_Edit_Chapter?.GetDomProperty("value");
        }
        /// <summary>
        /// 在全書單位中，取得在目前章節chapter（篇）序號基礎上加1後的下一個章節chapter（篇）的序號控制項（元件）
        /// 全書單位，即 https://ctext.org/wiki.pl?if=gb&res= 這樣的網址頁面（res=後的數值即為書文本的編號，即「URN: ctp:wb」後綴的數值） 20260111
        /// </summary>
        /// <param name="currentSequence"></param>
        /// <returns></returns>
        internal static IWebElement NextSequence_Linkbox(string currentSequence)
        {
            if (!int.TryParse(currentSequence, result: out int result)) return null;
            string nextSequence = (++result).ToString(), nextSeq = nextSequence;
            string nextSequenceTextContent = WaitFindWebElementBySelector_ToBeClickable("#content > div.ctext > span:nth-child(" + nextSequence + ")").GetDomProperty("textContent");
            if (nextSequenceTextContent.IndexOf(".") == -1) return null;
            nextSequence = nextSequenceTextContent.Substring(0, nextSequenceTextContent.IndexOf("."));
            while (nextSequence != nextSeq)
            {
                nextSequenceTextContent = WaitFindWebElementBySelector_ToBeClickable("#content > div.ctext > span:nth-child(" + (++result).ToString() + ")").GetDomProperty("textContent");
                nextSequence = nextSequenceTextContent.Substring(0, nextSequenceTextContent.IndexOf("."));
            }
            return WaitFindWebElementBySelector_ToBeClickable("#content > div.ctext > span:nth-child(" + result.ToString() + ") > a");
        }
        /// <summary>
        /// 前往下一個章節chapter（篇）準備編輯頁面
        /// </summary>
        /// <param name="currentSeq">目前章節序號</param>
        /// <returns>失敗則為false</returns>
        internal static bool GotoNextSection_SequenceChapterPage(string currentSeq)
        {//Section 見:Please begin each line that should be a section title with a single asterisk * (e.g. "*學而")  https://ctext.org/wiki.pl?if=en&action=new
            //string currentSeq = Sequence_Edit_Chapter_Value;
            IWebElement iwe = NextSequence_Linkbox(currentSeq);
            if (iwe != null) { iwe.JsClick(); return true; }
            return false;
        }
        /// <summary>
        /// 取得CTP文字版篇章單元網頁（即URN: ctp:ws…… 所在頁面）中的「修改」（Edit）控制項
        /// </summary>
        internal static IWebElement Edit_linkbox
        {//如此網頁： https://ctext.org/wiki.pl?if=gb&chapter=872411 上的[修改][Edit]元件
            get =>
            Browser.WaitFindWebElementBySelector_ToBeClickable("#content > h2 > span > a:nth-child(2)");
        }
        /// <summary>
        /// 取得CTP文字版篇章單元網頁（即URN: ctp:ws…… 所在頁面）中的「修改」（Edit）控制項
        /// 如此網頁： https://ctext.org/wiki.pl?if=en&res=53207&amp;action=newchapter 上的[標題][書名]元件
        /// </summary>
        internal static IWebElement Head_linkbox_wikiitemtitle_newchapter
        {
            get =>
            Browser.WaitFindWebElementBySelector_ToBeClickable("#content > h2");
        }
        /// <summary>
        /// 取得目前章節file(原作chapter）（冊）的序號，以供Selector字串參照使用
        /// </summary>
        internal static string CurrentFileNum_Selector
        {
            get
            {
                string selector = FileCSSSelector;//"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a";
                var match = Regex.Match(selector, @"tr:nth-child\((\d+)\)");
                //if (match.Success)
                return match.Groups[1].Value;


                //string pattern = @"tr:nth-child\((\d+)\)";//@"nth-child\((\d+)\)";
                //MatchCollection matches = Regex.Matches(selector, pattern);
                //foreach (Match match in matches)
                //{
                //    // 提取括號中的數值
                //    int value = int.Parse(match.Groups[1].Value);
                //    //Console.WriteLine($"nth-child 的值: {value}");
                //    // 在這裡進行你的後續計算
                //    retun
                //}
            }
        }

        /// <summary>
        /// 指定要清除quick edit box 內容的引數值 "\t"（其實是有由tab鍵所按下的值，或其他亂碼字），此與 Word VBA 中國哲學書電子化計劃.新頁面 為速新章節單位的配置有關 碼詳：https://github.com/oscarsun72/TextForCtext/blob/f75b5da5a5e6eca69baaae0b98ed2d6c286a3aab/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L32
        /// </summary>
        internal static readonly string chkClearQuickedit_data_textboxTxtStr = " ";
        internal static bool confirm_that_you_are_human = false;
        /// <summary>
        /// 在Chrome瀏覽器的文字框(ctext.org 的 Quick edit ）中輸入文字,creedit//若 xIuput= " "則清除而不輸入
        /// </summary>
        /// <param name="Browser.driver">chromedriver</param>
        /// <param name="xInput">要貼入的文本</param>
        /// <param name="url">要貼入的網頁網址</param>
        /// <returns>執行成功則回傳true</returns>
        internal static bool 在Chrome瀏覽器的Quick_edit文字框中輸入文字(ChromeDriver driver, string xInput, string url)
        {
            #region 檢查網址
            Uri uri = new Uri(url);
            if (uri.Authority != "ctext.org") { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("想要輸入的網址並不是CTP網址"); return false; }
            if (Browser.driver.Url == "about:blank")
            {
                Browser.driver.Close();
                bool found = false; string urlDriver;
                Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles.Last());
                for (int i = Browser.driver.WindowHandles.Count - 1; i > -1; i--)
                {
                    urlDriver = ReplaceUrl_Box2Editor(Browser.driver.Url);
                    if (urlDriver == url || url.Contains(urlDriver))
                    {
                        Browser.driver.Url = url;
                        found = true; break;
                    }
                }
                if (!found) Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles.Last());
            }

            uri = new Uri(Browser.driver.Url);
            if (uri.Authority != "ctext.org") { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("目前 Browser.driver的網址並不是CTP網址"); return false; }

            if (url.IndexOf("edit") == -1 && Browser.driver.Url.IndexOf("edit") == -1)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("網址中不包含「edit」");
                return false;
            }


            if (url != Browser.driver.Url)
            {
                if (Browser.driver.Url.IndexOf(url.Replace("editor", "box")) == -1)
                    //if (url != Browser.driver.Url && Browser.driver.Url.IndexOf(url.Replace("editor", "box")) == -1)
                    // 使用Browser.driver導航到給定的URL
                    Browser.driver.Navigate().GoToUrl(url);
                //("https://ctext.org/library.pl?if=en&file=79166&page=85&editwiki=297821#editor");//("http://www.example.com");

                //Uri uri=new  Uri(url);

                string urlShort = url.EndsWith("#editor") ? url.Substring(0, url.IndexOf("#editor")) : url;
                if (IsValidUrl＿keyDownCtrlAdd(url) && IsValidUrl＿keyDownCtrlAdd(Browser.driver.Url) == false)
                {
                    bool found = false;
                    foreach (var item in Browser.driver.WindowHandles)
                    {
                        Browser.driver.SwitchTo().Window(item);
                        if (Browser.driver.Url.StartsWith(urlShort))
                        {
                            if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個頁面要進行輸入？") == DialogResult.OK) { found = true; break; }

                        }
                    }
                    if (!found)
                    {
                        Form1.PlaySound(Form1.SoundLike.error, true);
                        MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查textBox3內的值是否是有效的 Quick edit 編輯頁面的網址！！感恩感恩　南無阿彌陀佛");//20260127
                        return false;
                    }
                }
                else if (IsValidUrl＿keyDownCtrlAdd(url) && IsValidUrl＿keyDownCtrlAdd(Browser.driver.Url))
                {
                    if (!Browser.driver.Url.StartsWith(urlShort))
                    {
                        bool found = false;
                        foreach (var item in Browser.driver.WindowHandles)
                        {
                            Browser.driver.SwitchTo().Window(item);
                            if (Browser.driver.Url.StartsWith(urlShort))
                            {
                                if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個頁面要進行輸入？") == DialogResult.OK) { found = true; break; }

                            }
                        }
                        if (!found)
                        {
                            Form1.PlaySound(Form1.SoundLike.error, true);
                            MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查textBox3內的值是否是有效的 Quick edit 編輯頁面的網址！！感恩感恩　南無阿彌陀佛");//20260127
                            return false;
                        }
                    }
                }
                else
                {
                    MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查textBox3內的值是否是有效的 Quick edit 編輯頁面的網址！！感恩感恩　南無阿彌陀佛");//20260127
                    Debugger.Break();
                }
            }

            #endregion

            #region 查找名稱為"data"的文字框(textbox)或ID為"quickedit"的元件，須要用到元件者均不宜另跑線程。這些名稱，都由按下 F12 或 Ctrl + shift + i 開啟開發者模式中「Elements」分頁頁籤中取得
            selenium.IWebElement textbox;
            try
            {
                textbox = Browser.driver.FindElement(selenium.By.Name("data"));//("textbox"));                

            }
            catch (Exception)
            {
                selenium.IWebElement quickedit = null;
                try
                {
                    //如果沒有按下「Quick edit」就按下它以開啟
                    quickedit = Browser.driver.FindElement(selenium.By.Id("quickedit"));
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.IndexOf("no such window: target window already closed") > -1)//"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                            {
                                if (!url.EndsWith("#editor")) url = ActiveTabURL_Ctext_Edit_includingEditorStr;
                                GoToUrlandActivate(url);
                                return false;
                            }
                            //"no such element: Unable to locate element: {\"method\":\"css selector\",\"selector\":\"#quickedit\"}\n  (Session info: chrome=111.0.5563.147)"
                            else if (ex.Message.IndexOf("no such element: Unable to locate elementno") > -1)
                            {
                                GoToCurrentUserActivateTab();
                                quickedit = Browser.driver.FindElement(selenium.By.Id("quickedit"));
                            }
                            else
                            {
                                Console.WriteLine(ex.HResult + ex.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                Debugger.Break();
                            }
                            break;
                        default:
                            //cDrv.Navigate().GoToUrl(Form1.mainFromTextBox3Text ?? "https://ctext.org/account.pl?if=en");                    
                            //MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");
                            quickedit = Browser.driver.FindElement(selenium.By.Id("quickedit"));
                            //throw;
                            break;
                    }
                }
                if (quickedit == null) return false;
                quickedit.JsClick();//下面「submit.Click();」不必等網頁作出回應才執行下一步，但這裡接下來還要取元件操作，就得在同一線程中跑。感恩感恩　南無阿彌陀佛
                textbox = Browser.driver.FindElement(selenium.By.Name("data"));
                //throw;
            }
            Quickedit_data_textboxSetting(url, textbox);

            #endregion

            ////清除原來文字，準備貼上新的
            //textbox.Clear();//20240913作廢

            #region input to textbox（old : paste to textbox）
            // 在文字框中輸入文字
            //textbox.SendKeys(@xInput); //("Hello, World!");
            /*
             chatGPT ：
                "ChromeDriver only supports characters in the BMP" 這個訊息的意思是，ChromeDriver 只支援 Unicode 基本多文種平面 (BMP) 中的字元。

                Unicode 是一種國際標準，用來對各種語言的文字進行統一編碼。它包含了超過 100,000 個字元，但是只有前 65536 個字元 (也就是基本多文種平面或 BMP) 是常用的，包括大部分的西方語言和一些亞洲語言。

                ChromeDriver 是一個 Web 自動化工具，它可以自動控制 Google Chrome 瀏覽器，執行各種測試和任務。這個訊息表示，當你在使用 ChromeDriver 時，只能輸入 BMP 中的字元。如果你想要輸入其他的字元 (比如許多亞洲語言中使用的字元)，可能會遇到問題。
             */
            //檢查是否都在BMP內
            //if (isAllinBmp(xInput))
            //{
            //textbox.SendKeys(stringtoEscape_sequences_for_Unicode_character_sets(xInput));//(Keys.Control + "v");            
            //textbox.SendKeys(xInput);
            //}
            //若含BMP外的字則用系統貼上的方法
            //else//今一律用貼上省事便捷 20230102
            //{

            ////文字框取得焦點
            //textbox.Click(); //20240913取消


            //chrome取得焦點
            //Form1 f = new Form1();
            //f.appActivateByName();

            #region 測試無誤////////……此行即可清除，不知為何多此一舉
            //////////////Browser.driver.SwitchTo().Window(Browser.driver.CurrentWindowHandle); //https://stackoverflow.com/questions/23200168/how-to-bring-selenium-browser-to-the-front#_=_
            // 讓 Chrome 瀏覽器成為作用中的程式
            //Browser.driver.Manage().Window.Maximize();//creedit chatGPT
            //Browser.driver.Manage().Window.Position = new Point(0, 0);
            #endregion

            //確定要送出文本時為true
            bool submitting = false;
            //清除內容不輸入(前已有textbox.Clear();）
            if (xInput != chkClearQuickedit_data_textboxTxtStr)//" ")// "\t")//是否清除當前頁面中的內容？（其實是有由tab鍵所按下的值)
                                                               // 建立 Actions 物件
                                                               //Actions actions = new Actions(Browser.driver);//creedit
                                                               // 貼上剪貼簿中的文字
                                                               //actions.MoveToElement(textbox).Click().Perform();
                                                               //actions.SendKeys(OpenQA.Selenium.Keys.Control + "v").Build().Perform();
                                                               //actions.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert).Build().Perform();
            {
                if (Quickedit_data_textbox_Txt != xInput)
                    if (!SetQuickedit_data_textboxTxt(xInput))
                    {
                        ActiveForm1.TextBox3Text = Browser.driver.Url;
                        if (!SetQuickedit_data_textboxTxt(xInput))
                            Debugger.Break();
                        if (Quickedit_data_textboxTxt != xInput)
                            Debugger.Break();
                        else
                            submitting = true;
                        //waitFindWebElementBySelector_ToBeClickable("#savechangesbutton")?.Click();
                    }
                    else
                        submitting = true;
                //20240913 改寫：以下作廢
                /*
                //Sendkeys(textbox, xInput);
                //while (!Form1.isClipBoardAvailable_Text()) { }
                try
                {
                    Clipboard.SetText(xInput);
                }
                catch (Exception)
                {
                    //Thread.Sleep(1500);
                    //Clipboard.Clear();
                    //Clipboard.SetText("x");
                    //Form1.PlaySound(Form1.soundLike.error, true);
                    //Clipboard.SetText(xInput);
                }
                //textbox.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert);
                textbox.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);*/
            }

            //SendKeys.Send("^v{tab}~");
            #endregion
            //}
            //Task.WaitAll();
            //System.Windows.Forms.Application.DoEvents();

            //內容經過編輯才送出，否則直接翻到下一頁或停留在此頁
            if (submitting)
            {
                #region 送出


                //selm.IWebElement submit = Browser.driver.FindElement(selm.By.Id("savechangesbutton"));//("textbox"));
                selenium.IWebElement submit = WaitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
                /* creedit 我問：在C#  用selenium 控制 chrome 瀏覽器時，怎麼樣才能不必等待網頁作出回應即續編處理按下來的程式碼 。如，以下程式碼，請問，如何在按下 submit.Click(); 後不必等這個動作完成或作出回應，即能繼續執行之後的程式碼呢 感恩感恩　南無阿彌陀佛
                            chatGPT他答：你可以將 submit.Click(); 放在一個 Task 中去執行，並立即返回。
                 */
                if (submit == null)
                {
                    submit = Browser.WaitFindWebElementBySelector_ToBeClickable("#savechangesbutton");
                    if (submit == null)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查頁面中的 Quict edit 是否可用，再按下確定繼續！");
                        //submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
                        submit = Browser.driver.FindElement(By.XPath("/html/body/div[2]/div[4]/form/div/input"));
                    }
                }
                LastValidWindow = Browser.driver.CurrentWindowHandle;
                //20250218取消多線程（多執行緒操作）
                //Task.Run(() =>//接下來不用理會，也沒有元件要操作、沒有訊息要回應，就可以給另一個線程去處理了。
                //{
                //reSubmit:
                try
                {
                    if (submit == null)
                        //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查頁面中的 Quict edit 是否可用，再按下確定繼續！");
                        //submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
                        //submit = waitFindWebElementBySelector_ToBeClickable("#savechangesbutton");
                        submit = Browser.driver.FindElement(By.XPath("/html/body/div[2]/div[4]/form/div/input"));
                    if (submit == null)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查頁面中的 Quict edit 是否可用!!!!!！");
                        if (Form1.InstanceForm1.FastMode)
                            Form1.InstanceForm1.FastModeSwitcher();
                        return false;
                    }
                    if (ActiveForm1.KeyinTextMode || int.Parse(ActiveForm1.CurrentPageNum) < 3)
                        ////submit.Click();
                        ////submit.Submit();
                        //if (submit != null)
                        //{//https://copilot.microsoft.com/shares/HMYjVyzi4Hz6WkCnCKgd6
                        //    ((IJavaScriptExecutor)Browser.driver).ExecuteScript("arguments[0].click();", submit);
                        //    return true;
                        //}
                        //else
                        //    return false;
                        return submit.JsClick(); //Click(submit);
                    else
                    {
                        if (!CheckPageNumBeforeSubmitSaveChanges(Browser.driver, submit))
                            return false;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.HelpLink + ex.Message);
                    //chatGPT：
                    // 等待網頁元素出現，最多等待 3 秒//應該不用這個，因為會貼上時，不太可能「savechangesbutton」按鈕還沒出現，除非網頁載入不完整……
                    submit = WaitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);  //Browser.driver.FindElement(selm.By.Id("savechangesbutton"));
                                                                                                                //WebDriverWait wait = new WebDriverWait(Browser.driver, TimeSpan.FromSeconds(3));
                                                                                                                ////安裝了 Selenium.WebDriver 套件，才說沒有「ExpectedConditions」，然後照Visual Studio 2022的改正建議又用NuGet 安裝了 Selenium.Suport 套件，也自動「 using OpenQA.Selenium.Support.UI;」了，末學自己還用物件瀏覽器找過了 「OpenQA.Selenium.Support.UI」，可就是沒有「ExpectedConditions」靜態類別可用，即使官方文件也說有 ： https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Support_UI_ExpectedConditions.htm 20230109 未知何故 阿彌陀佛
                                                                                                                //wait.Until(ExpectedConditions.ElementToBeClickable(submit));
                    /*chatGPT 您好，謝謝您將您的程式碼提供給我，我現在有更多的資訊可以幫助我了解您遇到的問題。按照您的程式碼，我可以確認您已經在您的項目中安裝了 Selenium.WebDriver 和 Selenium.Support NuGet 套件，並且在您的程式碼中使用了 using OpenQA.Selenium.Support.UI; 的聲明。
                     * 然而，我注意到您正在使用 .NET Framework 4.8，而非 .NET Core。根據 Selenium 文件，ExpectedConditions 類別在 .NET Framework 中只支援 .NET Core。
                     * 因此，如果您想在 .NET Framework 中使用 ExpectedConditions 類別，則您需要使用 .NET Core 來建立您的項目。如果您無法更改您的項目類型， 我現在繼續提供您有關解決方法的更多資訊。
                     * 如果您無法更改您的項目類型，則可以使用不同的方法來等待網頁元素的出現。例如，您可以使用以下方法之一：
                     * 使用 Thread.Sleep() 函式等待指定的時間。
                     * 使用 while 迴圈和 DateTime.Now 來等待網頁元素的出現。
                     * 使用 WebDriverWait 類別和 Until() 方法來等待網頁元素的出現。下面是使用第 3 種方法的示例程式碼：……
                     * 末學我回：菩薩您的解答終於、應該是對的了 是 Core 有 而Framework 不支援 才對 否則真的不知道是何緣故了。感恩感恩　讚歎讚歎　南無阿彌陀佛
                     * --然而--
                     * 不用更改 我找到了 謝謝您的回答 以後再來請教您。我剛才成功解決的是，如下所述： 在Visual Studio 2022 中的NuGet 套件不要裝「SeleniumExtras.WaitHelpers」要裝「DotNetSeleniumExtras.WaitHelpers」就可以成功安裝，再用「using SeleniumExtras.WaitHelpers;」則「wait.Until(ExpectedConditions.ElementToBeClickable(submit));」這一行程式碼就不再出錯了，也沒有紅蚯蚓了。現在我已正常編譯，……感恩感恩　讚歎讚歎　南無阿彌陀佛
                     */
                    // 在網頁元素載入完畢後，執行 Click 方法
                    if (submit != null)
                        try
                        {
                            return CheckPageNumBeforeSubmitSaveChanges(Browser.driver, submit);
                        }
                        catch (Exception)
                        {
                        }
                    else
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請手動檢查資料是否有正確送出。");
                        Browser.driver.Manage().Timeouts().PageLoad += new TimeSpan(0, 0, 3);
                        //LastValidWindow = Browser.driver.CurrentWindowHandle;
                        //openNewTabWindow();
                        try
                        {
                            Browser.driver.Navigate().GoToUrl(url);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    //throw;
                }
                #region 送出後檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl
                if (IsConfirmHumanPage())
                {
                    //Debugger.Break();
                    Form1.PlaySound(Form1.SoundLike.waiting, true);
                    //if (ActiveForm1.FastMode) ActiveForm1.FastModeSwitcher();
                    try
                    {
                        Clipboard.SetText(xInput);//複製到剪貼簿備用
                    }
                    catch (Exception)
                    {
                    }

                    //點選輸入框
                    //waitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]")?.Click();
                    IWebElement iweConfirm = Browser.WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                    if (iweConfirm == null) Browser.driver.Navigate().Back();//因非同步，若已翻到下一頁
                    iweConfirm = Browser.WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                    if (iweConfirm == null)
                    {
                        //Debugger.Break();
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("似有網頁故障！請檢查之前輸入的資料是否有正確送出。感恩感恩　南無阿彌陀佛");
                        ActiveForm1.TopMost = false;
                        Browser.driver.SwitchTo().Window(Browser.driver.CurrentWindowHandle);
                        Form1.InstanceForm1.EndUpdate();
                        Form1.InstanceForm1.TopMost = false;
                        //Application.DoEvents();//20260116●●●●●●●●●●●●●●●●●●●●●●●
                        Browser.driver.SwitchTo().Window(Browser.driver.CurrentWindowHandle);
                        BringToFront("chrome");
                        //將焦點交給Chrome瀏覽器，在以滑鼠啟動視窗時所需
                        //clickCopybutton_GjcoolFastExperience(iwe.Location);
                        IWebElement element = Browser.WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                        if (element != null && Cursor.Position != element?.Location)
                        {
                            Cursor.Position = element.Location;
                            element?.Click();
                        }
                        //Application.DoEvents();//20260116●●●●●●●●●●●●●●●●●●●●●●●
                        return false;
                    }
                    else
                        iweConfirm.Click();
                    //20251228 一律改成停下手動輸入了。因為會有誤差，有時會有兩頁以上沒正確送出。殘念。
                    //if (DialogResult.Cancel ==
                    //    Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("Please confirm that you are human! 請輸入認證圖案"
                    //    + Environment.NewLine + Environment.NewLine + "請輸入完畢後再按「確定」！程式會幫忙按下「OK」送出"
                    //    + Environment.NewLine + Environment.NewLine + "★★！最好按下「取消」以回到前數頁檢查是否有正確送出，以免白做！！", string.Empty, false))
                    //{
                    //Debugger.Break();
                    ActiveForm1.TopMost = false;
                    Browser.driver.SwitchTo().Window(Browser.driver.CurrentWindowHandle);
                    Form1.InstanceForm1.EndUpdate();
                    //Application.DoEvents();//20260116●●●●●●●●●●●●●●●●●●●●●●●
                    BringToFront("chrome");

                    //將焦點交給Chrome瀏覽器，在以滑鼠啟動視窗時所需
                    //clickCopybutton_GjcoolFastExperience(iwe.Location);
                    IWebElement iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                    if (Cursor.Position != iwe?.Location)
                        Cursor.Position = iwe.Location;
                    iwe?.Click();
                    //Application.DoEvents();//20260116●●●●●●●●●●●●●●●●●●●●●●●
                    return false;
                    //}
                    //while (true)
                    //{
                    //    Browser.WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(3) > td:nth-child(2) > input[type=submit]")?.Click();
                    //    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否重試？")) break;
                    //}
                    //Browser.driver.Navigate().Back();
                    //while (Browser.driver.Url == "https://ctext.org/wiki.pl" || Browser.driver.Url == "https://ctext.org/wiki.pl?if=en")
                    //{
                    //    Browser.driver.Navigate().Back();
                    //}
                    //if (Browser.driver.Url != url)
                    //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("網址並非 " + url + " 請檢查後再按下確定");
                    //if (Browser.driver.Url == url)
                    //{
                    //    SetQuickedit_data_textboxTxt(xInput);
                    //    goto reSubmit;
                    //}

                    //else Debugger.Break();
                }
                #endregion
                //});


                //加速連續性輸入（不必檢視貼入的文本時，很有效）
                //if (ActiveForm1.AutoPasteToCtext && Form1.FastMode)
                //if (ActiveForm1.AutoPasteToCtext && Form1.fastMode && Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName)
                if (ActiveForm1.AutoPasteToCtext && Form1.InstanceForm1.FastMode && Form1.BrowsrOPMode == Form1.BrowserOPMode.appActivateByName)
                {
                    Thread.Sleep(10);//等待 submit = waitFin……完成
                    Browser.driver.Close(); //需要重啟檢視時，只要開啟前一個被關掉的分頁頁籤即可（快速鍵時 Ctrl + Shift + t）
                }
                #endregion
            }
            else//若文本沒有改變，不用送出，則播放音效
                Form1.PlaySound(Form1.SoundLike.notify, true);
            return true;
        }
        /// <summary>
        /// 在按下
        /// </summary>
        /// <param name="Browser.driver"></param>
        /// <param name="submit_saveChanges"></param>
        /// <returns></returns>
        internal static bool CheckPageNumBeforeSubmitSaveChanges(ChromeDriver driver, IWebElement submit_saveChanges = null)
        {
            if (!IsDriverInvalid && int.Parse(ActiveForm1.CurrentPageNum) > 2)
            {
                int currentPageNum = int.Parse(Form1.InstanceForm1.CurrentPageNum);
                if (ActiveForm1.AutoPasteToCtext && currentPageNum != Form1.InstanceForm1.GetPageNumFromUrl(Browser.driver.Url) ||
                    Math.Abs(int.Parse(ActiveForm1.CurrentPageNum) - int.Parse(WindowHandles["currentPageNum"])) != 1)
                {
                    if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("頁碼不同！請轉至頁面" +
                        "頁再按下「確定」以供輸入"))
                    {
                        //submit_saveChanges?.Click();//按下 Save changes button（「保存編輯」按鈕）
                        //submit_saveChanges?.Submit();
                        //if (submit_saveChanges != null)
                        //{//https://copilot.microsoft.com/shares/HMYjVyzi4Hz6WkCnCKgd6
                        //    //((IJavaScriptExecutor)Browser.driver).ExecuteScript("arguments[0].click();", submit_saveChanges);
                        //    return true;
                        //}
                        //else
                        //    return false;
                        return submit_saveChanges.JsClick();//Click(submit_saveChanges);

                    }
                    else
                        return false;
                }
                else
                {
                    Form1.InstanceForm1.PauseEvents();
                    //submit_saveChanges?.Click();//按下 Save changes button（「保存編輯」按鈕）
                    //submit_saveChanges?.Submit();//和 Click 方法一樣 若被最大化的視窗遮住都會失效（無法確實按下），但不會出錯。 20251229
                    //以下的方法則大成功！感謝Copilot大菩薩、Gemini大菩薩 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
                    if (submit_saveChanges.JsClick())//Click(submit_saveChanges))
                    {
                        Form1.InstanceForm1.ResumeEvents();
                        return true;
                    }
                    else
                        return false;
                }
            }
            else
                return false;
        }


        static internal bool IsAllinBmp(string xChk)
        {
            char[] c = xChk.ToCharArray();
            foreach (char item in c)
            {
                if (!IsInBmp(item)) return false;
            }
            return true;
        }
        static bool IsInBmp(char c)//creedit 2023/1/1
        {
            return (0 <= c && c <= 0xFFFF) && !char.IsSurrogate(c);
        }

        /// <summary>
        /// 取得現行Ctext 編輯時前景之分頁網址。尤其是為使用者手動切換者；若找不到則傳回""（空字串）
        /// </summary>
        public static string ActiveTabURL_Ctext_Edit
        {
            get
            {
                //string url = getUrl(ControlType.Edit).Trim();
                string url = GetUrlFirst_Ctext_Edit(ControlType.Edit).Trim();
                if (url == "")
                {
                    try
                    {
                        string urlDriver = Browser.driver.Url;
                    }
                    catch (Exception)
                    {
                        if (IsValidUrl＿keyDownCtrlAdd(ActiveForm1.TextBox3Text))
                        {
                            //如：https://ctext.org/library.pl?if=en&file=38675&page=1&editwiki=573099#editor
                            string shortUrl = ActiveForm1.TextBox3Text.Substring(0, ActiveForm1.TextBox3Text.IndexOf("#editor") == -1 ? ActiveForm1.TextBox3Text.Length : ActiveForm1.TextBox3Text.IndexOf("#editor"));
                            for (int i = Browser.driver.WindowHandles.Count - 1; i > -1; i--)
                            {
                                Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles[i]);
                                if (Browser.driver.Url.StartsWith(shortUrl)) break;
                            }
                        }
                    }

                    if (!IsValidUrl_ImageTextComparisonPage(Browser.driver.Url))
                    {
                        for (int i = Browser.driver.WindowHandles.Count - 1; i > -1; i--)
                        {
                            Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles[i]);
                            if (IsValidUrl_ImageTextComparisonPage(Browser.driver.Url)) break;
                        }
                        if (!IsValidUrl_ImageTextComparisonPage(Browser.driver.Url))
                        {
                            int windowsCount;// = 0;
                            try
                            {
                                windowsCount = Browser.driver.WindowHandles.Count;
                            }
                            catch (Exception)
                            {
                                windowsCount = GetValidWindowHandles(Browser.driver).Count;
                            }
                            if (windowsCount > 1)
                            {
                                if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("目前作用中的分頁並非有效的圖文對照頁面，是否要讓程式繼續比對？"
                                        , "ActiveTabURL_Ctext_Edit\n\r\n\rgetUrlFirst_Ctext_Edit=\"\"") == DialogResult.OK)
                                    url = GetUrl(ControlType.Edit).Trim();
                            }
                        }
                        else
                        {
                            url = Browser.driver.Url;
                            ActiveForm1.TextBox3Text = url;
                        }
                    }
                    else
                    {
                        url = Browser.driver.Url;
                        ActiveForm1.TextBox3Text = url;
                    }
                }
                if (url != "") url = url.StartsWith("https://") ? url : "https://" + url;
                return url;
            }
        }
        /// <summary>
        /// 取得現行Ctext 編輯時前景之分頁網址（須含有"#editor"尾綴）。尤其是為使用者手動切換者；若找不到則傳回""（空字串）
        /// </summary>
        public static string ActiveTabURL_Ctext_Edit_includingEditorStr
        {
            get
            {
                string url = GetUrlFirst_Ctext_Edit(ControlType.Edit, true).Trim();
                if (url == "")
                {
                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("目前作用中的分頁並非有效的圖文對照頁面，是否要讓程式繼續比對？") == DialogResult.OK)
                    {

                        url = GetUrl(ControlType.Edit).Trim();
                    }
                }
                if (url != "") url = url.StartsWith("https://") ? url : "https://" + url;
                return url;
            }
        }

        /// <summary>
        /// 取得「簡單修改模式」的網址
        /// </summary>
        /// <returns>傳回「簡單修改模式」的網址</returns>
        internal static string GetQuickeditUrl()
        {
            string url = "";
            if (Browser.driver == null) Browser.driver = Browser.DriverNew();
            IWebElement ie = QuickeditLinkIWebElement;
            if (ie != null) url = ie.GetAttribute("href");
            return url;
            /*
             OpenQA.Selenium.IWebElement quickEditLink = br.
                 waitFindWebElementBySelector_ToBeClickable("#quickedit > a");
                    if (quickEditLink != null)
                    {
                        quickEditLinkUrl = quickEditLink.GetAttribute("href");
                    }
             */
        }

        /// <summary>
        /// geturl 修改後的程式碼:20230308 creedit with NotionAI大菩薩
        /// 〈get url FindAll vs FindFirst〉https://www.notion.so/get-url-FindAll-vs-FindFirst-88505499d53e4557a45fe8e844f0ee4a
        /// </summary>
        /// <param name="controlType"></param>
        /// <param name="endwithEditorStr">是否要取得末綴為「#editor」的網址</param>
        /// <returns></returns>
        static string GetUrlFirst_Ctext_Edit(ControlType controlType, bool endwithEditorStr = false)
        {
            try
            {
                //Process[] procsBrowser = GetChromeProcessInstances;
                Process[] procsBrowser = Process.GetProcessesByName(browserName);
                if (procsBrowser.Length <= 0)
                {
                    MessageBox.Show(browserName + " " + "is not the source running browser" + "\n" + "來源流覽器");
                }
                else
                {
                    foreach (Process proc in procsBrowser)
                    {
                        // the chrome process must have a window
                        if (proc.MainWindowHandle == IntPtr.Zero)
                        {
                            continue;
                        }
                        /* 20230313 AutomationElement 可以用using (){} 來寫嗎：Bing大菩薩：
                         * `AutomationElement` 不實現 `IDisposable`，因為它們沒有非托管資源。只要刪除了對它的所有引用，內存就會被釋放 - 例如將每個指針設置為空或變量超出範圍⁷。
                         * 您所謂的「指針」就是指參考型別的變數，而「變量」是指實質型別的變數吧？設置為空，就是「=null」吧 是嗎？                         * 
                         * 是的，您說得對。在 C# 中，「指針」通常指的是參考型別的變數，而「變量」可以指實質型別或參考型別的變數。將一個參考型別的變數設置為空，就是將它賦值為 null。

                            來源: 與 Bing 的交談， 2023/3/13(1) When does AutomationElement get disposed?. https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/105794db-fa1a-41d0-827a-6993973abde9/when-does-automationelement-get-disposed?forum=windowsaccessibilityandautomation 已存取 2023/3/13.
                            (2) Invoke a Control Using UI Automation - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/invoke-a-control-using-ui-automation 已存取 2023/3/13.
                            (3) Obtaining UI Automation Elements - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/obtaining-ui-automation-elements 已存取 2023/3/13.
                            (4) How can I send a right-click event to an AutomationElement using WPF's UI automation? - Stack Overflow. https://stackoverflow.com/questions/6554494/how-can-i-send-a-right-click-event-to-an-automationelement-using-wpfs-ui-automa 已存取 2023/3/13.
                            (5) Invoke a Control Using UI Automation - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/invoke-a-control-using-ui-automation 已存取 2023/3/13.
                            (6) Using objects that implement IDisposable | Microsoft Learn. https://learn.microsoft.com/en-us/dotnet/standard/garbage-collection/using-objects 已存取 2023/3/13.
                            (7) Obtaining UI Automation Elements - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/obtaining-ui-automation-elements 已存取 2023/3/13.
                            (8) AutomationElement Class (System.Windows.Automation). https://learn.microsoft.com/en-us/dotnet/api/system.windows.automation.automationelement?view=windowsdesktop-8.0 已存取 2023/3/13.
                            (9) Obtaining UI Automation Elements - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/obtaining-ui-automation-elements 已存取 2023/3/13.
                            (10) c# - selecting combobox item using ui automation - Stack Overflow. https://stackoverflow.com/questions/5814779/selecting-combobox-item-using-ui-automation 已存取 2023/3/13.
                         */
                        AutomationElement elm = AutomationElement.FromHandle(proc.MainWindowHandle);
                        AutomationElement elmUrlBar = elm.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));

                        if (elmUrlBar != null)
                        {
                            string url = ((ValuePattern)elmUrlBar.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
                            //if ((url.StartsWith("http") || url.StartsWith("ctext")))
                            if (endwithEditorStr)
                            {
                                if ((url.StartsWith("ctext.org/") || url.StartsWith("https://ctext.org/")) && url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1 && url.EndsWith("#editor"))
                                {
                                    return url;
                                }
                            }
                            else
                            {
                                if ((url.StartsWith("ctext.org/") || url.StartsWith("https://ctext.org/")) && url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1)//&& url.EndsWith("#editor"))
                                {
                                    return url;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore exception
            }
            return "";//url;

        }
        /// <summary>
        /// 取得Chrome瀏覽器現前作用中的《中國哲學書電子化計劃》頁面網址
        /// </summary>
        public static string GetChromeActiveUrl
        {
            get { return GetUrlFirst_Ctext_Edit(ControlType.Edit).Trim(); }
        }

        internal static void Quickedit_data_textboxSetting(string url, IWebElement textbox = null, IWebDriver driver = null)
        {
            if (url.IndexOf("edit") > -1)
            {
                if (textbox != null) Quickedit_data_textbox = textbox;
                else
                    try
                    {
                        Quickedit_data_textbox = WaitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan, Browser.driver);
                    }
                    catch (Exception ex)
                    {
                        //分頁視窗若關閉則忽略、繼續
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.Message);
                        return;
                        //throw;
                    }
                //Quickedit_data_textbox = waitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan, Browser.driver);
                try
                {
                    //.Text屬性會清除起首的全形空格！！20240313
                    //quickedit_data_textboxTxt = Quickedit_data_textbox == null ? "" : Quickedit_data_textbox.Text;
                    if (Quickedit_data_textbox == null)
                        Quickedit_data_textbox_Txt = "";
                    else
                    {
                        Quickedit_data_textbox_Txt = Quickedit_data_textboxTxt;
                    }

                }
                catch (Exception)
                {
                    Quickedit_data_textbox_Txt = string.Empty;
                }
            }
        }

        /// <summary>
        /// 判斷是否與目前的drive在同一本書的同一頁
        /// </summary>
        /// <param name="url">要比對的網址</param>
        /// <returns></returns>
        internal static bool IsSameBookPageWithDrive(string url)
        {
            int bookidDrive = ActiveForm1.GetBookID_fromUrl(Browser.driver?.Url ?? string.Empty), pageNumDrive = ActiveForm1.GetPageNumFromUrl(Browser.driver.Url), bookid = ActiveForm1.GetBookID_fromUrl(url), pageNum = ActiveForm1.GetPageNumFromUrl(url);
            //if (bookidDrive != bookid && pageNumDrive != pageNum)
            if (bookidDrive == bookid && pageNumDrive == pageNum)
                return true;
            else return false;
        }

        /* 以下是我先寫來問chatGPT的，依其建議改如上
        internal static string getImageUrl() {

        Browser br = new Browser(System.Windows.Forms.Application.OpenForms[0] as Form1);
        ChromeDriver Browser.driver = new ChromeDriver();
        IWebElement scancont = Browser.driver.FindElement(By.Id("scancont"));
        return scancont.GetAttribute("src");

        }
        */

        #region Ctext 三種網頁模式判斷
        /// <summary>
        /// 由Url判斷是否是[簡單修改模式][Quick edit] 
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool IsQuickEditUrl(string url)
        {
            return url != "" && url.Length >= "https://ctext.org/".Length
                && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" && url.IndexOf("edit") > -1
                    && url.LastIndexOf("#editor") > -1 && url.Substring(url.LastIndexOf("#editor")) == "#editor";
            //if (url != "" && url.Length >= "https://ctext.org/".Length
            //    && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" && url.IndexOf("edit") > -1
            //        && url.LastIndexOf("#editor") > -1 && url.Substring(url.LastIndexOf("#editor")) == "#editor") return true;
            //else
            //    return false;
        }
        /// <summary>
        /// 由Url判斷是否是[編輯]頁面（chapter=……&action=editchapter）
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool IsEditChapterUrl(string url)
        {
            return url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
                    url.LastIndexOf("&action = editchapter") > -1;
            //if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
            //        url.LastIndexOf("&action = editchapter") > -1) return true;

            //else
            //return false;
        }
        /// <summary>
        /// 由Url判斷是否是瀏覽圖文對照頁面，非[簡單修改模式][Quick edit] 
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool IsFilePageView(string url)
        {
            //if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
            //url.IndexOf("edit") == -1) return true;
            return (url != "" && url.StartsWith("https://ctext.org/library.pl?") &&
                            url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1 &&
                            url.IndexOf("edit") == -1);
            //    return true;
            //else
            //    return false;
        }
        #endregion

        /// <summary>
        /// 焦點必須在textBox1！！20240313
        /// 依選取文字取得目前URL加該選取字為該頁之關鍵字的連結。如欲在此頁中標出「𢔶」字，即為：
        /// https://ctext.org/library.pl?if=gb&file=36575&page=53#𢔶
        /// Ctrl + k
        /// </summary>
        /// <returns></returns>
        internal static string GetPageUrlKeywordLink(string w, string url, bool reMovePunctuations = false)
        {
            //if (!ActiveForm1.Controls["textBox1"].Focused) return string.Empty;
            //TextBox tb = ActiveForm1.Controls["textBox1"] as TextBox;
            //if (tb.SelectionLength == 0) return string.Empty;            
            if (url == null) return string.Empty;
            int i = url.IndexOf("&page=");
            if (i == -1) return string.Empty;

            i = url.IndexOf("&", i + "&page=".Length + 1);
            if (i > -1) //20240102 Bard大菩薩：C# 找到字串中「=53」的結束位置
                url = url.Substring(0, i);
            else
            {
                i = url.IndexOf("&page=") + "&page=".Length + 1;
                // 從起始位置開始，逐個字元比較，直到找到非數字或字串結束
                int end = i;
                while (end < url.Length && char.IsDigit(url[end]))
                {
                    end++;
                }
                url = url.Substring(0, end);
            }
            //Clipboard.SetText(w);
            //return url + "#" + HttpUtility.UrlEncode(w) ;//VBA中文編碼好像還是有問題，先用這個，並先複製一個字進剪貼簿，可以利用 Win + v 的方式檢視調用
            //以上VBA bug 已排除
            w = w.Replace(Environment.NewLine, string.Empty);
            return url + "#" + (reMovePunctuations ? CnText.RemovePunctuationsNum(w) : w);//到VBA再轉碼，以便複製此字、不必再key也。況昨晚才經Bing大菩薩、StackOverflow AI大菩薩的加持，得以成功建置此生第1個 dll檔案，供Word VBA調用。感恩感恩　讚歎讚歎　南無阿彌陀佛
        }

        /// <summary>
        /// 改變CTP圖文對照網址的 Page 參數以供翻頁
        /// 20240920 Copilot大菩薩：更改 URL 参数以翻页：https://sl.bing.net/jZV8afaj85Q
        /// </summary>
        /// <param name="url">要改變的網址</param>
        /// <param name="newPageNumber">Page參數要成的數值</param>
        /// <returns>傳回改動後的網址</returns>
        public static string ChangePageParameter(string url, int newPageNumber)
        {
            var uri = new Uri(url);
            var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
            query.Set("page", newPageNumber.ToString());
            var uriBuilder = new UriBuilder(uri)
            {
                Query = query.ToString()
            };
            return uriBuilder.ToString();
        }

        /// <summary>
        /// 置換Url中的Box 為Editor 如 https://ctext.org/library.pl?if=gb&file=34873&page=78&editwiki=164323#editor
        /// https://ctext.org/library.pl?if=gb&file=34873&page=78&editwiki=164323#box(280,86,1,0) 20241101
        /// 與 FixUrl_ImageTextComparisonPage 可互參考
        /// </summary>
        /// <param name="url"></param>
        /// <returns>傳回清除後的結果</returns>
        internal static string ReplaceUrl_Box2Editor(string url)
        {
            if (!url.StartsWith("http")) return url;
            int s = url.IndexOf("#box"); string xClear;// = null;
            if (s > -1)
            {
                xClear = url.Substring(s, url.IndexOf(")", s) - s + 1);
                url = url.Substring(0, s) + url.Substring(s + xClear.Length, url.Length - (s + xClear.Length))
                    + (url.IndexOf("#editor") == -1 ? "#editor" : string.Empty);
            }
            return url;
        }
        /// <summary>
        /// 清除Url中的雜項，如 #box(280,86,1,0)等（etc） 20241101 20250126
        /// </summary>
        /// <param name="url"></param>
        /// <returns>傳回清除後的結果</returns>
        internal static string ClearUrl_BoxEtc(string url)
        {
            if (!url.StartsWith("http")) return url;
            int s = url.IndexOf("#box");
            if (s > -1)
            {
                string xClear = url.Substring(s, url.IndexOf(")", s) - s + 1);
                url = url.Substring(0, s) + url.Substring(s + xClear.Length, url.Length - (s + xClear.Length));
            }
            return url;
        }
        /// <summary>
        /// 開啟完整編輯頁面
        /// 從DirectlyReplacingCharacters獨立出來 20260102
        /// </summary>
        /// <returns>失敗則傳回false</returns>
        internal static bool OpenEditPage()
        {
            //if (DirectlyReplacingCharactersPageWindowHandle != string.Empty)
            //{
            //    try
            //    {
            //        Browser.driver.SwitchTo().Window(DirectlyReplacingCharactersPageWindowHandle);
            //        return;
            //    }
            //    catch (Exception)
            //    {
            //        DirectlyReplacingCharactersPageWindowHandle = string.Empty;
            //    }
            //}
            //string editUrl = ActiveTabURL_Ctext_Edit_includingEditorStr;
            //if (editUrl == string.Empty) return;
            //if (!isQuickEditUrl(editUrl))
            //{
            //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("目前分頁並非簡單修改模式，無法直接取代文字，請先切換到簡單修改模式的編輯頁面再試");
            //    return;
            //}
            //Browser.driver.Navigate().GoToUrl(editUrl);
            //DirectlyReplacingCharactersPageWindowHandle = Browser.driver.CurrentWindowHandle;

            //以上是建議的

            if (Form1.BrowsrOPMode == Form1.BrowserOPMode.appActivateByName) return false;

            try
            {
                if (LastValidWindow != Browser.driver.CurrentWindowHandle)
                    Browser.driver.SwitchTo().Window(LastValidWindow);
                //else
                //    LastValidWindow = Browser.driver.CurrentWindowHandle;
            }
            catch (Exception)
            {
                Browser.driver.SwitchTo().Window(LastValidWindow);
            }

            string editUrl;// = string.Empty;
                           //找到「編輯」超連結
            IWebElement iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
            if (iwe == null)
            {
                //iwe = Browser.driver.FindElement(By.XPath("//*[@id=\"content\"]/div[4]/div[2]/a[2]"));
                //iwe = Browser.driver.FindElement(By.XPath("/html/body/div[2]/div[4]/div[2]/a[2]"));

                Browser.driver.SwitchTo().Window(LastValidWindow);
                //找到「編輯」超連結
                iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                if (iwe == null)
                {
                    Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles.Last());
                    iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                    if (iwe == null)
                    {
                        Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles.LastOrDefault());
                        iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                        if (iwe == null)
                        {
                            string url = ActiveForm1.TextBox3Text;
                            if (IsValidUrl_ImageTextComparisonPage(url))
                            {
                                foreach (var item in Browser.driver.WindowHandles)
                                {
                                    if (ReplaceUrl_Box2Editor(Browser.driver.Url) == url)
                                    {
                                        iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                        if (iwe == null)
                                        {
                                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面；若是新頁面，請先儲存，再執行此功能。");
                                            return false;
                                        }
                                        else
                                            break;
                                    }
                                }
                                iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                if (iwe == null)
                                {
                                    bool found = false;
                                    if (IsValidUrl_ImageTextComparisonPage(ActiveForm1.TextBox3Text))
                                    {
                                        foreach (var item in Browser.driver.WindowHandles)
                                        {
                                            Browser.driver.SwitchTo().Window(item);
                                            if (ReplaceUrl_Box2Editor(Browser.driver.Url) == ActiveForm1.TextBox3Text) { found = true; break; }
                                        }
                                    }
                                    if (!found)
                                    {
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面");
                                        return false;
                                    }
                                    iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                    if (iwe == null)
                                    {
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面");
                                        return false;
                                    }

                                }

                            }
                            else
                            {
                                iwe = Edit_Linkbox_ImageTextComparisonPage;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                if (iwe == null)
                                {
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面");
                                    return false;
                                }
                            }

                        }
                    }
                }
                editUrl = iwe.GetAttribute("href");
            }
            else
            {//取得「編輯」頁面的URL
                editUrl = iwe.GetAttribute("href");
            }

            if (DirectlyReplacingCharactersPageWindowHandle == string.Empty)
            {
                foreach (var item in Browser.driver.WindowHandles)
                {

                    string url;
                    try
                    {
                        url = ReplaceUrl_Box2Editor(Browser.driver.SwitchTo().Window(item).Url);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    //if (url.StartsWith("https://ctext.org/wiki.pl?") && url.Contains("&action=editchapter"))
                    if (url == editUrl)
                    {
                        DirectlyReplacingCharactersPageWindowHandle = Browser.driver.CurrentWindowHandle; break;
                    }
                }
            }
            else//如果 DirectlyReplacingCharactersPageWindowHandle 非空值
                if (!Browser.driver.WindowHandles.Contains(DirectlyReplacingCharactersPageWindowHandle))
                DirectlyReplacingCharactersPageWindowHandle = string.Empty; //goto reOpenEdittab; }
            reOpenEdittab:
            //如果分頁中沒有開啟「編輯」頁面
            if (DirectlyReplacingCharactersPageWindowHandle == string.Empty)
            {

                //開啟完整編輯頁面
                //openNewTabWindow();
                try
                {
                    //Browser.driver.SwitchTo().NewWindow(WindowType.Tab);
                    OpenNewTabWindow();
                }
                catch (Exception)
                {
                    Browser.driver.SwitchTo().Window(LastValidWindow);
                    try
                    {
                        Browser.driver.SwitchTo().NewWindow(WindowType.Tab);

                    }
                    catch (Exception)
                    {
                        Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles.Last());
                        //LastValidWindow = Browser.driver.WindowHandles.Last();
                        Browser.driver.SwitchTo().NewWindow(WindowType.Tab);
                    }

                }
                try
                {
                    Browser.driver.Navigate().GoToUrl(editUrl);
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.StartsWith("The HTTP request to the remote WebDriver server for URL "))
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("連線超時，請再重試。感恩感恩　南無阿彌陀佛");
                            }
                            else
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            }
                            return false;
                        default:
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            return false;
                    }
                }

                //取代區中的「名稱」欄名
                //while (null == waitFindWebElementBySelector_ToBeClickable("#content > table.restable > tbody > tr > td > table > tbody > tr:nth-child(1) > th:nth-child(1)", 0.2)) { }
                iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#content > table.restable > tbody > tr > td > table > tbody > tr:nth-child(1) > th:nth-child(1)", 10);
                if (iwe != null)
                    DirectlyReplacingCharactersPageWindowHandle = Browser.driver.CurrentWindowHandle;
                else
                    return false;
            }
            else
            {//如果現成的分頁有找到「編輯」頁面則切換到該頁面
                try
                {
                    Browser.driver.SwitchTo().Window(DirectlyReplacingCharactersPageWindowHandle);
                }
                catch (Exception err)
                {
                    DirectlyReplacingCharactersPageWindowHandle = string.Empty;
                    if (editUrl != string.Empty) goto reOpenEdittab;
                    else
                    { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(err.Message); return false; }
                }
            }

            return true;

        }
        /// <summary>
        /// 直接取代文字的編輯頁面
        /// </summary>
        internal static string DirectlyReplacingCharactersPageWindowHandle = string.Empty;
        /// <summary>
        /// 直接取代文字
        /// </summary>
        /// <param name="character">要直接被取代的單字（regexfrom）及取代成的單字（regexto）的字串陣列</param>
        /// <returns>成功則傳回true</returns>
        internal static bool DirectlyReplacingCharacters(StringInfo character)
        {
            #region 防呆

            if (character.LengthInTextElements != 2) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("指定的字元長度不對！請檢查"); return false; }
            if (character.SubstringByTextElements(0, 1) == character.SubstringByTextElements(1, 1)) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("所指定取代的字元相同，請重設！"); return false; }
            if (!Form1.IsChineseString(character.SubstringByTextElements(0, 1)) || !Form1.IsChineseString(character.SubstringByTextElements(1, 1))) { return false; }

            #endregion
            OpenEditPage();

            //「內容:」欄位文字方塊控制項
            IWebElement iwe = Textarea_data_Edit_textbox;//waitFindWebElementBySelector_ToBeClickable("#data");
            if (iwe == null)
            {
                DirectlyReplacingCharactersPageWindowHandle = string.Empty;
                Debugger.Break(); return false;
            }// goto reOpenEdittab; } //20260102●●●●●●●●●●●●●●●●●

            //輸入取代後的值//https://copilot.microsoft.com/shares/adtmSoMCVAJAebZxMzEke :GetDomProperty(name)	從 DOM 中的 property 讀取	適合取得即時屬性值，例如 checked、value、textContent	反映瀏覽器渲染後的最新狀態
            if (!SetIWebElementValueProperty(iwe, iwe.GetDomProperty("value").Replace(character.SubstringByTextElements(0, 1), character.SubstringByTextElements(1, 1)))) Debugger.Break();

            Browser.driver.SwitchTo().Window(LastValidWindow);
            return true;
        }

        /// <summary>       
        /// 20240430 Copilot大菩薩：下載網頁圖片的錯誤處理：
        /// 以下是一個使用 Selenium 來模擬「另存圖片」的基本範例。請注意，這個範例需要使用到 Actions 類別來模擬鼠標右鍵點擊和選擇「另存圖片」的選項，並且可能需要根據您的瀏覽器和操作系統的具體情況來調整。
        /// 段程式碼會打開圖片的網頁，然後模擬鼠標右鍵點擊圖片，並選擇「另存圖片」的選項。然而，這只是一個基本的範例，並且可能需要根據您的具體情況來調整。例如，處理「另存為」對話框可能需要使用到其他的工具或方法，例如 AutoIt 或 SendKeys。
        /// </summary>
        /// <param name="imageUrl">圖片所在網址</param>
        /// <param name="downloadImgFullName"></param>
        /// <param name="selectedInExplorer"></param>
        /// <returns>成功則傳回true</returns>
        internal static bool DownloadImage(string imageUrl, string downloadImgFullName)
        {
            //var Browser.driver = new ChromeDriver();
            OpenNewTabWindow();
            BringToFront("chrome");
        reGoto:
            try
            {
                Browser.driver.Navigate().GoToUrl(imageUrl);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://The HTTP request to the remote WebDriver server for URL http://localhost:5908/session/0b71d83809d531eca84ae9d77e0b4888/url timed out after 30.5 seconds.
                        if (ex.Message.EndsWith(" timed out after 30.5 seconds."))
                        {
                            Thread.Sleep(1500);
                            if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("下載書圖的網頁有問題，是否繼續？" +
                                Environment.NewLine + Environment.NewLine + "請確認網頁沒問題再按確定，否則請按取消。感恩感恩　南無阿彌陀佛") == DialogResult.Cancel)
                                return false;
                            goto reGoto;
                        }
                        else
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        return false;
                }
            }
            BringToFront("chrome");
            try
            {
                Browser.driver.SwitchTo().Window(Browser.driver.CurrentWindowHandle);
            }
            catch (Exception)
            {
                return false;
                //throw;
            }
            //IWebElement iw = waitFindWebElementBySelector_ToBeClickable("body > img");
            //Cursor.Position = (Point)iw?.Location;
            ////if (iw != null)  clickCopybutton_GjcoolFastExperience(iw.Location); 

            try
            {
                // 找到圖片元素
                var imageElement = Browser.driver.FindElement(By.TagName("img"));

                // 建立 Actions 物件
                var action = new Actions(Browser.driver);

                // 模擬鼠標右鍵點擊圖片

                action.ContextClick(imageElement).Perform();

                // 模擬按下「V」鍵，選擇「另存圖片」的選項
                // 注意：這可能需要根據您的瀏覽器和語言設定來調整
                action.SendKeys("v").Perform();
                //SendKeys.Send("{v 2}");
                SendKeys.SendWait("v");

                // TODO: 處理彈出的「另存為」對話框，輸入文件名並點擊「保存」
                // 這可能需要使用到其他的工具或方法，例如 AutoIt 或 SendKeys
                Clipboard.Clear();
                try
                {
                    Clipboard.SetText(downloadImgFullName);
                }
                catch (Exception)
                {
                }
                //Thread.Sleep(1190 + (
                Thread.Sleep(1900 + (
                    800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 : Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）
                                                                                                                                                                                              //Thread.Sleep(1200);
                                                                                                                                                                                              //Thread.Sleep(500);            


                //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
                SendKeys.SendWait("+{Insert}~~");//or "^v"
                                                 //Thread.Sleep(200);
                                                 //SendKeys.Send("{ENTER}");
                                                 //SendKeys.SendWait("%s");
                                                 //Clipboard.Clear();

                //Thread.Sleep(300);
            }
            catch (Exception)
            {
                return false;
            }

            try
            {
                Browser.driver.Close();
            }
            catch (Exception)
            {
                Browser.driver.SwitchTo().Window(LastValidWindow);//如果沒有切回關閉前的分頁，再打算開新分頁時Selenium就會出錯！20240720
                return false;
            }
            Browser.driver.SwitchTo().Window(LastValidWindow);//如果沒有切回關閉前的分頁，再打算開新分頁時Selenium就會出錯！20240720

            ////等待書圖檔下載完成
            //DateTime dt = DateTime.Now;
            //while (!File.Exists(downloadImgFullName))
            //{
            //    if (DateTime.Now.Subtract(dt).TotalSeconds > 28) return false;
            //}
            return true;
        }

        /// <summary>
        /// 在需要連續輸入截圖時 。按下Ctrl並按下滑鼠下一頁鍵時。今因《四庫全書》本《本草綱目》而設 20240510
        /// 須先畫出之截圖區域，然後按下Ctrl並按下滑鼠下一頁鍵時，會自動按下頁面中的[Input picture]連結並再按下 Replace page with this data 按鈕
        /// </summary>
        /// <returns>失敗則傳回false</returns>
        internal static bool Input_picture()
        {
            //按下頁面中的[Input picture]連結

            IWebElement iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#editor > a:nth-child(5)");
            if (iwe != null)
            {
                //iwe.Click();
                iwe.JsClick();
                //變更 quality="5" 為 quality="10" 以提高圖像質量
                iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#picturexml");
                SetIWebElementValueProperty(iwe, iwe.GetAttribute("value").Replace("quality=\"5\"", "quality=\"10\""));

                //再按下 Replace page with this data 按鈕
                iwe = Browser.WaitFindWebElementBySelector_ToBeClickable("#pictureinput > input[type=submit]");
                if (iwe != null)
                    //iwe.Click();
                    iwe.JsClick();
                else
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("頁面的【Replace page with this data 按鈕】沒找到。");
                    return false;
                }

            }
            else
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("頁面的【[Input picture]連結元件】沒找到。");
                return false;
            }
            return true;
        }

        /// <summary>
        /// 檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl 20240929 52生日
        /// <returns></returns>
        internal static bool IsConfirmHumanPage()
        {
            bool result = true; int retryCount = 0;
        retry:
            try
            {
                //result = confirm_that_you_are_human = Browser.driver.Url == "https://ctext.org/wiki.pl" || Please_confirm_that_you_are_human_Page != null;
                confirm_that_you_are_human = (Browser.driver.Url == "https://ctext.org/wiki.pl" || Please_confirm_that_you_are_human_Page != null);
                result = confirm_that_you_are_human;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.HResult + ex.Message);
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("tab crashed"))
                            RestartChromedriver();
                        break;
                    default:

                        break;
                }
                Console.WriteLine(WebDriverWaitTimeSpan.ToString());
                Console.WriteLine(Browser.driver.Manage().Timeouts().PageLoad.ToString());
                //Debugger.Break();
                if (Browser.driver.Manage().Timeouts().PageLoad < DriverManageTimeoutsPageLoad)
                    Browser.driver.Manage().Timeouts().PageLoad = DriverManageTimeoutsPageLoad;
                Thread.Sleep(1000);
                if (retryCount < 2) { retryCount++; goto retry; }

            }
            if (result)
            {
                SetPlease_confirm_that_you_are_human_Page_Occurrence_Interrupt_Message();
            }

            return result;
            //return confirm_that_you_are_human;

            //if (Browser.driver.Url == "https://ctext.org/wiki.pl" ||Please_confirm_that_you_are_human_Page!=null)
            //{
            //    if (Browser.WaitFindWebElementBySelector_ToBeClickable("#content > font")?.GetAttribute("textContent") == "Please confirm that you are human! 敬請輸入認證圖案")
            //    {
            //        confirm_that_you_are_human = true;
            //        return true;
            //    }
            //    else
            //        return false;
            //}
            //else
            //    return false;
        }

        internal static void SetPlease_confirm_that_you_are_human_Page_Occurrence_Interrupt_Message()
        {
            string theLetters = string.Empty, clipboardText = Clipboard.GetText();
            //if (!string.IsNullOrEmpty(clipboardText) &&//Leo AI 大菩薩 20260204
            //    clipboardText.Length <= 10 &&
            //    Regex.IsMatch(clipboardText, @"^[a-zA-Z0-9]+$"))
            if (Clipboard.GetText() is var clip &&
                clip.Length > 0 &&
                clip.Length <= 10 &&
                clip.All(c => char.IsLetterOrDigit(c)))
            {
                theLetters += ": " + clipboardText + " ";
            }
            Please_confirm_that_you_are_human_Page_Occurrence_Interrupt_Message =
                "<p>{{{⚠🚫✋⚡佛弟子文獻學博士孫守真任真甫按：🚧😵因認證碼機制（" +
                "\"Please confirm that you are human! 敬請輸入認證圖案\"" +
                theLetters +
                "）掣肘而致 TextForCtext 自動連續輸入中斷，屢次數處向站主反應投訴卻概不見報，愚為此干擾折騰隱忍洎今已逾年所，故請來者賢友諸仁注意協力檢查文本是否有經正確地輸入！愚莫復獨自承擔❤️💕日暮途遠，夫我則不暇矣。見原見諒⚠️☢️🈲感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主　哈利路亞 👼 " +
                DateTime.Now.ToString() + "}}}<p>";
        }

        /// <summary>
        /// 打開展開/收起閉合大綱標題（章節頁面）
        /// </summary>
        internal static void OutlineTitlesCloseOpenFoldExpandSwitcher()
        {
            if (Browser.driver == null) return;
            ActiveForm1.TopMost = false;
            if (IsDriverInvalid)
            {
                Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles.LastOrDefault());
            }
            if (!Browser.driver.Url.StartsWith("https://ctext.org/wiki.pl?if=gb&res="))
            {
                for (int i = Browser.driver.WindowHandles.Count - 1; i > -1; i--)
                {
                    string url = Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles[i]).Url;
                    //if (Browser.driver.SwitchTo().Window(Browser.driver.WindowHandles[i]).Url.StartsWith("https://ctext.org/wiki.pl?if=gb&res="))
                    if (url.StartsWith("https://ctext.org/wiki.pl?if=gb&res=") ||
                        url.StartsWith("https://ctext.org/wiki.pl?if=en&res="))
                        break;

                }

            }
            ReadOnlyCollection<IWebElement> iwes = Browser.driver.FindElements(By.TagName("DIV"));
            foreach (var item in iwes)
            {
                if (item.GetAttribute("title") == "+")
                    try
                    {
                        item.Click();
                    }
                    catch (Exception)
                    {

                    }
            }
            BringToFront("chrome");
        }

        /// <summary>
        /// 取代字串並取代分行符號為xml標記
        /// 待改名為「Replace_NewLine_Scanbreak」
        /// </summary>
        /// <param name="input"></param>
        /// <param name="oldStr"></param>
        /// <param name="newStr"></param>
        /// <param name="fileValue"></param>
        /// <returns></returns>
        internal static string Replace_NewLine_Scanbreak(string input, string fileValue = "")
        {//https://copilot.microsoft.com/shares/e3YDqo4V1tF253b1oem8a
            return fileValue == string.Empty ?
                input :
                input.Replace(Environment.NewLine, $"<scanbreak file=\"{fileValue}\" />");
        }
        /// <summary>
        /// 一次處理分行和全形空格的xml語法轉換
        /// </summary>
        /// <param name="input">要處理的字串</param>
        /// <param name="fileValue">指定 xlm 內容中的 file 值</param>
        /// <returns>轉換後的字串</returns>
        internal static string ReplaceBreaksAndFullWidthSpaces(string input, string fileValue)
        {//https://copilot.microsoft.com/shares/ar3XgcQ1vcSXRerZy7GCr
            // 1) 先切行，但自己控制換行輸出（避免 AppendLine 帶來多餘換行）
            var lines = input.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var sb = new StringBuilder();

            // 2) 處理第一行（沒有換行在前）
            if (lines.Length > 0)
            {
                sb.Append(ProcessInlineFullWidthSpaces(lines[0], fileValue));
            }

            // 3) 其餘行：在行前插入換行標記（scanbreak），視行首全形空格數量決定 y 與是否去除行首空格
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i];

                int leadingCount = 0;
                foreach (char c in line)
                {
                    if (c == '\u3000') leadingCount++;
                    else break;
                }

                if (leadingCount > 0)
                {
                    string trimmed = line.Substring(leadingCount);
                    sb.Append($"<scanbreak file=\"{fileValue}\" y=\"{leadingCount}\" />");
                    sb.Append(ProcessInlineFullWidthSpaces(trimmed, fileValue));
                }
                else
                {
                    sb.Append($"<scanbreak file=\"{fileValue}\" />");
                    sb.Append(ProcessInlineFullWidthSpaces(line, fileValue));
                }
            }

            return sb.ToString();
        }

        // 行中／行尾的全形空格群組，轉為 <scanskip file="..." y="N" />
        private static string ProcessInlineFullWidthSpaces(string line, string fileValue)
        {
            // 行首空格不在此處理；本方法僅處理非行首的全形空格
            // 用正則把非行首的 U+3000 連續群組替換為 scanskip
            // 先跳過行首的 U+3000（若有），讓 caller 決定行首邏輯
            int idx = 0;
            while (idx < line.Length && line[idx] == '\u3000') idx++;

            string head = line.Substring(0, idx);
            string tail = line.Substring(idx);

            // 將 tail 中的所有 \u3000 群組替換為 scanskip
            string replacedTail = Regex.Replace(tail, "\u3000+", m =>
            {
                int n = m.Value.Length;
                return $"<scanskip file=\"{fileValue}\" y=\"{n}\" />";
            });

            return head + replacedTail;
        }


        /// <summary>
        /// 下載書圖以供OCR用
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="imageUrl">書圖網址</param>
        /// <param name="pageUrl">書圖所在網頁網址（即圖文對照頁面，即textBox3.Text的值）</param>
        /// <param name="downloadImgFullName">下載路徑全檔名</param>
        /// <returns>若下載成功則傳回true</returns>
        internal static bool DownloadImage(ChromeDriver driver, string imageUrl, string pageUrl, out string downloadImgFullName)
        {//https://copilot.microsoft.com/shares/869qxNTvQ3AzbsSX2N8iG  https://copilot.microsoft.com/shares/869qxNTvQ3AzbsSX2N8iG 20260109
            downloadImgFullName = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "CtextTempFiles",
                "Ctext_Page_Image.png");

            Directory.CreateDirectory(Path.GetDirectoryName(downloadImgFullName));
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // 先試 Cookie 模式
            if (DownloadImage_WithSeleniumCookies(driver, imageUrl, pageUrl, downloadImgFullName))
                return true;

            //// 如果失敗，回退到瀏覽器 fetch 模式
            //return DownloadImage_ViaBrowserFetch(driver, imageUrl, downloadImgFullName);
            // 如果失敗，回退到瀏覽器 fetch 模式
            if (DownloadImage_ViaBrowserFetch(driver, imageUrl, downloadImgFullName))
                return true;

            // 再失敗，則回到原來開新分頁的方式下載
            return DownloadImage(imageUrl, downloadImgFullName);
        }

        private static bool DownloadImage_WithSeleniumCookies(ChromeDriver driver, string imageUrl, string pageUrl, string downloadImgFullName)
        {
            try
            {
                var seleniumCookies = driver.Manage().Cookies.AllCookies;
                var cookieContainer = new CookieContainer();

                foreach (var c in seleniumCookies)
                {
                    var domain = c.Domain;
                    if (!domain.StartsWith(".")) domain = "." + domain;
                    var path = string.IsNullOrEmpty(c.Path) ? "/" : c.Path;

                    try
                    {
                        cookieContainer.Add(new System.Net.Cookie(c.Name, c.Value, path, domain)
                        {
                            Secure = c.Secure,
                            HttpOnly = c.IsHttpOnly,
                            Expires = c.Expiry ?? DateTime.MinValue
                        });
                    }
                    catch { }
                }

                var handler = new HttpClientHandler
                {
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                    UseCookies = true,
                    CookieContainer = cookieContainer,
                    AllowAutoRedirect = true
                };

                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.Add("User-Agent",
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36");
                    client.DefaultRequestHeaders.Add("Accept", "image/png,image/*;q=0.8,*/*;q=0.5");
                    client.DefaultRequestHeaders.Add("Accept-Language", "zh-TW,zh;q=0.9,en;q=0.8");

                    if (!string.IsNullOrWhiteSpace(pageUrl))
                        client.DefaultRequestHeaders.Referrer = new Uri(pageUrl);

                    var response = client.GetAsync(imageUrl).GetAwaiter().GetResult();
                    if (!response.IsSuccessStatusCode) return false;

                    var bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                    if (IsHotlinkBlocked(response, bytes)) return false;

                    File.WriteAllBytes(downloadImgFullName, bytes);
                    return true;
                }
            }
            catch { return false; }
        }

        private static bool DownloadImage_ViaBrowserFetch(ChromeDriver driver, string imageUrl, string downloadImgFullName)
        {
            try
            {
                string script = @"
var callback = arguments[arguments.length - 1];
(async function(url) {
  try {
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) { callback(null); return; }
    const blob = await res.blob();
    const arrayBuf = await blob.arrayBuffer();
    let binary = '';
    const bytes = new Uint8Array(arrayBuf);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) { binary += String.fromCharCode(bytes[i]); }
    callback(btoa(binary));
  } catch (e) { callback(null); }
})(arguments[0]);
";
                var base64 = (string)((OpenQA.Selenium.IJavaScriptExecutor)driver).ExecuteAsyncScript(script, imageUrl);
                if (string.IsNullOrEmpty(base64)) return false;

                var bytes = Convert.FromBase64String(base64);
                if (bytes.Length < 7000) return false;

                File.WriteAllBytes(downloadImgFullName, bytes);
                return true;
            }
            catch { return false; }
        }

        private static bool IsHotlinkBlocked(HttpResponseMessage response, byte[] bytes)
        {
            if (response.Headers.TryGetValues("X-Sendfile", out var xs))
                foreach (var v in xs)
                    if (v.IndexOf("hotlink.png", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;

            if (bytes == null || bytes.Length < 7000) return true;
            return false;
        }

        /* 20230408 Bing大菩薩 ： 您可以使用正則表達式來簡化您的 if 判斷句。例如，您可以將條件提取到一個單獨的函數中，並使用正則表達式來檢查 url 是否包含特定字符串：
         */
        /// <summary>
        /// 檢查要輸入簡單修改模式頁面的指定網址是否合法
        /// </summary>
        /// <param name="url">要檢查的網址字串值</param>
        /// <returns>回傳網址是否合法</returns>
        internal static bool IsValidUrl＿keyDownCtrlAdd(string url)
        {
            url = ClearUrl_BoxEtc(url);
            //return Regex.IsMatch(url, @"(#editor|&page=|ctext\.org)");
            //return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*#editor");
            //也有可能是這種網址：https://ctext.org/library.pl?if=gb&file=34195&page=142&editwiki=826120#box(140,120,2,0)
            //return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*&edit");
            return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*#edit") || Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*&editwiki=.*");//20250126
            /*
             * Bing大菩薩：是的，在正則表達式中，小數點「.」是一個特殊字符，它匹配任何單個字符（除了換行符）。如果您想在正則表達式中匹配字面上的小數點，則需要在前面加上反斜杠「\」來對其進行轉義。
             * 在 C# 中，由於反斜杠「\」本身也是一個轉義字符，所以您需要使用兩個反斜杠「\\」來表示一個字面上的反斜杠。因此，在 C# 中的正則表達式中，要匹配字面上的小數點，您需要寫成「\\.」。
                希望這對您有所幫助！*/
        }
        /// <summary>
        /// 檢查是否是瀏覽圖文對照之頁面
        /// 可與 isQuickEditUrl 方法互參用
        /// </summary>
        /// <param name="url">要檢查的網址字串值</param>
        /// <returns></returns>
        internal static bool IsValidUrl_ImageTextComparisonPage(string url)
        {
            return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=");
        }
        /// <summary>
        /// 將圖文對照網址修整、規範之
        /// 20240813 creedit with Copilot大菩薩：改進C#程式碼：圖文對照網址修整：https://sl.bing.net/f2S0RcHJLyK
        /// 與 ReplaceUrl_Box2Editor 可互參考
        /// </summary>
        /// <param name="url">要被修整、規範化的圖文對照網址</param>
        /// <param name="editor">是否要在末尾改綴上"#editor"字串</param>
        /// <param name="driverGoToUrl">是否要移至這個網址</param>
        /// <returns>回傳修整過、規範的圖文對照網址</returns>
        internal static string FixUrl_ImageTextComparisonPage(string url, bool editor = false, bool driverGoToUrl = false)
        {
            #region 防呆
            if (!IsValidUrl_ImageTextComparisonPage(url) || BrowsrOPMode == BrowserOPMode.appActivateByName || driver == null) return null;
            #endregion

            // 使用正則表達式檢查和替換網址中的特定字串
            url = System.Text.RegularExpressions.Regex.Replace(url, "#box\\(.*?\\)", editor ? "#editor" : string.Empty);

            try
            {
                if (driverGoToUrl) driver.Navigate().GoToUrl(url);
            }
            catch (Exception ex)
            {
                // 記錄詳細的錯誤訊息
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly($"Error: {ex.HResult} - {ex.Message}");
            }

            return url;
        }

        /// <summary>
        /// 是否是圖文對照的頁面
        /// </summary>
        /// <returns></returns>
        internal static bool IsImageTextComparisonPage()
        {
            return Div_generic_TextBoxFrame != null;
        }



    }
    /// <summary>
    /// 有關 XML 處理的靜態類別；凡是完整編輯等處要處理 XML 內容的功能都放在這裡 20260111
    /// </summary>
    public static class XML
    {
        /// <summary>
        /// 存放已經編輯（修改）過的頁碼清單，供翻頁時核對避免重複處理
        /// </summary>
        public static HashSet<int> EditedPagesCache = new HashSet<int>();//https://gemini.google.com/share/1064e057a6f8

        /// <summary>
        /// 將編輯區內容指定頁碼之後的部分搬到下一個單位後回存
        /// 在Selenium模式表單顯示時，且其內容不少於30字元，按下Ctrl+Shift+Alt 可以啟動
        /// 在textBox1第一行/段輸入「修改摘要」欄位的值，開頭須是「本書與 URN: ctp:」以供程式判斷
        /// 如：本書與 URN: ctp:wb951851 同版，今據此版並以末學自製於GitHub開源、免費免安裝之TextForCtext應用程式及其內之WordVBA對應迻入，（討論區與末學YouTube頻道有實境演示影片可供參考），以俟後賢精校。各本後亦可同步更新。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
        /// </summary>
        internal static void Move2NextSectionChapter()
        {
            if (Browser.IsDriverInvalid) return;
            #region 將編輯區內容指定頁碼之後的部分搬到下一個單位後回存 20260111
            //20260111 為方便編輯同版書籍作對應頁故，因《天下郡國利病書》此二本之迻錄 https://ctext.org/wiki.pl?if=gb&res=5115261 https://ctext.org/wiki.pl?if=gb&res=951851
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            Form1.InstanceForm1.PauseEvents();
            Form1.InstanceForm1.TextBox3Text = driver.Url;
            Form1.InstanceForm1.ResumeEvents();
            //如果不是圖文對照之頁面就離開，不處理
            if (Div_generic_TextBoxFrame == null) return;
            string pageNum = CurrentPageNum_textbox_Value;
            if (pageNum.IsNullOrEmpty()) return;

            PlaySound(SoundLike.exam);
            if (Edit_Linkbox_ImageTextComparisonPage != null)
            {
                //進入編輯區
                Edit_Linkbox_ImageTextComparisonPage.JsClick();
                WaitFindWebElementBySelector_ToBeClickable("#data", 5);
                if (Textarea_data_Edit_textbox == null) return;
                string data = Textarea_data_Edit_textboxTxt;//Textarea_data_Edit_textbox.GetDomProperty("value");
                if (data.IsNullOrEmpty()) return;

                // 命名結果版本 https://copilot.microsoft.com/shares/gSnkjNrwocdyBVayFzFsg
                SplitResult part = SplitContentAtPageNum(data, pageNum);
                if (part.SplitPos == -1)
                {
                    // 未分割：result.FirstPart = 原字串；result.SecondPart = null
                    return;
                }
                else
                {
                    //// 已分割
                    //var first = result.FirstPart;
                    //var second = result.SecondPart;
                }
                //先設定一次；如果後續出錯的話還可以手動操作貼上，如碰到要輸入驗證時
                SetIWebElementValueProperty(Textarea_data_Edit_textbox, part.FirstPart);
                string sequence = Sequence_Edit_Chapter_Value;//取得目前章節單位的序號
                Clipboard.SetText(part.SecondPart);//存入剪貼簿中，以備萬一，若網頁壞掉或須驗證碼，則便手動搬到下一個單位來貼上
                if (sequence.IsNullOrEmpty()) return;
                if (Commit == null) return;
                Commit.JsClick();//保存編輯
                if (Title_Chapter_BookName == null) return;
                Title_Chapter_BookName.JsClick();//返回全書單位_章節列表
                                                 //進入下一單位
                if (!GotoNextSection_SequenceChapterPage(sequence)) return;//進入下一章節文字版內容
                if (Edit_linkbox == null)
                {
                    MessageBoxShowOKExclamationDefaultDesktopOnly("找不到「修改」元件。當是此編輯頁面中的「標題／篇名:」欄位置誤植入XML片段所致！請加以改正。感恩感恩　南無阿彌陀佛");
                    return;
                }
                Edit_linkbox.JsClick();//進入編輯頁面
                if (Textarea_data_Edit_textbox == null) return;
                //string moved = part[1];//Clipboard.SetText(part[1]);
                SetIWebElementValueProperty(Textarea_data_Edit_textbox, part.SecondPart);//搬入內容，取代原有內容（未經校對之亂碼內容）                
                SetIWebElementValueProperty(Description_Edit_textbox, GetDescription("本書與 URN: ctp:", "本書與本站另一本同版，今據此版並以末學自製於GitHub開源、免費免安裝之TextForCtext應用程式及其內之WordVBA對應迻入（討論區與末學YouTube頻道有實境演示影片可供參考），以俟後賢精校。各本後亦可同步更新。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主"));
                /*
                 if (!description.StartsWith("本書與 URN: ctp:"))
                        description = new Document(GetSecondFormText()).GetCurrentParagraph().Text;
                    if (!description.StartsWith("本書與 URN: ctp:"))
                        description = "本書與本站另一本同版，今據此版並以末學自製於GitHub開源、免費免安裝之TextForCtext應用程式及其內之WordVBA對應迻入（討論區與末學YouTube頻道有實境演示影片可供參考），以俟後賢精校。各本後亦可同步更新。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主";            
                 */
                Commit?.JsClick();
                PlaySound(SoundLike.done);
                Clipboard.Clear();
            }
            #endregion
        }
        /// <summary>
        /// 取得要輸入到 Description:（修改摘要:）欄位的值
        /// 以From1或Form2的textBox1.Text值為據，取其插入點所在段落之文字
        /// </summary>        
        /// <param name="refers">要用以比對的關鍵詞彙字句</param>
        /// <param name="promp">若找不到的預設值</param>
        /// <returns></returns>
        internal static string GetDescription(string refers, string promp)
        {
            //先取得textBox1的值，若不吻合，再取得Form2的textBox1的值
            string description = Form1.InstanceForm1.Document.GetCurrentParagraph().Text;//Form1.InstanceForm1.TextBox1_Text;
            if (!description.StartsWith(refers))
            {
                description = new Document(GetSecondFormText()).GetCurrentParagraph().Text;
                //description = GetSecondFormText();//https://copilot.microsoft.com/shares/kudAKddNP7MztHtzaz2o4
                //Console.WriteLine("第2個表單的文字：" + text);
                if (!description.StartsWith(refers))
                    description = promp;
            }
            return description;
        }


        /* creedit with Gemini大菩薩、Leo AI大菩薩 20260114：
         * 代碼邏輯說明
    動態提取 File ID：代碼會自動尋找最後一個 <scanend> 標記，並抓取其中的 file="238685"。這樣即使您的 file 編號改變了，程式碼依然有效。
    定位插入點：使用 lastMatch.Index 找到最後一個標記的起始位置，在那裡插入 \r\n\r\n，完美符合您「在 <scanend 前加上換行」的需求。
    字串插值 (String Interpolation)：使用 $"" 語法，可以非常直觀地將變數 uboundPageNum 嵌入到標記字串中。
    執行結果預覽
    假設輸入您的原始文本，uboundPageNum = 204，輸出將會是：
        ……
    <scanbreak file="238685" />|<scanbreak file="238685" />《經義述聞》弟二十二

    <scanend file="238685" page="202" /><scanbegin file="238685" page="204" />●<scanend file="238685" page="204" />

        溫馨小提示
    如果您的 Windows Forms 環境需要處理大量這類文本，建議確保 input 字串不是 null。
    如果您需要處理的 file 屬性不是純數字，只需將正則表達式中的 \d+ 改為 [^"]+ 即
         */

        // https://chatgpt.com/share/696261b5-652c-800b-bcc5-81f4bb36efb8
        /// <summary>
        /// 要調整 <scanbegin> / <scanend> 中的屬性值時使用
        /// xml編輯 完整編輯頁面時會用到 20260110
        /// </summary>
        public static class ScanPageAdjuster
        {
            /// <summary>
            /// 安全地調整 scanbegin / scanend 的 page，
            /// 以及 picture location 中的頁碼（第二段）
            /// </summary>
            /// <param name="input"></param>
            /// <param name="offset">頁碼差（=目的-來源）</param>
            /// <param name="minPage"></param>
            /// <returns></returns>
            public static string ShiftScanPages(string input, int offset = -1, int minPage = 1)
            {
                if (string.IsNullOrEmpty(input))
                    return input;

                StringBuilder output = new StringBuilder();

                XmlReaderSettings readerSettings = new XmlReaderSettings
                {
                    ConformanceLevel = ConformanceLevel.Fragment,
                    IgnoreComments = false,
                    IgnoreWhitespace = false
                };

                XmlWriterSettings writerSettings = new XmlWriterSettings
                {
                    ConformanceLevel = ConformanceLevel.Fragment,
                    OmitXmlDeclaration = true
                };

                XmlReader reader = null;
                XmlWriter writer = null;

                try
                {
                    reader = XmlReader.Create(new StringReader(input), readerSettings);
                    writer = XmlWriter.Create(output, writerSettings);

                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            writer.WriteStartElement(reader.Name);

                            if (reader.HasAttributes)
                            {
                                while (reader.MoveToNextAttribute())
                                {
                                    // 1. scanbegin / scanend 的 page
                                    if ((reader.Name == "page") &&
                                        (reader.Value.Length > 0) &&
                                        IsAllDigits(reader.Value))
                                    {
                                        int page = int.Parse(reader.Value);
                                        int newPage = Math.Max(page + offset, minPage);

                                        writer.WriteAttributeString(
                                            reader.Name,
                                            newPage.ToString()
                                        );
                                    }
                                    // 2. picture 的 location
                                    else if (reader.Name == "location")
                                    {
                                        writer.WriteAttributeString(
                                            reader.Name,
                                            AdjustPictureLocation(reader.Value, offset, minPage)
                                        );
                                    }
                                    else
                                    {
                                        // 其他屬性（包含 file）原樣保留
                                        writer.WriteAttributeString(reader.Name, reader.Value);
                                    }
                                }

                                reader.MoveToElement();
                            }

                            if (reader.IsEmptyElement)
                            {
                                writer.WriteEndElement();
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.Text)
                        {
                            writer.WriteString(reader.Value);
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement)
                        {
                            writer.WriteEndElement();
                        }
                        else if (reader.NodeType == XmlNodeType.CDATA)
                        {
                            writer.WriteCData(reader.Value);
                        }
                        else if (reader.NodeType == XmlNodeType.Whitespace ||
                                 reader.NodeType == XmlNodeType.SignificantWhitespace)
                        {
                            writer.WriteWhitespace(reader.Value);
                        }
                    }
                }
                finally
                {
                    writer?.Close();
                    reader?.Close();
                }

                return output.ToString();
            }

            /// <summary>
            /// 調整 picture location="檔號:頁碼(:座標)"
            /// 只改第二段頁碼
            /// </summary>
            private static string AdjustPictureLocation(string location, int offset, int minPage)
            {
                if (string.IsNullOrEmpty(location))
                    return location;

                string[] parts = location.Split(':');
                if (parts.Length < 2)
                    return location;

                if (!int.TryParse(parts[1], out int page))
                    return location;

                parts[1] = Math.Max(page + offset, minPage).ToString();
                return string.Join(":", parts);
            }

            /// <summary>
            /// 判斷是否全為數字（避免 LINQ，C# 7.3 安全）
            /// </summary>
            private static bool IsAllDigits(string text)
            {
                for (int i = 0; i < text.Length; i++)
                {
                    if (text[i] < '0' || text[i] > '9')
                        return false;
                }
                return true;
            }

            /* 20260111 完整編輯頁面會用到 由 editPageNumOffset_PageNumModifier發想而來
             https://copilot.microsoft.com/shares/TuLUScKoFWypCnvBzpCfS
             */
            /// <summary>
            /// 分割含有「<scanbegin file=」的XML為兩部分，依據指定的 pageNum
            ///  Tuple 版本
            /// </summary>
            /// <param name="content">要分割的XML內容</param>
            /// <param name="pageNum">要據以分割的頁碼</param>
            /// <param name="splitPos">可取得分割點的字串位置</param>
            /// <returns>傳回Tuple<string,string>物件，其第1個元素即分割出來的第1部分。第2個即第2部分</returns>
            internal static Tuple<string, string> SplitContentAtPageNum(string content, string pageNum, out int splitPos)
            {
                // 防呆：檢查輸入
                if (string.IsNullOrEmpty(content) || string.IsNullOrEmpty(pageNum))
                {
                    splitPos = -1;
                    return Tuple.Create<string, string>(content, (string)null);
                    // 或：return new Tuple<string, string>(content, null);
                }

                // 找到 pageNum 的位置
                int pageIndex = content.IndexOf(" page=\"" + pageNum + "\"", StringComparison.Ordinal);
                if (pageIndex == -1)
                {
                    splitPos = -1;
                    return Tuple.Create<string, string>(content, (string)null);
                }

                // 找到對應的 <scanbegin file= 標記
                splitPos = content.LastIndexOf("<scanbegin file=", pageIndex, StringComparison.Ordinal);
                if (splitPos == -1)
                {
                    return Tuple.Create<string, string>(content, (string)null);
                }

                // 分割字串
                string firstPart = content.Substring(0, splitPos);
                string secondPart = content.Substring(splitPos);

                return Tuple.Create<string, string>(firstPart, secondPart);
            }

            //https://copilot.microsoft.com/shares/qkNvEXe11Xuo8CbsYFJhu

            /// <summary>            
            /// 使用命名結果的版本
            /// 分割含有「<scanbegin file=」的XML為兩部分，依據指定的 pageNum
            /// </summary>
            /// <param name="content">要分割的XML內容</param>
            /// <param name="pageNum">要據以分割的頁碼</param>
            /// <returns>傳回 SplitResult 物件，其第1個元素即分割出來的第1部分。第2個即第2部分</returns>
            internal static SplitResult SplitContentAtPageNum(string content, string pageNum)
            {
                if (string.IsNullOrEmpty(content) || string.IsNullOrEmpty(pageNum))
                    return SplitResult.NotSplit(content);

                int pageIndex = content.IndexOf(" page=\"" + pageNum + "\"", StringComparison.Ordinal);
                if (pageIndex == -1)
                    return SplitResult.NotSplit(content);

                int splitPos = content.LastIndexOf("<scanbegin file=", pageIndex, StringComparison.Ordinal);
                if (splitPos == -1)
                    return SplitResult.NotSplit(content);

                return new SplitResult(
                    content.Substring(0, splitPos),
                    content.Substring(splitPos),
                    splitPos
                );
            }

            //https://copilot.microsoft.com/shares/34kpDrMYhMRELXecXRctw

            /// <summary>
            /// 如果你不喜歡 Item1/Item2，可以用一個簡單的類別包起來，讓語意更直觀 
            /// </summary>
            internal sealed class SplitResult
            {
                public string FirstPart { get; private set; }
                public string SecondPart { get; private set; }
                public int SplitPos { get; private set; }

                public SplitResult(string firstPart, string secondPart, int splitPos)
                {
                    FirstPart = firstPart;
                    SecondPart = secondPart;
                    SplitPos = splitPos;
                }

                public static SplitResult NotSplit(string originalContent)
                {
                    return new SplitResult(originalContent, null, -1);
                }
            }


            /// <summary>
            /// Xml重新編頁：編輯區內容頁碼減頁差後回存。
            /// 頁差由textBox1第一段設定
            /// （頁差=「來源-目的」頁碼；意謂來源到目的；或「來源~目的」-蓋在手動輸入模式下是禁止輸入「-」，因為數字鍵盤中的「-」常用來作自動段落標記功能用）
            /// 如果用「~」而省略來源頁碼，則取目前頁面之頁碼；不可用「-」，會與負/減混
            /// 在Selenium模式有效頁面下，網頁停在要開始編輯頁碼之首頁圖文對照頁面上，按下Ctrl + Shift 再顯示Form1（TextForCtext 主介面主表單）即可全程自動化
            /// </summary>
            internal static void EditPageNumOffset_PageNumModifier(out int offset_, string refers = "")
            {
                #region 編輯區內容頁碼減1後回存 20260110
                //20260110 為方便編輯同版書籍作對應頁故，因《天下郡國利病書》此二本之迻錄 https://ctext.org/wiki.pl?if=gb&res=5115261 https://ctext.org/wiki.pl?if=gb&res=951851
                string originalText = Clipboard.GetText();//記錄剪貼簿內文字資料
                string pageNum = string.Empty; offset_ = 0; int off = 0;
                void shiftPage()
                {
                    if (refers.IsNullOrEmpty() == false)
                    {
                        refers = refers.Substring(0, refers.IndexOf(Environment.NewLine) == -1 ? refers.Length : refers.IndexOf(Environment.NewLine));
                        if (refers.IndexOf('-') > 0 || refers.IndexOf('~') > -1)//Leo AI 大菩薩
                        {
                            //string input = "178-28";
                            //如果省略來源頁碼，則取目前頁面之頁碼
                            if (refers.IndexOf('~') == 0)
                            {
                                if (pageNum.IsNullOrEmpty()) Debugger.Break();//return;
                                refers = pageNum + refers;
                            }
                            char separator = '-';
                            //string[] parts = refers.Split('-');
                            string[] parts = refers.Split(separator);
                            if ((parts.Length) == 1) { separator = '~'; parts = refers.Split(separator); }
                            SplitResult sr = new SplitResult(parts[0], parts[1], refers.IndexOf(separator));
                            int result = int.Parse(sr.SecondPart) - int.Parse(sr.FirstPart);
                            refers = result.ToString();
                        }
                        if (int.TryParse(refers, result: out int offset))
                        {
                            Clipboard.SetText(ScanPageAdjuster.ShiftScanPages(originalText, offset));
                            off = offset;
                        }
                    }
                    else
                        Clipboard.SetText(ScanPageAdjuster.ShiftScanPages(originalText));
                }


                PlaySound(SoundLike.exam);
                if (originalText.Contains(@"<scanbegin file="""))
                {
                    shiftPage();
                    PlaySound(SoundLike.done, true);
                }
                else if (!IsDriverInvalid)
                {
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                    Form1.InstanceForm1.PauseEvents();
                    Form1.InstanceForm1.TextBox3Text = driver.Url;
                    Form1.InstanceForm1.ResumeEvents();
                    if (Div_generic_TextBoxFrame != null)
                    {
                        pageNum = CurrentPageNum_textbox_Value;//PageNum_textbox.GetDomProperty("value");
                        if (pageNum.IsNullOrEmpty()) return;
                        if (Edit_Linkbox_ImageTextComparisonPage != null)
                        {
                            Edit_Linkbox_ImageTextComparisonPage.JsClick();
                            WaitFindWebElementBySelector_ToBeClickable("#data", 5);
                            if (Textarea_data_Edit_textbox == null) return;
                            if (!originalText.Contains(@"<scanbegin file="""))
                            {
                                string data = Textarea_data_Edit_textboxTxt;//Textarea_data_Edit_textbox.GetDomProperty("value");
                                if (data.IsNullOrEmpty()) return;
                                //int splitPos = data.LastIndexOf("<scanbegin file=", data.IndexOf(" page=\"" + pageNum + "\""));
                                ////<scanbegin file="175494" page="85" />
                                //Clipboard.SetText(data.Substring(splitPos));
                                //originalText = Clipboard.GetText();
                                SplitResult part = SplitContentAtPageNum(data, pageNum);
                                //List<string> part = SplitContentAtPageNum(data, pageNum, out _);
                                originalText = part.SecondPart;//data.Substring(splitPos);
                                                               //先設定一次；如果後續出錯的話還可以手動操作貼上，如碰到要輸入驗證時
                                SetIWebElementValueProperty(Textarea_data_Edit_textbox, part.FirstPart);//data.Substring(0, splitPos));
                                SetIWebElementValueProperty(Description_Edit_textbox, "調整頁碼頁碼對齊-以末學自製於GitHub開源免費免安裝之應用程式TextForCtext調正之。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");
                            }
                            //檢視重新編頁後的結果
                            shiftPage();
                        }
                    }
                    if (Textarea_data_Edit_textbox != null)
                    {
                        SetIWebElementValueProperty(Textarea_data_Edit_textbox,
                            ScanPageCleaner.RemoveTrulyEmptyScanPages(
                                Textarea_data_Edit_textbox.GetDomProperty("value") + Clipboard.GetText()));
                    }
                    if (Commit != null)
                    {
                        Commit.JsClick();
                        Clipboard.Clear();
                    }
                    PlaySound(SoundLike.done);
                }
                offset_ = off;
                #endregion
            }


            //https://chatgpt.com/s/t_69650945a1e4819185e49644c1b463ac

            /// <summary>
            /// 清理空頁碼標記            
            /// </summary>
            public static class ScanPageCleaner
            {
                private static readonly Regex EmptyScanPairRegex =
                    new Regex(
                        @"<scanbegin\b[^>]*\bpage=""(\d+)""[^>]*/><scanend\b[^>]*\bpage=""\1""[^>]*/>",
                        RegexOptions.Compiled
                    );

                /// <summary>
                /// 只移除「scanbegin 與 scanend 緊貼、且頁碼相同」的空頁
                /// </summary>
                public static string RemoveTrulyEmptyScanPages(string input)
                {
                    if (string.IsNullOrEmpty(input))
                        return input;

                    Regex EmptyScanPairRegex =
                    new Regex(
                        @"<scanbegin\b[^>]*\bpage=""(\d+)""[^>]*/><scanend\b[^>]*\bpage=""\1""[^>]*/>",
                        RegexOptions.Compiled
                    );
                    return EmptyScanPairRegex.Replace(input, string.Empty);
                }


                /// <summary>
                /// XML中分段標記<p>位置清理+配置頁1的XML內容：清除頁前的分段標記、調適星號前的分段標記
                /// </summary>
                /// <returns>失敗或不需處理皆傳回false</returns>
                internal static bool FixXMLParagraphMarkPosition_and_Page1Content()
                {
                    if (!IsDriverInvalid)
                    {
                        if (CtextPageClassifier.ParseUrl(driver.Url).PageType == CtextPageType.LibraryFile
                            || CtextPageClassifier.ParseUrl(driver.Url).PageType == CtextPageType.LibraryFileEditWiki)
                            //if (CTP.IsImageTextComparisonPage())
                            Edit_Linkbox_ImageTextComparisonPage?.JsClick();
                        //現在用 FixXMLParagraphMarkPosition_and_Page1Content(); 程式化可以背景執行了，則手動的部分，則加標題，故將插入點位置移到標題欄位中以供輸入，也不怕其他欄位受影響，因為都是程式在幕後操作了。 20260117                        
                        Title_textBox?.Click();//焦點要放在這裡，就要用 Click，不能用 JsClick

                        if (Textarea_data_Edit_textbox != null)
                        {
                            string xml = Textarea_data_Edit_textboxTxt;//Clipboard.GetText();
                                                                       //if (xml.IsNullOrEmpty()) xml = Clipboard.GetText();
                            if (xml.IsNullOrEmpty() || xml.Contains("<scanbegin file=\"") == false) return false;
                            //Clipboard.SetText(xml);                            
                            if (TextForCtext.CtextCleaner.FixXMLParagraphMarkPosition_SetPage1Content(ref xml))
                            //if (xml != Clipboard.GetText())
                            {
                                SetIWebElementValueProperty(Description_Edit_textbox, GetDescription("將星號前的分段符號移置前段之末", "將星號前的分段符號移置前段之末 & 清除頁前的分段符號{佛弟子文獻學博士孫守真任真甫按：以末學自製於GitHub開源免費免安裝之 TextForCtext 應用程式加速輸入與排版。討論區與末學YouTube頻道有演示影片可資參考。感恩感恩　讚歎讚歎　南無阿彌陀佛"));
                                //SetIWebElementValueProperty(Textarea_data_Edit_textbox, Clipboard.GetText());
                                SetIWebElementValueProperty(Textarea_data_Edit_textbox, xml);
                                if (Commit == null) return false;
                                return Commit.JsClick();
                            }
                            else
                            {
                                MessageBoxShowOKCancelExclamationDefaultDesktopOnly("不需處理~");
                                return false;
                            }
                        }
                    }
                    return false;
                }

                //準備將WordVBA中「Sub 清除頁前的分段符號()」搬到這裡來跑了 https://chatgpt.com/s/t_6965bb4984e08191b33eae051683c086 20260113蔣經國前總統逝世紀念
                //以上成功之作，以下未竟圓滿


                /// <summary>
                /// 正則只負責「定位頁面」
                /// </summary>
                private static readonly Regex PageBlockRegex =
    new Regex(
        @"<scanbegin\b[^>]*page=""(\d+)""\s*/>(.*?)<scanend\b[^>]*page=""\1""\s*/>",
        RegexOptions.Singleline | RegexOptions.Compiled
    );
                /// <summary>
                /// 頁前垃圾的「極保守」定義
                /// </summary>
                private static readonly Regex LeadingPageGarbageRegex =
    new Regex(
        @"^(?:\s*<(scanbreak|scanskip)\b[^>]*/>\s*)+",
        RegexOptions.Compiled
    );


                // 可以此頁來測試：https://ctext.org/wiki.pl?if=gb&chapter=547844

                /// <summary>
                /// 主清理方法；即WordVBA中的Sub 清除頁前的分段符號()
                /// </summary>
                /// <param name="input"></param>
                /// <returns></returns>
                public static string RemoveLeadingPageBreaks_chatGPT(string input)
                {
                    if (string.IsNullOrEmpty(input))
                        return input;

                    #region 熱重載方便故
                    //                    Regex PageBlockRegex =
                    //new Regex(
                    //    @"<scanbegin\b[^>]*page=""(\d+)""\s*/>(.*?)<scanend\b[^>]*page=""\1""\s*/>",
                    //    RegexOptions.Singleline | RegexOptions.Compiled
                    //);
                    //                /// <summary>
                    //                /// 頁前垃圾的「極保守」定義
                    //                /// </summary>
                    //                Regex LeadingPageGarbageRegex =
                    //    new Regex(
                    //        @"^(?:\s*<(scanbreak|scanskip)\b[^>]*/>\s*)+",
                    //        RegexOptions.Compiled
                    //    );

                    #endregion

                    return PageBlockRegex.Replace(input, m =>
                    {
                        string page = m.Groups[1].Value;
                        string body = m.Groups[2].Value;

                        // 只處理「頁首」
                        string cleanedBody = LeadingPageGarbageRegex.Replace(body, string.Empty);

                        return m.Value.Replace(body, cleanedBody);
                    });
                }
                /// <summary>
                /// 即WordVBA中的Sub 清除頁前的分段符號()
                /// </summary>
                /// <param name="input"></param>
                /// <returns></returns>
                public static string RemoveLeadingPageBreaks(string input)
                {
                    // 首頁補「●」
                    input = SetPage1Code(input);

                    // 分割成各頁
                    string[] pages = input.Split(new string[] { "<scanbegin" }, StringSplitOptions.None);

                    for (int i = 0; i < pages.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(pages[i])) continue;

                        // 檢查頁首是否有兩個以上的分段符號
                        var match = Regex.Match(pages[i], @"^(?:\r\n|\n){2,}");
                        //if (match.Success && i > 0)
                        //{
                        //    // 搬到上一頁尾巴
                        //    pages[i - 1] += match.Value;
                        //    // 移除當前頁首的分段符號
                        //    pages[i] = pages[i].Substring(match.Length);
                        //}
                        //https://copilot.microsoft.com/shares/Ym8kPpyjwdULFHPWGFpvG
                        if (match.Success && i > 0)
                        {
                            // 找到上一頁的 scanend 標記
                            int scanendIndex = pages[i - 1].LastIndexOf("<scanend");
                            if (scanendIndex > -1)
                            {
                                // 在 scanend 標記前插入分段符號
                                pages[i - 1] = pages[i - 1].Insert(scanendIndex, match.Value);
                            }
                            else
                            {
                                // 如果沒找到 scanend，就退而求其次，加到尾巴
                                pages[i - 1] += match.Value;
                            }

                            // 移除當前頁首的分段符號
                            pages[i] = pages[i].Substring(match.Length);
                        }

                    }

                    return string.Join("<scanbegin", pages);
                }

                private static string SetPage1Code(string xmlText)
                {
                    if (!xmlText.Contains("page=\"1\""))
                    {
                        var match = Regex.Match(xmlText, "page=\"(\\d+)\"");
                        if (match.Success)
                        {
                            int pageNum = int.Parse(match.Groups[1].Value);
                            if (pageNum < 10)
                            {
                                var fileMatch = Regex.Match(xmlText, "file=\"([^\"]+)\"");
                                if (fileMatch.Success)
                                {
                                    string bID = fileMatch.Groups[1].Value;
                                    string page1Stub =
                                        $"<scanbegin file=\"{bID}\" page=\"1\" />●<scanend file=\"{bID}\" page=\"1\" />";
                                    xmlText = page1Stub + xmlText;
                                }
                            }
                        }
                    }
                    else
                    {
                        string page1 = GetPageContent(xmlText, "1");
                        if (page1.Contains("}}<scanskip "))
                        {
                            xmlText = xmlText.Replace(page1, "●");
                        }
                        else
                        {
                            if (!Page1Exam_NotContainsRegex(page1))
                            {
                                if (Page1Exam_ContainsRegex(page1))
                                {
                                    xmlText = xmlText.Replace(page1, "●");
                                }
                                else
                                {
                                    // 簡化：自動清除
                                    xmlText = xmlText.Replace(page1, "●");
                                }
                            }
                        }
                    }
                    return xmlText;
                }

                private static string GetPageContent(string xmlText, string pageNum)
                {
                    var re = new Regex(
                        $"<scanbegin[^>]*page=\"{pageNum}\"[^>]*>([\\s\\S]*?)<scanend[^>]*page=\"{pageNum}\"[^>]*>",
                        RegexOptions.IgnoreCase);
                    var match = re.Match(xmlText);
                    return match.Success ? match.Groups[1].Value : "";
                }

                private static bool Page1Exam_NotContainsRegex(string text)
                {
                    var re = new Regex("[《》]");
                    return re.IsMatch(text);
                }

                private static bool Page1Exam_ContainsRegex(string text)
                {
                    string[] patterns = { @"\}\}\r\n\r\n<scanbreak" };
                    string combined = "(" + string.Join("|", patterns) + ")";
                    var re = new Regex(combined, RegexOptions.IgnoreCase);
                    return re.IsMatch(text);
                }




            }





        }

        public static class XmlLookup
        {

            //C# XML 頁碼連貫性檢查 https://gemini.google.com/share/4606b14ecee8 https://gemini.google.com/share/ad8899527e5e 

            /// <summary>
            /// 記錄編輯過的頁面頁碼等資訊，以供檢核
            /// 直接傳入網頁原始碼 HTML 進行解析與同步
            /// </summary>
            /// <param name="htmlSource">從 Selenium 或 HttpClient 取得的網頁原始碼</param>
            /// <returns>失敗傳回false</returns>
            public static bool ProcessHtmlAndSyncCache(string htmlSource)
            {
                if (string.IsNullOrWhiteSpace(htmlSource)) return false;
                //MatchCollection allPageMatches;

                //// 1. 定義範圍：只抓右邊「修改後」的內容 // 兼顧中英文界面標籤
                //// 如果您想精確只抓右邊「修改後」的欄位：
                //// 先切出修改後的 <td> 內容
                ////int midIndex =  htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">修改後</th>");
                ////int midIndex = htmlSource.Contains("<th width=\"50%\" class=\"colhead\">修改後</th>") ?
                ////        htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">修改後</th>") :
                ////        htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">New</th>");
                //int midIndex = -1;
                //if (htmlSource.Contains("<th width=\"50%\" class=\"colhead\">修改後</th>"))
                //    midIndex = htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">修改後</th>");
                //else if (htmlSource.Contains("<th width=\"50%\" class=\"colhead\">New</th>"))
                //    midIndex = htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">New</th>");

                //// 1. 定義範圍：鎖定右側「修改後」的內容
                //string modifiedRightPart = htmlSource;
                //if (midIndex != -1)
                //{
                //    modifiedRightPart = htmlSource.Substring(midIndex);
                //}
                ////string modifiedRightPart = htmlSource; 
                ////if (midIndex != -1)
                ////{
                ////    modifiedRightPart = htmlSource.Substring(midIndex);
                ////    // 然後只在這個部分搜尋頁碼
                ////    allPageMatches = Regex.Matches(modifiedRightPart, @"page=""(\d+)""");
                ////    // ... 後續邏輯同上 ...
                ////}
                ////else// 備援方案：掃描全頁
                ////    // 1. 定義「修改後」欄位的範圍（避免抓到「修改前」的重複頁碼）
                ////    // 在 ctext 的 diff 頁面中，修改後的內容通常在第二個 <td> 或特定 class 之後
                ////    // 但為了保險與簡便，我們直接抓取所有 page="(\d+)" 並去重即可
                ////    allPageMatches = System.Text.RegularExpressions.Regex.Matches(htmlSource, @"page=""(\d+)""");
                // --- 1. 範圍選取與初始化 ---
                int midIndex = htmlSource.Contains("<th width=\"50%\" class=\"colhead\">修改後</th>") ?
                               htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">修改後</th>") :
                               htmlSource.IndexOf("<th width=\"50%\" class=\"colhead\">New</th>");

                string modifiedRightPart = (midIndex != -1) ? htmlSource.Substring(midIndex) : htmlSource;

                EditedPagesCache.Clear();
                List<int> allPagesForReport = new List<int>();
                List<string> orphanModifications = new List<string>(); // 存放那些找不到頁碼的孤兒片段內容

                // --- 2. 呼叫核心解析邏輯 (區域函式) ---
                AnalyzeModificationBlocks();

                //// 2. 存入快取與暫存清單 獲取【所有出現過】的頁碼 (供目測參考)
                //// 2. 獲取【所有出現過】的頁碼 (供目測參考)
                //// 這裡我們直接使用 MatchCollection
                //allPageMatches = Regex.Matches(modifiedRightPart, @"page=""(\d+)""");
                ////List<int> allPagesForReport = new List<int>();

                //foreach (System.Text.RegularExpressions.Match m in allPageMatches)
                //{
                //    //if (int.TryParse(m.Groups[1].Value, out int p))
                //    //{
                //    //    // 只要頁碼出現在這個 diff 頁面，就代表它已經被變動/編輯過
                //    //    if (!EditedPagesCache.Contains(p))
                //    //    {
                //    //        pages.Add(p);
                //    //        EditedPagesCache.Add(p);
                //    //    }
                //    //}
                //    int p = int.Parse(m.Groups[1].Value);
                //    // 使用一個暫時的 HashSet 或簡單判斷來去重
                //    if (!allPagesForReport.Contains(p)) allPagesForReport.Add(p);
                //}

                //if (pages.Count == 0) return;

                //// 排序
                //pages = pages.OrderBy(p => p).ToList();

                //// 2. 產生成報告並寫入剪貼簿 (調用先前寫好的邏輯)
                //// 這裡可以呼叫您之前整理報告的 StringBuilder 部分...
                //// (為了節省篇幅，此處省略 StringBuilder 部分)

                //if (allPagesForReport.Count == 0)
                //{
                //    MessageBox.Show("找不到任何頁碼資訊。", "檢查結束");
                //    return;
                //}


                //// 3. 【核心升級】精確偵測真正有修改 (diffadd/diffdel) 的頁碼
                //EditedPagesCache.Clear();

                //// 找出所有 diffadd 或 diffdel 的位置 // 找出所有真正動到文字的位置
                //var diffMatches = Regex.Matches(modifiedRightPart, @"<span class=""diff(add|del)"">");

                //foreach (Match dm in diffMatches)
                //{
                //    // 從這個修改點 dm.Index 開始，往前回溯尋找最近的一個 page="(\d+)"
                //    // 從修改點位置 dm.Index 開始，向左搜尋最近的一個頁碼宣告
                //    // RegexOptions.RightToLeft 是處理此類「向上追溯」邏輯的神兵利器
                //    string textBeforeDiff = modifiedRightPart.Substring(0, dm.Index);
                //    var lastPageBefore = Regex.Match(textBeforeDiff, @"page=""(\d+)""", RegexOptions.RightToLeft);

                //    if (lastPageBefore.Success)
                //    {
                //        int realModifiedPage = int.Parse(lastPageBefore.Groups[1].Value);
                //        // 建議寫法，確保邏輯一致性
                //        if (!EditedPagesCache.Contains(realModifiedPage))
                //            EditedPagesCache.Add(realModifiedPage);
                //        //EditedPagesCache.Add(realModifiedPage); // 只有真正動到文字的才入防護快取// 存入全域快取防護機制
                //    }
                //}

                // 3. 檢查是否有抓到頁碼 (移動到解析完之後)
                if (allPagesForReport.Count == 0)
                {
                    MessageBox.Show("找不到任何頁碼資訊。", "檢查結束");
                    return false;
                }

                // 4. 排序與區間計算
                // 排序與去重 排序與整理
                //pages = pages.Distinct().OrderBy(p => p).ToList();
                allPagesForReport = allPagesForReport.Distinct().OrderBy(p => p).ToList();
                //allPagesForReport = allPagesForReport.OrderBy(p => p).ToList(); // HashSet 已保證唯一性，直接排序即可
                List<int> realEditedList = EditedPagesCache.OrderBy(p => p).ToList();

                //// 5. 計算「未出現」的頁碼清單
                //List<int> missingPages = new List<int>();
                //for (int i = allPagesForReport.First(); i <= allPagesForReport.Last(); i++)
                //{
                //    //// 這裡改用 EditedPagesCache 檢查效能更好
                //    //if (!EditedPagesCache.Contains(i)) missingPages.Add(i);
                //    if (!allPagesForReport.Contains(i)) missingPages.Add(i);
                //}
                // 5. 計算「未出現」的頁碼清單 (在掃描到的最小與最大值之間)
                List<int> missingPages = new List<int>();
                int minPage = allPagesForReport.First();
                int maxPage = allPagesForReport.Last();

                // 建立一個快速查詢集提高效能
                HashSet<int> allPagesSet = new HashSet<int>(allPagesForReport);
                for (int i = minPage; i <= maxPage; i++)
                {
                    if (!allPagesSet.Contains(i)) missingPages.Add(i);
                }

                // 6. 定義內部工具函數  定義小工具：範圍字串化
                //    定義局部函數：範圍字串化(例如 1~10)
                string ToRangeString(List<int> nums)
                {
                    if (nums == null || nums.Count == 0) return "無";
                    var ranges = new List<string>();
                    int start = nums[0];
                    int end = nums[0];
                    for (int i = 1; i <= nums.Count; i++)
                    {
                        if (i < nums.Count && nums[i] == end + 1)
                        {
                            end = nums[i];
                        }
                        else
                        {
                            ranges.Add(start == end ? start.ToString() : $"{start}~{end}");
                            if (i < nums.Count) { start = nums[i]; end = nums[i]; }
                        }
                    }
                    return string.Join(", ", ranges);
                }

                // 7. 組合視覺化報告
                ////StringBuilder sb = new StringBuilder();
                ////sb.AppendLine("=== 頁碼檢查報告 ===");
                ////sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                ////sb.AppendLine($"已修改頁碼範圍：{allPagesForReport.First()} ~ {allPagesForReport.Last()}");
                ////sb.AppendLine();
                ////sb.AppendLine("-----------------------------------");
                ////sb.AppendLine("【已出現的頁碼】(已修改)：");
                ////sb.AppendLine(ToRangeString(allPagesForReport));
                ////sb.AppendLine();
                ////sb.AppendLine("【未出現的頁碼】(待檢查)：");
                ////sb.AppendLine(ToRangeString(missingPages));
                ////sb.AppendLine("-----------------------------------");
                ////sb.AppendLine();
                //StringBuilder sb = new StringBuilder();
                //sb.AppendLine("=== 網頁深度解析報告 (精確防護版) ===");
                //sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                //sb.AppendLine($"掃描範圍：第 {allPagesForReport.First()} ~ {allPagesForReport.Last()} 頁");
                //sb.AppendLine("-----------------------------------");
                //sb.AppendLine("【真正有修改的頁碼】(存入自動檢查機制)：");
                //sb.AppendLine(ToRangeString(realEditedList)); // 這是關鍵，EditedPagesCache 的來源
                //sb.AppendLine();
                //sb.AppendLine("【所有掃描到的頁碼】(供目測參考)：");
                //sb.AppendLine(ToRangeString(allPagesForReport));
                //sb.AppendLine();
                //sb.AppendLine("【未出現的頁碼】(待檢查)：");
                //sb.AppendLine(ToRangeString(missingPages));
                //sb.AppendLine("-----------------------------------");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("★★【真正有修改的頁碼】★★(已同步至自動煞車機制)：");
                sb.AppendLine(ToRangeString(realEditedList));
                sb.AppendLine();
                sb.AppendLine("=== 網頁深度解析報告 (精確防護版) ===");
                sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"掃描範圍：第 {minPage} ~ {maxPage} 頁");
                sb.AppendLine("-----------------------------------");
                sb.AppendLine("★【真正有修改的頁碼】★(已同步至自動煞車機制)：");
                sb.AppendLine(ToRangeString(realEditedList));
                sb.AppendLine();
                sb.AppendLine("【所有掃描到的頁碼】(供目測參考)：");
                sb.AppendLine(ToRangeString(allPagesForReport));
                sb.AppendLine();
                sb.AppendLine("【未出現的頁碼】(掃描區間內缺失)：");
                sb.AppendLine(ToRangeString(missingPages));

                // 8. 逐條列出不連貫明細
                List<string> gaps = new List<string>();
                for (int i = 0; i < allPagesForReport.Count - 1; i++)
                {
                    if (allPagesForReport[i + 1] != allPagesForReport[i] + 1)
                    {
                        gaps.Add($"   - {allPagesForReport[i]} 頁 之後跳到了 {allPagesForReport[i + 1]} 頁");
                    }
                }

                if (gaps.Count > 0)
                {
                    sb.AppendLine("-----------------------------------");
                    sb.AppendLine("⚠️ 發現不連貫處：");
                    foreach (var gap in gaps) sb.AppendLine(gap);
                }
                else
                {
                    sb.AppendLine("✅ 檢查結果：頁碼完全連貫。");
                }
                // --- 在此處加入 orphanModifications 的顯示 ---
                sb.AppendLine("-----------------------------------");
                if (orphanModifications.Count > 0)
                {
                    sb.AppendLine("-----------------------------------");
                    sb.AppendLine("★⚠️ 發現「無頁碼標籤」的修改內容 (請手動核對)：★");
                    foreach (var orphan in orphanModifications)
                    {
                        sb.AppendLine(orphan);
                    }
                }

                // 9. 回存剪貼簿並提示
                Clipboard.SetText(sb.ToString()); // 這樣解析完才能 Ctrl+V 貼出報告  // 加上這行，方便您稍後貼出參考

                ////MessageBox.Show($"網頁解析完成！已同步 {EditedPagesCache.Count} 個已修改頁碼至快取。{Environment.NewLine}報告已複製到剪貼簿。", "自動同步成功");
                //string summary = $"解析完成！{Environment.NewLine}" +
                //     $"總計出現：{allPagesForReport.Count} 頁{Environment.NewLine}" +
                //     $"真正修改：{EditedPagesCache.Count} 頁 (已納入煞車機制)";
                string summary = $"網頁解析完成！{Environment.NewLine}{Environment.NewLine}" +
                     $"● 總計偵測到：{allPagesForReport.Count} 頁{Environment.NewLine}" +
                     $"● 真正修改文字：{EditedPagesCache.Count} 頁 (已納入煞車機制){Environment.NewLine}{Environment.NewLine}" +
                     $"詳細報告已複製到剪貼簿。";
                MessageBoxShowOKExclamationDefaultDesktopOnly(summary, "深度解析成功");

                return true;

                // ======================================================================
                // 【核心解析區域函式】
                // ======================================================================
                void AnalyzeModificationBlocks()
                {
                    // 抓取所有 <td> 區塊
                    var tdBlocks = Regex.Matches(modifiedRightPart, @"<td[^>]*>(.*?)</td>", RegexOptions.Singleline);

                    foreach (Match td in tdBlocks)
                    {
                        string tdContent = td.Groups[1].Value;

                        // 獲取該區塊內所有出現過的頁碼 (供報告統計用)
                        var pagesInTd = Regex.Matches(tdContent, @"page=""(\d+)""");
                        foreach (Match pMatch in pagesInTd)
                        {
                            int pNum = int.Parse(pMatch.Groups[1].Value);
                            if (!allPagesForReport.Contains(pNum)) allPagesForReport.Add(pNum);
                        }

                        // 檢查是否有實際修改 (diffadd/del)
                        var diffs = Regex.Matches(tdContent, @"<span class=""diff(add|del)"">");
                        if (diffs.Count == 0) continue; // 沒修改就跳過

                        // 判斷是否為孤兒片段 (有修改但沒頁碼標籤)
                        if (pagesInTd.Count == 0)
                        {
                            // 提取中文文字，並轉譯 HTML
                            string pureText = Regex.Replace(tdContent, @"<[^>]*>", "");
                            pureText = System.Net.WebUtility.HtmlDecode(pureText).Replace("\n", "").Replace("\r", "").Trim();

                            // 尋找此區塊之前最近的頁碼作為參考點
                            string contextBefore = modifiedRightPart.Substring(0, td.Index);
                            var lastKnownPage = Regex.Match(contextBefore, @"page=""(\d+)""", RegexOptions.RightToLeft);
                            string rangeHint = lastKnownPage.Success ? $"{lastKnownPage.Groups[1].Value} 頁之後" : "開頭處";

                            orphanModifications.Add($"   [範圍：{rangeHint}] 內容：{pureText}");
                        }
                        else
                        {
                            // 正常有頁碼的區塊：執行「由修改點往回找頁碼」的嚴謹邏輯
                            foreach (Match dm in diffs)
                            {
                                string textBeforeDiff = tdContent.Substring(0, dm.Index);
                                var pageBefore = Regex.Match(textBeforeDiff, @"page=""(\d+)""", RegexOptions.RightToLeft);

                                if (pageBefore.Success)
                                {
                                    int p = int.Parse(pageBefore.Groups[1].Value);
                                    if (!EditedPagesCache.Contains(p)) EditedPagesCache.Add(p);
                                }
                                else
                                {
                                    // 若 diff 之前沒頁碼，抓該區塊第一個頁碼
                                    int p = int.Parse(pagesInTd[0].Groups[1].Value);
                                    if (!EditedPagesCache.Contains(p)) EditedPagesCache.Add(p);
                                }
                            }
                        }
                    }
                }
            }

            //https://gemini.google.com/share/a7154355aabe https://gemini.google.com/share/1c579cb86eac 

            /// <summary>
            /// 核心處理方法（整併報告與快取）
            /// 這個方法會一次處理完：產生報告文字、寫入剪貼簿、更新全域快取清單。
            /// </summary>            
            public static void ProcessXmlAndSyncCache()
            {
                // 1. 從剪貼簿取得 XML 文本
                string input = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(input)) return;

                // 2. 抓取 page="數字"
                var matches = System.Text.RegularExpressions.Regex.Matches(input, @"page=""(\d+)""");

                // 清空舊快取（存放已修改頁碼）
                EditedPagesCache.Clear();
                List<int> pages = new List<int>();

                foreach (System.Text.RegularExpressions.Match m in matches)
                {
                    if (int.TryParse(m.Groups[1].Value, out int p))
                    {
                        pages.Add(p);
                        EditedPagesCache.Add(p); // 【重要】這裡存入的是已修改的頁碼
                    }
                }

                if (pages.Count == 0)
                {
                    MessageBox.Show("找不到任何頁碼資訊。", "檢查結束");
                    return;
                }

                // 排序與去重
                pages = pages.Distinct().OrderBy(p => p).ToList();

                // 3. 計算「未出現」的頁碼清單
                List<int> missingPages = new List<int>();
                for (int i = pages.First(); i <= pages.Last(); i++)
                {
                    // 這裡改用 EditedPagesCache 檢查效能更好
                    if (!EditedPagesCache.Contains(i)) missingPages.Add(i);
                }

                // 4. 定義小工具：範圍字串化
                string ToRangeString(List<int> nums)
                {
                    if (nums == null || nums.Count == 0) return "無";
                    var ranges = new List<string>();
                    int start = nums[0];
                    int end = nums[0];
                    for (int i = 1; i <= nums.Count; i++)
                    {
                        if (i < nums.Count && nums[i] == end + 1)
                        {
                            end = nums[i];
                        }
                        else
                        {
                            ranges.Add(start == end ? start.ToString() : $"{start}~{end}");
                            if (i < nums.Count) { start = nums[i]; end = nums[i]; }
                        }
                    }
                    return string.Join(", ", ranges);
                }

                // 5. 組合報告
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("=== 頁碼檢查報告 ===");
                sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"已修改頁碼範圍：{pages.First()} ~ {pages.Last()}");
                sb.AppendLine();
                sb.AppendLine("-----------------------------------");
                sb.AppendLine("【已出現的頁碼】(已修改)：");
                sb.AppendLine(ToRangeString(pages));
                sb.AppendLine();
                sb.AppendLine("【未出現的頁碼】(待檢查)：");
                sb.AppendLine(ToRangeString(missingPages));
                sb.AppendLine("-----------------------------------");
                sb.AppendLine();

                // 6. 逐條列出不連貫明細
                List<string> gaps = new List<string>();
                for (int i = 0; i < pages.Count - 1; i++)
                {
                    if (pages[i + 1] != pages[i] + 1)
                    {
                        gaps.Add($"   - {pages[i]} 頁 之後跳到了 {pages[i + 1]} 頁");
                    }
                }

                if (gaps.Count > 0)
                {
                    sb.AppendLine("-----------------------------------");
                    sb.AppendLine("⚠️ 發現不連貫處：");
                    foreach (var gap in gaps) sb.AppendLine(gap);
                }
                else
                {
                    sb.AppendLine("✅ 檢查結果：頁碼完全連貫。");
                }

                // 7. 回存剪貼簿並提示
                Clipboard.SetText(sb.ToString());
                MessageBox.Show($"報告已產生並同步至系統快取！{Environment.NewLine}現在翻頁時將會自動比對這 {EditedPagesCache.Count} 個頁碼。", "同步成功");
            }

            #region 原未經測試者            
            //public static void ProcessXmlAndSyncCache()
            //{
            //    // 1. 從剪貼簿取得 XML 文本
            //    string input = Clipboard.GetText();
            //    if (string.IsNullOrWhiteSpace(input)) return;

            //    // 2. 抓取 page="數字"
            //    var matches = System.Text.RegularExpressions.Regex.Matches(input, @"page=""(\d+)""");

            //    // 清空舊快取，準備存入新的
            //    EditedPagesCache.Clear();
            //    List<int> pages = new List<int>();

            //    foreach (System.Text.RegularExpressions.Match m in matches)
            //    {
            //        if (int.TryParse(m.Groups[1].Value, out int p))
            //        {
            //            pages.Add(p);
            //            EditedPagesCache.Add(p); // 同步寫入全域快取，供 NextPages 檢查
            //        }
            //    }

            //    //if (pages.Count == 0) return;
            //    if (pages.Count == 0)
            //    {
            //        MessageBox.Show("找不到任何頁碼資訊。", "檢查結束");
            //        return;
            //    }

            //    // 排序與去重（供報告顯示用）
            //    pages = pages.Distinct().OrderBy(p => p).ToList();


            //    // 3. 計算「未出現」的頁碼清單
            //    List<int> missingPages = new List<int>();
            //    for (int i = pages.First(); i <= pages.Last(); i++)
            //    {
            //        if (!pages.Contains(i)) missingPages.Add(i);
            //    }

            //    // 4. 定義小工具：將數字清單轉為 "1~27, 30~37" 格式
            //    string ToRangeString(List<int> nums)
            //    {
            //        if (nums == null || nums.Count == 0) return "無";
            //        var ranges = new List<string>();
            //        int start = nums[0];
            //        int end = nums[0];
            //        for (int i = 1; i <= nums.Count; i++)
            //        {
            //            if (i < nums.Count && nums[i] == end + 1)
            //            {
            //                end = nums[i];
            //            }
            //            else
            //            {
            //                ranges.Add(start == end ? start.ToString() : $"{start}~{end}");
            //                if (i < nums.Count) { start = nums[i]; end = nums[i]; }
            //            }
            //        }
            //        return string.Join(", ", ranges);
            //    }


            //    //// 3. 產生您要的報告文字 (這部分維持昨天的格式)
            //    //StringBuilder sb = new StringBuilder();
            //    //sb.AppendLine("=== 頁碼檢查報告 ===");
            //    //sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            //    //sb.AppendLine($"已修改頁碼範圍：{pages.First()} ~ {pages.Last()}");
            //    //sb.AppendLine($"已載入快取頁數：{EditedPagesCache.Count} 頁");
            //    //sb.AppendLine();
            //    //// ... (中間省略 ToRangeString 的範圍顯示邏輯，請沿用昨天的程式碼) ...
            //    StringBuilder sb = new StringBuilder();
            //    sb.AppendLine("=== 頁碼檢查報告 ===");
            //    sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            //    sb.AppendLine($"已修改頁碼範圍：{pages.First()} ~ {pages.Last()}");
            //    sb.AppendLine();
            //    sb.AppendLine("-----------------------------------");
            //    sb.AppendLine("【已出現的頁碼】(已修改)：");
            //    sb.AppendLine(ToRangeString(pages));
            //    sb.AppendLine();
            //    sb.AppendLine("【未出現的頁碼】(待檢查)：");
            //    sb.AppendLine(ToRangeString(missingPages));
            //    sb.AppendLine("-----------------------------------");
            //    sb.AppendLine();

            //    // 6. 逐條列出不連貫明細
            //    List<string> gaps = new List<string>();
            //    for (int i = 0; i < pages.Count - 1; i++)
            //    {
            //        if (pages[i + 1] != pages[i] + 1)
            //        {
            //            gaps.Add($"   - {pages[i]} 頁 之後跳到了 {pages[i + 1]} 頁");
            //        }
            //    }

            //    if (gaps.Count > 0)
            //    {
            //        sb.AppendLine("-----------------------------------");
            //        sb.AppendLine("⚠️ 發現不連貫處：");
            //        foreach (var gap in gaps)
            //        {
            //            sb.AppendLine(gap);
            //        }
            //    }
            //    else
            //    {
            //        sb.AppendLine("✅ 檢查結果：頁碼完全連貫。");
            //    }


            //    // 4. 將報告回存剪貼簿，並提示成功
            //    Clipboard.SetText(sb.ToString());
            //    MessageBox.Show($"報告已產生並同步至系統快取！{Environment.NewLine}現在翻頁時將會自動比對這 {EditedPagesCache.Count} 個頁碼。", "同步成功");
            //}
            #endregion

            /// <summary>
            /// 更新/儲存/記錄已經編輯過的頁碼快取。由XML中的「page="數字"」標記擷取。
            /// C# XML 頁碼連貫性檢查機制
            /// </summary>
            public static void UpdateEditedPagesCacheFromClipboard(string input)//https://gemini.google.com/share/1064e057a6f8 20260115:C# XML 頁碼連貫性檢查
            {
                //string input = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(input)) return;

                // 抓取 page="數字"
                var matches = System.Text.RegularExpressions.Regex.Matches(input, @"page=""(\d+)""");

                EditedPagesCache.Clear();
                foreach (System.Text.RegularExpressions.Match m in matches)
                {
                    if (int.TryParse(m.Groups[1].Value, out int p))
                        EditedPagesCache.Add(p);
                }
            }

            //https://gemini.google.com/share/6972029e0e7a C# XML 頁碼連貫性檢查 https://gemini.google.com/share/73d33e77542b

            /// <summary>
            /// 匯出已修改之頁碼報告到剪貼簿
            /// 從剪貼簿讀取 XML 內容，分析 page="數字" 標記
            /// 由網址含有「&action=diff」這樣的頁面複製而來
            /// </summary>
            public static void PrintXmlModifiedPages()
            {
                // 1. 從剪貼簿讀取文本 由網址含有「&action=diff」這樣的頁面複製而來，如： https://ctext.org/wiki.pl?if=gb&action=diff&to=1419411&from=743797
                string input = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(input))
                {
                    MessageBox.Show("剪貼簿裡沒有文字內容喔！", "提示");
                    return;
                }

                // 2. 抓取頁碼 (page="數字")
                var matches = Regex.Matches(input, @"page=""(\d+)""");
                List<int> pages = matches.Cast<Match>()
                                         .Select(m => int.Parse(m.Groups[1].Value))
                                         .Distinct()
                                         .OrderBy(p => p)
                                         .ToList();

                if (pages.Count == 0)
                {
                    MessageBox.Show("找不到任何頁碼資訊。", "檢查結束");
                    return;
                }

                // 3. 計算「未出現」的頁碼清單
                List<int> missingPages = new List<int>();
                for (int i = pages.First(); i <= pages.Last(); i++)
                {
                    if (!pages.Contains(i)) missingPages.Add(i);
                }

                // 4. 定義小工具：將數字清單轉為 "1~27, 30~37" 格式
                string ToRangeString(List<int> nums)
                {
                    if (nums == null || nums.Count == 0) return "無";
                    var ranges = new List<string>();
                    int start = nums[0];
                    int end = nums[0];
                    for (int i = 1; i <= nums.Count; i++)
                    {
                        if (i < nums.Count && nums[i] == end + 1)
                        {
                            end = nums[i];
                        }
                        else
                        {
                            ranges.Add(start == end ? start.ToString() : $"{start}~{end}");
                            if (i < nums.Count) { start = nums[i]; end = nums[i]; }
                        }
                    }
                    return string.Join(", ", ranges);
                }

                // 5. 組合結果字串
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("=== 頁碼檢查報告 ===");
                sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"已修改頁碼範圍：{pages.First()} ~ {pages.Last()}");
                sb.AppendLine();
                sb.AppendLine("-----------------------------------");
                sb.AppendLine("【已出現的頁碼】(已修改)：");
                sb.AppendLine(ToRangeString(pages));
                sb.AppendLine();
                sb.AppendLine("【未出現的頁碼】(待檢查)：");
                sb.AppendLine(ToRangeString(missingPages));
                sb.AppendLine("-----------------------------------");
                sb.AppendLine();

                // 6. 逐條列出不連貫明細
                List<string> gaps = new List<string>();
                for (int i = 0; i < pages.Count - 1; i++)
                {
                    if (pages[i + 1] != pages[i] + 1)
                    {
                        gaps.Add($"   - {pages[i]} 頁 之後跳到了 {pages[i + 1]} 頁");
                    }
                }

                if (gaps.Count > 0)
                {
                    sb.AppendLine("-----------------------------------");
                    sb.AppendLine("⚠️ 發現不連貫處：");
                    foreach (var gap in gaps)
                    {
                        sb.AppendLine(gap);
                    }
                }
                else
                {
                    sb.AppendLine("✅ 檢查結果：頁碼完全連貫。");
                }

                // 7. 回存剪貼簿並提示
                Clipboard.SetText(sb.ToString());

                string status = gaps.Count > 0 ? $"發現 {gaps.Count} 處不連貫！" : "頁碼完全連貫！";
                MessageBox.Show($"{status}\n報告已存入剪貼簿，請直接貼上使用。", "處理完成");
            }
            public static void PrintXmlModifiedPages_OLD()
            {
                // 1. 從剪貼簿取得文字
                string input = Clipboard.GetText();

                if (string.IsNullOrWhiteSpace(input))
                {
                    MessageBox.Show("剪貼簿裡沒有文字內容喔！", "提示");
                    return;
                }

                // 2. 使用正則表達式抓取 page="數字"
                // 根據您提供的 v26.txt，頁碼格式為 page="數字"
                var matches = Regex.Matches(input, @"page=""(\d+)""");

                // 轉為數字列表、去重、排序
                List<int> pages = matches.Cast<Match>()
                                         .Select(m => int.Parse(m.Groups[1].Value))
                                         .Distinct()
                                         .OrderBy(p => p)
                                         .ToList();

                if (pages.Count == 0)
                {
                    MessageBox.Show("在剪貼簿內容中找不到任何頁碼標籤 (page=\"...\")。", "檢查結束");
                    return;
                }

                // 3. 檢查連貫性並整理結果文字
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine("=== 頁碼檢查報告 ===");
                sb.AppendLine($"處理時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"已修改頁碼範圍：{pages.First()} ~ {pages.Last()}");
                sb.AppendLine($"已出現的頁碼：{string.Join(", ", pages)}");
                sb.AppendLine("-----------------------------------");

                List<string> gaps = new List<string>();
                for (int i = 0; i < pages.Count - 1; i++)
                {
                    if (pages[i + 1] != pages[i] + 1)
                    {
                        gaps.Add($"{pages[i]} 頁 之後跳到了 {pages[i + 1]} 頁");
                    }
                }

                if (gaps.Count > 0)
                {
                    sb.AppendLine("⚠️ 發現不連貫處：");
                    foreach (var gap in gaps)
                    {
                        sb.AppendLine("   - " + gap);
                    }
                }
                else
                {
                    sb.AppendLine("✅ 檢查結果：頁碼完全連貫。");
                }

                // 4. 將結果回存到剪貼簿
                string finalResult = sb.ToString();
                Clipboard.SetText(finalResult);

                // 5. 提示使用者
                string briefSummary = gaps.Count > 0
                    ? $"發現 {gaps.Count} 處不連貫！"
                    : "頁碼非常連貫！";

                MessageBox.Show($"{briefSummary}\n\n詳細報告已存入剪貼簿，您可以直接貼到記事本查看。", "處理完成");
            }
        }
        public static class XmlProcessor
        {


            /// <summary>
            /// 根據現有的文本meta data，擷取其內容提交新文本，自動轉到對應的全書首頁，並開啟新增章節頁面，完成全書文本的新增。20260116
            /// create a new entry 見： https://ctext.org/wiki.pl?if=en
            /// </summary>
            /// <returns>失敗則傳回false</returns>
            internal static bool SubmitAnotherText_NewPage_Auto_action_newchapter_create_a_new_entry(string urlEditTextMetadata = "")
            {//由WordVBA「新頁面Auto_action_newchapter」其原理改寫而來//https://ctext.org/wiki.pl?if=en&action=new

                if (IsDriverInvalid) return false;
                #region 取得現有資源的meta data，以填入新文本的欄位
                //先到【修改原典後設資料】頁面，以取得現有資源的meta data
                //如果在編輯現有資源的「編輯文本資料」頁面，就先記錄其網址，以取得其meta data
                if (CtextPageClassifier.ParseUrl(driver.Url)?.PageType == CtextPageType.EditTextMetadata)
                {
                    if (urlEditTextMetadata.IsNullOrEmpty()) urlEditTextMetadata = driver.Url;
                }
                else if (!urlEditTextMetadata.IsNullOrEmpty() && driver.Url != urlEditTextMetadata)
                    driver.Url = urlEditTextMetadata;//轉到現有資源頁面，以取得其meta data
                                                     //如果不是在【修改原典後設資料】頁面，就不做
                if (CtextPageClassifier.ParseUrl(driver.Url)?.PageType != CtextPageType.EditTextMetadata) return false;
                //修改tag：
                string tags = Tags_textBox.GetDomProperty("value");
                string[] tagArray = tags.Split(new char[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                tags = string.Join(", ", tagArray.Where(t => t != "OCR_MATCH"));
                SetIWebElementValueProperty(Tags_textBox, tags);
                //讀取現有資源頁面的meta data 並複製到新文本頁面//如此頁：https://ctext.org/wiki.pl?if=gb&res=53207&action=edit
                //Dictionary<string,string>  metaDateDict=new Dictionary<string, string>();
                //metaDateDict.Add("Title", Title_textBox_editTextMetadata.GetDomProperty("value"));
                //metaDateDict.Add("Author", Author_textBox_editTextMetadata.GetDomProperty("value"));
                //metaDateDict.Add("Dynasty", Dynasty_selectListBox_editTextMetadata.GetDomProperty("value"));                
                //metaDateDict.Add("Base text:OtherEdition", OtherEdition_textBox_editTextMetadata.GetDomProperty("value"));
                //metaDateDict.Add("Alias", Alias_textBox_editTextMetadata.GetDomProperty("value"));
                //metaDateDict.Add("Description", Description_textBox_editTextMetadata.GetDomProperty("value"));
                //讀取現有資源頁面的meta data 並複製到dto物件
                var dto = ReadFromEditPage();//https://copilot.microsoft.com/shares/ryPqtnoS2Zg5zx3HfsEL5
                                             //取得title（書名）值：
                string title = dto.Title;//Title_textBox.GetDomProperty("value");
                #endregion

                //開新視窗頁籤到書籍首頁以取得第1冊的標題、首、末頁碼與file值
                //取得第1冊（現要處理之冊）的標題
                LastValidWindow = driver.CurrentWindowHandle;
                string editTextMetadataWindowHandle = LastValidWindow;
                Browser.OpenNewTabWindow();
                driver.Url = urlEditTextMetadata;
                if (!GotoBookHomepage(dto.OtherEdition)) return false;
                //未必從第1冊開始，但通常是
                //string titleFile = FirstFileItem_td_linkbox_BookHomepage.GetDomProperty("textContent");
                //取得現要處理之冊的標題（通常是第1冊，但未必，若被輸入驗證碼中斷或當機，可以從 CurrentFileSelector 取得現在要處理的冊數，此值亦可在textBox2中以「fn」前綴指定；若第3冊則為「fn3」）
                string titleFile = WaitFindWebElementBySelector_ToBeClickable(CurrentFileSelector).Text;

                //取得第1冊的首頁碼、末頁碼、file值
                if (!GotoFirstFile(dto.OtherEdition)) return false;
                CtextPageInfo pi = CtextPageClassifier.ParseUrl(driver.Url);
                int firstPage = (int)pi.PageNumber, lastPage = PageUBound, file = (int)pi.FileId;



                #region 轉入上傳新資料「提交新文本」頁面
                driver.Url = "https://ctext.org/wiki.pl?if=en&action=new";
                //如果不是上傳新資料「提交新文本」頁面，就不做
                if (CtextPageClassifier.ParseUrl(driver.Url)?.PageType != CtextPageType.SubmitNewText) return false;

                //將DTO物件的內容寫入上傳新資料「提交新文本」頁面
                WriteToNewTextPage(dto);
                SetIWebElementValueProperty(Tags_textBox, "OCR_MATCH, OCR_PRIMARY, OCR_CORRECTED(71)");
                SetIWebElementValueProperty(Description_Edit_textbox, GetDescription("原電子文本只是方便版，並非以原書圖為底本，故茲據"
                                        , "原電子文本只是方便版，並非以原書圖為底本，故據網路所得本輔以末學自製於GitHub開源免費免安裝之TextForCtext軟件排版對應錄入；討論區及末學YouTube頻道有實境演示影片。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主"));

                //如果「著作名稱:」（title）欄位是空的，就不做 ∵Title is required; leave all other fields blank if not known or not applicable
                if (Title_textBox.GetDomProperty("value").IsNullOrEmpty()) return false;
                //如果「其他版本」欄位是空的，就不做
                if (OtherEdition_textBox.GetDomProperty("value").IsNullOrEmpty()) return false;
                //先產生第一冊新頁面的XML標記字串
                //textBox1中第1行是首頁碼，第2行是末頁碼，第3行file值
                //var paras = Form1.InstanceForm1._document.GetParagraphs();
                //if (!int.TryParse(paras.First().Text, out int firstPage)) return msgboxParamErr();
                //if (!int.TryParse(paras[1].Text, out int lastPage)) return msgboxParamErr();
                //if (!int.TryParse(paras[2].Text, out int file)) return msgboxParamErr();
                //Console.WriteLine(ScanXmlGenerator.GenerateScanXml(firstPage, lastPage, file));
                string xml_newtext = ScanTagGenerator.GenerateXmlString(firstPage, lastPage, file);
                titleFile = GetTitleFile(titleFile, title);
                xml_newtext = "*" + titleFile + Environment.NewLine + xml_newtext;
                SetIWebElementValueProperty(XMLData_textarea_submitNewText, xml_newtext);
                //按下「檢查」按鈕
                Analyze_button_submitNewText?.JsClick();
                //按下「上傳資料」按鈕，提交新文本
                CreateResource_button_submitNewText?.JsClick();
                #endregion

                //轉到剛創建新資源的網頁，記下其網址
                string urlNewRes = driver.Url;
                //新增文本成功後，轉到新文本的編輯頁面，以取得其resource ID
                CtextPageInfo newTextPageInfo = CtextPageClassifier.ParseUrl(urlNewRes);//=driver.Url;
                if (newTextPageInfo == null || newTextPageInfo.PageType != CtextPageType.WikiResource)
                    return false;
                //取得新的resource ID
                int resID_NewText = (int)newTextPageInfo.ResId;
                LastValidWindow = driver.CurrentWindowHandle;
                //更新原先的【修改原典後設資料】頁面內容
                driver.SwitchTo().Window(editTextMetadataWindowHandle);
                SetIWebElementValueProperty(Tags_textBox,
                    tags.IsNullOrEmpty() ? "WORKSET(ctp:wb" + resID_NewText.ToString() + ")" :
                    tags + ", WORKSET(ctp:wb" + resID_NewText.ToString() + ")");//設定tag內容
                SetIWebElementValueProperty(Description_textBox, "原電子文本只是方便版，並非以原書圖為底本，故另建新的維基項目"
                    + " ctp:wb" + resID_NewText.ToString() + " "
                    + "與之對應，暫撤去「OCR_MATCH」標籤以免妨礙輸入。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");//設定description內容
                Submit_button_EditTextMetadata?.JsClick();//按下「保存」按鈕，提交修改
                driver.SwitchTo().Window(LastValidWindow);//回到新文本的編輯頁面

                #region 逐冊新增section [Add new section]
                int seq = 1;
                while (true)
                {
                    //到本書首頁
                    GotoBookHomepage();
                    //取得下一冊的連結元件
                    IWebElement iweCurrentFile = WaitFindWebElementBySelector_ToBeClickable(NextFileSelector);
                    if (iweCurrentFile == null) break;//如果沒有下一冊，就跳出迴圈結束程序
                    //取得下一冊的標題
                    titleFile = GetTitleFile(iweCurrentFile.Text, title);

                    //到下一冊                
                    iweCurrentFile.JsClick();
                    //前面已用過NextFileSelector了！
                    //if (GotoNextFile()) break;//如果沒有下一冊，就跳出迴圈結束程序

                    //在書籍首頁取得下一冊的始、末頁與file值
                    pi = CtextPageClassifier.ParseUrl(driver.Url);
                    firstPage = (int)pi.PageNumber; lastPage = PageUBound; file = (int)pi.FileId;
                    //由以上取得的3個參數以構建要輸入的新xml內容
                    xml_newtext = ScanTagGenerator.GenerateXmlString(firstPage, lastPage, file);

                    //新增章節頁面//[Add new section]                    
                    driver.Url = $"https://ctext.org/wiki.pl?if=en&res={resID_NewText}&action=newchapter";
                    //標題／篇名: Title:
                    SetIWebElementValueProperty(Title_textBox, titleFile);
                    //序號
                    SetIWebElementValueProperty(Sequence_Edit_Chapter, (++seq).ToString() + "00");
                    //content:內容:
                    SetIWebElementValueProperty(Textarea_data_Edit_textbox, xml_newtext);

                    SetIWebElementValueProperty(Description_Edit_textbox, GetDescription("原電子文本只是方便版，並非以原書圖為底本，故茲據"
                        , "原電子文本只是方便版，並非以原書圖為底本，故據網路所得本輔以末學自製於GitHub開源免費免安裝之TextForCtext軟件排版對應錄入；討論區及末學YouTube頻道有實境演示影片。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主"));
                    if (Commit?.JsClick() == false) break;//按下「保存」按鈕，提交新增章節

                    Thread.Sleep(2000);//休息一下 ：） 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
                }

                #endregion

                return true;
            }

            //bool msgboxParamErr()
            //{
            //    MessageBoxShowOKExclamationDefaultDesktopOnly("textBox1中第1行/段必須是第1冊的首頁碼，第2行是其末頁碼，第3行為第1冊的file值！感恩感恩　南無阿彌陀佛");
            //    return false;
            //}

            /// <summary>
            /// 取得titleFile（書名標題）以輸入新增section的 newchapter 「標題／篇名:」欄位
            /// </summary>
            /// <param name="titleFile">冊名（冊標題）</param>
            /// <param name="title">書名</param>
            /// <returns>傳回冊標名以供新增section時「Title:」（標題／篇名:）欄位用</returns>
            private static string GetTitleFile(string titleFile, string title)
            {
                titleFile = titleFile.Replace(title, string.Empty);
                titleFile = titleFile.Trim();
                titleFile = titleFile.Substring(0, 1) == "·" ? titleFile.Substring(1) : titleFile;//去掉開頭的「·」
                return titleFile;
            }





            /// <summary>
            /// 在Xml最後補上末頁標記
            /// Ctrl + Shift + Alt + l：在XML文末加上末頁的標記 （l:last page）
            /// </summary>
            /// <returns></returns>
            public static string AppendLastPage()
            {
                //public string AppendLastPage(string input, int uboundPageNum)
                if (PageNum_textbox == null) return null;
                int uboundPageNum = PageUBound;

                if (Edit_Linkbox_ImageTextComparisonPage == null) return null;
                Edit_Linkbox_ImageTextComparisonPage.JsClick();
                if (Textarea_data_Edit_textbox == null) return null;
                string input = Textarea_data_Edit_textboxTxt;


                // 1. 尋找最後一個 <scanend /> 標記的位置
                // 我們使用正則表達式來抓取最後一個標記，並同時提取其中的 file 屬性值
                string pattern = @"<scanend\s+file=""(?<file>\d+)""\s+page=""(?<page>\d+)""\s*/>";
                MatchCollection matches = Regex.Matches(input, pattern);

                if (matches.Count == 0) return input; // 如果沒找到標記，回傳原字串

                Match lastMatch = matches[matches.Count - 1]; // 取得最後一個匹配項
                int lastIndex = lastMatch.Index;
                string fileId = lastMatch.Groups["file"].Value; // 提取 file 變數，例如 "238685"

                // 2. 在最後一個 <scanend 之前插入兩個換行符 (\r\n\r\n)
                // 注意：這裡使用 Insert，會將原本的內容往後推
                string result = input.Insert(lastIndex, "\r\n\r\n").TrimEnd();//Leo AI 大菩薩 20260114

                // 3. 在整個字串末尾追加新的標記
                // 格式：<scanbegin file="xxx" page="uboundPageNum" />●<scanend file="xxx" page="uboundPageNum" />
                string newTags = $@"<scanbegin file=""{fileId}"" page=""{uboundPageNum}"" />●<scanend file=""{fileId}"" page=""{uboundPageNum}"" />";

                result += newTags;


                SetIWebElementValueProperty(Textarea_data_Edit_textbox, result);
                SetIWebElementValueProperty(Description_Edit_textbox, "加入末頁XML標記-以末學自製於GitHub開源免費免安裝之應用程式TextForCtext自動化加入。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");
                if (Commit == null) return null;
                Commit.JsClick();
                return result;
            }
        }

        ///======================================================================
        //https://copilot.microsoft.com/shares/kxu8fZCszjF9XpnpJvHRa 使用C#生成XML標記字串

        /// <summary>
        ///  物件導向風格 (OOP) 的設計。這樣不只是生成字串，而是把「掃描標記」抽象成一個類別，未來可以更容易擴充、序列化或轉換成其他格式。
        /// </summary>
        public class ScanTag
        {
            public int File { get; set; }
            public int Page { get; set; }

            public XElement Begin() => new XElement("scanbegin",
                new XAttribute("file", File),
                new XAttribute("page", Page));

            public XElement End() => new XElement("scanend",
                new XAttribute("file", File),
                new XAttribute("page", Page));
        }

        public static class ScanTagGenerator
        {
            public static IEnumerable<ScanTag> Create(int firstPage, int lastPage, int file) =>
                Enumerable.Range(firstPage, lastPage - firstPage + 1)
                          .Select(p => new ScanTag { File = file, Page = p });

            public static string GenerateXmlString(int firstPage, int lastPage, int file)
            {
                var tags = Create(firstPage, lastPage, file).ToList();
                return string.Concat(tags.Select((tag, idx) =>
                {
                    var sep = idx == 0 ? "●\t" : "\t";
                    return $"{tag.Begin()}{sep}{tag.End()}";
                }));
            }
        }

    }
}
