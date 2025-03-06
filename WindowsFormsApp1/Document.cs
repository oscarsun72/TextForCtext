using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media.TextFormatting;
using WindowsFormsApp1;

/* 20250124年假前最後上班日 GitHub Copilot大菩薩：
 * 了解了您的需求，我們可以創建一個 Document 類別來表示 textBox1，並創建一個 Paragraph 類別來表示段落。這些類別將能夠操作 textBox1 中的內容，並確保更改段落的文字後，textBox1 的文字也能對應正確地更新。
    首先，我們創建一個 Document 類別，該類別將包含 textBox1 並提供操作其內容的方法。然後，我們創建一個 Paragraph 類別，該類別將能夠操作 Document 中的段落。
    ……
 * 這樣，我們就創建了 Document 和 Paragraph 類別。Document 類別表示 textBox1，並提供操作其內容的方法。Paragraph 類別表示單個段落，並能夠操作 Document 中的段落。
    您可以在 Form1 中創建 Document 和 Paragraph 對象，並使用它們來操作 textBox1 中的內容。例如：
    private void button1_Click(object sender, EventArgs e)
    {
        var paragraphs = _document.GetParagraphs();
        if (paragraphs.Count > 0)
        {
            paragraphs[0].Text = "新的段落文本";
        }
    }
 * 這樣，當您點擊按鈕時，第一個段落的文本將被更改，並且 textBox1 中的內容也會相應更新。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
 */
namespace TextForCtext
{
    /// <summary>
    /// 表示文檔（對應於 textBox1）
    /// </summary>
    public class Document : IDisposable
    {
        private TextBox _textBox;
        private int currentParagraphIndex;
        private List<Paragraph> _cachedParagraphs; // 緩存段落集合
        /// <summary>
        /// 適用於前端介面操作
        /// </summary>
        /// <param name="textBox">要操作的介面textBox</param>
        public Document(ref TextBox textBox)
        {
            _textBox = textBox;
            Text = textBox.Text;
            Content = Range(0, textBox.Text.Length);
            _cachedParagraphs = null;
        }
        /// <summary>
        /// 後端伺服計算適用 20250225
        /// </summary>
        /// <param name="context">要計算的文字內容</param>
        public Document(ref string context)
        {
            _textBox = new TextBox();
            Text = context;//End = context.Length;已在Text屬性內設置
            if (Content == null)
                Content = Range(0, context.Length);
            _cachedParagraphs = null;
        }

        public Range Content { get; set; }
        public string Text
        {
            get => _textBox.Text;
            set
            {
                _textBox.Text = value;
                End = Start + value.Length;
                Content = Range(0, value.Length);
                _cachedParagraphs = null; // 文本改變時清空緩存

            }
        }

        public int SelectionStart
        {
            get => _textBox.SelectionStart;
            set => _textBox.SelectionStart = value;
        }

        public int SelectionLength
        {
            get => _textBox.SelectionLength;
            set => _textBox.SelectionLength = value;
        }

        //public int Start { get;private  set; }
        public int Start { get => 0; }
        public int End { get; private set; }

        internal int CurrentParagraphIndex { get => currentParagraphIndex; set => currentParagraphIndex = value; }

        public void InsertText(string text)
        {
            int selectionStart = _textBox.SelectionStart;
            _textBox.Text = _textBox.Text.Insert(selectionStart, text);
            _textBox.SelectionStart = selectionStart + text.Length;
            _cachedParagraphs = null; // 文本改變時清空緩存
        }

        public void ReplaceText(string text)
        {
            int selectionStart = _textBox.SelectionStart;
            _textBox.Text = _textBox.Text.Remove(selectionStart, _textBox.SelectionLength);
            _textBox.Text = _textBox.Text.Insert(selectionStart, text);
            _textBox.SelectionStart = selectionStart + text.Length;
            _cachedParagraphs = null; // 文本改變時清空緩存
        }

        public List<Paragraph> GetParagraphs(bool newOne = false, Range range = null)
        {
            ////在9百多萬字元、15萬多的行/段長時，有沒有以下判斷式效能差不多 20250225 若發現慢了再將此條件式還原
            if (_cachedParagraphs == null || newOne || !IsTextEqualToCachedParagraphs())
            {
                var paragraphs = new List<Paragraph>();
                var lines = _textBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                int start = 0;
                for (int i = 0; i < lines.Length; i++)
                {
                    var line = lines[i];
                    start = this.Text.IndexOf(line, start);

                    if (start == -1)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("有問題，請檢查！");
                        return null;
                    }

                    //paragraphs.Add(new Paragraph(line, this,new Range(this,Start,End), start, i));
                    paragraphs.Add(new Paragraph(line, this, range ?? Content, start, i));

                    start += (line.Length == 0 ? 1 : line.Length);
                    start = start > this.Text.Length ? this.Text.Length : start;
                }

                _cachedParagraphs = paragraphs; // 更新緩存
            }

            return _cachedParagraphs;
        }


        private bool IsTextEqualToCachedParagraphs()
        {
            var lines = _textBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            if (lines.Length != _cachedParagraphs.Count)
            {
                return false;
            }

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Length != _cachedParagraphs[i].Text.Length
                    || !lines[i].Equals(_cachedParagraphs[i].Text))
                {
                    return false;
                }
            }

            return true;
        }

        #region 取得插入點所在位置的段落，以及其上個段落和下個段落。
        public Paragraph GetCurrentParagraph()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount)
                {
                    return paragraphs[i];
                }
            }

            return null;
        }

        public Paragraph GetNextParagraph()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount && i + 1 < paragraphs.Count)
                {
                    return paragraphs[i + 1];
                }
            }

            return null;
        }

        public Paragraph GetPreviousParagraph()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount && i - 1 >= 0)
                {
                    return paragraphs[i - 1];
                }
            }

            return null;
        }
        #endregion

        public void UpdateParagraphFirstCharacter(int paragraphIndex, string newChar)
        {
            var paragraphs = GetParagraphs();
            if (paragraphs.Count > paragraphIndex)
            {
                var paragraph = paragraphs[paragraphIndex];
                if (paragraph.Text.Length > 0)
                {
                    if (char.IsHighSurrogate(paragraph.Text[0]) && paragraph.Text.Length > 1)
                    {
                        paragraph.Text = newChar + paragraph.Text.Substring(2);
                    }
                    else
                    {
                        paragraph.Text = newChar + paragraph.Text.Substring(1);
                    }
                }
            }
        }

        public void MergeParagraphsAtCaret()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int currentParagraphIndex = 0;
            int charCount = 0;
            int s = _textBox.SelectionStart;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount)
                {
                    currentParagraphIndex = i;
                    break;
                }
            }

            if (currentParagraphIndex + 2 >= paragraphs.Count)
            {
                throw new InvalidOperationException("沒有足夠的段落進行操作。");
            }

            var currentParagraph = paragraphs[currentParagraphIndex];
            var nextParagraph1 = paragraphs[currentParagraphIndex + 1];
            var nextParagraph2 = paragraphs[currentParagraphIndex + 2];

            string mergedText = "{{" + nextParagraph1.Text.Replace(Environment.NewLine, "") + nextParagraph2.Text.Replace(Environment.NewLine, "") + "}}";

            currentParagraph.Text += mergedText;

            paragraphs.RemoveAt(currentParagraphIndex + 2);
            paragraphs.RemoveAt(currentParagraphIndex + 1);

            _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));

            if (_textBox.Text.IndexOf(Environment.NewLine, s) > -1) s = _textBox.Text.IndexOf(Environment.NewLine, s);
            else s = _textBox.TextLength;

            if (currentParagraphIndex + 1 < paragraphs.Count)
            {
                var nextParagraph = paragraphs[currentParagraphIndex + 1];
                currentParagraph.Text += nextParagraph.Text.Replace(Environment.NewLine, "");
                paragraphs.RemoveAt(currentParagraphIndex + 1);
                _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));
            }
            _textBox.SelectionStart = s;
            _textBox.ScrollToCaret();
        }

        public void MergeParagraphsAtCaretWithShift()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int currentParagraphIndex = 0;
            int charCount = 0;
            int s = _textBox.SelectionStart;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount)
                {
                    currentParagraphIndex = i;
                    break;
                }
            }

            if (currentParagraphIndex + 2 >= paragraphs.Count)
            {
                throw new InvalidOperationException("沒有足夠的段落進行操作。");
            }

            var currentParagraph = paragraphs[currentParagraphIndex];
            var nextParagraph1 = paragraphs[currentParagraphIndex + 1];
            var nextParagraph2 = paragraphs[currentParagraphIndex + 2];

            string mergedText = "{{" + currentParagraph.Text.Replace(Environment.NewLine, "") + nextParagraph1.Text.Replace(Environment.NewLine, "") + "}}";

            currentParagraph.Text = mergedText;

            paragraphs.RemoveAt(currentParagraphIndex + 1);

            _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));

            if (currentParagraphIndex + 1 < paragraphs.Count)
            {
                var nextParagraph = paragraphs[currentParagraphIndex + 1];
                currentParagraph.Text += nextParagraph.Text.Replace(Environment.NewLine, "");
                paragraphs.RemoveAt(currentParagraphIndex + 1);
                _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));
            }

            _textBox.SelectionStart = _textBox.Text.IndexOf("}}", s) + 2;
            _textBox.ScrollToCaret();
        }

        public void AutoMarkTitles(int titleIndentationSpaces)
        {
            var paragraphs = GetParagraphs();
            var sb = new StringBuilder();
            bool inTitle = false;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                var text = paragraph.Text;
                int leadingSpaces = text.TakeWhile(c => c == '　').Count();

                if (leadingSpaces == titleIndentationSpaces)
                {
                    sb.Append(text.Insert(leadingSpaces, "*"));
                    for (int j = leadingSpaces + 1; j < text.Length; j++)
                    {
                        if (text[j] == '　')
                        {
                            sb.Append("􏿽");
                        }
                        else
                        {
                            sb.Append(text[j]);
                        }
                    }
                    inTitle = true;
                }
                else if (leadingSpaces < titleIndentationSpaces)
                {
                    for (int j = 0; j < text.Length; j++)
                    {
                        if (text[j] == '　')
                        {
                            sb.Append("􏿽");
                        }
                        else
                        {
                            sb.Append(text[j]);
                        }
                    }
                    sb.Append("<p>");
                    inTitle = false;
                }
                else
                {
                    sb.Append(text.Insert(leadingSpaces, "**"));
                    for (int j = leadingSpaces + 2; j < text.Length; j++)
                    {
                        if (text[j] == '　')
                        {
                            sb.Append("􏿽");
                        }
                        else
                        {
                            sb.Append(text[j]);
                        }
                    }
                    inTitle = true;
                }

                if (inTitle && (i + 1 >= paragraphs.Count || !paragraphs[i + 1].Text.StartsWith("　")))
                {
                    sb.Append("<p>");
                    inTitle = false;
                }

                sb.Append(Environment.NewLine);
            }

            _textBox.Text = sb.ToString();
        }

        public Range Range(int start = 0, int end = 0)
        {
            if (end == 0)
                return new Range(this, Start, End, this.Content);
            else
                return new Range(this, start, end, this.Content);
        }

        public void Dispose()
        {
            _textBox?.Dispose();
        }

        public List<Paragraph> GetSurroundingParagraphs(int index, int range = 3)
        {
            if (_cachedParagraphs == null || _textBox.Text != string.Join(Environment.NewLine, _cachedParagraphs.Select(p => p.Text)))
            {
                _cachedParagraphs = GetParagraphs();
            }

            var paragraphs = _cachedParagraphs;
            if (paragraphs == null || index < 0 || index >= paragraphs.Count)
            {
                return null;
            }

            int start = Math.Max(0, index - range);
            int end = Math.Min(paragraphs.Count - 1, index + range);

            return paragraphs.GetRange(start, end - start + 1);
        }

        public bool CheckAndRefreshParagraphs()
        {
            if (_cachedParagraphs == null || _textBox.Text != string.Join(Environment.NewLine, _cachedParagraphs.Select(p => p.Text)))
            {
                _cachedParagraphs = GetParagraphs();
                return true;
            }
            return false;
        }
    }


    /// <summary>
    /// 表示單個段落
    /// </summary>
    public class Paragraph
    {
        private Document _document;
        private string _text;
        private readonly string _text_beforeUpdate;
        private int _start;
        private readonly int _start_beforeUpdate;
        private int _end;
        //private readonly int _end_beforeUpdate;
        /// <summary>
        /// 段落的Range
        /// </summary>
        private Range _range;
        /// <summary>
        /// 段落集合上一層的Range
        /// </summary>
        private Range _parent;
        private int _index; // 新增的屬性
        /// <summary>
        /// 以documnent建立的段落集合
        /// </summary>
        /// <param name="text"></param>
        /// <param name="document"></param>
        /// <param name="start"></param>
        /// <param name="index"></param>
        public Paragraph(string text, Document document, Range range, int start, int index)
        {
            _text = text;
            _text_beforeUpdate = _text;
            _document = document;
            _parent = range;
            _start = start;
            _start_beforeUpdate = _start;
            //_end_beforeUpdate = _end;
            _end = _start + _text.Length;
            _index = index; // 初始化屬性
            //_range = document.Range(_start, _end);//此會造成overflow
            //UpdateDocumentRange();
        }
        ///// <summary>
        ///// 以range建立的段落集合
        ///// </summary>
        ///// <param name="text"></param>
        ///// <param name="range"></param>
        ///// <param name="start"></param>
        ///// <param name="index"></param>
        //public Paragraph(string text, Range range, int start, int index)
        //{
        //    _text = text;
        //    _text_beforeUpdate = _text;
        //    _Range = range;
        //    _document = range.Document;
        //    _start = start;
        //    _start_beforeUpdate = _start;
        //    _end = _start + _text.Length;
        //    _index = index; // 初始化屬性
        //    //UpdateDocumentRange();
        //}

        public string Text
        {
            get => _text;
            set
            {
                _text = value;
                _end = _start + _text.Length;//_start 是在建置或更新集合 Paragraphs時傳入的。即 public List<Paragraph> GetParagraphs(bool newOne = false) 中的 paragraphs.Add(new Paragraph(line, this, start, i));
                UpdateDocument();
                UpdateParentRange();//●●●●●●●●●●●●●
                //UpdateDocumentRange();//在UpdateDocument();中已有 _document.Text = …… 會更新
            }
        }

        public Range Parent { get => _parent; }

        /// <summary>
        /// 將段落所在的Range更新
        /// 如是則此Range即可與Paragraphs的內容一致，同時異動、連動。
        /// </summary>
        private void UpdateParentRange()
        {
            if (_document != null)
            {

                //var range = _document.Range(_start, _end);
                //range.Text = _text;
                //●●●●●●●●●●●
                _parent.Start = Math.Min(_parent.Start, _start);

                //_parent.End = Math.Max(_parent.End, _end) + (_end - _end_beforeUpdate);
                //_parent.End = Math.Max(_parent.End, _end) + ((_end-_start) - _text_beforeUpdate.Length);
                _parent.End = Math.Max(_parent.End, _end) + (Text.Length - _text_beforeUpdate.Length);
            }
        }

        private void UpdateDocumentRange()

        {

            //if (_document != null)

            //{

            //    //_document.Start = Math.Min(_document.Start, _start);

            //    //_document.End = Math.Max(_document.End, _end);

            //}

        }

        public string TextBeforeUpdate => _text_beforeUpdate;


        public int Start
        {
            get => _start;
            //set
            //{
            //    _start = value;
            //    _end = _start + _text.Length;
            //    UpdateDocument();
            //    UpdateDocumentRange();
            //}
        }

        public int StartBeforeUpdate => _start_beforeUpdate;


        public int End
        {
            get => _end;
            //set
            //{
            //    _end = value;
            //    _text = _document.Text.Substring(_start, _end - _start);
            //    UpdateDocument();
            //    UpdateDocumentRange();
            //}
        }

        //public int EndBeforeUpdate
        //{
        //    get => _start_beforeUpdate + _text_beforeUpdate.Length;
        //}
        public int EndBeforeUpdate => _start_beforeUpdate + _text_beforeUpdate.Length;

        public int Index // 新增的屬性
        {
            get => _index;
            set => _index = value;
        }


        private void UpdateDocument()
        {
            var paragraphs = _document.GetParagraphs();
            int index = _index; // 使用 Index 屬性

            if (index < 0 || index >= paragraphs.Count)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("找不到要更新段落"
                    + Environment.NewLine + Environment.NewLine
                    + this.Text + Environment.NewLine + Environment.NewLine
                    + this._text_beforeUpdate);
                return;
            }

            var lines = _document.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (index < lines.Length)
            {
                lines[index] = _text;
                _document.Text = string.Join(Environment.NewLine, lines);
            }
            else
            {
                throw new IndexOutOfRangeException("索引超出 lines 陣列的界限。");
            }
        }

        public Range Range
        {
            get
            {
                if (_range == null)
                {
                    _range = new Range(_document, _start, _end, _parent);
                    return _range;
                }
                return _range;
            }
        }
    }



    /* 20250204 GitHub　Copilot大菩薩：
      * 這樣，我們就實現了在 Document 類別中取得插入點所在位置的段落，以及其上個段落和下個段落的方法，並且創建了一個 Range 類別來模擬 MS Word VBA 的 Range 物件。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
      */

    /// <summary>
    /// 創建一個 Range 類別來模擬 MS Word VBA 的 Range 物件。
    /// </summary>
    public class Range
    {
        private Document _document;
        private int _start;
        private int _end;
        List<Paragraph> rangeParagraphs;
        private Range _parent;
        private readonly string _text_beforeUpdate;

        //TextBox _textBox;

        public Range(Document document, int start, int end, Range parent)
        //public Range(ref TextBox textBox, int start, int end)
        {
            _document = document;
            //_document = new Document(ref textBox);
            _start = start < 0 ? 0 : start;
            _end = end > document.Text.Length ? document.Text.Length : end;
            _text_beforeUpdate = _document.Text.Substring(start, end - start);
            rangeParagraphs = Paragraphs;
            _parent = parent;
            //_textBox = textBox;
            //UpdateDocumentRange();
        }

        public string Text
        {
            get => Document.Text.Substring(_start, _end - _start);
            set
            {
                var text = Document.Text;
                Document.Text = text.Substring(0, _start) + value + text.Substring(_end);
                _end = _start + value.Length;
                UpdateDocument();
                //GitHub　Copilot大菩薩：我們在設置 Text 屬性時調用 UpdateDocument 方法，以確保文檔內容的更新。
                //UpdateDocumentRange();
                UpdateParentRange();
            }
        }

        private void UpdateParentRange()
        {
            if (_document != null)
            {

                //var range = _document.Range(_start, _end);
                //range.Text = _text;
                //●●●●●●●●●●●
                _parent.Start = Math.Min(_parent.Start, _start);

                //_parent.End = Math.Max(_parent.End, _end) + (_end - _end_beforeUpdate);
                //_parent.End = Math.Max(_parent.End, _end) + ((_end-_start) - _text_beforeUpdate.Length);
                //_parent.End = Math.Max(_parent.End, _end) + (Text.Length - _text_beforeUpdate.Length);
                _parent.End += (Text.Length - _text_beforeUpdate.Length);
            }
        }
        /// <summary>
        /// 在 Range 類別中添加一個 UpdateDocument 方法，並在設置 Text 屬性時調用它。以下是具體實現：
        /// </summary>
        private void UpdateDocument()
        {
            var paragraphs = Document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount)
                {
                    if (paragraphs[i].Text != Document.Text.Substring(charCount - paragraphs[i].Text.Length - Environment.NewLine.Length, paragraphs[i].Text.Length))
                        paragraphs[i].Text = Document.Text.Substring(charCount - paragraphs[i].Text.Length - Environment.NewLine.Length, paragraphs[i].Text.Length);//●●●●●●●●●●●●●●20250213
                }
            }
        }


        private void UpdateDocumentRange()
        {
            if (Document != null)
            {
                //_document.Start = Math.Min(_document.Start, _start);
                //_document.End = Math.Max(_document.End, _end);
            }
        }


        /*
         GitHub　Copilot大菩薩：新增 Start 和 End 屬性，這些屬性將允許讀取和設置範圍的起始和結束位置。以下是具體實現：20250204
         */
        /// <summary>
        /// 就新增了 Start 和 End 屬性，允許讀取和設置範圍的起始和結束位置。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
        /// </summary>
        public int Start
        {
            get => _start;
            set
            {
                if (value < 0) value = 0;
                else if (value > Document.End) value = Document.End;
                _start = value;
                //UpdateDocumentRange();
            }

        }
        /// <summary>
        /// 就新增了 Start 和 End 屬性，允許讀取和設置範圍的起始和結束位置。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
        /// </summary>
        public int End
        {
            get => _end;
            set
            {
                _end = value > Document.End ? Document.End : value;
                //UpdateDocumentRange();//由Text屬性更新即可
            }
        }
        public void Select()
        {
            Document.SelectionStart = _start;
            Document.SelectionLength = _end - _start;
        }

        public Paragraph GetCurrentParagraph()
        {
            var paragraphs = Document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount)
                {
                    return paragraphs[i];
                }
            }

            return null;
        }

        public Paragraph GetNextParagraph()
        {
            var paragraphs = Document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount && i + 1 < paragraphs.Count)
                {
                    return paragraphs[i + 1];
                }
            }

            return null;
        }

        public Paragraph GetPreviousParagraph()
        {
            var paragraphs = Document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount && i - 1 >= 0)
                {
                    return paragraphs[i - 1];
                }
            }

            return null;
        }

        //public List<Paragraph> Paragraphs
        //{
        //    get
        //    {
        //        if (rangeParagraphs == null || rangeParagraphs.Count == 0)
        //        {
        //            rangeParagraphs = new List<Paragraph>();
        //            var paragraphs = Document.GetParagraphs();
        //            if (_start == Document.Start && _end == Document.End)
        //            {
        //                rangeParagraphs = paragraphs;
        //            }
        //            else
        //            {
        //                int charCount = 0;
        //                //GitHub　Copilot大菩薩：這樣，Paragraphs 屬性將始終返回最新的段落集合。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。 20250204
        //                //故不要改為欄位儲存值，以免造成段落集合不是最新的問題。
        //                foreach (var paragraph in paragraphs)
        //                {
        //                    charCount += paragraph.Text.Length + Environment.NewLine.Length;
        //                    if (charCount > _start && charCount - paragraph.Text.Length - Environment.NewLine.Length < _end)
        //                    {
        //                        rangeParagraphs.Add(paragraph);
        //                    }
        //                }
        //            }
        //        }
        //        return rangeParagraphs;
        //    }
        //}

        //private List<Paragraph> GetParagraphs()
        //{
        //    var paragraphs = new List<Paragraph>();
        //    var lines = this.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
        //    int start = 0;
        //    for (int i = 0; i < lines.Length; i++)
        //    {
        //        var line = lines[i];
        //        start = this.Text.IndexOf(line, start) + _start;//●●●●●●●●●●●●●●

        //        if (start == -1)
        //        {
        //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("有問題，請檢查！");
        //            return null;
        //        }

        //        paragraphs.Add(new Paragraph(line, this, start, i));

        //        start += (line.Length == 0 ? 1 : line.Length);
        //        start = start > this.Text.Length ? this.Text.Length : start;
        //    }


        //    return paragraphs;
        //}
        //public List<Paragraph> Paragraphs
        //{
        //    get
        //    {
        //        if (rangeParagraphs == null || rangeParagraphs.Count == 0)
        //            rangeParagraphs = GetParagraphs();
        //        return rangeParagraphs;
        //    }
        //}

        public List<Paragraph> Paragraphs
        {
            get
            {
                var paragraphs = _document.GetParagraphs(true, this);
                var rangeParagraphs = new List<Paragraph>();
                int charCount = 0;
                //GitHub　Copilot大菩薩：這樣，Paragraphs 屬性將始終返回最新的段落集合。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。 20250204
                //故不要改為欄位儲存值，以免造成段落集合不是最新的問題。
                foreach (var paragraph in paragraphs)
                {
                    charCount += paragraph.Text.Length + Environment.NewLine.Length;
                    if (charCount > _start && charCount - paragraph.Text.Length - Environment.NewLine.Length < _end)
                    {
                        rangeParagraphs.Add(paragraph);
                    }
                }

                return rangeParagraphs;
            }
        }

        public Document Document { get => _document; }// set => _document = value; }

    }
}
