using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
    public class Document
    {
        private TextBox _textBox;
        private int currentParagraphIndex;
        public Document(TextBox textBox)
        {
            _textBox = textBox;
        }

        public string Text
        {
            get => _textBox.Text;
            set => _textBox.Text = value;
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
        internal int CurrentParagraphIndex { get => currentParagraphIndex; set => currentParagraphIndex = value; }

        public void InsertText(string text)
        {
            int selectionStart = _textBox.SelectionStart;
            _textBox.Text = _textBox.Text.Insert(selectionStart, text);
            _textBox.SelectionStart = selectionStart + text.Length;
        }

        public void ReplaceText(string text)
        {
            int selectionStart = _textBox.SelectionStart;
            _textBox.Text = _textBox.Text.Remove(selectionStart, _textBox.SelectionLength);
            _textBox.Text = _textBox.Text.Insert(selectionStart, text);
            _textBox.SelectionStart = selectionStart + text.Length;
        }

        public List<Paragraph> GetParagraphs()
        {
            var paragraphs = new List<Paragraph>();
            var lines = _textBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            foreach (var line in lines)
            {
                paragraphs.Add(new Paragraph(line, this));
            }

            return paragraphs;
        }
        //public List<Paragraph> GetParagraphs(Document document)
        //{
        //    var paragraphs = new List<Paragraph>();
        //    var lines = this.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

        //    foreach (var line in lines)
        //    {
        //        paragraphs.Add(new Paragraph(line, this));
        //    }

        //    return paragraphs;
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraphIndex"></param>
        /// <param name="newChar"></param>
        public void UpdateParagraphFirstCharacter(int paragraphIndex, string newChar)
        {/* 好的，我們可以根據您的需求來編寫測試代碼。以下是實現這些步驟的代碼：
            1.	取得 textBox1 內容的所有段落。
            2.	遍歷各個段落，找出符合條件的段落。
            3.	更改符合條件的段落的下一個段落的第一個字元。
            4.	確保更改反映在 textBox1 的內容中。
            首先，我們需要在 Document 類別中添加一個方法來更新段落的第一個字元：
          */

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
    }

    /// <summary>
    /// 表示單個段落
    /// </summary>
    public class Paragraph
    {
        private string _text;
        private Document _document;

        public Paragraph(string text, Document document)
        {
            _text = text;
            _document = document;
        }

        public string Text
        {
            get => _text;
            set
            {
                _text = value;
                UpdateDocument();
            }
        }

        private void UpdateDocument()
        {
            /*
            var paragraphs = _document.GetParagraphs();
            int index = paragraphs.IndexOf(this);
            var lines = _document.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            lines[index] = _text;
            _document.Text = string.Join(Environment.NewLine, lines);
            */

            var paragraphs = _document.GetParagraphs();
            //var paragraphs = _document.GetParagraphs(_document);

            int index = paragraphs.IndexOf(this);            

            if (index == -1)
            {
                //return;                
                // 當前的 Paragraph 對象不在 paragraphs 列表中，處理這種情況
                // 可以選擇拋出異常或記錄錯誤信息
                //throw new InvalidOperationException("當前的 Paragraph 對象不在 paragraphs 列表中。");

                index = _document.CurrentParagraphIndex;
            }
            var lines = _document.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (index < lines.Length)
            {
                lines[index] = _text;
                _document.Text = string.Join(Environment.NewLine, lines);
            }
            else
            {
                // 處理 index 超出 lines 範圍的情況
                throw new IndexOutOfRangeException("索引超出 lines 陣列的界限。");
            }
        }
    }
}
