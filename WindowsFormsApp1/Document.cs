using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

/* 20250124�~���e�̫�W�Z�� GitHub Copilot�j���ġG
 * �F�ѤF�z���ݨD�A�ڭ̥i�H�Ыؤ@�� Document ���O�Ӫ�� textBox1�A�óЫؤ@�� Paragraph ���O�Ӫ�ܬq���C�o�����O�N����ާ@ textBox1 �������e�A�ýT�O���q������r��AtextBox1 ����r�]��������T�a��s�C
    �����A�ڭ̳Ыؤ@�� Document ���O�A�����O�N�]�t textBox1 �ô��Ѿާ@�䤺�e����k�C�M��A�ڭ̳Ыؤ@�� Paragraph ���O�A�����O�N����ާ@ Document �����q���C
    �K�K
 * �o�ˡA�ڭ̴N�ЫؤF Document �M Paragraph ���O�CDocument ���O��� textBox1�A�ô��Ѿާ@�䤺�e����k�CParagraph ���O��ܳ�Ӭq���A�ï���ާ@ Document �����q���C
    �z�i�H�b Form1 ���Ы� Document �M Paragraph ��H�A�èϥΥ��̨Ӿާ@ textBox1 �������e�C�Ҧp�G
    private void button1_Click(object sender, EventArgs e)
    {
        var paragraphs = _document.GetParagraphs();
        if (paragraphs.Count > 0)
        {
            paragraphs[0].Text = "�s���q���奻";
        }
    }
 * �o�ˡA��z�I�����s�ɡA�Ĥ@�Ӭq�����奻�N�Q���A�åB textBox1 �������e�]�|������s�C�Ʊ�o��z�������U�C�P���P���A�n�L��������C
 */
namespace TextForCtext
{
    /// <summary>
    /// ��ܤ��ɡ]������ textBox1�^
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
        {/* �n���A�ڭ̥i�H�ھڱz���ݨD�ӽs�g���եN�X�C�H�U�O��{�o�ǨB�J���N�X�G
            1.	���o textBox1 ���e���Ҧ��q���C
            2.	�M���U�Ӭq���A��X�ŦX���󪺬q���C
            3.	���ŦX���󪺬q�����U�@�Ӭq�����Ĥ@�Ӧr���C
            4.	�T�O���ϬM�b textBox1 �����e���C
            �����A�ڭ̻ݭn�b Document ���O���K�[�@�Ӥ�k�ӧ�s�q�����Ĥ@�Ӧr���G
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
    /// ��ܳ�Ӭq��
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
                // ��e�� Paragraph ��H���b paragraphs �C���A�B�z�o�ر��p
                // �i�H��ܩߥX���`�ΰO�����~�H��
                //throw new InvalidOperationException("��e�� Paragraph ��H���b paragraphs �C���C");

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
                // �B�z index �W�X lines �d�򪺱��p
                throw new IndexOutOfRangeException("���޶W�X lines �}�C���ɭ��C");
            }
        }
    }
}
