namespace TextForCtext
{
    /// <summary>
    /// 表示單個段落
    /// </summary>
    public class Paragraph
    {
        public string Text { get; set; }

        public Paragraph(string text)
        {
            Text = text;
        }

        // 可以在這裡添加更多段落相關的屬性和方法
    }

    /// <summary>
    /// 表示段落的集合
    /// </summary>
    public class Paragraphs
    {
        private List<Paragraph> _paragraphs;

        public Paragraphs()
        {
            _paragraphs = new List<Paragraph>();
        }

        public int Count => _paragraphs.Count;

        public Paragraph this[int index] => _paragraphs[index];

        public void Add(Paragraph paragraph)
        {
            _paragraphs.Add(paragraph);
        }

        public void RemoveAt(int index)
        {
            _paragraphs.RemoveAt(index);
        }

        public void Clear()
        {
            _paragraphs.Clear();
        }
        
        // 可以在這裡添加更多集合相關的操作方法
        
        /// <summary>
        /// 查找包含指定文本的段落索引
        /// </summary>
        /// <param name="text">要查找的文本</param>
        /// <returns>包含指定文本的段落索引列表</returns>
        public List<int> FindParagraphsContainingText(string text)
        {/* 在 Paragraphs 類別中添加一個方法來查找特定文本的段落，可以通過遍歷 _paragraphs 集合並檢查每個段落的文本是否包含指定的文本。這裡有一個示例方法 FindParagraphsContainingText，它將返回包含指定文本的段落索引列表。
            這個方法 FindParagraphsContainingText 接受一個字符串參數 text，並返回一個包含指定文本的段落索引列表。它遍歷 _paragraphs 集合，檢查每個段落的 Text 屬性是否包含指定的文本，如果包含，則將該段落的索引添加到結果列表中。
            您可以根據需要進一步擴展此方法，例如添加區分大小寫的選項或返回包含指定文本的段落對象列表。
          */
            List<int> indices = new List<int>();

            for (int i = 0; i < _paragraphs.Count; i++)
            {
                if (_paragraphs[i].Text.Contains(text))
                {
                    indices.Add(i);
                }
            }

            return indices;
        }

        /// <summary>
        /// 獲取所有段落的文本
        /// </summary>
        /// <returns>包含所有段落文本的列表</returns>
        public List<string> GetAllParagraphTexts()
        {/* 在 Paragraphs 類別中添加一個方法來獲取所有段落的文本，可以通過遍歷 _paragraphs 集合並將每個段落的 Text 屬性添加到一個列表中。這裡有一個示例方法 GetAllParagraphTexts，它將返回包含所有段落文本的列表。
            這個方法 GetAllParagraphTexts 遍歷 _paragraphs 集合，並將每個段落的 Text 屬性添加到一個 List<string> 中，最後返回這個列表。這樣，您就可以獲取所有段落的文本。
          */
            List<string> texts = new List<string>();

            foreach (var paragraph in _paragraphs)
            {
                texts.Add(paragraph.Text);
            }

            return texts;
        }
    }
}
