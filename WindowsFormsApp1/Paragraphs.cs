namespace TextForCtext
{
    /// <summary>
    /// ��ܳ�Ӭq��
    /// </summary>
    public class Paragraph
    {
        public string Text { get; set; }

        public Paragraph(string text)
        {
            Text = text;
        }

        // �i�H�b�o�̲K�[��h�q���������ݩʩM��k
    }

    /// <summary>
    /// ��ܬq�������X
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
        
        // �i�H�b�o�̲K�[��h���X�������ާ@��k
        
        /// <summary>
        /// �d��]�t���w�奻���q������
        /// </summary>
        /// <param name="text">�n�d�䪺�奻</param>
        /// <returns>�]�t���w�奻���q�����ަC��</returns>
        public List<int> FindParagraphsContainingText(string text)
        {/* �b Paragraphs ���O���K�[�@�Ӥ�k�Ӭd��S�w�奻���q���A�i�H�q�L�M�� _paragraphs ���X���ˬd�C�Ӭq�����奻�O�_�]�t���w���奻�C�o�̦��@�ӥܨҤ�k FindParagraphsContainingText�A���N��^�]�t���w�奻���q�����ަC��C
            �o�Ӥ�k FindParagraphsContainingText �����@�Ӧr�Ŧ�Ѽ� text�A�ê�^�@�ӥ]�t���w�奻���q�����ަC��C���M�� _paragraphs ���X�A�ˬd�C�Ӭq���� Text �ݩʬO�_�]�t���w���奻�A�p�G�]�t�A�h�N�Ӭq�������޲K�[�쵲�G�C���C
            �z�i�H�ھڻݭn�i�@�B�X�i����k�A�Ҧp�K�[�Ϥ��j�p�g���ﶵ�Ϊ�^�]�t���w�奻���q����H�C��C
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
        /// ����Ҧ��q�����奻
        /// </summary>
        /// <returns>�]�t�Ҧ��q���奻���C��</returns>
        public List<string> GetAllParagraphTexts()
        {/* �b Paragraphs ���O���K�[�@�Ӥ�k������Ҧ��q�����奻�A�i�H�q�L�M�� _paragraphs ���X�ñN�C�Ӭq���� Text �ݩʲK�[��@�ӦC���C�o�̦��@�ӥܨҤ�k GetAllParagraphTexts�A���N��^�]�t�Ҧ��q���奻���C��C
            �o�Ӥ�k GetAllParagraphTexts �M�� _paragraphs ���X�A�ñN�C�Ӭq���� Text �ݩʲK�[��@�� List<string> ���A�̫��^�o�ӦC��C�o�ˡA�z�N�i�H����Ҧ��q�����奻�C
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
