Attribute VB_Name = "Keywords"
Option Explicit
Rem ��������r�˯��B���Ѭ������ݩʡB�ѷӰO��

Rem �ΥH�ˬd�O�_�����ǽd�򤧤��e��
Property Get ����KeywordsToCheck() As Variant 'string()
    ����KeywordsToCheck = Array(VBA.ChrW(-10119), VBA.ChrW(-8742), VBA.ChrW(-30233), VBA.ChrW(-10164), VBA.ChrW(-8698), VBA.ChrW(-31827), VBA.ChrW(-10132), VBA.ChrW(-8313), VBA.ChrW(20810), VBA.ChrW(-10167), VBA.ChrW(-8698), VBA.ChrW(-26587), VBA.ChrW(21093), VBA.ChrW(14615), VBA.ChrW(20089), VBA.ChrW(26080), "�k", VBA.ChrW(26083), "��" _
        , "��", "�P", VBA.ChrW(20089), "��", "��", "�p�b", "�i", "�{", "�[", "�j�L", "�[", "��", "�_", "����", "�N", "��", "��", "�X", "�P�H", "�j��", "��", "�_", "��", "��", "�^", "��", "��", "�L�k", "�j�b", "�v", "��", "�H", "��", "�[", "�w", "��", "�l", "�q", "�_", "��", "����", "�Q", "�j��", "�[", "�l", "��", "�k�f", "�p�L", "��", "���i", "��", "��", "��", "��", "�J��", "����", "�a�H", "��", "�x", "��", "�S", "�I", "�", "��", "��", "��", "�A", "�`", "�ӷ�", "����", "���", "�H", "ν", _
        "�ѳ�", "�Ѷ�", "�ֳ�", "�ֶ�", "�")
End Property
Rem �ΥH���ѩ�������r��
Property Get ����KeywordsToMark() As Variant 'string()�]�� Array Returns a Variant containing an array,�ҥH����g�� as string()
    ����KeywordsToMark = Array("��", "�P��", "���g", "�j��", "���g", "���g", "�C�g", "�Q�T�g", "�", _
        "��", "�`��", "����", _
        "��", _
        "�t��", "ô��", "����", "����", "ô��", "����", "�Ǩ�", _
            "����", "�Ԩ�", "����", "�娥", "���[", "����", "�Q�s", "�v�O", _
        "�b", "�[", "��", "�q���r", "�q�[�r", "���B�[", "�q���B�[�r", "����", "�N", VBA.ChrW(20089), "�J��", VBA.ChrW(26083) & "��", "����", "�Q�l", _
        "�j" & VBA.ChrW(22766), _
        "��E", "�E�G", "�E�T", "�E�|", "�E��", "�W�E", VBA.ChrW(19972) & "�E", "�ΤE", "�줻", "���G", "���T", "���|", "����", "�W��", "�Τ�", _
        "�e��", "����", "�ӷ�", "�L��", _
            "�H��", "�q�H�r��", "�H��", "�H��", "�H��", "�q�j�H�r", "�p�H", "�H�q", "ν", _
            "��", "�[", "��", "����", "�I", "��", "�l", "�S", VBA.ChrW(14514), "��", VBA.ChrW(21093), "�Q", "�j��", "���i", "��" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "�p�b", "�j�b", "��", "�", "�A", VBA.ChrW(28067), "��", "��", "�k�f", "�p�L", "�j��", "�j�L", "�q���r", "�q�_�r", "�q�l�r", "�q�q�r", "�q�١r", "��", "�q�^�k�r", "�q�_�r", "�q�_�r", _
            "�ѳ�", "�Ѷ�", "�ֳ�", "�ֶ�", "����", "�ٵ�", _
            "�S", "�w��", "�w��", _
        "�L�k", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "�k", _
        "�L�S", VBA.ChrW(26080) & "�S", "�ѩS", "����", _
        "�H�ɤ��q", "������", "�]����", "��q�J��", "�F��", "����", "�Ӥ���", "�p�b�ѤW", "����", "���f", "�ޤ�", "�T��", "�g��", "����", "�g��", "�q�H����", "�q�H��~", "�g��o�D", "�Q��n", "�~���w��", "�ѤU�j��", "�q�ʦ�", "��i�Læ", "��i" & VBA.ChrW(26080) & "æ", "�W�S", "�W" & VBA.ChrW(14514), "�b��", "�W�_", "�~��", "�s��", "����", "���[", "����", "���U��", "�X���Q�s", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "���Q�s", "�񤧭�H", "�i�s", "�s�F", "�i�D�Z�Z", "�s�N", "����", "��W����", "�ҥ��U��", "���ӱo", _
        "���`", "��" & VBA.ChrW(20158), "��" & VBA.ChrW(20838), "�ɸq", "����", "�����ӥ~��", "�����~��", "�~���Ӥ���", "�~������", "��²", "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "���_", "�}������", "�a������", "��X���`", "���`��X", "�����h�E", "���L�h��", "�E����L", "�i��", "���Y", "�@���@��", _
        "��")
        
End Property
Rem �Y����r���e������O 20240914
Property Get ����KeywordsToMark_ExamPrecededAvoid() As Scripting.Dictionary
    
    
    Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
    ' �K�[��ƨ�r�� creedit_with_Copilot�j���ġGhttps://sl.bing.net/goDF239cQVw
    dict.Add "��", Array("��", "������", "��", "���x�B��", "���x��", "���H�M", "�@", "ź", "�@��@", "��M", "�s����", "����", "�Z�B", "�B", "��������������", "�����{", "�Ƥ[�h�{", "�թT", "�ߪ���]", "�ߪ�" & ChrW(25143) & "�]", "ť����", "�p�M", "�ۤ���", "�i", "�H��", "��", "��", ChrW(-28903), "��", "�L��M", "����", "��", "�e", "��", "��", "��", "��", "�T", "��", "��", "��", "�Υi", "��", "��", "�@", "��", "�C", "�U", "�Z", "��", "��", "�y", "��", "��", "��", "²", "�թ~", "�~", "�L", "��", "�I", "��", "����", "�@", "��", "��")
    dict.Add "��", Array("��", "��", "����", "��", "��", "��", "��", "��", "�ڽ�", "����", "�s��", "�f", "�\", "�M", "��", "��", "��", "�~�j��")
    dict.Add "���[", Array("�p��", "�_���s")
    dict.Add "��", Array("���w��", "�v", "�t", "��", "��", "�B", "�S", "�h", VBA.ChrW(-10143) & VBA.ChrW(-8996))
    dict.Add "��", Array("��", "�r", "�i", "������", "���")
    dict.Add VBA.ChrW(21093), Array("�i", "��", "������", "���")
    dict.Add "�[", Array("�P��", "�{", "��", "����", "����", "�r��")
    dict.Add VBA.ChrW(-26587), Array("�P��", "�{", "��", "����", "����")
    dict.Add "�S", Array("��", "�C��", "��")
    dict.Add VBA.ChrW(14514), Array("��", "�C��", "��")
    dict.Add "�j�L", Array("�i�L")
    dict.Add "�N", Array("�Ѹ�", "�ҡv�@�u", "�ҧ@")
    dict.Add "�", Array("��", "�B", "��", "�_")
    dict.Add "�[", Array("��")
    dict.Add "����", Array("�P")
    dict.Add "�p�L", Array("�O��")
    dict.Add "�A", Array("�Z", "Ĭ")
    dict.Add VBA.ChrW(28067), Array("�Z", "Ĭ")
    dict.Add "��", Array("��", "�m", "��", "�i�H��", "�D", "�~��", "�U", "����", "�פ�", "�C�J")
    dict.Add "��", Array("���")
    dict.Add "��", Array("���v�@�u", "���@")
    dict.Add "�j��", Array("����", "���", "�Ӥ�", "�ᦷ")
    dict.Add "�J��", Array("�H", "�I")
    dict.Add "�E�G", Array("���|", "��", "���G")
    dict.Add "�E�T", Array("��", "���@")
    dict.Add "���G", Array("��", "���@", "���G")
    dict.Add "���T", Array("��")
    dict.Add "���|", Array("��", "���@", "�K�@�@")
    dict.Add "�H��", Array("������")
    dict.Add "�娥", Array("�P�W")
    dict.Add "�S", Array("������", "�еL�S", "�ФV�L�S", "�o", "�l", "��", "��", "��", "�M", "����", "�R��", "���H��", "�ા��")
    dict.Add "����", Array("��")
    dict.Add "�s�N", Array("��")
    dict.Add "���[", Array("����")
    dict.Add "�L�S", Array("��", "�ФV")
    dict.Add "����", Array("�i")
    dict.Add "���_", Array("�i")
    dict.Add "��²", Array("Ĭ")
    dict.Add "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("Ĭ")
    
    Set ����KeywordsToMark_ExamPrecededAvoid = dict

        
End Property
Rem �Y����r���᭱����O 20240914
Property Get ����KeywordsToMark_ExamFollowedAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "��", Array("��", "�v", "�r��", "��", "�o�S�S", "�l�ҧ@", "���ͯ�", "�o��", "����", "�ȩ���", "�G��", "���@��", "�ۤ�", "��Q", "�H���I", "�����", "�}", "�w�~�h", "�W��", "�W����", "��", "�o��", "�Q�m", "���h��", "�n����", "������", "�U��", "��", "�C", "�o����", "����", "��", "�Xĳ��", "�t", "��", "�h���`", "��", "�n", "��", "��", "��", "�U", "�R��", "��", "��", "�}��", "��", "�m", "�ȥ�", "��c", "���ѳ�", "�m", "��", "��", "�u", "��", "�{", "����", "�p��", "�p½", "�p�B", "�e", "��", "�l��", "��", "��", "����", "�ƺ�", "������")
    dict.Add "��", Array("�}", "��", "�ؿv")
    dict.Add "�b", Array("�K")
    dict.Add "��", Array("�\", "�b", _
        "�D��", "�D���~", "��", "��", _
        "�M��")
    dict.Add "���[", Array("���N�H")
    dict.Add "��", Array("��", "��", "�{", "�v", "��", "�@", "��", VBA.ChrW(20675), "��", "�D", _
        "���n", "�ݬ�����", "��" & VBA.ChrW(29234) & "����")
    dict.Add "��", Array("�K" & VBA.ChrW(-31631), "�K��", "��", "�d", "��", "�k", "��", "��", "�h�z", "�H��", "���H��")
    dict.Add VBA.ChrW(21093), Array("��", "�d", "��", "�k", "��", "��", "�h�z", "�H��", "���H��")
    dict.Add "�", Array("�")
    dict.Add "�A", Array("�M", "��", "�E��")
    dict.Add VBA.ChrW(28067), Array("�M", "��", "�E��")
    dict.Add "�[", Array("�M��", "����", "�r��", "�Z��", "�r�Z", "�U����")
    dict.Add VBA.ChrW(-26587), Array("�M��", "����", "�r��", "�Z��", "�r�Z", "�U����")
    dict.Add "��", Array("����", "�_�@")
    dict.Add "��", Array("��", "��")
    dict.Add "�N", Array("����")
    dict.Add "�S", Array("��", "�r��", "�礣��", "��K", "����")
    dict.Add VBA.ChrW(14514), Array("��", "�r��", "�礣��", "��K", "����")
    dict.Add "��", Array("�ۮI", "���B", "��ͤl", "�r", "�o��", "�D", "����", "����", "����", "�̬O�]", "��", VBA.ChrW(23280))
    dict.Add "��", Array("�v�@�u��", "�@��")
    dict.Add "��", Array("��")
    dict.Add "�Q", Array("�@�ׯu")
    dict.Add "�p�b", Array("��")
    dict.Add "�j��", Array("����", "�O", "ĵ��")
    dict.Add "�j�L", Array("�H��")
    dict.Add "�p�L", Array("�y�e", "�y" & VBA.ChrW(23515))
    dict.Add "��E", Array("��")
    dict.Add "�줻", Array("��")
    dict.Add "�S", Array("��", "��")
    dict.Add "���_", Array("�Ǥh")

    
    Set ����KeywordsToMark_ExamFollowedAvoid = dict

End Property
Rem �Y����r����b�Y�ӻy�y�̭� 20240914
Property Get ����KeywordsToMark_ExamInPhraseAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "��", Array("�����w", "�~����", "Ĭ��²", "����²", "�i����", "������", "�H����", "�H���h", "������", _
        "���s���D", "�`�ѩ���", "�Z���{", "������", "�֩���", "�����@", "�W����", "�����o", "�ۦܩ�", "�h����", "�̩��s", "�m����", "���B����", "�Ʃ���", "�����H�a�A", "�ѩ��H", "�����]", "�����k", "�i�H���@��", "�����", "�M������", _
        "�ө���", "�ө��_", "�ө���", "�̩��o", "������", "�H���Q��", "�y���G", "�Ʃ��B", "�̩��V", "�ߩ���", "���ߩ���", "�G�仡���t", "�l�����O", "�D���X", "�ө���", "�C�ө���", "�ө��v", "�֩���", "���ө���", _
        "�H����", "�H���B��", "�]������", "�H�����̩���", "�]��������", "�h���o", "�ک��o", _
        "�ө���", "���h����", "�̩���", "�U�h����", "�����a", "��H���a", "�L�H����", "�󦳩�����", "�ө���", _
        "�ܩ���", "�����]", "�e�F�B���w", "���w�B�Q��", "�e�F���w", "�ܩ���", "���w�Q��", "�X���v", "�̩���", _
        "�e���@", "�f����", "�H���q", "���������", "���t�� ", "�����Y", "�ҥH����", "�̩���", "�O�����]", "������", "�����H����", "�]������", "�ҩ�����", "�ҩ�" & ChrW(-24892) & "��", "�ɩ��H�s", _
        "�h����", "�V���ӷ~", "������", "��������", "���۩���", "�D���A", _
        "�]���W��", "������", _
        "�h���J��", _
        "�D����", _
        "�Z����")
    dict.Add "��", Array("�K���s")
    dict.Add "��", Array("�j���q", "�Ű���")
    dict.Add "��", Array("�ѭ�ӵo��", "�y�鬥��")
    dict.Add VBA.ChrW(21093), Array("��" & VBA.ChrW(21093) & "�ӵo��", "�y" & VBA.ChrW(21093) & "����")
    dict.Add "��", Array("�H�ݪ�")
    dict.Add "�", Array("���", "���")
    dict.Add "�[", Array("���[��", "���[�j�f")
    dict.Add VBA.ChrW(-26587), Array("��" & VBA.ChrW(-26587) & "��", "��" & VBA.ChrW(-26587) & "�j�f")
    dict.Add "��", Array("�����ͯ�")
    dict.Add "��", Array("�H�ۯd�H", "�H�ۦ�", "���۪�", "�H�ۤj", "���۪�")
    dict.Add "�N", Array("�j�N��")
    dict.Add "�S", Array("�Y�S�]", "�F�S�u")
    dict.Add VBA.ChrW(14514), Array("�Y" & VBA.ChrW(14514) & "�]", "�F" & VBA.ChrW(14514) & "�u")
    dict.Add "�P�H", Array("���P�H��", "���P�H��", "���P�H��")
    dict.Add "�j��", Array("�Ƥj����")
    dict.Add "�j�L", Array("�L�j�L�c")
    dict.Add "�L�k", Array("�H�L�k��", "���L�k�O")
    dict.Add VBA.ChrW(26080) & "�k", Array("�H" & VBA.ChrW(26080) & "�k��", "��" & VBA.ChrW(26080) & "�k�O")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("�H" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "��", "��" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "�O")
    dict.Add "��E", Array("����E��")
    dict.Add "�E�G", Array("�@�E�G��")
    dict.Add "�E�T", Array("�ܤE�T�Q")
    dict.Add "�s��", Array("�ߦs�۷q")
        
    Set ����KeywordsToMark_ExamInPhraseAvoid = dict

End Property
Rem �˴�����r
'Function ����KeywordsToMark_Exam()
'
'End Function