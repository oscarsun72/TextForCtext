Attribute VB_Name = "Keywords"
Option Explicit
Rem ��������r�˯��B���Ѭ������ݩʡB�ѷӰO��

Property Get ����KeywordsToMark() As Variant
    ����KeywordsToMark = Array("��", "�P��", "���g", "�j��", "���g", "�C�g", "�Q�T�g", _
        "��", "�`��", "����", _
        "��", _
        "�t��", "ô��", "����", "����", "ô��", "����", "�Ǩ�", _
            "����", "�Ԩ�", "����", "�娥", "���[", "����", "�Q�s", "�v�O", _
        "�b", "�[", "��", "�q���r", "�q�[�r", "���B�[", "�q���B�[�r", "����", "�N", VBA.ChrW(20089), "�J��", VBA.ChrW(26083) & "��", "����", "�Q�l", _
        "�j" & VBA.ChrW(22766), _
        "��E", "�E�G", "�E�T", "�E�|", "�E��", "�W�E", VBA.ChrW(19972) & "�E", "�ΤE", "�줻", "���G", "���T", "���|", "����", "�W��", "�Τ�", _
        "�e��", "����", "�ӷ�", "�L��", _
            "�H��", "�q�H�r��", "�H��", "�H��", "�H��", "�q�j�H�r", "�p�H", "�H�q", "ν", _
            "��", "�[", "��", "����", "�I", "��", "�l", "�S", VBA.ChrW(14514), "��", "�Q", "�j��", "���i", "��" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "�p�b", "�j�b", "��", "�", "�A", VBA.ChrW(28067), "��", "��", "�k�f", "�p�L", "�j��", "�j�L", "�q���r", "�q�_�r", "�q�l�r", "�q�q�r", "�q�١r", "��", "�q�^�k�r", "�q�_�r", "�q�_�r", _
            "�ѳ�", "�Ѷ�", "�ֳ�", "�ֶ�", "����", "�ٵ�", _
            "�S", "�w��", "�w��", _
        "�L�k", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "�k", _
        "�L�S", VBA.ChrW(26080) & "�S", "�ѩS", _
        "�H�ɤ��q", "������", "�]����", "��q�J��", "�F��", "����", "�Ӥ���", "�p�b�ѤW", "����", "���f", "�ޤ�", "�T��", "�g��", "����", "�g��", "�q�H����", "�q�H��~", "�g��o�D", "�Q��n", "�~���w��", "�ѤU�j��", "�q�ʦ�", "��i�Læ", "��i" & VBA.ChrW(26080) & "æ", "�W�S", "�W" & VBA.ChrW(14514), "�b��", "�W�_", "�~��", "�s��", "����", "���[", "����", "���U��", "�X���Q�s", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "���Q�s", "�񤧭�H", "�i�s", "�s�F", "�i�D�Z�Z", "�s�N", "����", _
        "���`", "��" & VBA.ChrW(20158), "��" & VBA.ChrW(20838), "�ɸq", "����", "�����ӥ~��", "�����~��", "�~���Ӥ���", "�~������", "��²", "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "���_", "�}������", "�a������", "��X���`", "���`��X", "�����h�E", "���L�h��", "�E����L", "�i��", "���Y", "�@���@��", _
        "��")
        
End Property
Rem �Y����r���e������O 20240914
Property Get ����KeywordsToMark_ExamPrecededAvoid() As Scripting.Dictionary
    
    
    Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
    ' �K�[��ƨ�r�� creedit_with_Copilot�j���ġGhttps://sl.bing.net/goDF239cQVw
    dict.Add "��", Array("��", "��������������", "�թT", "�ߪ���]", "�ߪ�" & ChrW(25143) & "�]", "ť����", "�ۤ���", "�i", "�H��", "��", "��", ChrW(-28903), "��", "�L��M", "����", "��", "�e", "��", "��", "��", "��", "�T", "��", "��", "��", "�Υi", "��", "��", "�@", "��", "�C", "�U", "�Z", "��", "��", "�y", "��", "��", "��", "²", "�թ~", "�~", "�L", "��", "�I", "��", "����", "�@", "��", "��")
    dict.Add "��", Array("��", "��", "����", "��", "��", "��", "��", "��", "�ڽ�", "����", "�s��", "�f", "�\", "�M", "��", "��", "��", "�~�j��")
    dict.Add "��", Array("���w��", "�v", "�t", "��", "��", "�B", "�S", "�h", VBA.ChrW(-10143) & VBA.ChrW(-8996))
    dict.Add "��", Array("��", "������")
    dict.Add "�[", Array("�P��", "�{", "��")
    dict.Add "�j�L", Array("�i�L")
    dict.Add "�S", Array("��", "��", "�M")
    dict.Add "�s�N", Array("��")
    dict.Add "��²", Array("Ĭ")
    dict.Add "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("Ĭ")
    
    Set ����KeywordsToMark_ExamPrecededAvoid = dict

        
End Property
Rem �Y����r���᭱����O 20240914
Property Get ����KeywordsToMark_ExamFollowedAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "��", Array("�o�S�S", "�w�~�h", "�o��", "���h��", "�n����", "������", "�U��", "��", "�C", "�o����", "����", "��", "�Xĳ��", "�t", "��", "�h���`", "��", "�n", "��", "��", "��", "�U", "�R��", "��", "��", "�}��", "��", "�m", "�ȥ�", "���ѳ�", "�m", "��", "��", "�u", "��", "�{", "����", "�p��", "�p½", "�p�B", "�e", "��", "�l��", "�", "��")
    dict.Add "��", Array("�}", "��", "�ؿv")
    dict.Add "�b", Array("�K")
    dict.Add "��", Array("�\", "�b")
    dict.Add "��", Array("���n", _
        "��", "��", "�@", "��", VBA.ChrW(20675), "��", "�{", "��", "�D")
    dict.Add "��", Array("�d", "��")
    dict.Add "��", Array("�M", "��")
    dict.Add "�[", Array("�M��")
    dict.Add "�S", Array("��")
    dict.Add VBA.ChrW(14514), Array("��")
    dict.Add "�j�L", Array("�H��")
    dict.Add "�p�L", Array("�y�e", "�y" & VBA.ChrW(23515))
    dict.Add "�S", Array("��", "��")
    
    
    Set ����KeywordsToMark_ExamFollowedAvoid = dict

End Property
Rem �Y����r����b�Y�ӻy�y�̭� 20240914
Property Get ����KeywordsToMark_ExamInPhraseAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "��", Array("�����w", "�~����", _
        "���s���D", "�`�ѩ���", "�W����", "�ۦܩ�", "�����]", "�����k", "�i�H���@��", _
        "�ө���", "�ө��_", "�ө���", "�̩��o", "�̩��V", "�ߩ���", "���ߩ���", "�G�仡���t", "�l�����O", "�D���X", "�ө���", "�C�ө���", "�ө��v", "�֩���", _
        "�H����", "�H���B��", _
        "�ө���", "���h����", "�̩���", "�U�h����", "�����a", "��H���a", "�L�H����", "�󦳩�����", _
        "�ܩ���", "�����]", _
        "�e���@", "�f����", "�H���q", "���������", "���t�� ", "�����Y", "�ҥH����", "�̩���", "�O�����]", _
        "�h����", _
        "�h���J��", _
        "�D����", _
        "�Z����")
    dict.Add "��", Array("�K���s")
    dict.Add "��", Array("�ѭ�ӵo��")
    dict.Add "��", Array("�H�ݪ�")
    dict.Add "�[", Array("���[��")
    dict.Add "�P�H", Array("���P�H��", "���P�H��", "���P�H��")
    dict.Add "�L�k", Array("�H�L�k��", "���L�k�O")
    dict.Add VBA.ChrW(26080) & "�k", Array("�H" & VBA.ChrW(26080) & "�k��", "��" & VBA.ChrW(26080) & "�k�O")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("�H" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "��", "��" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "�O")
        
    Set ����KeywordsToMark_ExamInPhraseAvoid = dict

End Property
Rem �˴�����r
Function ����KeywordsToMark_Exam()
    
End Function
