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
        "��", "�`��", "����", "�ٻX", "��" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "�١B�X", "�١B" & VBA.ChrW(-10132) & VBA.ChrW(-8313), _
        "��", _
        "�t��", "ô��", "����", "����", "ô��", "����", "�Ǩ�", _
            "����", "�Ԩ�", "����", "�娥", "���[", "����", "�Q�s", "�v�O", _
        "�b", "�[", "�����j�l", "�[�@����", "���H����", "�[�H²��", "��", "�q���r", "�q�[�r", "���B�[", "�q���B�[�r", "����", "�N�_�~", "�N��~", "�~�N", "���N", "�N", VBA.ChrW(20089), "�J��", VBA.ChrW(26083) & "��", "����", "�Q�l", _
        "�j" & VBA.ChrW(22766), _
        "��E", "�E�G", "�E�T", "�E�|", "�E��", "�W�E", VBA.ChrW(19972) & "�E", "�ΤE", "�줻", "���G", "���T", "���|", "����", "�W��", "�Τ�", _
        "�e��", "����", "�ӷ�", "�L��", "���", _
            "�H��", "�q�H�r��", "�H��", "�H��", "�H��", "�q�j�H�r", "�p�H", "�H�q", "�|�H", VBA.ChrW(-10145) & VBA.ChrW(-9156), "�H�G", "�H��", _
            "ν", _
             "���I", "��", "�[", "�P�H�_�v", "�P�H", "��", "����", "�I", "��", "�l", "�S", VBA.ChrW(14514), "��", VBA.ChrW(21093), "�Q�@�L�e", "�Q�@" & ChrW(26080) & "�e", "�Q", "�j��", "���i", "��" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "�p�b", "�j�b", "��", "�", "�A", VBA.ChrW(28067), "��", "��", "�k�f", "�p�L", "�j��", "�j�L", "�q���r", "�q�_�r", "�q�l�r", "�q�q�r", "�q��", "�X�r", VBA.ChrW(-10132) & VBA.ChrW(-8313) & "�r", "�ݤj", "��", "�q�^�k�r", "�q�_�r", "�q�_�r", "�q�ݡr", _
            "�ѳ�", "�Ѷ�", "�ֳ�", "�ֶ�", "����", "�ٵ�", _
            "�S", "�w��", "�w��", "�j�l", _
        "�L�k", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "�k", _
        "�L�S", VBA.ChrW(26080) & "�S", "�ѩS", "����", "�٨�I", "�צ�", _
        "�H�ɤ��q", "������", "�]����", "��q�J��", "�F��", "����", "�Ӥ���", "�p�b�ѤW", "����", "���f", "�ޤ�", "�T��", "�g��", "����", "�g��", "�q�H����", "�q�H��~", "�g��o�D", "�Q��n", "�~���w��", "�ѤU�j��", "�q�ʦ�", "��i�Læ", "��i" & VBA.ChrW(26080) & "æ", "�W�S", "�W" & VBA.ChrW(14514), "�b��", "�W�_", "�~��", "�s��", "����", "���[", "����", "���U��", "�X���Q�s", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "���Q�s", "�񤧭�H", "�i�s", "�s�F", "�i�D�Z�Z", "�s�N", "����", "��W����", "�ҥ��U��", "���ӱo", _
        "�L�e", ChrW(26080) & "�e", "���`", "��" & VBA.ChrW(20158), "��" & VBA.ChrW(20838), "�ɸq", "����", "�����ӥ~��", "�����~��", "�~���Ӥ���", "�~������", "��²", "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "���_", "�}������", "�a������", "��X���`", "���`��X", "�����h�E", "���L�h��", "�E����L", "�i��", "���Y", "�@���@��", "�ڦ��n��", "������", "���t�H���D�|", "���l�Ӯv", "�̤l�֤r", "��ΦӤ���", "���D�A", "��l�ϲ�", "�I�M����", "�P�ӹE�q", "�B�q", "�B�r", "�e���b��", "�e���b" & VBA.ChrW(-30650), "�i��", "�i��", "���{", "�{�j�g", "�q�Ӧ���", VBA.ChrW(-24871) & "�Ӧ���", "�����ӫH", "�s�G�w��", "�q�ѤU����", "�i��", "�~���̵�", "���̨���", "���̨���", "���̨���", _
        "�j�s", "�p�s", "�ҥX�G�_", "�ҥX��_", "�ҥX�_�_", "�P�ɰ���", "�յ�", "��" & VBA.ChrW(-31142), "��" & VBA.ChrW(-10119) & VBA.ChrW(-8991), "��" & VBA.ChrW(-31145), _
        "��", "��" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "��" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "����", "�W��", "�W" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "�W��", "����", "����")
        
End Property
Rem �Y����r���e������O 20240914
Property Get ����KeywordsToMark_ExamPrecededAvoid() As Scripting.Dictionary
    
    
    Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
    ' �K�[��ƨ�r�� creedit_with_Copilot�j���ġGhttps://sl.bing.net/goDF239cQVw
    dict.Add "��", Array("��", "�e", "��", "��", "��", "��", "�T", "��", "��", "��", "�Υi", "��", "��", "�@", "��", "�C", "�U", "�Z", "��", "��", "�y", "��", "��", "��", "²", "�թ~", "�~", "�L", "��", "�I", "��", "����", "�@", "��", "��", _
        "��", "������", "�E", "�C���@", "��", "���x�B��", "���x��", "���H�M", "�@", "ź", "�@��@", "��M", _
        "�s����", "����", "�Z�B", "�B", "�мY��", "������", "��������������", "�����{", "�Ƥ[�h�{", "�թT", "�ߪ���]", "�ߪ�" & VBA.ChrW(25143) & "�]", _
        "ť����", "�a�u�q", "�a" & VBA.ChrW(30494) & "�q", "�p�M", "�ۤ���", "�i", "�H��", "��", "��", VBA.ChrW(-28903), "��", "�L��M", "����", "�j�֥�", _
        "�����", "�߸`��B", "�߸`��", "�������")
    dict.Add "��", Array("��", "��", "����", "��", "��", "��", "��", "��", "�ڽ�", "����", "�s��", "�f", "�\", "�M", "��", "��", "��", "�~�j��")
    dict.Add "���[", Array("�p��", "�_���s")
    dict.Add "��", Array("���w��", "�v", "�t", "��", "��", "��", "��", "�B", "�S", "�h", VBA.ChrW(-10143) & VBA.ChrW(-8996), "�D", "�n")
    dict.Add "��", Array("��", "��", "�i", "��", "�r", "������", "���", "���_")
    dict.Add VBA.ChrW(21093), Array("��", "��", "�i", "��", "������", "���", "���_")
    dict.Add "�[", Array("�P��", "�{", "��", "��", "��", "����", "����", "�r��", "�i", _
                    "�W��", "����", "���", _
                    "�H�ĻP")
    dict.Add VBA.ChrW(-26587), Array("�P��", "�{", "��", "��", "��", "����", "����", "�r��", "�i", _
                    "�W��", "����", "���", _
                    "�H�ĻP")
    dict.Add "�S", Array("��", "�C��", "��", "�a��", "��", "�B")
    dict.Add VBA.ChrW(14514), Array("��", "�C��", "��", "�a��", "��", "�B")
    dict.Add "�I", Array("��")
    dict.Add VBA.ChrW(20817), Array("��")
    dict.Add VBA.ChrW(20810), Array("��")
    dict.Add "�j�L", Array("�i�L")
    dict.Add "�N", Array("�s", "�]", "�Ѹ�", "�ҡv�@�u", "�ҧ@")
    dict.Add "�", Array("��", "�B", "��", "�_")
    dict.Add "�[", Array("��")
    dict.Add "����", Array("�P", "��")
    dict.Add "�p�L", Array("�O��")
    dict.Add "�A", Array("�Z", "Ĭ", "�{", "�`", "���@", "�h", "��", VBA.ChrW(-24892), "��")
    dict.Add VBA.ChrW(28067), Array("�Z", "Ĭ", "�{", "�`", "���@", "�h", "��", VBA.ChrW(-24892), "��")
    dict.Add "��", Array("��", "��", "�m", "��", "�i�H��", "�D", "�~��", "�U", "����", "�פ�", "�C�J", "�g")
    dict.Add "��", Array("��", "�c", "�ޥG��", "�x")
    dict.Add "��", Array("���v�@�u", "���@")
    dict.Add "�Q", Array("�ݵM����")
    dict.Add "�P�H", Array("��", "�x", "�s��j")
    dict.Add "�j��", Array("����", "���", "�Ӥ�", "�ᦷ", "��")
    dict.Add "����", Array("�O")
    dict.Add "�J��", Array("�H", "�I")
    dict.Add "��E", Array("��")
    dict.Add "�E�G", Array("���|", "��", "���G", "�@�@", "�@��")
    dict.Add "�E�T", Array("��", "���@", "�@��")
    dict.Add "�E�|", Array("��", "���@", "�@��", "�i")
    dict.Add "�E��", Array("��", "���@", "�@��")
    dict.Add "�줻", Array("��")
    dict.Add "���G", Array("��", "���@", "���G")
    dict.Add "���T", Array("��")
    dict.Add "���|", Array("��", "���@", "�K�@�@", "�]�E")
    dict.Add "ν", Array("�i")
    dict.Add "�H��", Array("��", "��", "������")
    dict.Add "�娥", Array("�P�W")
    dict.Add "�b", Array("��", "�A")
    dict.Add "�S", Array("�o", "�k", "��", "�L", "�l", "��", "��", "��", "��", "�o", "��", "�еL", "�ФV�L", "��", "����", "����", "����", _
        "�M", "��", "�N��", "��") ', "����", "�R��", "���H��", "�ા��", "������"
    dict.Add "�L�k", Array("�ۦ�")
    dict.Add VBA.ChrW(26080) & "�k", Array("�ۦ�")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("�ۦ�")
    dict.Add "����", Array("��", "�U")
    dict.Add "�s�N", Array("��")
    dict.Add "���[", Array("����")
    dict.Add "�L�S", Array("��", "�ФV")
    dict.Add "����", Array("�i", "��", "�r")
    dict.Add "���_", Array("�i", "�r")
    dict.Add "����", Array("��", "��", "�r")
    dict.Add "�s��", Array("���")
    dict.Add "��²", Array("Ĭ")
    dict.Add "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("Ĭ")
    dict.Add "���ӱo", Array("�Z�a")
    dict.Add "�i��", Array("�B")
    
    Set ����KeywordsToMark_ExamPrecededAvoid = dict

        
End Property
Rem �Y����r���᭱����O 20240914
Property Get ����KeywordsToMark_ExamFollowedAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "��", Array("�", "��", "���H�ѫ�", "�w�~�h", "�m", "��", "��", "��", "�{", "����", "�}", "�}��", "��", "�m", "�W����", "�R��", "��", "��", _
        "�u", "��", "��", "�n", "�e", "��", "��", "��", "��B", "�W��", "��", "�o����", "�C", "��", "��", "��", "��", "�U", "�t", "��", "�v", "�U��", "��", "�ȩ���", "�G��", _
         "��", "�p��", "���@��", "�ۤ�", "��Q", "����", "���å", "�H���I", "�����", "��������", "�H�ʤH", "�o��", "�Q�m", "����Ϊ̬�", "���h��", "����", _
        "�n����", "������", "�r��", "����", "���̡C�D�����̤]", "���̫D�����̤]", "�Xĳ��", "�h���`", "�ȥ�", "��c", "���ѳ�", "�e����", _
        "�p½", "�p�B", "�l��", "����", "�ƺ�", "������", "���Ӥ���", "�º]", "�o�S�S", "�l�ҧ@", "���ͯ�", "�o��", "�p�R��")
    dict.Add "��", Array("�}", "��", "�ؿv")
    dict.Add "�b", Array("�K")
    dict.Add "��", Array("�b", "��", "��", "��", _
        "�D��", "�D���~", "�D�G�~", "�D�T�~", "�D�|�~", "�D���~", _
        "�M��", "�\", "�[�M��", "�[���M��", "�N")
    dict.Add "���[", Array("���N�H", "�M��", "���M��")
    dict.Add "��", Array("��", "��", "��", "¤", "�h��", "�{", "�v", "��", "�@", "��", VBA.ChrW(20675), "��", "�D", "��", _
        "���n", "�ݬ�����", "��" & VBA.ChrW(29234) & "����", "�l")
    dict.Add "��", Array("��", "�d", "��", "�k", "��", "��", "��", "�a", "���", "�h�z", "�H��", "���H��", "�Ө���", "�K" & VBA.ChrW(-31631), "�K��", "��")
    dict.Add VBA.ChrW(21093), Array("��", "�d", "��", "�k", "��", "��", "��", "�a", "���", "�h�z", "�H��", "���H��", "�Ө���", "�K" & VBA.ChrW(-31631), "�K��", "��")
    dict.Add "�", Array("�", "�j", "��", "��", "�P", "����", "�ݵM", "�w��", "�B", "�Y")
    dict.Add "�A", Array("�M", "��", "�E��")
    dict.Add VBA.ChrW(28067), Array("�M", "��", "�E��")
    dict.Add "�[", Array("��")
    dict.Add "�[", Array("�M��", "����", "�r��", "�Z��", "�r�Z", "�����", "�g", "��", "�u", "�U����", "�ҥH��")
    dict.Add VBA.ChrW(-26587), Array("�M��", "����", "�r��", "�Z��", "�r�Z", "�����", "�g", "��", "�u", "�U����", "�ҥH��")
    dict.Add "��", Array("��", "����", "�_�@", "�j�H", "����", "��a", "��", "�Ѯw", "���@��", "��@��")
    dict.Add "��", Array("��", "��", "�N")
    dict.Add "�N", Array("��", "����")
    dict.Add "�S", Array("��", "��", "�r��", "�礣��", "��K", "����", "�վ�", "�ѬF", "�e��")
    dict.Add VBA.ChrW(14514), Array("��", "��", "�r��", "�礣��", "��K", "����", "�վ�", "�ѬF", "�e��")
    dict.Add "��", Array("��", "���", "�ۮI", "���B", "��ͤl", "�r", "�o��", "�D", "����", "����", "����", "�̬O�]", "��", VBA.ChrW(23280), "�Ҥ��H")
    dict.Add "��", Array("�v�@�u��", "�@��")
    dict.Add "��", Array(VBA.ChrW(22728), "��")
    dict.Add "�Q", Array("��", VBA.ChrW(-28679), "�N", "�@�ׯu")
    dict.Add "�I", Array("�R")
    dict.Add VBA.ChrW(20817), Array("�R")
    dict.Add VBA.ChrW(20810), Array("�R")
    dict.Add "�p�b", Array("��")
    dict.Add "�P�H", Array("���")
    dict.Add "����", Array("�I�l")
    dict.Add "�j��", Array("�\", "�O", "����", "ĵ��", "�w�x", "����", "����")
    dict.Add "�j�L", Array("�H��")
    dict.Add "�p�L", Array("�y�e", "�y" & VBA.ChrW(23515))
    dict.Add "��E", Array("��")
    dict.Add "�줻", Array("��")
    dict.Add "���|", Array("�^")
    dict.Add "�W��", Array("�Q��")
    dict.Add "�ֳ�", Array("�B")
    dict.Add "�S", Array("��", "��", "��", "��", "��")
    dict.Add "���_", Array("�Ǥh")
    dict.Add "���{", Array("��", "�u")
    dict.Add "����", Array("�e��")
    dict.Add "�ޤ�", Array("�L��", "�U��")

    
    Set ����KeywordsToMark_ExamFollowedAvoid = dict

End Property
Rem �Y����r����b�Y�ӻy�y�̭� 20240914
Property Get ����KeywordsToMark_ExamInPhraseAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "��", Array("�����w", "�~����", "Ĭ��²", "����²", "�i����", "������", "�H����", "�H���h", "������", _
        "���s���D", "�`�ѩ���", "�Z���o", "�Z���{", "�g�l����", "�u����", "�����", "������", "�o�ө���", "�H�ө��p", "�H�ө�" & VBA.ChrW(-10155) & VBA.ChrW(-8352), "�H����", "�L�H���]", "������", "�֩���", "�����@", "�W����", "�����o", "�ۦܩ�", "�h����", "�̩��s", "�m����", "���B����", "�Ʃ���", "�����H�a�A", "�ѩ��H", "�����]", "�����k", "�i�H���@��", "�����", "�M������", _
        "�ө���", "�ө��_", "�ө���", "�Щ���", "�h����", "�����䤶", "������]", "�̩��o", "������", "�H���Q��", "�y���G", "�Ʃ��B", "�̩��V", "�ߩ���", "���ߩ���", "�G�仡���t", "�l�����O", "�D���X", "�ө���", "�C�ө���", "�ө��v", "�֩���", "���ө���", _
        "�H����", "�H���B��", "�]������", "�H�����̩���", "�]��������", "�ؤo���H", "�h���o", "�ک��o", "�ȩ���", "�D����", _
        "�ө���", "���h����", "�����e", "�`�w���e", "����e", "�̩���", "�U�h����", "�����Ż���", "�����a", "��H���a", "�󦳩�����", "�ө���", _
        "�ܩ���", "�����]", "�e�F�B���w", "���w�B�Q��", "�e�F���w", "�ܩ���", "���w�Q��", "�X���v", "�̩���", _
        "�e���@", "�f����", "�H���q", "���������", "���t�� ", "�����Y", "�ҥH����", "�̩���", "�O�����]", "������", "�����H����", "�]������", "�ҩ�����", "�ҩ�" & ChrW(-24892) & "��", "�ɩ��H�s", _
        "�h����", "�ѩ��Ӥl", "�V���ӷ~", "������", "��������", "���۩���", "�D���A", "�c����", "�@����", _
        "�]���W��", "�H�r���W", "Ţ����", "������", "�M����H", "�j�֥���", "�H�����~", "�H������", "�������]", _
        "�h���J��", "�ҩ��P��", "�K���g", "�L���i", "����h", "��ө���", _
        "�D����", _
        "�Z����")
    dict.Add "��", Array("�K���s", "�K����")
    dict.Add "��", Array("�j���q", "�Ű���")
    dict.Add "��", Array("�ѭ�ӵo��", "�y�鬥��")
    dict.Add VBA.ChrW(21093), Array("��" & VBA.ChrW(21093) & "�ӵo��", "�y" & VBA.ChrW(21093) & "����")
    dict.Add "��", Array("�ƿݫh", "�H�ݪ�", "�ӿݨD", "��ݰf", "�ݮg��", "��ݹ�")
    dict.Add "�", Array("���", "���")
    dict.Add "�[", Array("���[��", "���[�j�f", "���[���x", "���[��")
    dict.Add VBA.ChrW(-26587), Array("��" & VBA.ChrW(-26587) & "��", "��" & VBA.ChrW(-26587) & "�j�f")
    dict.Add "��", Array("�����ͯ�")
    dict.Add "��", Array("�H�ۯd�H", "�H�ۦ�", "���۪�", "�H�ۤj", "���۪�")
    dict.Add "��", Array("�ȵѤ�")
    dict.Add "�N", Array("�j�N��")
    dict.Add "�S", Array("�Y�S�]", "�F�S�u", "�ܴS�G", "���S��", "���S�q")
    dict.Add VBA.ChrW(14514), Array("�Y" & VBA.ChrW(14514) & "�]", "�F" & VBA.ChrW(14514) & "�u", "��" & VBA.ChrW(14514) & "��", "��" & VBA.ChrW(14514) & "�q")
    dict.Add "�P�H", Array("���P�H��", "���P�H��", "���P�H��")
    dict.Add "�A", Array("���A��")
    dict.Add VBA.ChrW(28067), Array("��" & VBA.ChrW(28067) & "��")
    dict.Add "�j��", Array("�Ƥj����")
    dict.Add "�j�L", Array("�L�j�L�c")
    dict.Add "�J��", Array("�e�J�ٯu")
    dict.Add VBA.ChrW(26083) & "��", Array("�e" & VBA.ChrW(26083) & "�ٯu")
    dict.Add "�L�k", Array("�H�L�k��", "���L�k�O")
    dict.Add VBA.ChrW(26080) & "�k", Array("�H" & VBA.ChrW(26080) & "�k��", "��" & VBA.ChrW(26080) & "�k�O")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("�H" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "��", "��" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "�O")
    dict.Add "��E", Array("����E��")
    dict.Add "�E�G", Array("�@�E�G��")
    dict.Add "�E�T", Array("�ܤE�T�Q")
    dict.Add "�E��", Array("�@�E���E")
    dict.Add "�ΤE", Array("���ΤE��")
    dict.Add "���G", Array("�@���G�|", "�@���G��")
    dict.Add "�W��", Array("�H�W����", "�w�W����")
    dict.Add "�Τ�", Array("�¥Τ���")
    dict.Add "�s��", Array("�ߦs�۷q")
        
    Set ����KeywordsToMark_ExamInPhraseAvoid = dict

End Property
Rem �˴�����r
'Function ����KeywordsToMark_Exam()
'
'End Function
