Attribute VB_Name = "Keywords"
Option Explicit
Rem ��������r�˯��B���Ѭ������ݩʡB�ѷӰO��
Dim zhouyiguaShapeNameSequence As Scripting.Dictionary '�o�ӵ��ƬO�T�w���A�ҥH�i�H�p��
Dim zhouyiguaNameShapeSequence As Scripting.Dictionary '�o�ӵ��ƬO�T�w���A�ҥH�i�H�p��
Dim yiVariants As Scripting.Dictionary '�P������r�r��
Dim preceded_Avoid As Scripting.Dictionary '�{�b�٦b�H�ɷs�W���A�G���y�g��'�{�b���D�į�A�٬O���g�A�ϥ�����Word�N�|��s�A�B�ݭn��s�ɥi�H�b�Y�ɹB���������J���O�M���w�������e 20241019
Dim followed_Avoid As Scripting.Dictionary '�{�b�٦b�H�ɷs�W���A�G���y�g��
Dim inPhrase_Avoid As Scripting.Dictionary '�{�b�٦b�H�ɷs�W���A�G���y�g��

Sub ClearDicts_YiKeywords()
    Set preceded_Avoid = Nothing
    Set followed_Avoid = Nothing
    Set inPhrase_Avoid = Nothing

End Sub
Rem �m���n�ǲ���r���/�m����
Property Get ���ǲ���r��() As Scripting.Dictionary
    If yiVariants Is Nothing Then
        Set ���ǲ���r�� = New Scripting.Dictionary
        ���ǲ���r��.Add VBA.ChrW(20089), "��"
        ���ǲ���r��.Add VBA.ChrW(22531), "�["
        ���ǲ���r��.Add VBA.ChrW(-10132) & VBA.ChrW(-8313), "�X"
        ���ǲ���r��.Add VBA.ChrW(-10151) & VBA.ChrW(-9004), "��"
        ���ǲ���r��.Add VBA.ChrW(-29764), "�^"
        ���ǲ���r��.Add VBA.ChrW(24072), "�v"
        ���ǲ���r��.Add VBA.ChrW(-29658), "��"
        ���ǲ���r��.Add VBA.ChrW(-26993), "�H"
        ���ǲ���r��.Add VBA.ChrW(20020), "�{"
        ���ǲ���r��.Add VBA.ChrW(-30270), "�["
        ���ǲ���r��.Add VBA.ChrW(-29390), "�N"
        ���ǲ���r��.Add VBA.ChrW(21093), "��"
        ���ǲ���r��.Add "�`", "�_"
        
        ���ǲ���r��.Add "�L�k", VBA.ChrW(26080) & "�k"
        ���ǲ���r��.Add "�L" & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "�k"
        ���ǲ���r��.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "�k"
        
        ���ǲ���r��.Add VBA.ChrW(-26587), "�["
        ���ǲ���r��.Add "�j" & VBA.ChrW(-28729), "�j�L"
        ���ǲ���r��.Add "��", "��"
        ���ǲ���r��.Add "��", "��"
        
        ���ǲ���r��.Add VBA.ChrW(26187), "��"
        ���ǲ���r��.Add VBA.ChrW(-10164) & VBA.ChrW(-8698), "��"
        
        ���ǲ���r��.Add "��" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "���i"
        
        ���ǲ���r��.Add "��", "��"
        
        ���ǲ���r��.Add VBA.ChrW(-30233), "��"
        ���ǲ���r��.Add VBA.ChrW(25439), "�l"
        ���ǲ���r��.Add VBA.ChrW(28176), "��"
        ���ǲ���r��.Add VBA.ChrW(24402) & "�f", "�k�f"
        ���ǲ���r��.Add "��", "��"
        
        ���ǲ���r��.Add VBA.ChrW(-26520), "�S"
        ���ǲ���r��.Add VBA.ChrW(14514), "�S"
        
        ���ǲ���r��.Add VBA.ChrW(20817), "�I"
        ���ǲ���r��.Add VBA.ChrW(20810), "�I"
        
        ���ǲ���r��.Add VBA.ChrW(28067), "�A"
        ���ǲ���r��.Add VBA.ChrW(-32126), "�`"
        ���ǲ���r��.Add "�p" & VBA.ChrW(-28729), "�p�L"
        
        ���ǲ���r��.Add "�J" & VBA.ChrW(27982), "�J��"
        ���ǲ���r��.Add VBA.ChrW(26083) & "��", "�J��"
        
        ���ǲ���r��.Add "��" & VBA.ChrW(27982), "����"
        
        
        Set yiVariants = ���ǲ���r��
    Else
        Set ���ǲ���r�� = yiVariants
    End If
End Property
Rem key ,string()
Property Get �P������_���W_����() As Scripting.Dictionary
    If zhouyiguaShapeNameSequence Is Nothing Then
        Set �P������_���W_���� = New Scripting.Dictionary
        �P������_���W_����.Add VBA.ChrW(19904), Array("��", 1)
        �P������_���W_����.Add VBA.ChrW(19905), Array("�[", 2)
        �P������_���W_����.Add VBA.ChrW(19906), Array("��", 3)
        �P������_���W_����.Add VBA.ChrW(19907), Array("�X", 4)
        �P������_���W_����.Add VBA.ChrW(19908), Array("��", 5)
        �P������_���W_����.Add VBA.ChrW(19909), Array("�^", 6)
        �P������_���W_����.Add VBA.ChrW(19910), Array("�v", 7)
        �P������_���W_����.Add VBA.ChrW(19911), Array("��", 8)
        �P������_���W_����.Add VBA.ChrW(19912), Array("�p�b", 9)
        �P������_���W_����.Add VBA.ChrW(19913), Array("�i", 10)
        �P������_���W_����.Add VBA.ChrW(19914), Array("��", 11)
        �P������_���W_����.Add VBA.ChrW(19915), Array("�_", 12)
        �P������_���W_����.Add VBA.ChrW(19916), Array("�P�H", 13)
        �P������_���W_����.Add VBA.ChrW(19917), Array("�j��", 14)
        �P������_���W_����.Add VBA.ChrW(19918), Array("��", 15)
        �P������_���W_����.Add VBA.ChrW(19919), Array("��", 16)
        �P������_���W_����.Add VBA.ChrW(19920), Array("�H", 17)
        �P������_���W_����.Add VBA.ChrW(19921), Array("��", 18)
        �P������_���W_����.Add VBA.ChrW(19922), Array("�{", 19)
        �P������_���W_����.Add VBA.ChrW(19923), Array("�[", 20)
        �P������_���W_����.Add VBA.ChrW(19924), Array("����", 21)
        �P������_���W_����.Add VBA.ChrW(19925), Array("�N", 22)
        �P������_���W_����.Add VBA.ChrW(19926), Array("��", 23)
        �P������_���W_����.Add VBA.ChrW(19927), Array("�_", 24)
        �P������_���W_����.Add VBA.ChrW(19928), Array(VBA.ChrW(26080) & "�k", 25)
        �P������_���W_����.Add VBA.ChrW(19929), Array("�j�b", 26)
        �P������_���W_����.Add VBA.ChrW(19930), Array("�[", 27)
        �P������_���W_����.Add VBA.ChrW(19931), Array("�j�L", 28)
        �P������_���W_����.Add VBA.ChrW(19932), Array("��", 29)
        �P������_���W_����.Add VBA.ChrW(19933), Array("��", 30)
        �P������_���W_����.Add VBA.ChrW(19934), Array("�w", 31)
        �P������_���W_����.Add VBA.ChrW(19935), Array("��", 32)
        �P������_���W_����.Add VBA.ChrW(19936), Array("�Q", 33)
        �P������_���W_����.Add VBA.ChrW(19937), Array("�j��", 34)
        �P������_���W_����.Add VBA.ChrW(19938), Array("��", 35)
        �P������_���W_����.Add VBA.ChrW(19939), Array("���i", 36)
        �P������_���W_����.Add VBA.ChrW(19940), Array("�a�H", 37)
        �P������_���W_����.Add VBA.ChrW(19941), Array("��", 38)
        �P������_���W_����.Add VBA.ChrW(19942), Array("�", 39)
        �P������_���W_����.Add VBA.ChrW(19943), Array("��", 40)
        �P������_���W_����.Add VBA.ChrW(19944), Array("�l", 41)
        �P������_���W_����.Add VBA.ChrW(19945), Array("�q", 42)
        �P������_���W_����.Add VBA.ChrW(19946), Array("�[", 43)
        �P������_���W_����.Add VBA.ChrW(19947), Array("�l", 44)
        �P������_���W_����.Add VBA.ChrW(19948), Array("��", 45)
        �P������_���W_����.Add VBA.ChrW(19949), Array("��", 46)
        �P������_���W_����.Add VBA.ChrW(19950), Array("�x", 47)
        �P������_���W_����.Add VBA.ChrW(19951), Array("��", 48)
        �P������_���W_����.Add VBA.ChrW(19952), Array("��", 49)
        �P������_���W_����.Add VBA.ChrW(19953), Array("��", 50)
        �P������_���W_����.Add VBA.ChrW(19954), Array("�_", 51)
        �P������_���W_����.Add VBA.ChrW(19955), Array("��", 52)
        �P������_���W_����.Add VBA.ChrW(19956), Array("��", 53)
        �P������_���W_����.Add VBA.ChrW(19957), Array("�k�f", 54)
        �P������_���W_����.Add VBA.ChrW(19958), Array("��", 55)
        �P������_���W_����.Add VBA.ChrW(19959), Array("��", 56)
        �P������_���W_����.Add VBA.ChrW(19960), Array("�S", 57)
        �P������_���W_����.Add VBA.ChrW(19961), Array("�I", 58)
        �P������_���W_����.Add VBA.ChrW(19962), Array("�A", 59)
        �P������_���W_����.Add VBA.ChrW(19963), Array("�`", 60)
        �P������_���W_����.Add VBA.ChrW(19964), Array("����", 61)
        �P������_���W_����.Add VBA.ChrW(19965), Array("�p�L", 62)
        �P������_���W_����.Add VBA.ChrW(19966), Array("�J��", 63)
        �P������_���W_����.Add VBA.ChrW(19967), Array("����", 64)
        Set zhouyiguaShapeNameSequence = �P������_���W_����
    Else
        Set �P������_���W_���� = zhouyiguaShapeNameSequence
    End If
End Property
Rem key ,string()
Property Get �P�����W_����_����() As Scripting.Dictionary
    On Error GoTo eH:
    If zhouyiguaNameShapeSequence Is Nothing Then
        Set �P�����W_����_���� = New Scripting.Dictionary
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19904), 1)
        �P�����W_����_����.Add "�[", Array(VBA.ChrW(19905), 2)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19906), 3)
        �P�����W_����_����.Add "�X", Array(VBA.ChrW(19907), 4)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19908), 5)
        �P�����W_����_����.Add "�^", Array(VBA.ChrW(19909), 6)
        �P�����W_����_����.Add "�v", Array(VBA.ChrW(19910), 7)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19911), 8)
        �P�����W_����_����.Add "�p�b", Array(VBA.ChrW(19912), 9)
        �P�����W_����_����.Add "�i", Array(VBA.ChrW(19913), 10)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19914), 11)
        �P�����W_����_����.Add "�_", Array(VBA.ChrW(19915), 12)
        �P�����W_����_����.Add "�P�H", Array(VBA.ChrW(19916), 13)
        �P�����W_����_����.Add "�j��", Array(VBA.ChrW(19917), 14)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19918), 15)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19919), 16)
        �P�����W_����_����.Add "�H", Array(VBA.ChrW(19920), 17)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19921), 18)
        �P�����W_����_����.Add "�{", Array(VBA.ChrW(19922), 19)
        �P�����W_����_����.Add "�[", Array(VBA.ChrW(19923), 20)
        �P�����W_����_����.Add "����", Array(VBA.ChrW(19924), 21)
        �P�����W_����_����.Add "�N", Array(VBA.ChrW(19925), 22)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19926), 23)
        �P�����W_����_����.Add "�_", Array(VBA.ChrW(19927), 24)
        �P�����W_����_����.Add VBA.ChrW(26080) & "�k", Array(VBA.ChrW(19928), 25)
        �P�����W_����_����.Add "�j�b", Array(VBA.ChrW(19929), 26)
        �P�����W_����_����.Add "�[", Array(VBA.ChrW(19930), 27)
        �P�����W_����_����.Add "�j�L", Array(VBA.ChrW(19931), 28)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19932), 29)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19933), 30)
        �P�����W_����_����.Add "�w", Array(VBA.ChrW(19934), 31)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19935), 32)
        �P�����W_����_����.Add "�Q", Array(VBA.ChrW(19936), 33)
        �P�����W_����_����.Add "�j��", Array(VBA.ChrW(19937), 34)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19938), 35)
        �P�����W_����_����.Add "���i", Array(VBA.ChrW(19939), 36)
        �P�����W_����_����.Add "�a�H", Array(VBA.ChrW(19940), 37)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19941), 38)
        �P�����W_����_����.Add "�", Array(VBA.ChrW(19942), 39)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19943), 40)
        �P�����W_����_����.Add "�l", Array(VBA.ChrW(19944), 41)
        �P�����W_����_����.Add "�q", Array(VBA.ChrW(19945), 42)
        �P�����W_����_����.Add "�[", Array(VBA.ChrW(19946), 43)
        �P�����W_����_����.Add "�l", Array(VBA.ChrW(19947), 44)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19948), 45)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19949), 46)
        �P�����W_����_����.Add "�x", Array(VBA.ChrW(19950), 47)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19951), 48)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19952), 49)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19953), 50)
        �P�����W_����_����.Add "�_", Array(VBA.ChrW(19954), 51)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19955), 52)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19956), 53)
        �P�����W_����_����.Add "�k�f", Array(VBA.ChrW(19957), 54)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19958), 55)
        �P�����W_����_����.Add "��", Array(VBA.ChrW(19959), 56)
        �P�����W_����_����.Add "�S", Array(VBA.ChrW(19960), 57)
        �P�����W_����_����.Add "�I", Array(VBA.ChrW(19961), 58)
        �P�����W_����_����.Add "�A", Array(VBA.ChrW(19962), 59)
        �P�����W_����_����.Add "�`", Array(VBA.ChrW(19963), 60)
        �P�����W_����_����.Add "����", Array(VBA.ChrW(19964), 61)
        �P�����W_����_����.Add "�p�L", Array(VBA.ChrW(19965), 62)
        �P�����W_����_����.Add "�J��", Array(VBA.ChrW(19966), 63)
        �P�����W_����_����.Add "����", Array(VBA.ChrW(19967), 64)
        Set zhouyiguaNameShapeSequence = �P�����W_����_����
    Else
        Set �P�����W_����_���� = zhouyiguaNameShapeSequence
    End If
    Exit Property
eH:
    Select Case Err.Number
        Case Else
            Debug.Print Err.Number & Err.Description
            Stop
    End Select
End Property
Rem �ΥH�ˬd�O�_�����ǽd�򤧤��e��
Property Get ����Keywords_ToCheck() As Variant 'string()
    ����Keywords_ToCheck = Array(VBA.ChrW(-10119), VBA.ChrW(-8742), VBA.ChrW(-30233), VBA.ChrW(-10164), VBA.ChrW(-8698), VBA.ChrW(-31827), VBA.ChrW(-10132), VBA.ChrW(-8313), VBA.ChrW(20810), VBA.ChrW(-10167), VBA.ChrW(-8698), VBA.ChrW(-26587), VBA.ChrW(21093), VBA.ChrW(14615), VBA.ChrW(20089), VBA.ChrW(26080), "�k", VBA.ChrW(26083), "��" _
        , "��", "�P", VBA.ChrW(20089), "��", "��", "�p�b", "�i", "�{", "�[", "�j�L", "�[", "��", "�_", "����", "�N", "��", "��", "�X", "�P�H", "�j��", "��", "�_", "��", "��", "�^", "��", "��", "�L�k", "�j�b", "�v", "��", "�H", "��", "�[", "�w", "��", "�l", "�q", "�_", "��", "����", "�Q", "�j��", "�[", "�l", "��", "�k�f", "�p�L", "��", "���i", "��", "��", "��", "��", "�J��", "����", "�a�H", "��", "�x", "��", "�S", "�I", "�", "��", "��", "��", "�A", "�`", "�ӷ�", "����", "���", "�H", VBA.ChrW(-10145) & VBA.ChrW(-9156), "ν", _
        "�ѳ�", "�Ѷ�", "�ֳ�", "�ֶ�", "�")
End Property
Rem �ΥH���ѩ�������r��
Property Get ����Keywords_ToMark() As Variant 'string()�]�� Array Returns a Variant containing an array,�ҥH����g�� as string()
    ����Keywords_ToMark = Array("��", "�P��", "���g", "�j��", "���g", "���g", "�C�g", "�Q�T�g", "�", _
        "��", "�`��", "����", "�{��", "�ٻX", "��" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "�١B�X", "�١B" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "����", "�[��", "�^��", "�v��", "���", "�i��", "����", "�_��", "����", "�H��", "�[��", "�_��", "�ߧ�", "����", "�w��", "���", "�ʨ�", "�a�H��", "�Ѩ�", "�l��", "�q��", "�ɨ�", "�x��", "����", "����", "����", "�_��", "����", "�ר�", "�Ȩ�", "�`��", "��", "�t��", "ô��", "����", "����", "ô��", "����", "�Ǩ�", "����", "�Ԩ�", "����", "�娥", "���[", "����", "�Q�s", "�v�O", "�b", "�[", "�����j�l", "�[�@����", "���H����", "�[�H²��", "��", "�q���r", "�q�[�r", "���B�[", "�q���B�[�r", "����", "�N�_�~", "�N��~", "�~�N", "���N", "�N", VBA.ChrW(20089), "�J��", VBA.ChrW(26083) & "��", "����", "�Q�l", "�j" & VBA.ChrW(22766), _
        "��E", "�E�G", "�E�T", "�E�|", "�E��", "�W�E", VBA.ChrW(19972) & "�E", "�ΤE", "�줻", "���G", "���T", "���|", "����", "�W��", "�Τ�", "�e��", "����", "�ӷ�", "�L��", "���", _
            "�H��", "�q�H�r��", "�H��", "�H��", "�H��", VBA.ChrW(-10145) & VBA.ChrW(-9156) & "��", "�j�H", "�j" & VBA.ChrW(-10145) & VBA.ChrW(-9156), "�p�H", "�H�q", "�|�H", "�H�G", "�H��", "ν", _
             "���I", "��", "�[", "�P�H�_�v", "�P�H", "��", "����", "�I", VBA.ChrW(20817), VBA.ChrW(20810), "��", "�l", "�S", VBA.ChrW(14514), VBA.ChrW(-26520), "��", VBA.ChrW(21093), "�Q�@�L�e", "�Q�@" & ChrW(26080) & "�e", "�Q", "�j��", "���i", "��" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "�p�b", "�j�b", "��", "�", "�A", VBA.ChrW(28067), "��", "��", "�k�f", "�p�L", "�j��", "�j�L", "�q���r", "�q�_�r", "�q�l�r", "�q�q�r", "�q��", "�X�r", VBA.ChrW(-10132) & VBA.ChrW(-8313) & "�r", "�ݤj", "��", "�q�^�k�r", "�q�_�r", "�q�_�r", "�q�ݡr", "�ѳ�", "��" & VBA.ChrW(-27006), "�Ѷ�", "�ֳ�", "��" & VBA.ChrW(-27006), "�ֶ�", "����", "�ٵ�", _
            "�S", "����", VBA.ChrW(24451) & "��", "���[", VBA.ChrW(24451) & "�[", "�w��", "�w��", "�j�l", "�H�]��", VBA.ChrW(-10145) & VBA.ChrW(-9156) & "�]��", "�w���E��", "����", "�F�F", "�F" & VBA.ChrW(-26973), "��" & VBA.ChrW(-10155) & VBA.ChrW(-8630), "��" & VBA.ChrW(-10131) & VBA.ChrW(-8268), "�����n", "�ѤU�p��", "���B", "���s�b��", "ۤڮ", VBA.ChrW(-31656) & "ڮ", _
        "���a��", "�L�k", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "�k", "�L�S", VBA.ChrW(26080) & "�S", "�ѩS", "����", "�٨�I", "�צ�", "����", "�Ѧ氷", "�Ѧ�" & VBA.ChrW(24484), "���X", "��" & VBA.ChrW(-10123) & VBA.ChrW(-8628), "��" & VBA.ChrW(14444), "���l��", "���l��", "�����G�@", "�����G" & VBA.ChrW(21323), "�����G" & VBA.ChrW(19991), "�����G�W", "�Ѥ@�a�G", "�Y��", VBA.ChrW(-31867) & "��", VBA.ChrW(-31867) & VBA.ChrW(-30650), "�Y" & VBA.ChrW(-30650), "�F��", "�Τ�", "�Y�Q", "���{", "����", "����", _
        "�H�ɤ��q", "������", "�]����", "��q�J��", "�q�G�ީ]", _
        "�F��", "����", "�Ӥ���", "�p�b�ѤW", "����", "���f", "�ޤ�", "�T��", "�g��", "����", "����", "�g��", "�q�H" & VBA.ChrW(-10114) & VBA.ChrW(-8896) & "��", "�q�H��" & VBA.ChrW(20869), "�q�H����", "�q�H��~", "�g��o�D", "�Q��n", "�~���w��", "�ѤU�j��", "�q�ʦ�", "��i�Læ", "��i" & VBA.ChrW(26080) & "æ", "�W�S", "�W" & VBA.ChrW(14514), "�b��", "�W�_", "�~��", "�s��", "����", "���[", "����", "���U��", VBA.ChrW(-10139) & VBA.ChrW(-8938) & "�U��", VBA.ChrW(-10139) & VBA.ChrW(-8937) & "�U��", "�X���Q�s", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "���Q�s", "�񤧭�H", "�i�s", "�s�F", "�i�D�Z�Z", "�s�N", "�s��", "����", "��W����", "���Ƥ��J", "���Ƥ�" & VBA.ChrW(30694), "���|���", "��" & VBA.ChrW(23577) & "���", VBA.ChrW(-25895) & VBA.ChrW(23577) & "���", VBA.ChrW(-25895) & "�|���", "�ҥ��U��", "���ӱo", "����" & VBA.ChrW(-10167) & VBA.ChrW(-8906), "�s���r", "�ߦ���", "�P�a��", _
        "�A�n", "�L�e", VBA.ChrW(26080) & "�e", "���`", "��" & VBA.ChrW(20158), "��" & VBA.ChrW(20838), "�ɸq", "����", "�����ӥ~��", "�����~��", "�~���Ӥ���", "�~������", "��²", "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "���_", "�}������", "�a������", "��X���`", "���`��X", "��X", "�����h�E", "���L�h��", "�E����L", "�i��", "���Y", "�@" & VBA.ChrW(-27006) & "�@��", "�@���@��", "�ڦ��n��", "������", "���t�H���D�|", "���l�Ӯv", "�̤l�֤r", "��ΦӤ���", "��Τ���", "���D�A", "��l�ϲ�", "�M����", "�P�ӹE�q", "�B�q", "�B�r", "�e���b��", "�e���b" & VBA.ChrW(-30650), "�i��", "�i��", "���{", "�{�j�g", "�ܤƤ�", VBA.ChrW(-10164) & VBA.ChrW(-9163) & "�Ƥ�", "���D�]��", "���D�]" & VBA.ChrW(25934), "�q�Ӧ���", VBA.ChrW(-24871) & "�Ӧ���", "�����ӫH", "�s�G�w��", "�q�ѤU����", "�i��", "�~���̵�", "���̨���", "���̨���", "���̨���", "�Z��", VBA.ChrW(-10114) & VBA.ChrW(-9019) & "��", "�z�]", "����", "�T�����D", _
        "�j�s", "�p�s", "�ҥX�G�_", "�ҥX��_", "�ҥX�_�_", "�P�ɰ���", "�յ�", "��" & VBA.ChrW(-31142), "��" & VBA.ChrW(-10119) & VBA.ChrW(-8991), "��" & VBA.ChrW(-31145), "�i��", "��䭭", "�D��", "�C��", "�C��", "�]�X", "���X", "�X�N", "�]" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "��" & VBA.ChrW(-10132) & VBA.ChrW(-8313), VBA.ChrW(-10132) & VBA.ChrW(-8313) & "�N", "�T�G�䤣�i��", "���G�䤣�i��", _
        "�Ѧb�s��", "�h�ѫe��" & VBA.ChrW(24451) & "��", "�h�ѫe������", "��", "��`", "��" & VBA.ChrW(-29005), "����", "�s�A�q��", "���", "�~��", "�s�w", "�V���y", VBA.ChrW(24892) & "���y", "�`����", "��R�N", VBA.ChrW(25914) & "�R�N", "�}��Ӯa", "�¤�", VBA.ChrW(26087) & "��", VBA.ChrW(-10149) & VBA.ChrW(-8300) & "��", "����", _
        "���R", "���R", "���R", VBA.ChrW(29234) & "�R", "�N��C��", "�N�_�C��", "�N�����", "�N�_����", "�N��" & VBA.ChrW(-10176) & VBA.ChrW(-9207) & "��", "�N�_" & VBA.ChrW(-10176) & VBA.ChrW(-9207) & "��", "���ѫ��", "������W", "�਩", VBA.ChrW(17966) & "��", VBA.ChrW(-10172) & VBA.ChrW(-9052) & "��", VBA.ChrW(-10124) & VBA.ChrW(-8660) & "��", VBA.ChrW(-10173) & VBA.ChrW(-8748) & "��", VBA.ChrW(20007) & "��", VBA.ChrW(-10173) & VBA.ChrW(-8650) & "��", "����", "��" & VBA.ChrW(-32119), "Ĳÿ", "Ĳ��", "�ߨ��", "�߸�", "����߸�", "����߸�", _
        "�N��" & VBA.ChrW(-10176) & VBA.ChrW(-9204) & "��", "�N�_" & VBA.ChrW(-10176) & VBA.ChrW(-9204) & "��", "�N��" & VBA.ChrW(-10143) & VBA.ChrW(-8559) & "��", "�N�_" & VBA.ChrW(-10143) & VBA.ChrW(-8559) & "��", "�d��", "�}�s", "����", "�Q��", "��", "����", _
        "����", "̱̱", "�l�U�H�q�W", "�l�U�q�W", "�l�U�ӯq�W", "�L�Φ�", "�Ǭ��ۼ�", "�Q���j�H", "����{", "�P�k�Ӯ��", "�@�P�Ӧʼ{", "�P�k���", "�@�P�ʼ{", "��𬰪�", "�̨���", "�Ƥ��K", "�Ƥ�" & VBA.ChrW(23483), "�਩", VBA.ChrW(-10173) & VBA.ChrW(-8748) & "��", "��s", _
        "����", "�C" & VBA.ChrW(19487) & "����", "��" & VBA.ChrW(19487) & "����", "���p", "�I�p", "�B�`", "�A��s", VBA.ChrW(28067) & "��s", "�A��" & VBA.ChrW(32675), VBA.ChrW(28067) & "��" & VBA.ChrW(32675), "�ҤT��", "���T��", "�ɨ䰪��", "�ɨ�" & VBA.ChrW(-25895) & "��", "�ѹD����", "����", "�鮬", "����", "��", _
        "��", "��" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "��" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "����", "�W��", "�W" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "�W��", "����", "����")
        
End Property
Rem �Y����r���e������O 20240914
Property Get ����KeywordsToMark_Exam_Preceded_Avoid() As Scripting.Dictionary
    If preceded_Avoid Is Nothing Then
        
        Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
        ' �K�[��ƨ�r�� creedit_with_Copilot�j���ġGhttps://sl.bing.net/goDF239cQVw
        dict.Add "��", Array("��", "�e", "��", "��", "��", "��", "�T", "��", "��", "��", "�Υi", "��", "��", "�@", "��", "�C", "�U", "�Z", "��", "��", "�y", "��", "��", "��", "²", "�թ~", "�~", "�L", "��", "�I", "��", "����", "�@", "��", "��", _
            "��", "�g", "������", "�E", "�C���@", "��", "�g", "�Ǽg���K", "���x�B��", "���x��", "���H�M", "�@", "ź", "�@��@", "��M", _
            "�s����", "����", "�W�U��", "�Z�B", "�B", "�мY��", "������", "��������������", "�����{", "�Ƥ[�h�{", "�թT", "�ߪ���]", "�ߪ�" & VBA.ChrW(25143) & "�]", _
            "ť����", "�a�u�q", "�M�S", "�r�M", "�a" & VBA.ChrW(30494) & "�q", "�p�M", "�ۤ���", "�i", "�H��", "��", "��", VBA.ChrW(-28903), "��", "�L��M", "����", "�j�֥�", _
            "�����", "�߸`��B", "�߸`��", "�������", "���w", "���ɵL", "��ť�̤�", "������", "�A����", "�ӥ~�M", "�P�ר���", "�H�ѤU�P�H", "�פ��P")
            
        dict.Add "��", Array("��", "��", "��", "����", "��", "��", "��", "��", "��", "�ڽ�", "����", "�s��", "�f", "�\", "�M", "��", "��", "��", "�~�j��")
        dict.Add "���[", Array("�p��", "�_���s", "���u")
        dict.Add "��", Array("���w��", "�v", "�t", "��", "��", "��", "��", "��", "�B", "�S", "�h", VBA.ChrW(-10143) & VBA.ChrW(-8996), "�D", "�n", "����", "��", VBA.ChrW(25135))
        dict.Add "��", Array("��", "��", "�i", "��", "�r", "������", "���", "���_")
        dict.Add VBA.ChrW(21093), Array("��", "��", "�i", "��", "������", "���", "���_")
        dict.Add "�[", Array("�P��", "�{", "��", "��", "��", "����", "����", "�r��", "�i", "��", _
                        "�W��", "����", "���", _
                        "�H�ĻP")
        dict.Add VBA.ChrW(-26587), Array("�P��", "�{", "��", "��", "��", "����", "����", "�r��", "�i", "��", _
                        "�W��", "����", "���", _
                        "�H�ĻP")
                        
        dict.Add "�S", Array("��", "�C��", "��", "�a��", "��", "�B", "��", "����")
        dict.Add VBA.ChrW(14514), Array("��", "�C��", "��", "�a��", "��", "�B", "��", "����")
        
        dict.Add "�I", Array("��")
        dict.Add VBA.ChrW(20817), Array("��")
        dict.Add VBA.ChrW(20810), Array("��")
        
        dict.Add "�j�L", Array("�i�L", "����L")
        dict.Add "�N", Array("��", "�s", "�]", "�Ѹ�", "�ҡv�@�u", "�ҧ@", "�F����")
        dict.Add "�", Array("��", "�B", "��", "�_", "������")
        dict.Add "�[", Array("��")
        dict.Add "����", Array("�P", "��")
        dict.Add "�p�L", Array("�O��")
        
        dict.Add "�A", Array("�Z", "Ĭ", "�{", "�`", "���@", "�h", "��", VBA.ChrW(-24892), "��", "�A", "��", "��")
        dict.Add VBA.ChrW(28067), Array("�Z", "Ĭ", "�{", "�`", "���@", "�h", "��", VBA.ChrW(-24892), "��", "�A", "��", "��")
        
        dict.Add "��", Array("��", "��", "�m", "��", "�i�H��", "�D", "�~��", "�U", "����", "�פ�", "�C�J", "�g", "��", "��")
        dict.Add "��", Array("��", "�c", "�ޥG��", "�x", "����", "��")
        dict.Add "��", Array("���v�@�u", "���@")
        dict.Add "��", Array("�H��")
        dict.Add "�Q", Array("�ݵM����", "�l�y�t")
        dict.Add "�P�H", Array("��", "�x", "�s��j")
        dict.Add "�j��", Array("����", "���", "�Ӥ�", "�ᦷ", "��")
        dict.Add "����", Array("�O")
        dict.Add "�J��", Array("�H", "�I")
        dict.Add "��E", Array("��")
        dict.Add "�E�G", Array("���|", "��", "���G", "�@�@", "�@��")
        dict.Add "�E�T", Array("��", "���@", "�@��")
        dict.Add "�E�|", Array("��", "���@", "�@��", "�i")
        dict.Add "�E��", Array("��", "���@", "�@��", "�@��")
        dict.Add "�줻", Array("��")
        dict.Add "���G", Array("��", "���@", "���G")
        dict.Add "���T", Array("��")
        dict.Add "���|", Array("��", "���@", "�K�@�@", "�]�E")
        dict.Add "����", Array("��", "���@", "�K�@�@", "�]�E")
        dict.Add "ν", Array("�i")
        dict.Add "�H��", Array("��", "��", "������")
        dict.Add "�娥", Array("�P�W", "�B�W")
        dict.Add "�b", Array("��", "�A")
        dict.Add "�S", Array("�o", "�k", "��", "�L", "�l", "��", "��", "��", "��", "�o", "��", "�еL", "�ФV�L", "��", "����", "����", "����", "�ת̬�", _
            "�M", "��", "�N��", "��", "�}�y", "�A�{��", "���ޤ�", "�e�o�K") ', "����", "�R��", "���H��", "�ા��", "������"
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
        
        dict.Add "�յ�", Array("��뤧", "�ʥ@")
        dict.Add "��" & VBA.ChrW(-31142), Array("��뤧", "�ʥ@")
        dict.Add "��" & VBA.ChrW(-10119) & VBA.ChrW(-8991), Array("��뤧", "�ʥ@")
        dict.Add "��" & VBA.ChrW(-31145), Array("��뤧", "�ʥ@")
        
        dict.Add "�s��", Array("���")
        dict.Add "�L��", Array("�P��", "�@�t", "�A��")
        dict.Add "��²", Array("Ĭ")
        dict.Add "��" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("Ĭ")
        dict.Add "���ӱo", Array("�Z�a")
        dict.Add "����" & VBA.ChrW(-10167) & VBA.ChrW(-8906), Array("�Z�a")
        dict.Add "�i��", Array("�B")
        dict.Add "���B", Array(VBA.ChrW(32675) & "��")
        
        Set preceded_Avoid = dict
        Set ����KeywordsToMark_Exam_Preceded_Avoid = dict
    Else
        Set ����KeywordsToMark_Exam_Preceded_Avoid = preceded_Avoid
    End If
        
End Property
Rem �Y����r���᭱����O 20240914
Property Get ����KeywordsToMark_Exam_Followed_Avoid() As Scripting.Dictionary
    If followed_Avoid Is Nothing Then
        Dim dict As New Scripting.Dictionary
        dict.Add "��", Array("�", "��", "���H�ѫ�", "�w�~�h", "�m", "��", "��", "��", "�{", "����", "�}", "�}��", "��", "�m", "�W����", "�R��", "��", "��", _
            "�u", "��", "��", "�n", "�e", "��", "��", "��", "��", "��B", "�W��", "��", "�o����", "�C", "��", "��", "��", "��", "�U", "�t", "��", "�v", "�U��", "��", "�ȩ���", "�G��", _
             "��", "�p��", "���@��", "�Ǧ�", "�ۤ�", "��Q", "����", "���å", "�@����", "�H���I", "�����", "��������", "�H�ʤH", "�o��", "�Q�m", "����Ϊ̬�", "���h��", "����", _
            "�n����", "�����w", "������", "�r��", "����", "�N�ӧ�", "��V��", VBA.ChrW(24422) & "��", "�۳�", "�H��", "���̡C�D�����̤]", "���̫D�����̤]", "�Xĳ��", "�h���`", "�ȥ�", "��c", "���ѳ�", "�e����", "�@�ӫ�", "���k", _
            "�p½", "�p�B", "�l��", "����", "�a�Ҧw", "���D��", "���׶�", "������", "�ƺ�", "������", "���Ӥ���", "�º]", "�o�S�S", "�l�ҧ@", "���ͯ�", "�o��", "�p�R��", "�m�䦸", "�ܪq��", "��]")
        
        dict.Add "�P��", Array("è")
        dict.Add "��", Array("�}", "��", "�ؿv")
        dict.Add "�b", Array("�K")
        dict.Add "��", Array("�b", "��", "��", "��", _
            "�D��", "�D���~", "�D�G�~", "�D�T�~", "�D�|�~", "�D���~", "�D�C��", _
            "�M��", "�\", "�[�M��", "�[���M��", "�N", "���@" & VBA.ChrW(-28146) & "��")
        dict.Add "���[", Array("���N�H", "�M��", "���M��")
        dict.Add "��", Array("��", "��", "��", "¤", "�h��", "�{", "�v", "��", "�@", "��", VBA.ChrW(20675), "��", "�D", "��", "�~��", _
            "���n", "�ݬ�����", "��" & VBA.ChrW(29234) & "����", "�l", "�R�a")
        dict.Add "��", Array("��", "�d", "��", "�k", "��", "��", "��", "�a", "���", "�h�z", "�H��", "���H��", "�Ө���", "�K" & VBA.ChrW(-31631), "�K��", "��")
        dict.Add VBA.ChrW(21093), Array("��", "�d", "��", "�k", "��", "��", "��", "�a", "���", "�h�z", "�H��", "���H��", "�Ө���", "�K" & VBA.ChrW(-31631), "�K��", "��")
        dict.Add "�", Array("�", "�j", "��", "��", "�P", "�Ǩ�", "����", "�ݵM", "�w��", "�B", "�Y")
        dict.Add "�A", Array("�M", "��", "�E��", "�����O", "�B��")
        dict.Add VBA.ChrW(28067), Array("�M", "��", "�E��", "�����O", "�B��")
        dict.Add "�[", Array("��")
        dict.Add "�[", Array("�M��", "����", "�r��", "�Z��", "�r�Z", "�����", "�g", "��", "�u", "��", "�U����", "�ҥH��")
        dict.Add VBA.ChrW(-26587), Array("�M��", "����", "�r��", "�Z��", "�r�Z", "�����", "�g", "��", "�u", "��", "�U����", "�ҥH��")
        dict.Add "��", Array("��", "����", "�_�@", "�j�H", "����", "��a", "��", "�Ѯw", "���@��", "��@��", "��", "�F��", "�|", "�s�w��", "�ʤ��P�v", "������")
        dict.Add "��", Array("��", "��", "�N")
        dict.Add "�N", Array("��", "����", "�x���N", "Ū����", "�r", "�x")
        
        dict.Add "�S", Array("��", VBA.ChrW(23891), "��", "�r��", "�礣��", "��K", "����", "�վ�", "�ѬF", "�e��")
        dict.Add VBA.ChrW(14514), Array("��", "��", "�r��", "�礣��", "��K", "����", "�վ�", "�ѬF", "�e��")
        
        dict.Add "��", Array("�b�H��", "��", "���", "�ۮI", "���B", "��ͤl", "�r", "�o��", "�D", "����", "����", "����", "�̬O�]", "��", VBA.ChrW(23280), "�Ҥ��H")
        dict.Add "��", Array("�v�@�u��", "�@��", "�j", "��")
        dict.Add "��", Array(VBA.ChrW(22728), "��")
        dict.Add "�Q", Array("��", VBA.ChrW(-28679), "�N", "�@�ׯu", "�@����")
        dict.Add "�I", Array("�R")
        dict.Add VBA.ChrW(20817), Array("�R")
        dict.Add VBA.ChrW(20810), Array("�R")
        dict.Add "�p�b", Array("��")
        dict.Add "�P�H", Array("���")
        dict.Add "����", Array("�I�l")
        dict.Add "�j��", Array("�\", "�O", "����", "ĵ��", "�w�x", "����", "����", "�Ʀb", "�j�n", "�]��")
        dict.Add "�L�k", Array("�Q")
        dict.Add VBA.ChrW(26080) & "�k", Array("�Q")
        dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("�Q")
        dict.Add "�j�L", Array("�H��")
        dict.Add "�p�L", Array("�y�e", "�y" & VBA.ChrW(23515))
        dict.Add "��E", Array("��")
        dict.Add "�줻", Array("��")
        dict.Add "���|", Array("�^")
        dict.Add "�W��", Array("�Q��")
        dict.Add "�娥", Array("���t")
        dict.Add "�ֳ�", Array("�B")
        dict.Add "�S", Array("��", "��", "��", "��", "��", "�b�H����", VBA.ChrW(-10172) & VBA.ChrW(-8632) & "�H����", "�x����")
        dict.Add "���_", Array("�Ǥh")
        dict.Add "���{", Array("��", "�u")
        dict.Add "�b��", Array("��")
        dict.Add "����", Array("�e��")
        dict.Add "�ޤ�", Array("�L��", "�U��")
        dict.Add "�j�l", Array("��")

        
        Set ����KeywordsToMark_Exam_Followed_Avoid = dict
        Set followed_Avoid = dict
    Else
        Set ����KeywordsToMark_Exam_Followed_Avoid = followed_Avoid
    End If
End Property
Rem �Y����r����b�Y�ӻy�y�̭� 20240914
Property Get ����KeywordsToMark_Exam_InPhrase_Avoid() As Scripting.Dictionary
    If inPhrase_Avoid Is Nothing Then
        Dim dict As New Scripting.Dictionary
        dict.Add "��", Array("�����w", "�~����", "Ĭ��²", "����²", "�i����", "������", "�H����", "�H���h", "������", "�̩��c", "�H������v", "�H�v����", _
            "���s���D", "�`�ѩ���", "�P��è", "�Z���o", "�Z���{", "�g�l����", "�u����", "�����", "������", "���H����", "�o�ө���", "�H�ө��p", "�H�ө�" & VBA.ChrW(-10155) & VBA.ChrW(-8352), "�H����", "�L�H���]", "������", "�֩���", "�����@", "�W����", "�����o", "�ۦܩ�", "�h����", "�̩��s", "�m����", "���B����", "�Ʃ���", "�����H�a�A", "�ѩ��H", "�����]", "�����k", "�i�H���@��", "�����", "�M������", "�����", "�ߩ���", "ļ�ө��E", "�L�ө���", _
            "�ө���", "�ө��_", "²�ө���", "�ө���", "�Щ���", "�h����", "�����䤶", "������]", "�̩��o", "�䱡����", "������", "�H���Q��", "�y���G", "�Ʃ��B", "�̩��V", "�ߩ���", "���ߩ���", "�G�仡���t", "�l�����O", "�D���X", "�ө���", "�C�ө���", "�ө��v", "�֩���", "���ө���", "��������", "�̩��o��", _
            "�H����", "�H���B��", "�ϩ��ѩ�", "�]������", "�H�����̩���", "�]��������", "�ؤo���H", "�ȩ��D��", "�@����g", "�h���o", "�ک��o", "�ȩ���", "�H�����", "�D����", "������", "���H������", "�������", "²�n����", _
            "�ө���", "�ҩ����H", "�Ʀb���ӨD", "���h����", "�����e", "������j", "�`�w���e", "����e", "�N���^", "�̩���", "�U�h����", "�����Ż���", "�����a", "��H���a", "�󦳩�����", "�ө���", _
            "�ܩ���", "�����]", "�e�F�B���w", "���w�B�Q��", "�Z������", "�e�F���w", "�ܩ���", "���v����", "���w�Q��", "�X���v", "�̩���", _
            "�e���@", "�f����", "�����`��", "�L���A�O�H���]", "�H���q", "��H����", "���������", "�����", "�������", "���t�� ", "�����Y", "�ҥH����", "�̩���", "�O�����]", "������", "�����H����", "�]������", "�ҩ�����", "�ҩ�" & ChrW(-24892) & "��", "�ɩ��H�s", _
            "��H���ѤU", "�h����", "������g�l", "�ѩ��Ӥl", "�V���ӷ~", "�g�v����", "������", "��������", "�k�M����", "���۩���", "�D���A", "�c����", "�@����", "�u����a", "�T���䭵", _
            "�]���W��", "�H�r���W", "���ө��o", "Ţ����", "�פH���o", "�G����]", "������", "�c�X���K", "�M����H", "�H�����", "�j�֥���", "�H�����~", "�H������", "�������]", "��ө���", _
            "�h���J��", "�ҩ��P��", "�������q", "����S", "�K���g", "�L���i", "����h", "��ө���", "������", "�W�_����", "�ˤ_����", "�۫H�A���H���\", _
            "�D����", "�����󤺪�", "���̩���", VBA.ChrW(-30681) & "�̩���", "��������", "�ө����ݸ�", "�H���ةm", "�ߦ~���W", "�����", "������o", VBA.ChrW(27507) & "����o", _
            "�Z����", "�H�H����", "�H�����j", "�ع��̩��_��W", "�h���H���_", "�Ө����" & VBA.ChrW(-10175) & VBA.ChrW(-8614))
        
        dict.Add "��", Array("�K���s", "�K����")
        dict.Add "��", Array("�p����r")
        dict.Add "��", Array("�O����", "�j���q", "�Ű���", "�B����", "���u���[")
        dict.Add "��", Array("�ѭ�ӵo��", "�y�鬥��")
        dict.Add VBA.ChrW(21093), Array("��" & VBA.ChrW(21093) & "�ӵo��", "�y" & VBA.ChrW(21093) & "����")
        dict.Add "��", Array("�ƿݫh", "�d�v�ݨo", "�H�ݪ�", "�ӿݨD", "��ݰf", "�ݮg��", "��ݹ�", "�ݦ{�C�ݡC�Τ]")
        dict.Add "�", Array("���", "���", "��Hr", "ź���")
        
        dict.Add "�[", Array("���[��", "���[�j�f", "���[���x", "���[��", "���[�Ӯ�")
        dict.Add VBA.ChrW(-26587), Array("��" & VBA.ChrW(-26587) & "��", "��" & VBA.ChrW(-26587) & "�j�f", "��" & VBA.ChrW(-26587) & "�Ӯ�")
        
        dict.Add "��", Array("�����ͯ�")
        dict.Add "��", Array("�H�ۯd�H", "�H�ۦ�", "���۪�", "�H�ۤj", "���۪�", "�����v")
        dict.Add "��", Array("�ȵѤ�")
        dict.Add "�N", Array("�j�N��", "���ѶN�r", "�s�N��", "�]�N��")
        
        dict.Add "�S", Array("�Y�S�]", "�F�S�u", "�ܴS�G", "���S��", "���S�q")
        dict.Add VBA.ChrW(14514), Array("�Y" & VBA.ChrW(14514) & "�]", "�F" & VBA.ChrW(14514) & "�u", "��" & VBA.ChrW(14514) & "��", "��" & VBA.ChrW(14514) & "�q")
        
        dict.Add "�P�H", Array("���P�H��", "���P�H��", "���P�H��", "�ҫ��P�H�A����", "�n���H�v�P�H�h")
        dict.Add "�[", Array("�Q�C�[��")
        
        dict.Add "�A", Array("���A��", "�O�A��", "�w�A��")
        dict.Add VBA.ChrW(28067), Array("��" & VBA.ChrW(28067) & "��", "�O" & VBA.ChrW(28067) & "��", "�w" & VBA.ChrW(28067) & "��")
        
        dict.Add "�j��", Array("�Ƥj����", "�Ӥj���")
        dict.Add "�j��", Array("���j����")
        dict.Add "�j�L", Array("�L�j�L�c")
        dict.Add "�p�L", Array("�Ѥp�L�H��", VBA.ChrW(14592) & "�p�L�H��")
        
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
        dict.Add "���G", Array("�@���G�|", "�@���G��", "�����G��")
        dict.Add "���|", Array("�@���|�K")
        dict.Add "�W��", Array("�H�W����", "�w�W����", "�H�W���Q")
        dict.Add "�Τ�", Array("�¥Τ���")
        dict.Add "�娥", Array("�Ӧ��娥��", "�j�娥�y�`", "�j�娥�`")
        dict.Add "�s��", Array("�ߦs�۷q")
        dict.Add "���X", Array("�����X��")
                    
        Set ����KeywordsToMark_Exam_InPhrase_Avoid = dict
        Set inPhrase_Avoid = dict
    Else
        Set ����KeywordsToMark_Exam_InPhrase_Avoid = inPhrase_Avoid
    End If
End Property
Rem �˴�����r
'Function ����KeywordsToMark_Exam()
'
'End Function

