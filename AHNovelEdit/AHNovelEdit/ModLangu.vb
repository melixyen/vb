Module ModLangu
    Public NVTest(19) As String
    Public DotCM(40) As String '���I�Ÿ�
    'Public RecogLeft() As String = {"�u", "�y", "��", "�i", "�m", "�]", "��", Chr(-24153)}
    'Public RecogRight() As String = {"�v", "�z", "��", "�j", "�n", "�^", "��", Chr(-24152)}
    Public RecogLeft() As String = {"�u", "�y", "�i", "�m", "�]", Chr(-24153)}
    Public RecogRight() As String = {"�v", "�z", "�j", "�n", "�^", Chr(-24152)}

    '��r��
    Public Lang_Total = "�`�@"
    Public Lang_At = "�b"
    Public Lang_Area1 = "�a��"
    Public Lang_Ga = "��"
    Public Lang_Finded = "���"
    Public Lang_And1 = "�M"
    Public Lang_And2 = "�åB"
    Public Lang_And3 = "��"
    Public Lang_Replace = "���N"

    'Message ��
    Public Msg_DoYouWantSave As String = "�o�g�p���|���s�ɡA�O�_�n�x�s�O�H"
    Public Msg_YouHadOpenThisNovel As String = "�z�w�g�}�ҳo�g�p���F"
    Public Msg_SearchAtLast As String = "�䤣��U�@�ӤF�A�n���n�q�Y��_�O�H"
    Public Msg_InputJeTell As String = "�п�J�z�n�]�w�������`�����D�ΦW�١A�H���]�w�п�J�z�Q��ܦb�C���W�r�A�p�G�A�Q����Ф��n��J����r" & vbCrLf & vbCrLf & "�p�G��J�k�ܦ��Ǧ�خصL�k������A�Ы��@���������Ϣ�"
    Public Msg_InputJeEnd1 As String = "�ڨS�ݨ�A��J�����`�W�١A����s�W�o"
    Public Msg_DeleteLsJe1 As String = "�A���w�n�R�����аO�L�k�b�峹������A�Э����A�R��"
    Public Msg_DeleteLsJe2 As String = "����Ϫ����`�аO�O���O�A�Q�n�R�����O�H"
    Public Msg_JumpLsJe1 As String = "��A�ڧ䤣��z�ҿ�ܪ��o�Ӥ��`�аO�O"
    Public Msg_OpenTxtBak1 As String = "��A�ڰ�����o���ɮצ��@�ӳƥ���(.nbk)�s�b�A�i��O�e�������`�����{���y�����A�{�b�z���ݰ��@�ӿ��" & vbCrLf & vbCrLf & "��ܬO(YES)�A�ڭ̱N���J�ƥ��ɪ����e�A���i��O�������A�������@�b�����" & vbCrLf & vbCrLf & "��ܧ_(NO)�A�ڭ̱N���J�z����l�ɮסA���i�ఱ�d�b�z�̫�@���s�ɪ����e"
    Public Msg_InputShuQian1 As String = "�аݱz�s�W�����ҭn�Τ���W�r�H"
    Public Msg_BSExpertTitleDo As String = "�O�_�n�[�W���� HTML ���Y<html><body>�������H" & vbCrLf & vbCrLf & "���[���ܶȷ|��X�H���]�w���y�k"
    Public Msg_DeleteArtJe As String = "�A�T�w�n�R�����o�Ӥ��`�Ҧ����e�μ��ҶܡH" & vbCrLf & vbCrLf & "�p�G�n�ֳt�R���i�H�b�I����`�C�����ث���U��L��Delete�䪽���R��"

    'ToolTip��
    Public FormEdit_BtnJeNewTP As String = "�|�b�z����ЩҦb��m���J�@�Ӥ��`�q���O��"
    Public FormEdit_BtnJeDelTP As String = "�N�R���A�b���`�C��W�ҿ�ܪ��Ӥ��`�q���O��"
    Public FormEdit_BtnJeRefreshTP As String = "�p�G�A�b�峹�����F��ʡA�Ы����s��s���`�C��"
    Public FormEdit_BtnJeSelTP As String = "���U�ڥi�H��A��ܪ����`�q���ϰ�峹����_��"
    Public FormEdit_LsJeJumpTP As String = "�I��U�i�H���U�A���D��A�ҿ�ܪ����`�q���h" & vbCrLf & vbCrLf & "�I���A������(�u����)�i�����q���`���һP����" & vbCrLf & vbCrLf & "�I������L�� Delete ��i�H�N��q���`���e(�t���`����)�R��"
    Public FormEdit_HSTabPGTP As String = "�վ���s��ϩM���U�O�Ҥ��t���e�����"
    Public FormEdit_LongSelectToolStripButtonTP As String = "��ثe��Ц�m���U���I����s��إ���@��m�æA�I�@�������s�N�|�N�o�@�q���_��"
    Public FormEdit_FormatAllToolStripButtonTP As String = "�S��������ܷ|�N�峹���Ҧ��r���M��m�Τ@���ثe�t�γ]�w���˦�" & vbCrLf & "��������ܥu�|�N����������r���M��m���w�榡���t�γ]�w���˦�"
    Public FormEdit_RecognitionToolStripButtonTP As String = "�������U�z��ثe�峹�ϥΤ޸������ФW�C����ش���"
    Public FormEdit_BtnPanelCountWordHide As String = "���}�������r�ƭp��A�ݪ`�N���}�ɨϥζǲΤ����J�k�|���L�����r�����"
    Public FormEdit_LbTWCTP As String = "�h���ťթM����r���᪺�`�r��"
    Public FormEdit_LbNWCTP As String = "�h���ťթM����r���H�Τ��`���ҫ᪺�`�r��"
    Public FormEdit_LbSWCTP As String = "����d��h���ťթM����r���᪺�`�r��"

    Public FormSet_CkAutoBakTP As String = "�Ŀ�ڪ��ܡA���s�ɦ�m���ɮרC" & AutoBakSec.ToString & "��b�A��ɾ��|�ƥ��@������ɮ׸�Ƨ��U�����ɦW.nbk��"
    Public FormSet_CkSTimeTP As String = "�Ŀ�ڪ��ܷ|�b�����r�ƭp��ܱz�Ұʥ��n���w�g���奴�F�h�[���ɶ�"
    Public FormSet_BtnBSBK1LTP As String = "�N�ƥζ���1���H���]�w���خ榡���N���ثe�ϥΪ��榡"
    Public FormSet_BtnBSBK1CTP As String = "�洫�A���ƥζ���1���榡�����ϥΪ��榡�A�M���ثe�ϥΪ��榡�s�i�ƥζ���1"
    Public FormSet_BtnBSBK1RTP As String = "�N�ثe�ϥΪ��H���]�w���خ榡�s�i�ƥζ���1"
    Public FormSet_BtnBSBK2LTP As String = "�N�ƥζ���2���H���]�w���خ榡���N���ثe�ϥΪ��榡"
    Public FormSet_BtnBSBK2CTP As String = "�洫�A���ƥζ���2���榡�����ϥΪ��榡�A�M���ثe�ϥΪ��榡�s�i�ƥζ���2"
    Public FormSet_BtnBSBK2RTP As String = "�N�ثe�ϥΪ��H���]�w���خ榡�s�i�ƥζ���2"
    Public FormSet_BtnBSBK3LTP As String = "�N�ƥζ���3���H���]�w���خ榡���N���ثe�ϥΪ��榡"
    Public FormSet_BtnBSBK3CTP As String = "�洫�A���ƥζ���3���榡�����ϥΪ��榡�A�M���ثe�ϥΪ��榡�s�i�ƥζ���3"
    Public FormSet_BtnBSBK3RTP As String = "�N�ثe�ϥΪ��H���]�w���خ榡�s�i�ƥζ���3"




    Public Sub ADDTxTe()
        NVTest(0) = "�A����Label�W�٬OlabEqu�A�M����Label�W�٬OlabClear��L�����ܰT���Ϊ�Label�ڱN�W�٧אּNot�}�l���W�١A�ҦpNotlab�F�A�bForm_Load���Q��addhandler�N�����Q�n�B�z���ƥ�[��B�z�`��"
        Dim i As Integer
        For i = 0 To 18
            'NVTest(i).Text = i.ToString & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text
            NVTest(i) = CStr(i)
        Next
    End Sub

    Public Sub LoadLangu()
        '�w��y���n���J�����ئb���]�w��Τ@���J
        LoadLangu_DotCM()
    End Sub

    Public Sub LoadLangu_DotCM()
        DotCM(0) = "�A" : DotCM(1) = "�C" : DotCM(2) = "�B" : DotCM(3) = "�D"
        DotCM(4) = "�F" : DotCM(5) = "��" : DotCM(6) = "�G" : DotCM(7) = "�H"
        DotCM(8) = "�I" : DotCM(9) = "�K" : DotCM(10) = "��" : DotCM(11) = "�L"
        DotCM(12) = "��" : DotCM(13) = "�V" : DotCM(14) = "��" : DotCM(15) = "��"
        DotCM(16) = "��" : DotCM(17) = "��" : DotCM(18) = "��" : DotCM(19) = "��"
        DotCM(20) = "��" : DotCM(21) = "��" : DotCM(22) = "��" : DotCM(23) = "��"
        DotCM(24) = "��" : DotCM(25) = "�\" : DotCM(26) = "��" : DotCM(27) = "��"
        DotCM(28) = "��" : DotCM(29) = "��" : DotCM(30) = "��" : DotCM(31) = "��"
        DotCM(32) = "�i�j" : DotCM(33) = "�m�n" : DotCM(34) = "�u�v" : DotCM(35) = "�y�z"
        DotCM(36) = "����" : DotCM(37) = "����" : DotCM(38) = "�]�^" : DotCM(39) = Chr(-24153) & Chr(-24152)
    End Sub

    Public Function Make_arBtn_TP(ByVal j As Integer) As String
        Make_arBtn_TP = ""
        If j = 0 Then Make_arBtn_TP = "Ctrl + ,"
        If j = 1 Then Make_arBtn_TP = "Ctrl + ."
        If j = 4 Then Make_arBtn_TP = "Ctrl + ;"
        If j = 39 Then Make_arBtn_TP = "Ctrl + """
        If j = 34 Then Make_arBtn_TP = "Ctrl + Q"
        If j = 35 Then Make_arBtn_TP = "Ctrl + W"
        If j = 2 Then Make_arBtn_TP = "Ctrl + 3"
        If j = 9 Then Make_arBtn_TP = "Ctrl + 4"
    End Function
End Module
