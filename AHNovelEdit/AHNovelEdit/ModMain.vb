Module ModMain
    Public Structure NovelFile
        Public AskSave As Boolean '�����ܥ��s�ɮɫK�]�� True �����s��
        Public NeedBak As Boolean '�����ܻݳƥ��ɫK�]�� True �۰ʳƥ�
        Public FilePath As String '�ɮ��x�s���|
        Public FileName As String '�ɮ��x�s�ɦW
        Public PMenu As ToolStripItem '���ҥ\���
        Public FmtU8 As Boolean '��H UTF8 �榡�x�s�ɳ]�� True
    End Structure

    '����������s����
    Public WithEvents Sck1 As System.Net.Sockets.Socket
    'Public Htp1 As System.Windows.Forms.WebBrowser = New WebBrowser
    Public CanLinkUpVer As Boolean = False
    Public VerDetHost As String = "blog.yam.com"
    Public VerDetPath As String = "/melix/article/25910610"
    Public VerDetURL As String = "http://" & VerDetHost & VerDetPath '��������ɮ׺��}URL
    Public VerDetSplit1 As String = "YGYGYGYGYGYGAHNE" '��������1
    Public VerDetSplit2 As String = "TNTNTNTNTNTN" '����������2
    '��k:���YG�Ҧb�B,����TN������,�A�N�o�r��H�_�����}�Y�i

    Public RegPath As String = "Software\AuroraHacker\AHNocelEdit" '�n���ɥؿ�

    Public SystemSet As Boolean = False '��o�A�t���ܧ�ɱN���L��True,�ƫ�A��^�ӥH�קK�@�ǵ{���~�P
    Public NSplitL As String = "<ahne[tag:split]" '�峹���Τ��`�аO���_
    Public NSplitR As String = "[end:split]>" '�峹���Τ��`�аO�k��
    '���νd�� <ahne[tag:split]001�^��5�`[end:split]> �����Y���`�W

    '����
    Public ShuQian(0) As ToolStripItem '����
    Public BkMk As String = "" '�w�]���Ҧr����
    Public QLoad(0) As ToolStripItem '�ָ��\���

    '�H���]�w�����ܼ�
    Public BSItem As String = FormResUse.TxBodySet.Text '�H�]���عw�]��
    Public BSItemBK1 As String = FormResUse.TxBodySet.Text '�H�]���سƥέ�1
    Public BSItemBK2 As String = FormResUse.TxBodySet.Text '�H�]���سƥέ�2
    Public BSItemBK3 As String = FormResUse.TxBodySet.Text '�H�]���سƥέ�3
    Public BSqtL As String = "�i" '�H���]�w���ؤ��e���_�Ÿ�
    Public BSqtR As String = "�j" '�H���]�w���ؤ��e�k��Ÿ�
    Public BSqtM As String = "�G" '�H���]�w���ئW�٤��O��
    Public BSHide As String = "*" '�H���]�w���ػ����ḛ̂O��
    Public HtmlHeadA As String = "<html><head><title>�H���]�w</title><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""></head><body>" & vbCrLf & "<!-- Start Expert -->" & vbCrLf
    Public HtmlDLast As String = vbCrLf & "</body></html>" & vbCrLf


    '�M��Τ����ܼ�
    Public FindWord As String
    Public ReplaceWord As String
    Public FindUp As Boolean
    Public LastFind() As String = {""}
    Public LastReplace() As String = {""}

    '�]�w�����۰ʤ�
    Public AutoBakSec As Integer = 100 '�]�w�̵u��Ƶo��Ĳ�o�ɦ۰ʳƥ��@��
    Public AutoBakWork As Boolean = False '�]�w�O�_�n����۰ʳƦs���u�@
    Public AutoUpVerDet As Boolean = True '�]�w�O�_�n�۰��ˬd��s����
    Public PEyeMin As Integer = 30 '�O�@���O����������
    Public PEyeWork As Boolean = True '�O�@���O�O�_�}��
    Public STime As Boolean = True '����p�ɾ��O�_���

    '�̪�ϥά���
    Public RecentFile As String = "" '�̪񪺤����ɮ�



    'Public RTX_Font = New Font("Lucida Console�s�ө���Fixedsys", 12, FontStyle.Regular)
    Public RTX_Font As Font = New Font("�ө���", 10, FontStyle.Regular)
    Public RTX_FrColor As Color = Color.Black
    Public RTX_BkColor As Color = Color.White
    Public RTX_ResColor As Color = Color.FromArgb(255, 15, 150, 35)
    Public RTX_Left As Integer = 5 '174
    Public RTX_Top As Integer = 6 '106
    Public RTX_AutoWordSelection As Boolean = False
    Public RTX_Height_OFFSET As Integer = 35
    Public RTX_Width_OFFSET As Integer = 20 '400
    Public RTX_ImeMode As ImeMode = Windows.Forms.ImeMode.On

    'Public NVE_Font = New Font("Lucida Console", 12, FontStyle.Regular)
    'Public NVE_Left = 5 '174
    'Public NVE_Top = 6 '106
    'Public NVE_AutoWordSelection = False
    'Public NVE_Height_OFFSET = 40
    'Public NVE_Width_OFFSET = 20 '400
    'Public NVE_ImeMode = Windows.Forms.ImeMode.On
    Public UBD_Font As Font = New Font("�ө���", 10, FontStyle.Regular)
    Public UBD_FrColor As Color = Color.Black
    Public UBD_BkColor As Color = Color.White

    Public Sub MOD_START()
        'Mod �����ݥ��Ұʰ��檺�{��
    End Sub

    Public Sub ReadSetting()
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "TT1", Nothing) Is Nothing Then
            WriteMemSetting()
        Else
            Dim RF_Name As String = "�ө���", RF_Size As Integer = 10, RF_Style As Integer = 0
            Dim UF_Name As String = "�ө���", UF_Size As Integer = 12, UF_Style As Integer = 0
            RF_Name = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RTX_Font_Name", RTX_Font.Name)
            RF_Size = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RTX_Font_Size", RTX_Font.Size)
            RF_Style = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RTX_Font_Style", RTX_Font.Style)
            RTX_Font = New Font(RF_Name, RF_Size, CType(RF_Style, FontStyle))
            RTX_FrColor = Color.FromArgb(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RTX_FrColor", RTX_FrColor))
            RTX_BkColor = Color.FromArgb(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RTX_BkColor", RTX_BkColor))
            RTX_ResColor = Color.FromArgb(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RTX_ResColor", RTX_ResColor))


            UF_Name = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBD_Font_Name", UBD_Font.Name)
            UF_Size = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBD_Font_Size", UBD_Font.Size)
            UF_Style = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBD_Font_Style", UBD_Font.Style)
            UBD_Font = New Font(UF_Name, UF_Size, CType(UF_Style, FontStyle))
            UBD_FrColor = Color.FromArgb(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBD_FrColor", UBD_FrColor))
            UBD_BkColor = Color.FromArgb(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBD_BkColor", UBD_BkColor))

            AutoBakSec = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "AutoBakSec", AutoBakSec)
            AutoBakWork = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "AutoBakWork", AutoBakWork)
            AutoUpVerDet = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "AutoUpVerDet", AutoUpVerDet)
            PEyeMin = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "PEyeMin", PEyeMin)
            PEyeWork = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "PEyeWork", PEyeWork)
            STime = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "STime", STime)

            BSItem = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "BSItem", BSItem)
            BSItemBK1 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "BSItemBK1", BSItemBK1)
            BSItemBK2 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "BSItemBK2", BSItemBK2)
            BSItemBK3 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "BSItemBK3", BSItemBK3)
        End If
    End Sub

    Public Sub WriteMemSetting()
        Dim FRPath As String = "HKEY_CURRENT_USER\" & RegPath
        My.Computer.Registry.CurrentUser.CreateSubKey(RegPath)
        My.Computer.Registry.SetValue(FRPath, "TT1", "yes")
        My.Computer.Registry.SetValue(FRPath, "RTX_Font_Name", RTX_Font.Name)
        My.Computer.Registry.SetValue(FRPath, "RTX_Font_Size", CInt(RTX_Font.Size))
        My.Computer.Registry.SetValue(FRPath, "RTX_Font_Style", CInt(RTX_Font.Style))
        My.Computer.Registry.SetValue(FRPath, "RTX_FrColor", RTX_FrColor.ToArgb)
        My.Computer.Registry.SetValue(FRPath, "RTX_BkColor", RTX_BkColor.ToArgb)
        My.Computer.Registry.SetValue(FRPath, "RTX_ResColor", RTX_ResColor.ToArgb)

        My.Computer.Registry.SetValue(FRPath, "UBD_Font_Name", UBD_Font.Name)
        My.Computer.Registry.SetValue(FRPath, "UBD_Font_Size", CInt(UBD_Font.Size))
        My.Computer.Registry.SetValue(FRPath, "UBD_Font_Style", CInt(UBD_Font.Style))
        My.Computer.Registry.SetValue(FRPath, "UBD_FrColor", UBD_FrColor.ToArgb)
        My.Computer.Registry.SetValue(FRPath, "UBD_BkColor", UBD_BkColor.ToArgb)

        My.Computer.Registry.SetValue(FRPath, "AutoBakSec", AutoBakSec)
        My.Computer.Registry.SetValue(FRPath, "AutoBakWork", AutoBakWork)
        My.Computer.Registry.SetValue(FRPath, "AutoUpVerDet", AutoUpVerDet)
        My.Computer.Registry.SetValue(FRPath, "PEyeMin", PEyeMin)
        My.Computer.Registry.SetValue(FRPath, "PEyeWork", PEyeWork)
        My.Computer.Registry.SetValue(FRPath, "STime", STime)

        My.Computer.Registry.SetValue(FRPath, "BSItem", BSItem)
        My.Computer.Registry.SetValue(FRPath, "BSItemBK1", BSItemBK1)
        My.Computer.Registry.SetValue(FRPath, "BSItemBK2", BSItemBK2)
        My.Computer.Registry.SetValue(FRPath, "BSItemBK3", BSItemBK3)

    End Sub

    Public Sub ReadReginfo()
        '�Ψ�Ū���ﶵ�H�~����
        RecentFile = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RecentFile", RecentFile) '�̪񪺤����ɮ�
        FormEdit.TxUBD.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBDText", "") '���U�O��r
        FormEdit.TxUBD.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBDVis", FormEdit.TxUBD.Visible) '���U�O�O�_���
        FormEdit.WindowState = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "WindowStateA", FormEdit.WindowState) '�����O�_���̤j
        FormEdit.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "WindowWidthA", FormEdit.Width) '�����e��
        FormEdit.Height = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "WindowHeightA", FormEdit.Height) '��������
        FormEdit.TabPG.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "TabpgWidthA", FormEdit.TabPG.Width) 'TabPG �e��
        BkMk = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "BkMk", BkMk) '����
    End Sub

    Public Sub WriteReginfo()
        Dim FRPath As String = "HKEY_CURRENT_USER\" & RegPath
        My.Computer.Registry.SetValue(FRPath, "RecentFile", RecentFile)
        My.Computer.Registry.SetValue(FRPath, "UBDText", FormEdit.TxUBD.Text)
        My.Computer.Registry.SetValue(FRPath, "UBDVis", FormEdit.TxUBD.Visible)
        My.Computer.Registry.SetValue(FRPath, "WindowStateA", CInt(FormEdit.WindowState))
        My.Computer.Registry.SetValue(FRPath, "WindowWidthA", CInt(FormEdit.Width))
        My.Computer.Registry.SetValue(FRPath, "WindowHeightA", CInt(FormEdit.Height))
        My.Computer.Registry.SetValue(FRPath, "TabpgWidthA", CInt(FormEdit.TabPG.Width))
        My.Computer.Registry.SetValue(FRPath, "BkMk", BkMk)
    End Sub

    Public Sub RightMenuQuickAR(Optional ByVal Rmv As Boolean = False)
        Dim FRPath As String = "HKEY_CLASSES_ROOT\*\shell\AHNovelEdit"
        Dim FGPath As String = FRPath & "\command"
        Dim AppP As String = Application.ExecutablePath
        Dim ApName As String = "�� AH Novel Edit ���}"
        If Rmv = False Then
            My.Computer.Registry.SetValue(FRPath, "", ApName) '�g�J�w�]��
            My.Computer.Registry.SetValue(FRPath & "\command", "", """" & AppP & """ ""%1""") '�g�J�w�]��
        Else
            Try
                My.Computer.Registry.ClassesRoot.DeleteSubKeyTree("*\shell\AHNovelEdit")
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub VerDet(Optional ByVal ByHand As Boolean = False)
        'AddHandler Htp1.DocumentCompleted, AddressOf HTP_RECODE
        'Htp1.Navigate(VerDetURL)
        '==================================�W�� Browser �j�k , �U�� Socket ��k=======================

        Dim SHeader As String
        Sck1 = New System.Net.Sockets.Socket(Net.Sockets.AddressFamily.InterNetwork, Net.Sockets.SocketType.Stream, Net.Sockets.ProtocolType.IP)
        Dim UHtml As String = ""
        Dim URlng As Integer
        Dim URstt(1023) As Byte
        Dim i As Integer, RcvT As String = ""
        Sck1.ReceiveTimeout = 100

        '��s Firefox HTTP Header
        SHeader = "GET " & VerDetPath & " HTTP/1.1" & vbCrLf
        SHeader = SHeader & "Host: " & VerDetHost & vbCrLf
        SHeader = SHeader & "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-TW; rv:1.9.0.15) Gecko/2009101601 Firefox/3.0.15" & vbCrLf
        SHeader = SHeader & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" & vbCrLf
        SHeader = SHeader & "Accept-Language: zh-tw,en-us;q=0.7,en;q=0.3" & vbCrLf
        SHeader = SHeader & "Accept-Encoding: gzip,deflate" & vbCrLf
        SHeader = SHeader & "Accept-Charset: Big5,utf-8;q=0.7,*;q=0.7" & vbCrLf
        SHeader = SHeader & "Keep-Alive: 300" & vbCrLf
        SHeader = SHeader & "Connection: keep-alive" & vbCrLf
        SHeader = SHeader & "" & vbCrLf

        Try

            Sck1.Connect(VerDetHost, 80)
            If Sck1.Connected = True Then
                Sck1.Send(System.Text.Encoding.Default.GetBytes(SHeader))
            End If

        Catch ex As Exception

        End Try
        Try
            Do
                URlng = Sck1.Receive(URstt, URstt.Length, Net.Sockets.SocketFlags.None)
                RcvT = System.Text.Encoding.UTF8.GetString(URstt)
                UHtml = UHtml & (RcvT)
                Application.DoEvents()
            Loop While (URlng > 4)
        Catch ex As Exception

        End Try
        Sck1.Close()

        If UHtml <> "" Then
            Try
                Dim SystemVerAry() As String = My.Application.Info.Version.ToString.Split(".")
                Dim SystemVer As Double = CDbl(SystemVerAry(0) & "." & SystemVerAry(1))
                Dim UTo1 As Integer = UHtml.IndexOf(VerDetSplit1)
                Dim UTo2 As Integer = UHtml.IndexOf(VerDetSplit2, UTo1)
                Dim UMRsv As String = UHtml.Substring(UTo1, UTo2 - UTo1)
                Dim USPLF() As String = UMRsv.Split(":")
                Dim CkVerDA As Double = CDbl(USPLF(1))
                Dim CompaV As Double = CkVerDA - SystemVer
                Dim RSQTX As String = "�z�t�Υثe���������G" & SystemVer.ToString & vbCrLf & vbCrLf & "�����W�̷s���������G" & USPLF(1) & vbCrLf & vbCrLf
                If CompaV > 0 Then
                    MessageBox.Show(RSQTX & "��ĳ�z�ߧY���@�����к��@ BLOG�@��s�̷s�����I")
                Else
                    If ByHand = True Then MessageBox.Show(RSQTX & "�z�ثe���������ݭn��s�I")
                End If

            Catch ex As Exception
                If ByHand = True Then MessageBox.Show("�۰��ˬd��s�s�u���ѡI")
            End Try
        End If


    End Sub

    Public Sub HTP_RECODE(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        '�o�@�ӬO�������ˬd�� Webbrowser ��@�ɥΪ� function
        'If CanLinkUpVer = True Then
        '    Dim getReader As New System.IO.StreamReader(Htp1.DocumentStream, System.Text.Encoding.UTF8)
        '    Htp1.Stop()
        '    Dim UHtml As String = getReader.ReadToEnd
        '    Dim UTo1 As Integer = UHtml.IndexOf(VerDetSplit1)
        '    Dim UTo2 As Integer = UHtml.IndexOf(VerDetSplit2, UTo1)
        '    Dim UMRsv As String = UHtml.Substring(UTo1, UTo2 - UTo1)
        '    Dim USPLF() As String = UMRsv.Split(":")
        '    Dim CkVerDA As Double = CDbl(USPLF(1))
        '    Dim CompaV As Double = CkVerDA - SystemVer
        '    Dim RSQTX As String = "�z�t�Υثe���������G" & SystemVer.ToString & vbCrLf & vbCrLf & "�����W�̷s���������G" & USPLF(1) & vbCrLf & vbCrLf
        '    If CompaV > 0 Then
        '        MessageBox.Show(RSQTX & "��ĳ�z�ߧY���@�����к��@ BLOG�@��s�̷s�����I")
        '    Else
        '        MessageBox.Show(RSQTX & "�z�ثe���������ݭn��s�I")
        '    End If
        '    CanLinkUpVer = False
        'End If
    End Sub

End Module

