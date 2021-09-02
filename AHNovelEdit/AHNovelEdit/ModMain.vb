Module ModMain
    Public Structure NovelFile
        Public AskSave As Boolean '當有改變未存檔時便設為 True 提醒存檔
        Public NeedBak As Boolean '當有改變需備份時便設為 True 自動備份
        Public FilePath As String '檔案儲存路徑
        Public FileName As String '檔案儲存檔名
        Public PMenu As ToolStripItem '頁籤功能表
        Public FmtU8 As Boolean '當以 UTF8 格式儲存時設為 True
    End Structure

    '網路版本更新偵測
    Public WithEvents Sck1 As System.Net.Sockets.Socket
    'Public Htp1 As System.Windows.Forms.WebBrowser = New WebBrowser
    Public CanLinkUpVer As Boolean = False
    Public VerDetHost As String = "blog.yam.com"
    Public VerDetPath As String = "/melix/article/25910610"
    Public VerDetURL As String = "http://" & VerDetHost & VerDetPath '版本比對檔案網址URL
    Public VerDetSplit1 As String = "YGYGYGYGYGYGAHNE" '分切標籤1
    Public VerDetSplit2 As String = "TNTNTNTNTNTN" '分切內標籤2
    '找法:找到YG所在處,取到TN結尾區,再將這字串以冒號切開即可

    Public RegPath As String = "Software\AuroraHacker\AHNocelEdit" '登錄檔目錄

    Public SystemSet As Boolean = False '當牽涉系統變更時將布林為True,事後再改回來以避免一些程式誤判
    Public NSplitL As String = "<ahne[tag:split]" '文章切割分節標記左起
    Public NSplitR As String = "[end:split]>" '文章切割分節標記右止
    '切割範例 <ahne[tag:split]001回第5節[end:split]> 中間即為節名

    '書籤
    Public ShuQian(0) As ToolStripItem '書籤
    Public BkMk As String = "" '預設書籤字串檔
    Public QLoad(0) As ToolStripItem '快載功能表

    '人物設定相關變數
    Public BSItem As String = FormResUse.TxBodySet.Text '人設項目預設值
    Public BSItemBK1 As String = FormResUse.TxBodySet.Text '人設項目備用值1
    Public BSItemBK2 As String = FormResUse.TxBodySet.Text '人設項目備用值2
    Public BSItemBK3 As String = FormResUse.TxBodySet.Text '人設項目備用值3
    Public BSqtL As String = "【" '人物設定項目內容左起符號
    Public BSqtR As String = "】" '人物設定項目內容右止符號
    Public BSqtM As String = "：" '人物設定項目名稱分記號
    Public BSHide As String = "*" '人物設定項目需隱藏者記號
    Public HtmlHeadA As String = "<html><head><title>人物設定</title><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""></head><body>" & vbCrLf & "<!-- Start Expert -->" & vbCrLf
    Public HtmlDLast As String = vbCrLf & "</body></html>" & vbCrLf


    '尋找用公用變數
    Public FindWord As String
    Public ReplaceWord As String
    Public FindUp As Boolean
    Public LastFind() As String = {""}
    Public LastReplace() As String = {""}

    '設定提醒自動化
    Public AutoBakSec As Integer = 100 '設定最短秒數發生觸發時自動備份一次
    Public AutoBakWork As Boolean = False '設定是否要執行自動備存的工作
    Public AutoUpVerDet As Boolean = True '設定是否要自動檢查更新版本
    Public PEyeMin As Integer = 30 '保護視力提醒分鐘數
    Public PEyeWork As Boolean = True '保護視力是否開啟
    Public STime As Boolean = True '打文計時器是否顯示

    '最近使用相關
    Public RecentFile As String = "" '最近的五個檔案



    'Public RTX_Font = New Font("Lucida Console新細明體Fixedsys", 12, FontStyle.Regular)
    Public RTX_Font As Font = New Font("細明體", 10, FontStyle.Regular)
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
    Public UBD_Font As Font = New Font("細明體", 10, FontStyle.Regular)
    Public UBD_FrColor As Color = Color.Black
    Public UBD_BkColor As Color = Color.White

    Public Sub MOD_START()
        'Mod 中必需先啟動執行的程式
    End Sub

    Public Sub ReadSetting()
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "TT1", Nothing) Is Nothing Then
            WriteMemSetting()
        Else
            Dim RF_Name As String = "細明體", RF_Size As Integer = 10, RF_Style As Integer = 0
            Dim UF_Name As String = "細明體", UF_Size As Integer = 12, UF_Style As Integer = 0
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
        '用來讀取選項以外的值
        RecentFile = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "RecentFile", RecentFile) '最近的五個檔案
        FormEdit.TxUBD.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBDText", "") '輔助板文字
        FormEdit.TxUBD.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "UBDVis", FormEdit.TxUBD.Visible) '輔助板是否顯示
        FormEdit.WindowState = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "WindowStateA", FormEdit.WindowState) '視窗是否放到最大
        FormEdit.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "WindowWidthA", FormEdit.Width) '視窗寬度
        FormEdit.Height = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "WindowHeightA", FormEdit.Height) '視窗高度
        FormEdit.TabPG.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "TabpgWidthA", FormEdit.TabPG.Width) 'TabPG 寬度
        BkMk = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & RegPath, "BkMk", BkMk) '書籤
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
        Dim ApName As String = "用 AH Novel Edit 打開"
        If Rmv = False Then
            My.Computer.Registry.SetValue(FRPath, "", ApName) '寫入預設值
            My.Computer.Registry.SetValue(FRPath & "\command", "", """" & AppP & """ ""%1""") '寫入預設值
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
        '==================================上為 Browser 大法 , 下為 Socket 方法=======================

        Dim SHeader As String
        Sck1 = New System.Net.Sockets.Socket(Net.Sockets.AddressFamily.InterNetwork, Net.Sockets.SocketType.Stream, Net.Sockets.ProtocolType.IP)
        Dim UHtml As String = ""
        Dim URlng As Integer
        Dim URstt(1023) As Byte
        Dim i As Integer, RcvT As String = ""
        Sck1.ReceiveTimeout = 100

        '仿製 Firefox HTTP Header
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
                Dim RSQTX As String = "您系統目前的版本為：" & SystemVer.ToString & vbCrLf & vbCrLf & "網站上最新的版本為：" & USPLF(1) & vbCrLf & vbCrLf
                If CompaV > 0 Then
                    MessageBox.Show(RSQTX & "建議您立即按　說明－維護 BLOG　更新最新版本！")
                Else
                    If ByHand = True Then MessageBox.Show(RSQTX & "您目前的版本不需要更新！")
                End If

            Catch ex As Exception
                If ByHand = True Then MessageBox.Show("自動檢查更新連線失敗！")
            End Try
        End If


    End Sub

    Public Sub HTP_RECODE(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        '這一個是給版本檢查用 Webbrowser 實作時用的 function
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
        '    Dim RSQTX As String = "您系統目前的版本為：" & SystemVer.ToString & vbCrLf & vbCrLf & "網站上最新的版本為：" & USPLF(1) & vbCrLf & vbCrLf
        '    If CompaV > 0 Then
        '        MessageBox.Show(RSQTX & "建議您立即按　說明－維護 BLOG　更新最新版本！")
        '    Else
        '        MessageBox.Show(RSQTX & "您目前的版本不需要更新！")
        '    End If
        '    CanLinkUpVer = False
        'End If
    End Sub

End Module

