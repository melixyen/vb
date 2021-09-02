Public Class FormSet

    Public FRTX_Font As Font
    Public FRTX_FrColor As Color
    Public FRTX_BkColor As Color
    Public FRTX_ResColor As Color
    Public FUBD_Font As Font
    Public FUBD_FrColor As Color
    Public FUBD_BkColor As Color

    Private Sub FormSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DefLoad()
        ColorRdemo()
    End Sub

    Private Sub DefLoad()
        FRTX_Font = RTX_Font
        FRTX_FrColor = RTX_FrColor
        FRTX_BkColor = RTX_BkColor
        FRTX_ResColor = RTX_ResColor
        FUBD_Font = UBD_Font
        FUBD_FrColor = UBD_FrColor
        FUBD_BkColor = UBD_BkColor
        CkAutoBak.Checked = AutoBakWork
        TxAutoBakSec.Text = AutoBakSec.ToString
        CkAutoUpVerDet.Checked = AutoUpVerDet
        CkPEye.Checked = PEyeWork
        TxPEyeMin.Text = PEyeMin.ToString
        CkSTime.Checked = STime
        TxBodySet.Text = BSItem
        TxBSBK1.Text = BSItemBK1
        TxBSBK2.Text = BSItemBK2
        TxBSBK3.Text = BSItemBK3

        ToolTip1.SetToolTip(CkAutoBak, FormSet_CkAutoBakTP) '自動備存
        ToolTip1.SetToolTip(CkSTime, FormSet_CkSTimeTP) '顯示計時器
        '人物設定區
        ToolTip1.SetToolTip(BtnBSBK1L, FormSet_BtnBSBK1LTP)
        ToolTip1.SetToolTip(BtnBSBK1C, FormSet_BtnBSBK1CTP)
        ToolTip1.SetToolTip(BtnBSBK1R, FormSet_BtnBSBK1RTP)
        ToolTip1.SetToolTip(BtnBSBK2L, FormSet_BtnBSBK2LTP)
        ToolTip1.SetToolTip(BtnBSBK2C, FormSet_BtnBSBK2CTP)
        ToolTip1.SetToolTip(BtnBSBK2R, FormSet_BtnBSBK2RTP)
        ToolTip1.SetToolTip(BtnBSBK1L, FormSet_BtnBSBK3LTP)
        ToolTip1.SetToolTip(BtnBSBK3C, FormSet_BtnBSBK3CTP)
        ToolTip1.SetToolTip(BtnBSBK3R, FormSet_BtnBSBK3RTP)
    End Sub
    Private Sub ColorRdemo()
        RDemo.SelectAll()
        RDemo.SelectionFont = FRTX_Font
        RDemo.DeselectAll()

        RDemo.ForeColor = FRTX_FrColor
        RDemo.BackColor = FRTX_BkColor

        Dim i As Integer, AwL As String, AwR As String, SIdx As Integer, EndIdx As Integer

        For i = 0 To RecogLeft.Length - 1
            AwL = RecogLeft(i)
            AwR = RecogRight(i)
            SIdx = -1
            Do
                SIdx = RDemo.Text.IndexOf(AwL, SIdx + 1)
                If SIdx <> -1 Then
                    RDemo.SelectionStart = SIdx
                    EndIdx = RDemo.Text.IndexOf(AwR, SIdx)
                    If EndIdx = -1 Then
                        RDemo.SelectionLength = RDemo.Text.Length - SIdx
                        RDemo.SelectionColor = FRTX_ResColor
                    Else
                        RDemo.SelectionLength = EndIdx - SIdx + 1
                        RDemo.SelectionColor = FRTX_ResColor
                    End If
                End If
            Loop Until SIdx = -1 Or EndIdx = -1
        Next
        RDemo.DeselectAll()

        '輔助板設定
        TxUDemo.Font = FUBD_Font
        TxUDemo.ForeColor = FUBD_FrColor
        TxUDemo.BackColor = FUBD_BkColor
    End Sub

    Private Sub BtnFontSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFontSet.Click
        FontD.Font = FRTX_Font
        FontD.ShowDialog()
        FRTX_Font = FontD.Font
        ColorRdemo()
    End Sub

    Private Sub BtnDefSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDefSet.Click
        FRTX_Font = New Font("細明體", 10, FontStyle.Regular)
        FRTX_FrColor = Color.Black
        FRTX_BkColor = Color.White
        FRTX_ResColor = Color.FromArgb(255, 15, 150, 35)
        FUBD_Font = New Font("細明體", 10, FontStyle.Regular)
        FUBD_FrColor = Color.Black
        FUBD_BkColor = Color.White
        CkAutoBak.Checked = False
        CkAutoUpVerDet.Checked = True
        CkPEye.Checked = True
        CkSTime.Checked = True
        ColorRdemo()
    End Sub

    Private Sub BtnFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFC.Click
        ColorD.Color = FRTX_FrColor
        ColorD.ShowDialog()
        FRTX_FrColor = ColorD.Color
        ColorRdemo()
    End Sub

    Private Sub BtnBC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBC.Click
        ColorD.Color = FRTX_BkColor
        ColorD.ShowDialog()
        FRTX_BkColor = ColorD.Color
        ColorRdemo()
    End Sub

    Private Sub BtnRC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRC.Click
        ColorD.Color = FRTX_ResColor
        ColorD.ShowDialog()
        FRTX_ResColor = ColorD.Color
        ColorRdemo()
    End Sub

    Private Sub BtnUBDFontSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUBDFontSet.Click
        FontD.Font = FUBD_Font
        FontD.ShowDialog()
        FUBD_Font = FontD.Font
        ColorRdemo()
    End Sub

    Private Sub BtnUBDFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUBDFC.Click
        ColorD.Color = FUBD_FrColor
        ColorD.ShowDialog()
        FUBD_FrColor = ColorD.Color
        ColorRdemo()
    End Sub

    Private Sub BtnUBDBC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUBDBC.Click
        ColorD.Color = FUBD_BkColor
        ColorD.ShowDialog()
        FUBD_BkColor = ColorD.Color
        ColorRdemo()
    End Sub

    Private Sub BtnSaveSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSaveSet.Click
        RTX_Font = FRTX_Font
        RTX_FrColor = FRTX_FrColor
        RTX_BkColor = FRTX_BkColor
        RTX_ResColor = FRTX_ResColor
        UBD_Font = FUBD_Font
        UBD_FrColor = FUBD_FrColor
        UBD_BkColor = FUBD_BkColor
        AutoBakWork = CkAutoBak.Checked
        AutoBakSec = CInt(TxAutoBakSec.Text)
        AutoUpVerDet = CkAutoUpVerDet.Checked
        PEyeWork = CkPEye.Checked
        PEyeMin = CInt(TxPEyeMin.Text)
        STime = CkSTime.Checked
        BSItem = TxBodySet.Text
        BSItemBK1 = TxBSBK1.Text
        BSItemBK2 = TxBSBK2.Text
        BSItemBK3 = TxBSBK3.Text

        FormEdit.ShowRTXinPG()
        FormEdit.TxUBDSet()
        WriteMemSetting()
        Me.Close()
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub RDemo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RDemo.TextChanged

    End Sub

    Private Sub BtnBSBK1L_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK1L.Click
        TxBodySet.Text = TxBSBK1.Text
    End Sub

    Private Sub BtnBSBK1C_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK1C.Click
        Dim TYT = TxBodySet.Text : TxBodySet.Text = TxBSBK1.Text : TxBSBK1.Text = TYT
    End Sub

    Private Sub BtnBSBK1R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK1R.Click
        TxBSBK1.Text = TxBodySet.Text
    End Sub

    Private Sub BtnBSBK2L_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK2L.Click
        TxBodySet.Text = TxBSBK2.Text
    End Sub

    Private Sub BtnBSBK2C_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK2C.Click
        Dim TYT = TxBodySet.Text : TxBodySet.Text = TxBSBK2.Text : TxBSBK2.Text = TYT
    End Sub

    Private Sub BtnBSBK2R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK2R.Click
        TxBSBK2.Text = TxBodySet.Text
    End Sub

    Private Sub BtnBSBK3L_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK3L.Click
        TxBodySet.Text = TxBSBK3.Text
    End Sub

    Private Sub BtnBSBK3C_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK3C.Click
        Dim TYT = TxBodySet.Text : TxBodySet.Text = TxBSBK3.Text : TxBSBK3.Text = TYT
    End Sub

    Private Sub BtnBSBK3R_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBSBK3R.Click
        TxBSBK3.Text = TxBodySet.Text
    End Sub

    Private Sub BtnDefBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDefBS.Click

        TxBodySet.Text = FormResUse.TxBodySet.Text
    End Sub

End Class