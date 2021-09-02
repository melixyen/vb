<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSet
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.TabSet = New System.Windows.Forms.TabControl
        Me.TPgRTX = New System.Windows.Forms.TabPage
        Me.BtnRC = New System.Windows.Forms.Button
        Me.BtnBC = New System.Windows.Forms.Button
        Me.BtnFC = New System.Windows.Forms.Button
        Me.BtnFontSet = New System.Windows.Forms.Button
        Me.RDemo = New System.Windows.Forms.RichTextBox
        Me.TpgUBD = New System.Windows.Forms.TabPage
        Me.BtnUBDBC = New System.Windows.Forms.Button
        Me.BtnUBDFC = New System.Windows.Forms.Button
        Me.BtnUBDFontSet = New System.Windows.Forms.Button
        Me.TxUDemo = New System.Windows.Forms.TextBox
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.CkAutoUpVerDet = New System.Windows.Forms.CheckBox
        Me.CkSTime = New System.Windows.Forms.CheckBox
        Me.TxPEyeMin = New System.Windows.Forms.TextBox
        Me.CkPEye = New System.Windows.Forms.CheckBox
        Me.TxAutoBakSec = New System.Windows.Forms.TextBox
        Me.CkAutoBak = New System.Windows.Forms.CheckBox
        Me.TpgBodySet = New System.Windows.Forms.TabPage
        Me.BtnDefBS = New System.Windows.Forms.Button
        Me.FLBSBK3 = New System.Windows.Forms.FlowLayoutPanel
        Me.BtnBSBK3L = New System.Windows.Forms.Button
        Me.BtnBSBK3C = New System.Windows.Forms.Button
        Me.BtnBSBK3R = New System.Windows.Forms.Button
        Me.FLBSBK2 = New System.Windows.Forms.FlowLayoutPanel
        Me.BtnBSBK2L = New System.Windows.Forms.Button
        Me.BtnBSBK2C = New System.Windows.Forms.Button
        Me.BtnBSBK2R = New System.Windows.Forms.Button
        Me.FLBSBK1 = New System.Windows.Forms.FlowLayoutPanel
        Me.BtnBSBK1L = New System.Windows.Forms.Button
        Me.BtnBSBK1C = New System.Windows.Forms.Button
        Me.BtnBSBK1R = New System.Windows.Forms.Button
        Me.LbBSBK3 = New System.Windows.Forms.Label
        Me.LbBSBK2 = New System.Windows.Forms.Label
        Me.LbBSBK1 = New System.Windows.Forms.Label
        Me.TxBSBK3 = New System.Windows.Forms.TextBox
        Me.TxBSBK2 = New System.Windows.Forms.TextBox
        Me.TxBSBK1 = New System.Windows.Forms.TextBox
        Me.TxBodySet = New System.Windows.Forms.TextBox
        Me.FontD = New System.Windows.Forms.FontDialog
        Me.ColorD = New System.Windows.Forms.ColorDialog
        Me.BtnSaveSet = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BtnDefSet = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TabSet.SuspendLayout()
        Me.TPgRTX.SuspendLayout()
        Me.TpgUBD.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TpgBodySet.SuspendLayout()
        Me.FLBSBK3.SuspendLayout()
        Me.FLBSBK2.SuspendLayout()
        Me.FLBSBK1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabSet
        '
        Me.TabSet.Controls.Add(Me.TPgRTX)
        Me.TabSet.Controls.Add(Me.TpgUBD)
        Me.TabSet.Controls.Add(Me.TabPage1)
        Me.TabSet.Controls.Add(Me.TpgBodySet)
        Me.TabSet.Location = New System.Drawing.Point(12, 12)
        Me.TabSet.Name = "TabSet"
        Me.TabSet.SelectedIndex = 0
        Me.TabSet.Size = New System.Drawing.Size(630, 353)
        Me.TabSet.TabIndex = 0
        '
        'TPgRTX
        '
        Me.TPgRTX.Controls.Add(Me.BtnRC)
        Me.TPgRTX.Controls.Add(Me.BtnBC)
        Me.TPgRTX.Controls.Add(Me.BtnFC)
        Me.TPgRTX.Controls.Add(Me.BtnFontSet)
        Me.TPgRTX.Controls.Add(Me.RDemo)
        Me.TPgRTX.Location = New System.Drawing.Point(4, 21)
        Me.TPgRTX.Name = "TPgRTX"
        Me.TPgRTX.Padding = New System.Windows.Forms.Padding(3)
        Me.TPgRTX.Size = New System.Drawing.Size(622, 328)
        Me.TPgRTX.TabIndex = 0
        Me.TPgRTX.Text = "編輯區設定"
        Me.TPgRTX.UseVisualStyleBackColor = True
        '
        'BtnRC
        '
        Me.BtnRC.Location = New System.Drawing.Point(30, 119)
        Me.BtnRC.Name = "BtnRC"
        Me.BtnRC.Size = New System.Drawing.Size(96, 24)
        Me.BtnRC.TabIndex = 4
        Me.BtnRC.Text = "變更識別色"
        Me.BtnRC.UseVisualStyleBackColor = True
        '
        'BtnBC
        '
        Me.BtnBC.Location = New System.Drawing.Point(30, 90)
        Me.BtnBC.Name = "BtnBC"
        Me.BtnBC.Size = New System.Drawing.Size(96, 23)
        Me.BtnBC.TabIndex = 3
        Me.BtnBC.Text = "變更底色"
        Me.BtnBC.UseVisualStyleBackColor = True
        '
        'BtnFC
        '
        Me.BtnFC.Location = New System.Drawing.Point(30, 60)
        Me.BtnFC.Name = "BtnFC"
        Me.BtnFC.Size = New System.Drawing.Size(96, 23)
        Me.BtnFC.TabIndex = 2
        Me.BtnFC.Text = "變更字型色彩"
        Me.BtnFC.UseVisualStyleBackColor = True
        '
        'BtnFontSet
        '
        Me.BtnFontSet.Location = New System.Drawing.Point(30, 31)
        Me.BtnFontSet.Name = "BtnFontSet"
        Me.BtnFontSet.Size = New System.Drawing.Size(96, 23)
        Me.BtnFontSet.TabIndex = 1
        Me.BtnFontSet.Text = "變更字型"
        Me.BtnFontSet.UseVisualStyleBackColor = True
        '
        'RDemo
        '
        Me.RDemo.ImeMode = System.Windows.Forms.ImeMode.Close
        Me.RDemo.Location = New System.Drawing.Point(198, 18)
        Me.RDemo.Name = "RDemo"
        Me.RDemo.Size = New System.Drawing.Size(149, 135)
        Me.RDemo.TabIndex = 0
        Me.RDemo.Text = "極光駭客" & Global.Microsoft.VisualBasic.ChrW(10) & "Aurora Hacker" & Global.Microsoft.VisualBasic.ChrW(10) & "向外分享必需附上原始碼" & Global.Microsoft.VisualBasic.ChrW(10) & "GPL" & Global.Microsoft.VisualBasic.ChrW(10) & "他說：「一2參ⅣFive」結果卻是『柒8』 nine"
        '
        'TpgUBD
        '
        Me.TpgUBD.Controls.Add(Me.BtnUBDBC)
        Me.TpgUBD.Controls.Add(Me.BtnUBDFC)
        Me.TpgUBD.Controls.Add(Me.BtnUBDFontSet)
        Me.TpgUBD.Controls.Add(Me.TxUDemo)
        Me.TpgUBD.Location = New System.Drawing.Point(4, 21)
        Me.TpgUBD.Name = "TpgUBD"
        Me.TpgUBD.Padding = New System.Windows.Forms.Padding(3)
        Me.TpgUBD.Size = New System.Drawing.Size(622, 328)
        Me.TpgUBD.TabIndex = 1
        Me.TpgUBD.Text = "輔助板設定"
        Me.TpgUBD.UseVisualStyleBackColor = True
        '
        'BtnUBDBC
        '
        Me.BtnUBDBC.Location = New System.Drawing.Point(33, 90)
        Me.BtnUBDBC.Name = "BtnUBDBC"
        Me.BtnUBDBC.Size = New System.Drawing.Size(96, 23)
        Me.BtnUBDBC.TabIndex = 6
        Me.BtnUBDBC.Text = "變更底色"
        Me.BtnUBDBC.UseVisualStyleBackColor = True
        '
        'BtnUBDFC
        '
        Me.BtnUBDFC.Location = New System.Drawing.Point(33, 60)
        Me.BtnUBDFC.Name = "BtnUBDFC"
        Me.BtnUBDFC.Size = New System.Drawing.Size(96, 23)
        Me.BtnUBDFC.TabIndex = 5
        Me.BtnUBDFC.Text = "變更字型色彩"
        Me.BtnUBDFC.UseVisualStyleBackColor = True
        '
        'BtnUBDFontSet
        '
        Me.BtnUBDFontSet.Location = New System.Drawing.Point(33, 31)
        Me.BtnUBDFontSet.Name = "BtnUBDFontSet"
        Me.BtnUBDFontSet.Size = New System.Drawing.Size(96, 23)
        Me.BtnUBDFontSet.TabIndex = 4
        Me.BtnUBDFontSet.Text = "變更字型"
        Me.BtnUBDFontSet.UseVisualStyleBackColor = True
        '
        'TxUDemo
        '
        Me.TxUDemo.Location = New System.Drawing.Point(187, 33)
        Me.TxUDemo.Multiline = True
        Me.TxUDemo.Name = "TxUDemo"
        Me.TxUDemo.Size = New System.Drawing.Size(116, 104)
        Me.TxUDemo.TabIndex = 0
        Me.TxUDemo.Text = "一二三四五六七" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "周宇成" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "藍雲飛"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.CkAutoUpVerDet)
        Me.TabPage1.Controls.Add(Me.CkSTime)
        Me.TabPage1.Controls.Add(Me.TxPEyeMin)
        Me.TabPage1.Controls.Add(Me.CkPEye)
        Me.TabPage1.Controls.Add(Me.TxAutoBakSec)
        Me.TabPage1.Controls.Add(Me.CkAutoBak)
        Me.TabPage1.Location = New System.Drawing.Point(4, 21)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(622, 328)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "自動化設定"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CkAutoUpVerDet
        '
        Me.CkAutoUpVerDet.AutoSize = True
        Me.CkAutoUpVerDet.Location = New System.Drawing.Point(32, 130)
        Me.CkAutoUpVerDet.Name = "CkAutoUpVerDet"
        Me.CkAutoUpVerDet.Size = New System.Drawing.Size(180, 16)
        Me.CkAutoUpVerDet.TabIndex = 5
        Me.CkAutoUpVerDet.Text = "自動連線檢查是否有更新版本"
        Me.CkAutoUpVerDet.UseVisualStyleBackColor = True
        '
        'CkSTime
        '
        Me.CkSTime.AutoSize = True
        Me.CkSTime.Location = New System.Drawing.Point(32, 97)
        Me.CkSTime.Name = "CkSTime"
        Me.CkSTime.Size = New System.Drawing.Size(108, 16)
        Me.CkSTime.TabIndex = 4
        Me.CkSTime.Text = "顯示打文計時器"
        Me.CkSTime.UseVisualStyleBackColor = True
        '
        'TxPEyeMin
        '
        Me.TxPEyeMin.Location = New System.Drawing.Point(359, 62)
        Me.TxPEyeMin.Name = "TxPEyeMin"
        Me.TxPEyeMin.Size = New System.Drawing.Size(50, 22)
        Me.TxPEyeMin.TabIndex = 3
        '
        'CkPEye
        '
        Me.CkPEye.AutoSize = True
        Me.CkPEye.Location = New System.Drawing.Point(32, 64)
        Me.CkPEye.Name = "CkPEye"
        Me.CkPEye.Size = New System.Drawing.Size(324, 16)
        Me.CkPEye.TabIndex = 2
        Me.CkPEye.Text = "啟動保護視力定時提醒功能，請設定每隔幾分鐘提醒您："
        Me.CkPEye.UseVisualStyleBackColor = True
        '
        'TxAutoBakSec
        '
        Me.TxAutoBakSec.Location = New System.Drawing.Point(359, 28)
        Me.TxAutoBakSec.Name = "TxAutoBakSec"
        Me.TxAutoBakSec.Size = New System.Drawing.Size(50, 22)
        Me.TxAutoBakSec.TabIndex = 1
        '
        'CkAutoBak
        '
        Me.CkAutoBak.AutoSize = True
        Me.CkAutoBak.Location = New System.Drawing.Point(32, 28)
        Me.CkAutoBak.Name = "CkAutoBak"
        Me.CkAutoBak.Size = New System.Drawing.Size(312, 16)
        Me.CkAutoBak.TabIndex = 0
        Me.CkAutoBak.Text = "啟動自動備存，請設定每次備存至少幾秒後再次備存："
        Me.CkAutoBak.UseVisualStyleBackColor = True
        '
        'TpgBodySet
        '
        Me.TpgBodySet.Controls.Add(Me.BtnDefBS)
        Me.TpgBodySet.Controls.Add(Me.FLBSBK3)
        Me.TpgBodySet.Controls.Add(Me.FLBSBK2)
        Me.TpgBodySet.Controls.Add(Me.FLBSBK1)
        Me.TpgBodySet.Controls.Add(Me.LbBSBK3)
        Me.TpgBodySet.Controls.Add(Me.LbBSBK2)
        Me.TpgBodySet.Controls.Add(Me.LbBSBK1)
        Me.TpgBodySet.Controls.Add(Me.TxBSBK3)
        Me.TpgBodySet.Controls.Add(Me.TxBSBK2)
        Me.TpgBodySet.Controls.Add(Me.TxBSBK1)
        Me.TpgBodySet.Controls.Add(Me.TxBodySet)
        Me.TpgBodySet.Location = New System.Drawing.Point(4, 21)
        Me.TpgBodySet.Name = "TpgBodySet"
        Me.TpgBodySet.Size = New System.Drawing.Size(622, 328)
        Me.TpgBodySet.TabIndex = 3
        Me.TpgBodySet.Text = "人物設定項目"
        Me.TpgBodySet.UseVisualStyleBackColor = True
        '
        'BtnDefBS
        '
        Me.BtnDefBS.Location = New System.Drawing.Point(13, 270)
        Me.BtnDefBS.Name = "BtnDefBS"
        Me.BtnDefBS.Size = New System.Drawing.Size(75, 23)
        Me.BtnDefBS.TabIndex = 11
        Me.BtnDefBS.Text = "預設項目"
        Me.BtnDefBS.UseVisualStyleBackColor = True
        '
        'FLBSBK3
        '
        Me.FLBSBK3.Controls.Add(Me.BtnBSBK3L)
        Me.FLBSBK3.Controls.Add(Me.BtnBSBK3C)
        Me.FLBSBK3.Controls.Add(Me.BtnBSBK3R)
        Me.FLBSBK3.Location = New System.Drawing.Point(272, 182)
        Me.FLBSBK3.Name = "FLBSBK3"
        Me.FLBSBK3.Size = New System.Drawing.Size(42, 79)
        Me.FLBSBK3.TabIndex = 10
        '
        'BtnBSBK3L
        '
        Me.BtnBSBK3L.Location = New System.Drawing.Point(3, 3)
        Me.BtnBSBK3L.Name = "BtnBSBK3L"
        Me.BtnBSBK3L.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK3L.TabIndex = 0
        Me.BtnBSBK3L.Text = "<<"
        Me.BtnBSBK3L.UseVisualStyleBackColor = True
        '
        'BtnBSBK3C
        '
        Me.BtnBSBK3C.Location = New System.Drawing.Point(3, 28)
        Me.BtnBSBK3C.Name = "BtnBSBK3C"
        Me.BtnBSBK3C.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK3C.TabIndex = 1
        Me.BtnBSBK3C.Text = "<>"
        Me.BtnBSBK3C.UseVisualStyleBackColor = True
        '
        'BtnBSBK3R
        '
        Me.BtnBSBK3R.Location = New System.Drawing.Point(3, 53)
        Me.BtnBSBK3R.Name = "BtnBSBK3R"
        Me.BtnBSBK3R.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK3R.TabIndex = 2
        Me.BtnBSBK3R.Text = ">>"
        Me.BtnBSBK3R.UseVisualStyleBackColor = True
        '
        'FLBSBK2
        '
        Me.FLBSBK2.Controls.Add(Me.BtnBSBK2L)
        Me.FLBSBK2.Controls.Add(Me.BtnBSBK2C)
        Me.FLBSBK2.Controls.Add(Me.BtnBSBK2R)
        Me.FLBSBK2.Location = New System.Drawing.Point(271, 97)
        Me.FLBSBK2.Name = "FLBSBK2"
        Me.FLBSBK2.Size = New System.Drawing.Size(42, 79)
        Me.FLBSBK2.TabIndex = 9
        '
        'BtnBSBK2L
        '
        Me.BtnBSBK2L.Location = New System.Drawing.Point(3, 3)
        Me.BtnBSBK2L.Name = "BtnBSBK2L"
        Me.BtnBSBK2L.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK2L.TabIndex = 0
        Me.BtnBSBK2L.Text = "<<"
        Me.BtnBSBK2L.UseVisualStyleBackColor = True
        '
        'BtnBSBK2C
        '
        Me.BtnBSBK2C.Location = New System.Drawing.Point(3, 28)
        Me.BtnBSBK2C.Name = "BtnBSBK2C"
        Me.BtnBSBK2C.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK2C.TabIndex = 1
        Me.BtnBSBK2C.Text = "<>"
        Me.BtnBSBK2C.UseVisualStyleBackColor = True
        '
        'BtnBSBK2R
        '
        Me.BtnBSBK2R.Location = New System.Drawing.Point(3, 53)
        Me.BtnBSBK2R.Name = "BtnBSBK2R"
        Me.BtnBSBK2R.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK2R.TabIndex = 2
        Me.BtnBSBK2R.Text = ">>"
        Me.BtnBSBK2R.UseVisualStyleBackColor = True
        '
        'FLBSBK1
        '
        Me.FLBSBK1.Controls.Add(Me.BtnBSBK1L)
        Me.FLBSBK1.Controls.Add(Me.BtnBSBK1C)
        Me.FLBSBK1.Controls.Add(Me.BtnBSBK1R)
        Me.FLBSBK1.Location = New System.Drawing.Point(272, 12)
        Me.FLBSBK1.Name = "FLBSBK1"
        Me.FLBSBK1.Size = New System.Drawing.Size(42, 79)
        Me.FLBSBK1.TabIndex = 8
        '
        'BtnBSBK1L
        '
        Me.BtnBSBK1L.Location = New System.Drawing.Point(3, 3)
        Me.BtnBSBK1L.Name = "BtnBSBK1L"
        Me.BtnBSBK1L.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK1L.TabIndex = 0
        Me.BtnBSBK1L.Text = "<<"
        Me.BtnBSBK1L.UseVisualStyleBackColor = True
        '
        'BtnBSBK1C
        '
        Me.BtnBSBK1C.Location = New System.Drawing.Point(3, 28)
        Me.BtnBSBK1C.Name = "BtnBSBK1C"
        Me.BtnBSBK1C.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK1C.TabIndex = 1
        Me.BtnBSBK1C.Text = "<>"
        Me.BtnBSBK1C.UseVisualStyleBackColor = True
        '
        'BtnBSBK1R
        '
        Me.BtnBSBK1R.Location = New System.Drawing.Point(3, 53)
        Me.BtnBSBK1R.Name = "BtnBSBK1R"
        Me.BtnBSBK1R.Size = New System.Drawing.Size(38, 19)
        Me.BtnBSBK1R.TabIndex = 2
        Me.BtnBSBK1R.Text = ">>"
        Me.BtnBSBK1R.UseVisualStyleBackColor = True
        '
        'LbBSBK3
        '
        Me.LbBSBK3.Location = New System.Drawing.Point(332, 185)
        Me.LbBSBK3.Name = "LbBSBK3"
        Me.LbBSBK3.Size = New System.Drawing.Size(24, 62)
        Me.LbBSBK3.TabIndex = 7
        Me.LbBSBK3.Text = "備用項目 3"
        '
        'LbBSBK2
        '
        Me.LbBSBK2.Location = New System.Drawing.Point(332, 100)
        Me.LbBSBK2.Name = "LbBSBK2"
        Me.LbBSBK2.Size = New System.Drawing.Size(24, 62)
        Me.LbBSBK2.TabIndex = 6
        Me.LbBSBK2.Text = "備用項目 2"
        '
        'LbBSBK1
        '
        Me.LbBSBK1.Location = New System.Drawing.Point(332, 15)
        Me.LbBSBK1.Name = "LbBSBK1"
        Me.LbBSBK1.Size = New System.Drawing.Size(24, 62)
        Me.LbBSBK1.TabIndex = 5
        Me.LbBSBK1.Text = "備用項目 1"
        '
        'TxBSBK3
        '
        Me.TxBSBK3.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.TxBSBK3.Location = New System.Drawing.Point(362, 182)
        Me.TxBSBK3.Multiline = True
        Me.TxBSBK3.Name = "TxBSBK3"
        Me.TxBSBK3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxBSBK3.Size = New System.Drawing.Size(244, 79)
        Me.TxBSBK3.TabIndex = 4
        '
        'TxBSBK2
        '
        Me.TxBSBK2.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.TxBSBK2.Location = New System.Drawing.Point(362, 97)
        Me.TxBSBK2.Multiline = True
        Me.TxBSBK2.Name = "TxBSBK2"
        Me.TxBSBK2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxBSBK2.Size = New System.Drawing.Size(244, 79)
        Me.TxBSBK2.TabIndex = 3
        '
        'TxBSBK1
        '
        Me.TxBSBK1.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.TxBSBK1.Location = New System.Drawing.Point(362, 12)
        Me.TxBSBK1.Multiline = True
        Me.TxBSBK1.Name = "TxBSBK1"
        Me.TxBSBK1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxBSBK1.Size = New System.Drawing.Size(244, 79)
        Me.TxBSBK1.TabIndex = 2
        '
        'TxBodySet
        '
        Me.TxBodySet.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.TxBodySet.Location = New System.Drawing.Point(13, 12)
        Me.TxBodySet.Multiline = True
        Me.TxBodySet.Name = "TxBodySet"
        Me.TxBodySet.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TxBodySet.Size = New System.Drawing.Size(244, 251)
        Me.TxBodySet.TabIndex = 1
        Me.TxBodySet.WordWrap = False
        '
        'BtnSaveSet
        '
        Me.BtnSaveSet.Location = New System.Drawing.Point(188, 373)
        Me.BtnSaveSet.Name = "BtnSaveSet"
        Me.BtnSaveSet.Size = New System.Drawing.Size(75, 23)
        Me.BtnSaveSet.TabIndex = 1
        Me.BtnSaveSet.Text = "儲存(&S)"
        Me.BtnSaveSet.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Location = New System.Drawing.Point(269, 373)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 2
        Me.BtnClose.Text = "取消(&C)"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BtnDefSet
        '
        Me.BtnDefSet.Location = New System.Drawing.Point(350, 373)
        Me.BtnDefSet.Name = "BtnDefSet"
        Me.BtnDefSet.Size = New System.Drawing.Size(75, 23)
        Me.BtnDefSet.TabIndex = 3
        Me.BtnDefSet.Text = "預設(&D)"
        Me.BtnDefSet.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'FormSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(654, 408)
        Me.Controls.Add(Me.BtnDefSet)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.BtnSaveSet)
        Me.Controls.Add(Me.TabSet)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FormSet"
        Me.Text = "Setting"
        Me.TabSet.ResumeLayout(False)
        Me.TPgRTX.ResumeLayout(False)
        Me.TpgUBD.ResumeLayout(False)
        Me.TpgUBD.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TpgBodySet.ResumeLayout(False)
        Me.TpgBodySet.PerformLayout()
        Me.FLBSBK3.ResumeLayout(False)
        Me.FLBSBK2.ResumeLayout(False)
        Me.FLBSBK1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabSet As System.Windows.Forms.TabControl
    Friend WithEvents TPgRTX As System.Windows.Forms.TabPage
    Friend WithEvents TpgUBD As System.Windows.Forms.TabPage
    Friend WithEvents FontD As System.Windows.Forms.FontDialog
    Friend WithEvents ColorD As System.Windows.Forms.ColorDialog
    Friend WithEvents BtnFC As System.Windows.Forms.Button
    Friend WithEvents BtnFontSet As System.Windows.Forms.Button
    Friend WithEvents RDemo As System.Windows.Forms.RichTextBox
    Friend WithEvents BtnBC As System.Windows.Forms.Button
    Friend WithEvents BtnRC As System.Windows.Forms.Button
    Friend WithEvents BtnSaveSet As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnDefSet As System.Windows.Forms.Button
    Friend WithEvents BtnUBDBC As System.Windows.Forms.Button
    Friend WithEvents BtnUBDFC As System.Windows.Forms.Button
    Friend WithEvents BtnUBDFontSet As System.Windows.Forms.Button
    Friend WithEvents TxUDemo As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents CkAutoBak As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TxAutoBakSec As System.Windows.Forms.TextBox
    Friend WithEvents CkPEye As System.Windows.Forms.CheckBox
    Friend WithEvents TxPEyeMin As System.Windows.Forms.TextBox
    Friend WithEvents CkSTime As System.Windows.Forms.CheckBox
    Friend WithEvents TpgBodySet As System.Windows.Forms.TabPage
    Friend WithEvents TxBodySet As System.Windows.Forms.TextBox
    Friend WithEvents TxBSBK1 As System.Windows.Forms.TextBox
    Friend WithEvents LbBSBK3 As System.Windows.Forms.Label
    Friend WithEvents LbBSBK2 As System.Windows.Forms.Label
    Friend WithEvents LbBSBK1 As System.Windows.Forms.Label
    Friend WithEvents TxBSBK3 As System.Windows.Forms.TextBox
    Friend WithEvents TxBSBK2 As System.Windows.Forms.TextBox
    Friend WithEvents FLBSBK1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents BtnBSBK1L As System.Windows.Forms.Button
    Friend WithEvents BtnBSBK1C As System.Windows.Forms.Button
    Friend WithEvents BtnBSBK1R As System.Windows.Forms.Button
    Friend WithEvents FLBSBK3 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents BtnBSBK3L As System.Windows.Forms.Button
    Friend WithEvents BtnBSBK3C As System.Windows.Forms.Button
    Friend WithEvents BtnBSBK3R As System.Windows.Forms.Button
    Friend WithEvents FLBSBK2 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents BtnBSBK2L As System.Windows.Forms.Button
    Friend WithEvents BtnBSBK2C As System.Windows.Forms.Button
    Friend WithEvents BtnBSBK2R As System.Windows.Forms.Button
    Friend WithEvents BtnDefBS As System.Windows.Forms.Button
    Friend WithEvents CkAutoUpVerDet As System.Windows.Forms.CheckBox
End Class
