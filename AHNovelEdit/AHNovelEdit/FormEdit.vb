Public Class FormEdit

    'Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Dim MousePT As Point
    Dim TabKeyV As Integer = 0 '負責計算目前的新分頁要用的 key 編號
    Dim NowKeyV As Integer '目前 Tab 所用的 Key 編號
    Dim arrRTX(0) As System.Windows.Forms.RichTextBox
    Dim arrFILE(0) As NovelFile '儲存檔案的路徑
    Dim ProcBox As RichTextBox = New RichTextBox
    Dim TempRX As RichTextBox = New RichTextBox
    Dim arrBtnCM(40) As Button '標點符號大按鈕
    Dim RTX_OCHG As Boolean = False '當外力改變 RTX Text 值時為避免影響字形而做全選全設動作
    Dim IS_ACTIVE As Boolean = False
    Dim LongSelectStart As Integer = 0 '長選取起點
    Dim Un32der255 As Boolean = False '判斷輸入的是否為 ASCII 33 to 255
    Dim AutoBakTimer, BeforeBakTime, BeforePEyeTime As Integer '每秒累加計時用及前次時間紀錄用
    Dim DragOPFiles() As String, InDrag As Boolean = False '拖放檔案的陣列以及旗號
    Dim OnClosingFlag As Boolean = False '關閉時設為 True
    Dim StrForPrint As String = "" '列印用

    Private Sub FormEdit_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        IS_ACTIVE = True
        'Timer1.Enabled = True

    End Sub

    Private Sub FormEdit_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        'Timer1.Enabled = False
    End Sub

    Private Sub FormEdit_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim OPFiles() As String = CType(e.Data.GetData(DataFormats.FileDrop, True), String())
            DragOPFiles = OPFiles
            InDrag = True
        End If
    End Sub


    Private Sub FormEdit_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        OnClosingFlag = True
        Dim RETU As Integer
        For Each tbpTemp As TabPage In TabPG.TabPages
            RETU = RemoveRTXinPG(tbpTemp)
            If RETU = DialogResult.Cancel Then
                e.Cancel = True
                Exit For
            End If
        Next
        WriteReginfo()
    End Sub

    Private Sub FormEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '最優先載入(未繪製前):語系模組
        ReadSetting()
        LoadLangu()

        '次要載入
        ADDTxTe()
        MAKE_RTX(0)
        MAKE_BtnCM()
        AddToolTip1()
        ReadReginfo()

        '最後載入
        LAST_Give_Text()
        TxUBDSet()
        RecentFileProc()
        ShuQianProc()
        If AutoUpVerDet = True Then VerDet()

        If Command() IsNot Nothing Then OpenTxt(Command.Replace("""", ""))
        'MessageBox.Show(Command())
    End Sub

    Private Sub FormEdit_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        FPanelCM.Top = Me.Size.Height - 80
        TabPG.Height = Me.Size.Height - 160
        SplitC1.Height = Me.Size.Height - SplitC1.Location.Y - 50
        TxMsg1.Top = Me.Size.Height - TxMsg1.Height - 23
        TxUBD.Width = Me.Width - TabPG.Width - TabPG.Left - 15

        If IS_ACTIVE = True Then
            ASET_RTX()
        End If

    End Sub

    Private Sub TabPG_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TabPG.DragDrop
        Dim i As Integer
        '拖放開啟檔案
        For i = 0 To DragOPFiles.Length - 1
            OpenTxt(DragOPFiles(i))
        Next
    End Sub

    Private Sub TabPG_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TabPG.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TabPG_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TabPG.MouseDown

        MousePT.X = e.X
        MousePT.Y = e.Y
        If e.Button = Windows.Forms.MouseButtons.Middle Then
            Dim Pnt1 As Point = New Point(e.X, e.Y)
            Dim rec As Rectangle
            Dim ClPage As TabPage = TabPG.SelectedTab
            For Each tbpTemp As TabPage In TabPG.TabPages
                rec = TabPG.GetTabRect(TabPG.TabPages.IndexOf(tbpTemp))
                If rec.Contains(Pnt1.X, Pnt1.Y) Then
                    ClPage = tbpTemp
                End If
            Next
            RemoveRTXinPG(ClPage)

        End If
    End Sub

    Private Sub TabPG_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TabPG.MouseMove
        Dim MPT As Point
        Dim RTV As Control
        MPT.X = e.X
        MPT.Y = e.Y
        RTV = TabPG.GetChildAtPoint(MPT)

    End Sub

    Private Sub TabPG_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPG.Resize
        TxUBD.Left = TabPG.Width + TabPG.Left + 5
        TxUBD.Width = Me.Width - TabPG.Width - TabPG.Left - 15
        TxUBD.Height = TabPG.Height - 30
        RTX.Width = TabPG.Size.Width - RTX_Width_OFFSET
        HSTabPG.Maximum = 10000
        HSTabPG.Value = TabPG.Width
        BtnCloseTabPG.Left = TabPG.Right - BtnCloseTabPG.Width
    End Sub

    Private Sub TabPG_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPG.SelectedIndexChanged

        If TabPG.TabPages.Count > 0 Then
            NowKeyV = CInt(TabPG.SelectedTab.Name)
        Else
            NowKeyV = -1
        End If
    End Sub

    Private Sub TabPG_Page_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ShowRTXinPG()
    End Sub

    Public Sub ShowRTXinPG()
        If TabPG.TabPages.Count < 1 Then Exit Sub
        RTX = New RichTextBox
        RTX = arrRTX(CInt(TabPG.SelectedTab.Name))
        RTX.ContextMenuStrip = Popup1
        ASET_RTX()
        Me.TabPG.TabPages(TabPG.SelectedIndex).Controls.Add(RTX)
        RTX.Focus()
        ChangeResetAllDo()
        NowKeyV = CInt(TabPG.SelectedTab.Name)
        RefreshLsJe()
        CountWord()
        QLoadMenu()
        FmtChange()

    End Sub

    Private Function RemoveRTXinPG(ByVal RemovePage As TabPage) As Integer
        '要關閉分頁時一起移掉 RTX
        Dim RETV As Integer, TBPGR As Integer
        Dim NowUsePage As TabPage = TabPG.SelectedTab
        TabPG.SelectedTab = RemovePage
        TBPGR = TabPG.SelectedIndex
        RETV = ClosePreDo()
        RemoveRTXinPG = RETV
        If RETV = DialogResult.Yes Or RETV = DialogResult.No Then
            If arrFILE(NowKeyV).FilePath <> "" Then
                '正常關閉刪除暫存檔
                If My.Computer.FileSystem.FileExists(arrFILE(NowKeyV).FilePath & ".nbk") = True Then My.Computer.FileSystem.DeleteFile(arrFILE(NowKeyV).FilePath & ".nbk")
            End If
            arrRTX(NowKeyV).ResetText()
            arrRTX(NowKeyV) = Nothing
            PageToolStripMenuItem.DropDownItems.Remove(arrFILE(NowKeyV).PMenu)
            arrFILE(NowKeyV) = Nothing
            TabPG.TabPages.Remove(TabPG.SelectedTab)
        End If
        If TabPG.TabPages.Count > 1 Then
            If TBPGR > 0 Then TabPG.SelectedIndex = TBPGR - 1
            If TBPGR = 0 Then TabPG.SelectedIndex = 0
            If RemovePage.Name <> NowUsePage.Name Then TabPG.SelectedTab = NowUsePage
        End If
        ShowRTXinPG()

    End Function

    Private Sub RefreshLsJe()
        '重新整理分節列表
        If TabPG.TabPages.Count < 1 Then Exit Sub
        Dim TxNv As String = RTX.Text '先將整個Text載進變數
        Dim SplStAry() As String = {NSplitL} '將NSplitL設為字串陣列分割記號
        Dim ArSpSt() As String = TxNv.Split(SplStAry, StringSplitOptions.None) '用來切割儲放
        Dim ArL As Integer = ArSpSt.Length - 1
        Dim i As Integer, TmpName As String, EndIdx As Integer
        Dim URMG As Integer = RTX.SelectionStart, ZRMG As Integer = 0, TotalWord As Integer = 0
        Dim URMGFLAG As Boolean = False

        LsJe.Items.Clear()
        For i = 1 To ArL
            EndIdx = ArSpSt(i).IndexOf(NSplitR)
            TotalWord = TotalWord + ArSpSt(i).Length + NSplitL.Length
            If URMGFLAG = False Then
                If TotalWord > URMG Then
                    URMGFLAG = True
                    ZRMG = i - 1
                End If
            End If
            If EndIdx = -1 Then ArL = ArL - 1 : Exit For
            TmpName = ArSpSt(i).Substring(0, EndIdx)
            LsJe.Items.Add(TmpName)
        Next
        LbJeCount.Text = "0"
        Try
            LsJe.SelectedIndex = ZRMG
            LbJeCount.Text = ArL.ToString

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DeleteLsJe(ByVal DelSp As Integer)
        '刪除分節標記
        If TabPG.TabPages.Count < 1 Then Exit Sub
        Dim TxNv As String = RTX.Text '先將整個Text載進變數
        Dim SIdx As Integer = -1 '用來儲存每輪的 INDEX
        Dim i As Integer, DelReq As Integer, EndIdx As Integer

        For i = 1 To DelSp
            SIdx = TxNv.IndexOf(NSplitL, SIdx + 1)
            If SIdx = -1 Then MessageBox.Show(Msg_DeleteLsJe1) : Exit Sub
            If i = DelSp Then
                EndIdx = TxNv.IndexOf(NSplitR, SIdx)
                EndIdx = EndIdx + NSplitR.Length
                RTX.SelectionStart = SIdx
                RTX.SelectionLength = EndIdx - SIdx
                RTX.Focus()
                RTX.ScrollToCaret()
                DelReq = MessageBox.Show(Msg_DeleteLsJe2, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                If DelReq = DialogResult.Yes Then ProcBoxProc("", True)
                RefreshLsJe()
            End If
        Next

    End Sub

    Private Sub JumpLsJe(ByVal JumpSp As Integer)
        '跳至分節標記
        If TabPG.TabPages.Count < 1 Then Exit Sub
        Dim TxNv As String = RTX.Text '先將整個Text載進變數
        Dim SIdx As Integer = -1 '用來儲存每輪的 INDEX
        Dim i As Integer, EndIdx As Integer

        For i = 1 To JumpSp
            SIdx = TxNv.IndexOf(NSplitL, SIdx + 1)
            If SIdx = -1 Then MessageBox.Show(Msg_JumpLsJe1) : Exit Sub
            If i = JumpSp Then
                EndIdx = TxNv.IndexOf(NSplitR, SIdx)
                EndIdx = EndIdx + NSplitR.Length
                RTX.SelectionStart = SIdx
                RTX.SelectionLength = EndIdx - SIdx
                RTX.Focus()
                RTX.ScrollToCaret()
            End If
        Next

    End Sub

    Private Sub SelectLsJe(ByVal JumpSp As Integer, Optional ByVal FullSelect As Boolean = False)
        '複製分節標記
        If TabPG.TabPages.Count < 1 Then Exit Sub
        Dim TxNv As String = RTX.Text '先將整個Text載進變數
        Dim SIdx As Integer = -1 '用來儲存每輪的 INDEX
        Dim i As Integer, SSIdx, EndIdx As Integer

        For i = 1 To JumpSp
            SIdx = TxNv.IndexOf(NSplitL, SIdx + 1)
            If SIdx = -1 Then MessageBox.Show(Msg_JumpLsJe1) : Exit Sub
            If i = JumpSp Then
                EndIdx = TxNv.IndexOf(NSplitR, SIdx)
                EndIdx = EndIdx + NSplitR.Length
                RTX.SelectionStart = SIdx
                RTX.SelectionLength = EndIdx - SIdx
                If RTX.SelectedText <> NSplitL & LsJe.SelectedItem.ToString & NSplitR Then MessageBox.Show(Msg_JumpLsJe1) : Exit Sub
                If FullSelect = True Then SSIdx = SIdx
                SIdx = EndIdx '將尾端重新指記給Sidx作頭
                EndIdx = TxNv.IndexOf(NSplitL, SIdx)
                If EndIdx = -1 Then EndIdx = TxNv.Length
                If FullSelect = True Then
                    RTX.SelectionStart = SSIdx
                    RTX.SelectionLength = EndIdx - SSIdx
                Else
                    RTX.SelectionStart = SIdx
                    RTX.SelectionLength = EndIdx - SIdx
                End If
                RTX.Focus()
                RTX.ScrollToCaret()

            End If
        Next

    End Sub

    Private Function RemoveAllJe(Optional ByVal OpenNew As Boolean = False) As String
        RemoveAllJe = ""
        If LsJe.Items.Count = 0 Then Exit Function
        If TabPG.TabPages.Count < 1 Then Exit Function
        Dim SIdx As Integer = -1 '用來儲存每輪的 INDEX
        Dim VRTX As RichTextBox = New RichTextBox
        Dim EndIdx As Integer
        Dim i As Integer
        VRTX.Text = RTX.Text
        For i = 1 To LsJe.Items.Count
            SIdx = VRTX.Text.IndexOf(NSplitL, SIdx + 1)
            EndIdx = VRTX.Text.IndexOf(NSplitR, SIdx)
            EndIdx = EndIdx + NSplitR.Length
            VRTX.SelectionStart = SIdx
            VRTX.SelectionLength = EndIdx - SIdx
            VRTX.SelectedText = ""
        Next
        RemoveAllJe = VRTX.Text
        If OpenNew = True Then
            OpenTxt("", True)
            ProcBoxProc(RemoveAllJe, True)
        End If
    End Function

    Private Function AddJeNew() As Boolean
        '新增一組分節標記到文章中
        AddJeNew = False
        If TabPG.TabPages.Count < 1 Then Exit Function
        Dim IpSt As String
        IpSt = InputBox(Msg_InputJeTell, "Split!")
        If IpSt = "" Or IpSt = Nothing Then MessageBox.Show(Msg_InputJeEnd1) : Exit Function
        ProcBoxProc(vbCrLf & NSplitL & IpSt & NSplitR & vbCrLf, True)
        RefreshLsJe()
        RTX.Focus()
        AddJeNew = True

    End Function

    Public Sub MAKE_RTX(ByVal i As Integer)
        arrRTX(i) = New RichTextBox
        arrRTX(i).Left = RTX_Left '174
        arrRTX(i).Top = RTX_Top '106
        arrRTX(i).Font = RTX_Font
        arrRTX(i).AcceptsTab = True
        arrRTX(i).AutoWordSelection = RTX_AutoWordSelection
        arrRTX(i).Height = TabPG.Size.Height - RTX_Height_OFFSET
        arrRTX(i).Width = TabPG.Size.Width - RTX_Width_OFFSET '400
        arrRTX(i).Enabled = True
        arrRTX(i).Visible = True
        arrRTX(i).ImeMode = RTX_ImeMode

    End Sub

    Public Sub MAKE_BtnCM()
        Dim i As Integer
        Dim PDASD As Padding
        PDASD.All = 1
        For i = 0 To arrBtnCM.Length - 1
            arrBtnCM(i) = New Button
            arrBtnCM(i).Height = 20
            arrBtnCM(i).Width = 20
            arrBtnCM(i).Text = DotCM(i)
            If i > 31 Then arrBtnCM(i).Width = 38
            Me.FPanelCM.Controls.Add(arrBtnCM(i))
            arrBtnCM(i).FlatStyle = FlatStyle.Flat
            arrBtnCM(i).FlatAppearance.BorderSize = 0
            arrBtnCM(i).Margin = PDASD
            ToolTip1.SetToolTip(arrBtnCM(i), Make_arBtn_TP(i))
            AddHandler arrBtnCM(i).Click, AddressOf arBtnCM_Click
        Next

    End Sub

    Public Sub arBtnCM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim VCMStr As String, pos As Integer
        VCMStr = sender.ToString
        pos = VCMStr.IndexOf("Text:")
        pos = pos + 6
        VCMStr = VCMStr.Substring(pos)

        RTX_OCHG = True
        INPUTCM_RTX(VCMStr)
        'Dim POT As Integer ', AA As String, BB As String
        'POT = RTX.SelectionStart
        'If RTX.SelectionLength > 0 Then
        'RTX.SelectedText = ""
        'End If
        'ProcBox.Text = VCMStr
        'ProcBox.SelectAll()
        'ProcBox.SelectionFont = RTX.Font
        'RTX.SelectedRtf = ProcBox.SelectedRtf
        'RTX.SelectionStart = POT + 1
        'RTX.Focus()
    End Sub
    Public Sub INPUTCM_RTX(ByVal StrA As String, Optional ByVal OffsetA As Integer = 1)
        '插入字元模組
        Dim POT As Integer
        POT = RTX.SelectionStart
        If RTX.SelectionLength > 0 Then
            If StrA.Length = 2 Then StrA = StrA.Substring(0, 1) & RTX.SelectedText & StrA.Substring(1, 1)
            RTX.SelectedText = ""
        End If
        ProcBoxProc(StrA, True)
        RTX.SelectionStart = POT + OffsetA
        RTX.Focus()
    End Sub

    Public Sub ASET_RTX()
        On Error Resume Next
        SystemSet = True
        Dim MDMT As Integer '先將原本游標位置紀錄下來
        MDMT = RTX.SelectionStart
        RTX.Left = RTX_Left '174
        RTX.Top = RTX_Top '106
        RTX.Font = RTX_Font
        RTX.DetectUrls = False
        RTX.HideSelection = False
        RTX.ScrollBars = RichTextBoxScrollBars.ForcedVertical
        RTX.AllowDrop = True '啟用拖曳
        'RTX.ShowSelectionMargin = True
        RTX.AcceptsTab = True
        RTX.ForeColor = RTX_FrColor
        RTX.BackColor = RTX_BkColor

        RTX.LanguageOption = RichTextBoxLanguageOptions.DualFont
        RTX.AutoWordSelection = RTX_AutoWordSelection
        RTX.Height = TabPG.Size.Height - RTX_Height_OFFSET
        RTX.Width = TabPG.Size.Width - RTX_Width_OFFSET '400
        RTX.Enabled = True
        RTX.Visible = True
        RTX.ImeMode = RTX_ImeMode

        If RTX_OCHG = True Then
            RTX.SelectAll()
            RTX.SelectionFont = RTX_Font
            RTX.SelectionLength = 0
            RTX_OCHG = False
        End If

        RTX.SelectionStart = MDMT
        SystemSet = False

    End Sub

    Private Sub CountWord()
        If PanelCountWord.Visible = False Then Exit Sub
        Dim i As Integer
        Dim NVB As String
        Dim JL As Integer = CInt(LbJeCount.Text)
        Dim Tmp1 As String
        Dim RexT As String() = {vbCr, vbLf, " ", "　", "	", Chr(0)}

        NVB = RTX.Text
        For i = 0 To RexT.Length - 1
            NVB = NVB.Replace(RexT(i), Nothing)
        Next
        LbTWC.Text = NVB.Length
        For i = 0 To JL - 1
            Tmp1 = NSplitL & LsJe.Items.Item(i) & NSplitR
            NVB = NVB.Replace(Tmp1, Nothing)
        Next
        LbNWC.Text = NVB.Length
        If RTX.SelectionLength > 0 Then
            Dim SVB As String
            SVB = RTX.SelectedText
            For i = 0 To RexT.Length - 1
                SVB = SVB.Replace(RexT(i), Nothing)
            Next
            LbSWC.Text = SVB.Length
            'LbSWC.Text = RTX.SelectedText.Length
        Else
            LbSWC.Text = 0
        End If

    End Sub


    Public Sub HIDE_ALL_RTX()
        Exit Sub
        Dim i As Integer
        For i = 0 To arrRTX.Length - 1
            arrRTX(i).Visible = False
            arrRTX(i).AutoWordSelection = False
        Next
    End Sub
    Private Sub AddToolTip1()
        ToolTip1.SetToolTip(BtnJeNew, FormEdit_BtnJeNewTP)
        ToolTip1.SetToolTip(BtnJeDel, FormEdit_BtnJeDelTP)
        ToolTip1.SetToolTip(BtnJeRefresh, FormEdit_BtnJeRefreshTP)
        ToolTip1.SetToolTip(BtnJeSel, FormEdit_BtnJeSelTP)
        ToolTip1.SetToolTip(LsJe, FormEdit_LsJeJumpTP)
        ToolTip1.SetToolTip(HSTabPG, FormEdit_HSTabPGTP)
        ToolTip1.SetToolTip(LbTWC, FormEdit_LbTWCTP) '以下6個為字數計區域
        ToolTip1.SetToolTip(LbNWC, FormEdit_LbNWCTP)
        ToolTip1.SetToolTip(LbSWC, FormEdit_LbSWCTP)
        ToolTip1.SetToolTip(LbTWC1, FormEdit_LbTWCTP)
        ToolTip1.SetToolTip(LbNWC1, FormEdit_LbNWCTP)
        ToolTip1.SetToolTip(LbSWC1, FormEdit_LbSWCTP)
        ToolTip1.SetToolTip(BtnPanelCountWordHide, FormEdit_BtnPanelCountWordHide)

        LongSelectToolStripButton.ToolTipText = FormEdit_LongSelectToolStripButtonTP
        FormatAllToolStripButton.ToolTipText = FormEdit_FormatAllToolStripButtonTP
        RecognitionToolStripButton.ToolTipText = FormEdit_RecognitionToolStripButtonTP
    End Sub

    Private Sub BtnCM1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        RTX_OCHG = True
        Dim POT As Integer
        Dim SCRMM As RichTextBox = New RichTextBox
        POT = RTX.SelectionStart
        RTX.Focus()
    End Sub

    Private Sub RTX_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RTX.DragDrop
        Dim i As Integer
        '拖放開啟檔案
        For i = 0 To DragOPFiles.Length - 1
            OpenTxt(DragOPFiles(i))
        Next
    End Sub

    Private Sub RTX_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RTX.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub RTX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RTX.KeyDown

        '特殊鍵盤符號輸入
        If Control.ModifierKeys = Keys.Control Then
            If e.KeyValue = Keys.R Then e.Handled = True 'CTRL+R不靠右
            If e.KeyValue = Keys.E Then e.Handled = True 'CTRL+E不置中
            If e.KeyValue = Keys.J Then e.Handled = True 'CTRL+J不左右
            If e.KeyValue = Keys.L Then e.Handled = True 'CTRL+R不靠左
            If e.KeyCode = 188 Then INPUTCM_RTX(DotCM(0)) '按,輸入全形逗號
            If e.KeyCode = 190 Then INPUTCM_RTX(DotCM(1)) '按.輸入全形句號
            If e.KeyValue = Keys.OemSemicolon Then INPUTCM_RTX(DotCM(4)) '按;輸入全形分號
            If e.KeyValue = Keys.OemQuotes Then INPUTCM_RTX(DotCM(39)) '按"輸入全形美式雙引號
            If e.KeyValue = Keys.Q Then INPUTCM_RTX(DotCM(34)) '按Q輸入中文引號
            If e.KeyValue = Keys.W Then INPUTCM_RTX(DotCM(35)) '按W輸入中文雙引號
            If e.KeyValue = Keys.D3 Then INPUTCM_RTX(DotCM(2)) '按3輸入中文頓號
            If e.KeyValue = Keys.D4 Then INPUTCM_RTX(DotCM(9)) '按4輸入中文點點點
            
        End If

    End Sub

    Private Sub RTX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RTX.KeyPress
        'If (Asc(e.KeyChar) >= 32 And Asc(e.KeyChar) < 65) Or (Asc(e.KeyChar) >= 91 And Asc(e.KeyChar) < 97) Or (Asc(e.KeyChar) >= 123 And Asc(e.KeyChar) < 256) Then Un32der255 = True
        'If Asc(e.KeyChar) >= 32 And Asc(e.KeyChar) < 256 Then Un32der255 = True
    End Sub

    Private Sub RTX_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RTX.KeyUp
        If e.KeyValue = Keys.ControlKey Then
            SaveTxtBak(NowKeyV) '打Ctrl時也檢查自動備份
            RefreshLsJe() '順便重整列表
        End If
    End Sub

    Private Sub RTX_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles RTX.Leave
        SaveTxtBak(NowKeyV)
    End Sub

    Private Sub RTX_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RTX.SelectionChanged
        CountWord()
    End Sub

    Private Sub RTX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RTX.TextChanged
        If TabPG.TabPages.Count < 1 Then Exit Sub
        RTX.AutoWordSelection = False
        'If RTX.SelectionFont IsNot RTX_Font Then
        'RTX.SelectionFont = RTX_Font
        'End If
        If SystemSet = False Then
            arrFILE(NowKeyV).AskSave = True
            arrFILE(NowKeyV).NeedBak = True
            TabPG.SelectedTab.Text = "*" & TabPG.SelectedTab.Text.Replace("*", "")
        End If
        'On Error Resume Next
        Exit Sub
        'If Un32der255 = True Then '
        'Un32der255 = False
        'If RTX.SelectionStart > 1 Then
        'If AscW(Mid(RTX.Text, RTX.SelectionStart - 1, 1)) > 255 Then
        '   Dim RTXStart As Integer = RTX.SelectionStart
        '    Dim RTXLenth As Integer = RTX.SelectionLength
        '     RTX.Select(RTX.SelectionStart - 1, 1)
        '      RTX.SelectionFont = RTX_Font
        '       RTX.SelectionLength = 0
        '        RTX.SelectionStart = RTXStart
        '         RTX.SelectionLength = RTXLenth
        '      End If
        '   End If
        'End If
    End Sub

    Private Sub ChangeResetAllDo()
        '發生改變時把長選取等一切功能狀態旗標歸位
        LongSelectToolStripButton.Checked = False

    End Sub

    Private Sub LAST_Give_Text()
        CutPopup1.Text = CutToolStripMenuItem.Text
        CopyPopup1.Text = CopyToolStripMenuItem.Text
        PastePopup1.Text = PasteToolStripMenuItem.Text
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)

    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub LongSelectToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LongSelectToolStripButton.Click
        If LongSelectToolStripButton.Checked = False Then
            LongSelectToolStripButton.Checked = True
            LongSelectStart = RTX.SelectionStart
        Else
            LongSelectToolStripButton.Checked = False
            RTX.Select(LongSelectStart, RTX.SelectionStart - LongSelectStart)
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        If RTX.SelectionLength > 0 And RTX.Focused = True Then RTX.Copy()
        If TxUBD.SelectionLength > 0 And TxUBD.Focused = True Then TxUBD.Copy()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        If RTX.SelectedText <> "" And RTX.Focused = True Then RTX.Cut()
        If TxUBD.SelectedText <> "" And TxUBD.Focused = True Then TxUBD.Cut()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        If Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) = True Then
            If RTX.Focused = True Then
                Dim TYPY As RichTextBox = New RichTextBox
                TYPY.Paste()
                ProcBoxProc(TYPY.Text, True)
                'Dim AA1 As Integer, AA2 As Integer
                'AA1 = RTX.SelectionStart
                'RTX.Paste()
                'AA2 = RTX.SelectionStart
                'RTX.Select(AA1, AA2 - AA1)
                'RTX.SelectionFont = RTX_Font
                'RTX.SelectionColor = RTX_FrColor
                'RTX.SelectionLength = 0
                'RTX.SelectionStart = AA2
            End If
            If TxUBD.Focused = True Then TxUBD.Paste()
        End If

    End Sub

    Private Sub CutToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripButton.Click
        CutToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub CopyToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripButton.Click
        CopyToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub PasteToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripButton.Click
        PasteToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub Popup1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Popup1.Opening
        If Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) = True Then
            PastePopup1.Enabled = True
        Else
            PastePopup1.Enabled = False
        End If
    End Sub

    Private Sub CutPopup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutPopup1.Click
        CutToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub CopyPopup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyPopup1.Click
        CopyToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub PastePopup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PastePopup1.Click
        PasteToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub MenuStrip1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MenuStrip1.MouseDown

    End Sub

    Private Sub MenuStrip1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.MouseEnter
        If Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) = True Then
            PasteToolStripMenuItem.Enabled = True
        Else
            PasteToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        If RTX.CanUndo = True Then RTX.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        If RTX.CanRedo = True Then RTX.Redo()
    End Sub

    Private Sub UndoToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripButton.Click
        UndoToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub RedoToolStripButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RedoToolStripButton.Click
        RedoToolStripMenuItem_Click(sender, e)
    End Sub

    Private Function OPF1Txt() As String()
        'OPF1Txt = ""
        OPF1.Title = "開啟小說文字檔"
        OPF1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        OPF1.FileName = ""
        OPF1.ShowDialog()
        OPF1Txt = OPF1.FileNames
    End Function

    Private Function SVF1Txt(Optional ByVal FileNM As String = "新小說.txt") As String
        Dim RETV As Integer
        SVF1Txt = ""
        SVF1.FileName = FileNM
        SVF1.Title = "另存小說文字檔/HTML"
        SVF1.Filter = "txt files (*.txt)|*.txt|HTML files (*.htm)|*.htm|All files (*.*)|*.*"
        RETV = SVF1.ShowDialog()
        If RETV = DialogResult.Cancel Then
            SVF1Txt = ""
        Else
            SVF1Txt = SVF1.FileName
        End If
    End Function

    Private Sub OpenTxt(ByVal FileDIR As String, Optional ByVal FileNew As Boolean = False, Optional ByVal FuDangMing As String = ".txt")
        'On Error Resume Next
        Dim Vidx As Integer, VFName As String '最後一個斜線所在與檔案名稱
        Dim i, j As Integer, IsInUse As Boolean = False '是否已用程式開啟的旗號
        If FileNew = False Then
            If FileDIR = "" Or FileDIR = Nothing Then Exit Sub
            If My.Computer.FileSystem.FileExists(FileDIR) = False Then MessageBox.Show("NO FILE ON " & FileDIR) : Exit Sub
            Vidx = FileDIR.LastIndexOf("\") '找出最後的斜線
            VFName = FileDIR.Substring(Vidx + 1)

        Else
            VFName = "New" & TabKeyV.ToString & FuDangMing '".txt"
            FileDIR = ""
        End If

        If (FileDIR = "" Or FileDIR = Nothing) And FileNew = False Then Exit Sub

        For i = 0 To arrFILE.Length - 1 ' TabPG.TabPages.Count - 1
            If arrFILE(i).FilePath = FileDIR And FileDIR <> "" Then
                IsInUse = True
                For j = 0 To TabPG.TabPages.Count - 1
                    If TabPG.TabPages.Item(j).Name = i Then TabPG.SelectTab(j) : Exit Sub
                Next
            End If
        Next

        TabPG.TabPages.Add(TabKeyV, TabKeyV.ToString & ":" & VFName)
        AddHandler TabPG.Click, AddressOf TabPG_Page_Click
        'HIDE_ALL_RTX()
        ReDim Preserve arrRTX(arrRTX.Length)
        ReDim Preserve arrFILE(arrFILE.Length)
        MAKE_RTX(arrRTX.Length - 1)
        '當不是開新檔案時載入舊檔
        If FileNew = False Then
            If My.Computer.FileSystem.FileExists(FileDIR & ".nbk") = True Then
                Dim FOpenBK As Integer
                FOpenBK = MessageBox.Show(Msg_OpenTxtBak1, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If FOpenBK = DialogResult.Yes Then
                    arrRTX(TabKeyV).LoadFile(FileDIR & ".nbk", RichTextBoxStreamType.PlainText)
                    arrFILE(TabKeyV).AskSave = True
                Else
                    arrRTX(TabKeyV).LoadFile(FileDIR, RichTextBoxStreamType.PlainText)
                    arrFILE(TabKeyV).AskSave = False
                End If
            Else
                arrRTX(TabKeyV).LoadFile(FileDIR, RichTextBoxStreamType.PlainText)
                arrFILE(TabKeyV).AskSave = False
            End If
            Dim KKJJ() As Byte '判斷 239,187,191 為前三 Byte 時就表示為 UTF8
            KKJJ = My.Computer.FileSystem.ReadAllBytes(FileDIR)
            If KKJJ.Length >= 3 Then
                If KKJJ(0) = 239 And KKJJ(1) = 187 And KKJJ(2) = 191 Then
                    arrFILE(TabKeyV).FmtU8 = True
                Else
                    arrFILE(TabKeyV).FmtU8 = False
                End If
            End If
        Else
            '開新檔案就預設為 UTF8
            arrFILE(TabKeyV).FmtU8 = True
        End If
        arrRTX(TabKeyV).SelectAll()
        arrRTX(TabKeyV).SelectionFont = RTX_Font
        arrRTX(TabKeyV).DeselectAll()
        arrRTX(TabKeyV).Visible = True
        arrFILE(TabKeyV).NeedBak = False
        arrFILE(TabKeyV).FileName = VFName
        arrFILE(TabKeyV).FilePath = FileDIR
        TabPG.ShowToolTips = True
        TabPG.SelectedIndex = TabPG.TabPages.Count - 1
        TabPG.SelectedTab.ToolTipText = arrFILE(TabKeyV).FilePath

        arrFILE(TabKeyV).PMenu = PageToolStripMenuItem.DropDownItems.Add(CStr(TabKeyV) & ".|" & arrFILE(TabKeyV).FileName)
        arrFILE(TabKeyV).PMenu.ToolTipText = arrFILE(TabKeyV).FilePath
        arrFILE(TabKeyV).PMenu.Name = TabKeyV
        AddHandler PageToolStripMenuItem.DropDownItemClicked, AddressOf PDrP_Click
        ShowRTXinPG()

        RecentFileProc(arrFILE(TabKeyV).FilePath)
        TabKeyV = TabKeyV + 1

    End Sub

    Private Sub PDrP_Click(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles PageToolStripMenuItem.DropDownItemClicked
        TabPG.SelectTab(e.ClickedItem.Name)
    End Sub

    Private Function ClosePreDo() As Integer
        Dim SKTT As Integer = DialogResult.Yes
        If arrFILE(NowKeyV).AskSave = True Then
            SKTT = MessageBox.Show(Msg_DoYouWantSave, "Save ?", MessageBoxButtons.YesNoCancel)
            If SKTT = DialogResult.Yes Then
                SKTT = SaveTxt()
            End If
        End If
        ClosePreDo = SKTT
    End Function
    Private Function SaveTxt(Optional ByVal U8 As Boolean = True) As Integer
        If TabPG.TabPages.Count < 1 Then Exit Function
        If arrFILE(NowKeyV).FilePath = "" Then
            arrFILE(NowKeyV).FilePath = SVF1Txt(arrFILE(NowKeyV).FileName)
            arrFILE(NowKeyV).FileName = arrFILE(NowKeyV).FilePath.Substring(arrFILE(NowKeyV).FilePath.LastIndexOf("\") + 1)
        Else
            U8 = arrFILE(NowKeyV).FmtU8
        End If

        If arrFILE(NowKeyV).FilePath = "" Then
            SaveTxt = DialogResult.Cancel
        Else
            'RTX.SaveFile(arrFILE(NowKeyV).FilePath, RichTextBoxStreamType.PlainText)
            Dim SVNN As String = RTX.Text.Replace(vbCrLf, vbLf)
            SVNN = SVNN.Replace(vbLf, vbCrLf)
            If U8 = True Then
                System.IO.File.WriteAllText(arrFILE(NowKeyV).FilePath, SVNN, System.Text.Encoding.UTF8)
            Else
                System.IO.File.WriteAllText(arrFILE(NowKeyV).FilePath, SVNN, System.Text.Encoding.Default)
            End If
            SaveTxtBak(NowKeyV)
            arrFILE(NowKeyV).AskSave = False
            TabPG.SelectedTab.Text = TabPG.SelectedTab.Text.Replace("*", "")
            TabPG.SelectedTab.Text = TabPG.SelectedTab.Text.Substring(0, TabPG.SelectedTab.Text.IndexOf(":") + 1)
            TabPG.SelectedTab.Text = TabPG.SelectedTab.Text & arrFILE(NowKeyV).FileName
            TabPG.SelectedTab.ToolTipText = arrFILE(NowKeyV).FilePath
            SaveTxt = DialogResult.Yes
            RecentFileProc(arrFILE(NowKeyV).FilePath)
        End If
    End Function
    Private Sub SaveTxtAs()
        If TabPG.TabPages.Count < 1 Then Exit Sub
        Dim PreFilePath As String = arrFILE(NowKeyV).FilePath
        Dim RETV As Integer
        arrFILE(NowKeyV).FilePath = ""
        RETV = SaveTxt()
        If RETV = DialogResult.Cancel Then arrFILE(NowKeyV).FilePath = PreFilePath
    End Sub
    Private Sub SaveTxtBak(ByVal TbKv As Integer)
        If arrFILE(TbKv).FilePath = "" Then Exit Sub
        If AutoBakWork = True And arrFILE(TbKv).NeedBak = True Then
            Dim PTH As String = arrFILE(NowKeyV).FilePath & ".nbk"
            Dim ETIME As Integer = AutoBakTimer - BeforeBakTime
            If ETIME > AutoBakSec Then
                'arrRTX(TbKv).SaveFile(arrFILE(TbKv).FilePath & ".nbk", RichTextBoxStreamType.PlainText)
                Dim SVNN As String = RTX.Text.Replace(vbCrLf, vbLf)
                SVNN = SVNN.Replace(vbLf, vbCrLf)
                System.IO.File.WriteAllText(arrFILE(NowKeyV).FilePath & ".nbk", SVNN, System.Text.Encoding.UTF8)
                BeforeBakTime = AutoBakTimer
                arrFILE(TbKv).NeedBak = False
            End If
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim FileDIR As String()
        FileDIR = OPF1Txt()
        Dim i As Integer
        For i = 0 To FileDIR.Length - 1
            OpenTxt(FileDIR(i))
        Next i
    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click
        OpenToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        SaveTxt()
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        SaveToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        OpenTxt("", True)
    End Sub

    Private Sub NewToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripButton.Click
        NewToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        If TabPG.TabPages.Count < 1 Then Exit Sub
        RemoveRTXinPG(TabPG.SelectedTab)
        RTX.Focus()
    End Sub

    Private Sub BtnCloseTabPG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCloseTabPG.Click
        CloseToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveTxtAs()
    End Sub

    Public Sub ProcBoxProc(ByVal Stt As String, Optional ByVal GiveRTX As Boolean = False)
        '負責將文字格式處理框處理好,選擇性指定文字取代RTX選取文字
        Dim QDC As Color = RTX.SelectionColor
        ProcBox.Text = Stt
        ProcBox.SelectAll()
        ProcBox.SelectionFont = RTX_Font
        ProcBox.SelectionColor = QDC
        If GiveRTX = True Then RTX.SelectedRtf = ProcBox.SelectedRtf
    End Sub
    Public Sub FindReplace(ByVal FWord As String, Optional ByVal ReplaceFlag As Boolean = False, Optional ByVal RWord As String = "", Optional ByVal ReplaceAll As Boolean = False)
        '尋找取代RTX ,規則為 字串,是否為找下一個,是否取代,取代字串

        Dim FindIndex As Integer
        Dim RetValue As Integer
        Dim CountRP As Integer = 0

        If ReplaceAll = False Then
            FindIndex = RTX.SelectionStart
            If ReplaceFlag = True And RTX.SelectionLength > 0 Then ProcBoxProc(RWord, True) : RTX.SelectionLength = RWord.Length
            If FindUp = False Then
                If RTX.SelectionLength <> 0 Then FindIndex += RTX.SelectionLength
                If FindIndex > RTX.Text.Length Then FindIndex = RTX.Text.Length
                RetValue = RTX.Find(FWord, FindIndex, RichTextBoxFinds.None)
                If RetValue < FindIndex Then RetValue = -1 : RTX.SelectionStart = Math.Abs(FindIndex - RTX.SelectionLength)
            Else
                RetValue = RTX.Find(FWord, 0, FindIndex, RichTextBoxFinds.Reverse)
                If RetValue > FindIndex Then RetValue = -1 : RTX.SelectionStart = FindIndex
            End If
            If RetValue = -1 Then
                Dim QsValue
                QsValue = MessageBox.Show(Msg_SearchAtLast, "Problem!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
                If QsValue = DialogResult.Yes Then
                    RTX.SelectionLength = 0
                    If FindUp = False Then
                        RTX.SelectionStart = 0
                    Else
                        RTX.SelectionStart = RTX.Text.Length
                    End If
                    FindReplace(FWord, ReplaceFlag, RWord)
                Else
                    If FindUp = False Then RTX.SelectionStart = Math.Abs(FindIndex - RTX.SelectionLength)
                End If
            End If

        Else
            Dim Sta As Integer = 0
            Dim TmpStr As String
            Do
                FindIndex = RTX.SelectionStart
                If RTX.SelectionLength <> 0 Then FindIndex += RTX.SelectionLength
                If FindIndex > RTX.Text.Length Then FindIndex = RTX.Text.Length
                RetValue = RTX.Find(FWord, FindIndex, RichTextBoxFinds.None)
                If RetValue < FindIndex Then RetValue = -1
                CountRP += 1
            Loop Until RetValue = -1 Or RetValue = RTX.Text.Length
            TmpStr = RTX.Text.Replace(FWord, RWord)
            RTX.SelectAll()
            ProcBoxProc(TmpStr, True)
            MessageBox.Show(Lang_Total & Lang_At & " " & CountRP - 1 & " " & Lang_Ga & Lang_Area1 & Lang_Finded & Lang_And2 & Lang_Replace)
        End If

    End Sub

    Private Sub FindToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindToolStripMenuItem.Click
        If TabPG.TabPages.Count <= 0 Then Exit Sub
        If RTX.SelectionLength > 0 Then FindWord = RTX.SelectedText
        FormSearch.Show()
        FormSearch.Owner = Me
        FormSearch.Focus()
    End Sub

    Private Sub FindNextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindNextToolStripMenuItem.Click
        FindReplace(FindWord)
    End Sub

    Private Sub SplitC1_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitC1.SplitterMoved
        LsJe.Height = SplitC1.Panel1.Height - GroupBox1.Height - 4
    End Sub

    Private Sub BtnJeNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnJeNew.Click
        AddJeNew()
    End Sub

    Private Sub BtnJeDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnJeDel.Click
        If LsJe.Items.Count > 0 Then
            DeleteLsJe(LsJe.SelectedIndex + 1)
        End If
    End Sub

    Private Sub BtnJeRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnJeRefresh.Click
        RefreshLsJe()
    End Sub

    Private Sub LsJe_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LsJe.DoubleClick
        If LsJe.Items.Count > 0 Then JumpLsJe(LsJe.SelectedIndex + 1)
    End Sub

    Private Sub LsJe_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LsJe.KeyUp
        If LsJe.Items.Count > 0 Then
            If e.KeyValue = Keys.Delete Then
                SelectLsJe(LsJe.SelectedIndex + 1, True)
                RTX.SelectedText = ""
                RefreshLsJe()
                LsJe.Focus()
            End If
        End If
    End Sub

    Private Sub LsJe_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LsJe.MouseDown
        If e.Clicks = 1 And e.Button = Windows.Forms.MouseButtons.Middle Then
            If LsJe.Items.Count > 0 Then SelectLsJe(LsJe.SelectedIndex + 1, True)
        End If
    End Sub

    Private Sub BtnJeSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnJeSel.Click

        If LsJe.Items.Count > 0 Then SelectLsJe(LsJe.SelectedIndex + 1)
    End Sub

    Private Sub BtnUBDHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUBDHide.Click
        TxUBD.Visible = Not (TxUBD.Visible)
    End Sub

    Private Sub HSTabPG_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HSTabPG.Scroll
        TabPG.Width = HSTabPG.Value
    End Sub

    Private Sub BtnHSTabPG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHSTabPG.Click
        TabPG.Width = 600
    End Sub

    Private Sub FPanelCM_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles FPanelCM.Paint

    End Sub

    Private Sub CbCM1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbCM1.SelectedIndexChanged
        If CbCM1.SelectedIndex > 0 Then INPUTCM_RTX(CbCM1.Text) : CbCM1.SelectedIndex = 0
    End Sub

    Private Sub CbCM2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbCM2.SelectedIndexChanged
        If CbCM2.SelectedIndex > 0 Then INPUTCM_RTX(CbCM2.Text) : CbCM2.SelectedIndex = 0
    End Sub

    Private Sub CbCM3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbCM3.SelectedIndexChanged
        If CbCM3.SelectedIndex > 0 Then INPUTCM_RTX(CbCM3.Text) : CbCM3.SelectedIndex = 0
    End Sub

    Private Sub BtnPanelCountWordHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPanelCountWordHide.Click
        PanelCountWord.Visible = Not (PanelCountWord.Visible)
        CountWord()
        If TabPG.TabPages.Count > 0 Then RTX.Focus()
    End Sub

    Private Sub OptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionToolStripMenuItem.Click
        FormSet.ShowDialog()
    End Sub

    Public Sub TxUBDSet()
        TxUBD.Font = UBD_Font
        TxUBD.ForeColor = UBD_FrColor
        TxUBD.BackColor = UBD_BkColor
    End Sub

    Public Sub FormatTotalRTX()
        Dim SAL As Integer
        SAL = RTX.SelectionStart
        If RTX.SelectionLength = 0 Then RTX.SelectAll()
        RTX.SelectionFont = RTX_Font
        RTX.SelectionColor = RTX_FrColor
        RTX.DeselectAll()
        RTX.SelectionStart = SAL
        'RTX.ClearUndo()
    End Sub

    Private Sub FormatAllToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormatAllToolStripButton.Click
        FormatTotalRTX()
    End Sub

    Public Sub RecognitionRTX()
        If RTX.Text.Length < 2 Then Exit Sub
        Dim i As Integer, AwL As String, AwR As String, SIdx As Integer, EndIdx As Integer, SAL As Integer
        Dim YFST As String = RecognitionToolStripButton.Text
        RecognitionToolStripButton.Text = YFST & "(請稍等...)"
        Application.DoEvents()
        TempRX.Rtf = RTX.Rtf

        SAL = RTX.SelectionStart
        For i = 0 To RecogLeft.Length
            If i = RecogLeft.Length Then
                Dim FindIndex As Integer = -1, RetV1, RetV2 As Integer
                TempRX.SelectionStart = 0
                Do
                    FindIndex = TempRX.SelectionStart
                    If RTX.SelectionLength <> 0 Then FindIndex += TempRX.SelectionLength
                    If FindIndex > TempRX.Text.Length Then FindIndex = TempRX.Text.Length
                    RetV1 = TempRX.Find(NSplitL, FindIndex, RichTextBoxFinds.None)
                    TempRX.SelectionColor = RTX_ResColor
                    TempRX.SelectionStart = TempRX.SelectionStart + TempRX.SelectionLength : TempRX.SelectionLength = 1
                    TempRX.SelectionColor = RTX_FrColor
                    If RetV1 < FindIndex Then
                        RetV1 = -1
                    Else
                        RetV2 = TempRX.Find(NSplitR, RetV1, RichTextBoxFinds.None)
                        TempRX.SelectionColor = RTX_ResColor
                        TempRX.SelectionStart = TempRX.SelectionStart + TempRX.SelectionLength : TempRX.SelectionLength = 1
                        TempRX.SelectionColor = RTX_FrColor
                    End If
                Loop Until RetV1 = -1 Or RetV1 = RTX.Text.Length
            Else
                AwL = RecogLeft(i)
                AwR = RecogRight(i)
                SIdx = -1
                Do
                    SIdx = TempRX.Text.IndexOf(AwL, SIdx + 1)
                    If SIdx <> -1 Then
                        TempRX.SelectionStart = SIdx
                        EndIdx = TempRX.Text.IndexOf(AwR, SIdx)
                        If EndIdx = -1 Then
                            TempRX.SelectionLength = TempRX.Text.Length - SIdx
                        Else
                            TempRX.SelectionLength = EndIdx - SIdx + 1
                        End If
                        TempRX.SelectionColor = RTX_ResColor
                    End If
                Loop Until SIdx = -1 Or EndIdx = -1
            End If
        Next
        RTX.SelectAll()
        RTX.SelectedRtf = TempRX.Rtf
        RTX.DeselectAll()
        RTX.SelectionStart = SAL
        RecognitionToolStripButton.Text = YFST
    End Sub

    Private Sub RecognitionToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecognitionToolStripButton.Click
        RecognitionRTX()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick


        AutoBakTimer += Timer1.Interval / 1000 '10 '十秒加十

        LbKT.Text = CTimeT(AutoBakTimer)
        If PEyeWork = True Then
            If AutoBakTimer - BeforePEyeTime > (PEyeMin * 60) Then Timer2.Enabled = True
        End If
    End Sub

    Private Sub BSJoinMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSJoinMenu.Click
        Dim BSFlg As Boolean
        Dim YBS As String = BSItem & vbCrLf
        BSFlg = AddJeNew()
        If BSFlg = True Then ProcBoxProc(YBS, True)

    End Sub

    Private Function BSExpertHTML(Optional ByVal ExHide As Boolean = False, Optional ByVal OpenNew As Boolean = False) As String
        On Error Resume Next
        BSExpertHTML = ""
        Dim SplStAryL() As String = {NSplitL} '將NSplitL設為字串陣列分割記號
        Dim SplStAryR() As String = {NSplitR} '將NSplitR設為字串陣列分割記號
        Dim NBSV() As String = RTX.Text.Split(SplStAryL, StringSplitOptions.None)
        Dim KBody() As String, PBody() As String, TBody() As String
        Dim AllHtml As String = "", OneHtml As String
        Dim i, j, UsrCh As Integer
        For i = 1 To NBSV.Length - 1
            KBody = NBSV(i).Split(SplStAryR, StringSplitOptions.None)
            'Kbody(0)為標題名 , KBody(1)為內部設定
            PBody = KBody(1).Split(BSqtR)
            'PBody利用右括將內部設定每一項細分開來
            OneHtml = vbCrLf & NSplitL & KBody(0) & NSplitR & vbCrLf & "<div><h2>" & KBody(0) & "</h2><br /><table border=""1"">" & vbCrLf '先清空重置
            For j = 0 To PBody.Length - 2
                TBody = PBody(j).Split(BSqtL)
                TBody(0).Replace(vbCr, Nothing) : TBody(0).Replace(vbLf, Nothing)
                If ExHide = False And TBody(0).Substring(1, 1) = BSHide Then
                Else
                    TBody(1) = TBody(1).Replace(vbCr, Nothing)
                    TBody(1) = TBody(1).Replace(vbLf, "<br />")
                    OneHtml = OneHtml & "<tr><td>" & TBody(0) & "</td><td>" & TBody(1) & "</td></tr>" & vbCrLf
                End If
            Next
            OneHtml = OneHtml & "</table></div><br />" & vbCrLf
            AllHtml = AllHtml & OneHtml
        Next
        UsrCh = MessageBox.Show(Msg_BSExpertTitleDo, "Ask", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If UsrCh = DialogResult.Yes Then
            BSExpertHTML = HtmlHeadA & AllHtml & HtmlDLast
        Else
            BSExpertHTML = AllHtml
        End If
        If OpenNew = True Then
            OpenTxt("", True, ".htm")
            ProcBoxProc(BSExpertHTML, True)
        End If
        RefreshLsJe()
    End Function

    Private Sub BSExpertHTMLMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSExpertHTMLMenu.Click
        BSExpertHTML(False, True)
    End Sub

    Private Sub BSExpertAllHTMLMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSExpertAllHTMLMenu.Click
        BSExpertHTML(True, True)
    End Sub

    Public Function RemoveHideBodySetItem(Optional ByVal OpenNew As Boolean = False) As String
        On Error Resume Next
        RemoveHideBodySetItem = ""
        Dim SplStAryL() As String = {NSplitL} '將NSplitL設為字串陣列分割記號
        Dim SplStAryR() As String = {NSplitR} '將NSplitR設為字串陣列分割記號
        Dim NBSV() As String = RTX.Text.Split(SplStAryL, StringSplitOptions.None)
        Dim KBody() As String, PBody() As String, TBody() As String
        Dim AllBSTx As String = "", OneBSTx As String
        Dim i, j As Integer
        For i = 1 To NBSV.Length - 1
            KBody = NBSV(i).Split(SplStAryR, StringSplitOptions.None)
            'Kbody(0)為標題名 , KBody(1)為內部設定
            PBody = KBody(1).Split(BSqtR)
            'PBody利用右括將內部設定每一項細分開來
            OneBSTx = vbCrLf & NSplitL & KBody(0) & NSplitR & vbCrLf '先清空重置
            For j = 0 To PBody.Length - 1
                TBody = PBody(j).Split(BSqtL)
                TBody(0).Replace(vbCr, Nothing) : TBody(0).Replace(vbLf, Nothing)
                If TBody(0).Substring(1, 1) = BSHide Then
                Else
                    If TBody(1) = "" Or TBody(1) = Nothing Then
                    Else
                        OneBSTx = OneBSTx & TBody(0) & BSqtL & TBody(1) & BSqtR
                    End If
                End If
            Next
            AllBSTx = AllBSTx & OneBSTx & vbCrLf
        Next
        RemoveHideBodySetItem = AllBSTx
        If OpenNew = True Then
            OpenTxt("", True)
            ProcBoxProc(RemoveHideBodySetItem, True)
        End If
        RefreshLsJe()
    End Function

    Private Sub AllselectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllselectToolStripMenuItem.Click
        RTX.SelectAll()
    End Sub

    Private Sub 關於AToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 關於AToolStripMenuItem.Click
        FormAbout.ShowDialog()
    End Sub

    Private Sub WeblogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeblogToolStripMenuItem.Click
        'ShellExecute(0, "Open", "http://blog.yam.com/user/melix.html", "", "", 3)
        System.Diagnostics.Process.Start("http://blog.yam.com/melix/article/25769065")
    End Sub

    Private Function CTimeT(ByVal SecA As Integer) As String
        CTimeT = ""
        If STime = False Then Exit Function
        Dim HH As Integer, MM As Integer, RTVT As String = "已打了"
        If SecA >= 3600 Then
            HH = SecA \ 3600
            SecA = SecA - HH * 3600
            RTVT = RTVT & HH.ToString & "小時"
        End If
        If SecA >= 60 Then
            MM = SecA \ 60
            SecA = SecA - MM * 60
            RTVT = RTVT & MM.ToString & "分"
        End If
        CTimeT = RTVT & SecA.ToString & "秒"
    End Function

    Private Sub RemoveJeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveJeToolStripMenuItem.Click
        RemoveAllJe(True)
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        LbPEye.Visible = True
        If LbPEye.ForeColor = Color.Yellow Then
            LbPEye.ForeColor = Color.Red
        Else
            LbPEye.ForeColor = Color.Yellow
        End If
    End Sub

    Private Sub LbPEye_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LbPEye.Click
        LbPEye.Visible = False
        Timer2.Enabled = False
        BeforePEyeTime = AutoBakTimer
    End Sub

    Private Sub TransHTMLMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransHTMLMenu.Click
        Dim NVEE As String = RTX.Text.Replace(vbCrLf, vbLf)
        Dim RET As Integer
        RET = MessageBox.Show("是否要加檔頭和結尾？ ", "Do you want <html><body> ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If RET = DialogResult.Yes Then
            NVEE = NVEE.Replace(vbCr, "")
            NVEE = NVEE.Replace(vbLf, "</p>" & vbCrLf & "<p>")
            NVEE = "<p>" & NVEE & "</p>"
            NVEE = NVEE.Replace("<p></p>", "<br />")
            NVEE = NVEE.Replace("<p>" & NSplitL, NSplitL)
            NVEE = NVEE.Replace(NSplitR & "</p>", NSplitR)
            NVEE = HtmlHeadA & NVEE & HtmlDLast
        Else
            NVEE = NVEE.Replace(vbCr, "")
            NVEE = NVEE.Replace(vbLf, "<br />" & vbCrLf)
        End If
        OpenTxt("", True, ".htm")
        ProcBoxProc(NVEE, True)
        RefreshLsJe()
    End Sub

    Private Sub TransBBSMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransBBSMenu.Click
        Dim NVEE As String = RTX.Text.Replace(vbCrLf, vbLf)
        If NVEE.Length < 1 Then Exit Sub
        Dim i, CAC As Integer, KW As String = "", DDW As Char
        If NVEE.Length < 1 Then Exit Sub
        NVEE = NVEE.Replace(vbCr, "")
        NVEE = NVEE.Replace("	", "") 'Tab
        CAC = 0
        For i = 0 To NVEE.Length - 1
            DDW = NVEE.Substring(i, 1)
            If AscW(DDW) > 255 Or AscW(DDW) < 0 Then
                CAC += 2
                KW = KW & DDW
            ElseIf AscW(DDW) = 10 Then 'vbLf
                CAC = 99
            Else
                CAC += 1
                KW = KW & DDW
            End If
            If CAC > 76 Then
                KW = KW & vbCrLf
                CAC = 0
            End If
        Next
        KW = KW & vbCrLf
        OpenTxt("", True, ".txt")
        ProcBoxProc(KW, True)
        RefreshLsJe()
    End Sub

    Private Sub 內容CToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 內容CToolStripMenuItem.Click
        Dim YBS As String = FormResUse.TxHelp.Text & vbCrLf

        OpenTxt("", True, ".txt")
        ProcBoxProc(YBS, True)
        RefreshLsJe()
        RTX.SelectionStart = 0

    End Sub

    Private Sub TxUBD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxUBD.KeyDown
        '特殊鍵盤符號輸入
        If Control.ModifierKeys = Keys.Control Then
            If e.KeyValue = Keys.A Then TxUBD.SelectAll()
            If e.KeyCode = 188 Then TxUBD.SelectedText = (DotCM(0)) '按,輸入全形逗號
            If e.KeyCode = 190 Then TxUBD.SelectedText = (DotCM(1)) '按.輸入全形句號
            If e.KeyValue = Keys.OemSemicolon Then TxUBD.SelectedText = (DotCM(4)) '按;輸入全形分號
            If e.KeyValue = Keys.OemQuotes Then TxUBD.SelectedText = (DotCM(39)) '按"輸入全形美式雙引號
            If e.KeyValue = Keys.Q Then TxUBD.SelectedText = (DotCM(34)) '按Q輸入中文引號
            If e.KeyValue = Keys.W Then TxUBD.SelectedText = (DotCM(35)) '按W輸入中文雙引號
            If e.KeyValue = Keys.D3 Then TxUBD.SelectedText = (DotCM(2)) '按3輸入中文頓號
            If e.KeyValue = Keys.D4 Then TxUBD.SelectedText = (DotCM(9)) '按4輸入中文點點點
        End If
    End Sub

    Private Sub TxUBD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxUBD.TextChanged

    End Sub

    Private Sub HelpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripButton.Click
        內容CToolStripMenuItem_Click(sender, e)
    End Sub

    Private Function ConvertStr(ByVal JobA As Integer, Optional ByVal GiveRTX As Boolean = False) As String
        ConvertStr = ""
        'JobA值 1=半形轉全形, 2=全形轉半形, 3=簡轉繁, 4=繁轉簡
        Dim NVRTX As String, YT As Integer = RTX.SelectionStart
        If RTX.SelectionLength = 0 Then
            NVRTX = RTX.Text
        Else
            NVRTX = RTX.SelectedText
        End If

        Select Case (JobA)
            Case 1
                NVRTX = StrConv(NVRTX, VbStrConv.Wide)
            Case 2
                NVRTX = NVRTX.Replace("。", ".")
                NVRTX = NVRTX.Replace("「", """")
                NVRTX = NVRTX.Replace("」", """")
                NVRTX = StrConv(NVRTX, VbStrConv.Narrow)
            Case 3
                NVRTX = StrConv(NVRTX, VbStrConv.TraditionalChinese, 2052)
            Case 4
                NVRTX = StrConv(NVRTX, VbStrConv.SimplifiedChinese, 2052)

        End Select

        ConvertStr = NVRTX
        If GiveRTX = True Then
            If RTX.SelectionLength = 0 Then
                RTX.SelectAll()
            End If
            RTX.SelectedText = NVRTX
            RTX.SelectionStart = YT
        End If
    End Function

    Private Sub H2FMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles H2FMenu.Click
        ConvertStr(1, True)
    End Sub

    Private Sub F2HMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles F2HMenu.Click
        ConvertStr(2, True)
    End Sub

    Private Sub S2TMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles S2TMenu.Click
        ConvertStr(3, True)
    End Sub

    Private Sub T2SMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T2SMenu.Click
        ConvertStr(4, True)
    End Sub

    Private Sub RecentFileProc(Optional ByVal AddNew As String = "")
        If RecentFile = "" And AddNew = "" Then Exit Sub
        Dim ReF(0) As String, ReFL As Integer = -1, i, FStp As Integer
        Dim HRTT As String = "", RRTT As String = ""
        If RecentFile <> "" Then
            ReF = RecentFile.Split("|")
            ReFL = ReF.Length - 1
            If AddNew <> "" Then
                ReDim Preserve ReF(ReFL + 1)
                FStp = ReFL + 1
                For i = 0 To ReFL
                    If AddNew = ReF(i) Then
                        FStp = i
                    End If
                Next
                For i = FStp To 1 Step -1
                    ReF(i) = ReF(i - 1)
                Next
                ReF(0) = AddNew
                ReFL = ReF.Length - 1 '重新定位
            End If
        Else
            If AddNew <> "" Then
                'ReDim ReF(0)
                ReF(0) = AddNew
                ReFL = ReF.Length - 1 '重新定位
            End If
        End If
        If ReFL > 4 Then ReFL = 4
        For i = 0 To ReFL
            If ReF(i) = "" Then Exit For
            HRTT = HRTT & ReF(i) & "|"
            RRTT = ReF(i).Substring(ReF(i).LastIndexOf("\") + 1)
            If i = 0 Then R1FMenu.ToolTipText = ReF(i) : R1FMenu.Text = RRTT
            If i = 1 Then R2FMenu.ToolTipText = ReF(i) : R2FMenu.Text = RRTT
            If i = 2 Then R3FMenu.ToolTipText = ReF(i) : R3FMenu.Text = RRTT
            If i = 3 Then R4FMenu.ToolTipText = ReF(i) : R4FMenu.Text = RRTT
            If i = 4 Then R5FMenu.ToolTipText = ReF(i) : R5FMenu.Text = RRTT
        Next
        HRTT = HRTT.Substring(0, HRTT.Length - 1)
        RecentFile = HRTT
    End Sub

    Private Sub R1FMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R1FMenu.Click
        OpenTxt(R1FMenu.ToolTipText)
    End Sub

    Private Sub R2FMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R2FMenu.Click
        OpenTxt(R2FMenu.ToolTipText)
    End Sub

    Private Sub R3FMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R3FMenu.Click
        OpenTxt(R3FMenu.ToolTipText)
    End Sub

    Private Sub R4FMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R4FMenu.Click
        OpenTxt(R4FMenu.ToolTipText)
    End Sub

    Private Sub R5FMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R5FMenu.Click
        OpenTxt(R5FMenu.ToolTipText)
    End Sub

    Private Sub RTXcopyUBD(Optional ByVal CoverIt As Boolean = False)
        Dim NCPP As String
        If RTX.SelectionLength = 0 Then
            NCPP = RTX.Text
        Else
            NCPP = RTX.SelectedText
        End If
        NCPP = NCPP.Replace(vbCr, "")
        NCPP = NCPP.Replace(vbLf, vbCrLf)

        If CoverIt = False Then
            TxUBD.AppendText(vbCrLf & "==================" & vbCrLf & NCPP)
        Else
            TxUBD.Text = NCPP
        End If
    End Sub

    Private Sub ApndUBDPopup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApndUBDPopup.Click
        RTXcopyUBD()
    End Sub

    Private Sub CvrUBDPopup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CvrUBDPopup.Click
        RTXcopyUBD(True)
    End Sub

    Private Sub ShuQianProc(Optional ByVal AddNew As String = "", Optional ByVal RmvName As String = "")
        If AddNew Is Nothing Then Exit Sub
        If BkMk = "" And AddNew = "" Then Exit Sub
        Dim ReF(0) As String, ReY(1) As String, ReFL As Integer = -1, i As Integer
        Dim HRTT As String = "", RRTT As String = "", AddNewName As String = ""
        If AddNew <> "" Then
            For i = 0 To ShuQian.Length - 1
                Try
                    If ShuQian(i).ToolTipText = AddNew Then Exit Sub
                Catch ex As Exception

                End Try
            Next
            AddNewName = AddNew.Substring(AddNew.LastIndexOf("\") + 1)
            AddNewName = InputBox(Msg_InputShuQian1, "Question", AddNewName)
            If AddNewName = "" Then Exit Sub
            If BkMk <> "" Then BkMk = BkMk & "|"
            BkMk = BkMk & AddNew & "?" & AddNewName
        End If
        If RmvName <> "" Then
            Dim DELR As Integer '刪字INDEX
            DELR = BkMk.IndexOf(RmvName)
            If DELR >= 0 Then
                Dim ENDIDX As Integer
                ENDIDX = BkMk.IndexOf("|", DELR)
                If ENDIDX = -1 Then ENDIDX = BkMk.Length
                If DELR > 1 Then DELR -= 1
                BkMk = BkMk.Remove(DELR, ENDIDX - DELR)
                '刪書籤功能表陣列
                For i = 0 To ShuQian.Length - 1
                    If ShuQian(i).ToolTipText = RmvName Then ShuQian(i).ToolTipText = Nothing
                Next
            End If
        End If
        For i = 0 To ShuQian.Length - 1
            BookmarkToolStripMenuItem.DropDownItems.Remove(ShuQian(i))
        Next i
        If BkMk = "" Then Exit Sub
        '用?分隔前面為路徑後面為名稱,用|分隔每一筆書籤,並去頭尾的|
        If BkMk.Substring(0, 1) = "|" Then BkMk = BkMk.Remove(0, 1)
        If BkMk.Substring(BkMk.Length - 1, 1) = "|" Then BkMk = BkMk.Remove(BkMk.Length - 1, 1)
        ReF = BkMk.Split("|")
        ReFL = ReF.Length - 1
        ReDim ShuQian(ReFL)

        For i = 0 To ReFL
            If ReF(i) = "" Or ReF(i) Is Nothing Then Continue For
            ReY = ReF(i).Split("?")
            ShuQian(i) = BookmarkToolStripMenuItem.DropDownItems.Add(ReY(1))
            ShuQian(i).ToolTipText = ReY(0)
            ShuQian(i).Name = "BKMK" & CStr(i)
            HRTT = HRTT & ReY(0) & "?" & ReY(1) & "|"
            AddHandler ShuQian(i).MouseUp, AddressOf ShuQian_MouseUp
        Next
        AddHandler BookmarkToolStripMenuItem.DropDownItemClicked, AddressOf ShuQian_Click
        HRTT = HRTT.Substring(0, HRTT.Length - 1)
        BkMk = HRTT
    End Sub

    Private Sub ShuQian_Click(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles BookmarkToolStripMenuItem.DropDownItemClicked
        If e.ClickedItem.Name = "AddSQMenu" Then
            ShuQianProc(arrFILE(NowKeyV).FilePath) '按的是加入書籤
        Else
            OpenTxt(e.ClickedItem.ToolTipText)
        End If
    End Sub

    Private Sub ShuQian_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PageToolStripMenuItem.MouseUp
        Dim i As Integer, RET As Integer
        If e.Button = Windows.Forms.MouseButtons.Right Then
            RET = MessageBox.Show("真的要刪除 " & sender.ToString & " ?", "DELETE", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If RET = DialogResult.Yes Then
                For i = 0 To ShuQian.Length - 1
                    If sender.Equals(ShuQian(i)) = True Then
                        ShuQianProc(, ShuQian(i).ToolTipText)
                        Exit Sub
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub QLoadMenu()
        If arrFILE(NowKeyV).FilePath <> "" And OnClosingFlag = False Then
            Dim FLAry() As String, i As Integer = 0, TGU As Integer
            Dim PPth As String = arrFILE(NowKeyV).FilePath.Substring(0, arrFILE(NowKeyV).FilePath.LastIndexOf("\"))
            FLAry = System.IO.Directory.GetFiles(PPth, "*.txt", IO.SearchOption.TopDirectoryOnly)
            ReDim QLoad(FLAry.Length - 1)
            QuickLoadToolStripMenuItem.DropDownItems.Clear()
            TGU = 0
            For Each FN As String In FLAry
                QLoad(TGU) = QuickLoadToolStripMenuItem.DropDownItems.Add(FN.Substring(FN.LastIndexOf("\") + 1))
                QLoad(TGU).ToolTipText = FN
                TGU += 1
            Next
            AddHandler QuickLoadToolStripMenuItem.DropDownItemClicked, AddressOf QLoad_Click

        End If
    End Sub

    Private Sub QLoad_Click(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles QuickLoadToolStripMenuItem.DropDownItemClicked
        OpenTxt(e.ClickedItem.ToolTipText)
    End Sub

    Private Sub UpVerDetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpVerDetToolStripMenuItem.Click
        VerDet(True)
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        
        StrForPrintProc(RTX.Text)
        PDoc.Print()
    End Sub

    Public Sub StrForPrintProc(ByVal PrintStr As String)
        Dim i, w As Integer, SRP As String = ""
        Dim pf As New PointF(0, 0)
        Dim g As Graphics = Graphics.FromHwnd(RTX.Handle)
        Dim tmp As String = ""
        For i = 0 To PrintStr.Length - 1
            w = g.MeasureString(tmp + PrintStr(i), RTX_Font, pf, Nothing).Width
            If w >= 650 Then
                SRP = SRP + tmp + vbCrLf
                tmp = PrintStr(i)
            Else
                tmp = tmp & PrintStr(i)
            End If
            Application.DoEvents()
        Next
        StrForPrint = SRP
        g.Dispose()
    End Sub

    Private Sub PDoc_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PDoc.PrintPage
        Dim charactersOnPage As Integer = 0
        Dim linesPerPage As Integer = 0
        Dim FontAP As Font = RTX_Font

        Dim pta1 As New RectangleF(e.MarginBounds.Left - 40, e.MarginBounds.Top - 40, e.MarginBounds.Right, e.MarginBounds.Bottom)
        e.Graphics.MeasureString(StrForPrint, FontAP, pta1.Size, StringFormat.GenericTypographic, charactersOnPage, linesPerPage)
        e.Graphics.DrawString(StrForPrint, FontAP, Brushes.Black, pta1)
        StrForPrint = StrForPrint.Substring(charactersOnPage)
        e.HasMorePages = StrForPrint.Length > 0
    End Sub

    Private Sub PPreView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PPreView.Load

        PPreView.Document = PDoc

    End Sub

    Private Sub PrintViewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintViewToolStripMenuItem.Click

        StrForPrintProc(RTX.Text)
        
        PPreView.ShowDialog()
    End Sub

    Private Sub AddOutRightMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddOutRightMenu.Click
        RightMenuQuickAR()
    End Sub

    Private Sub RemoveOutRightMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveOutRightMenu.Click
        RightMenuQuickAR(True)
    End Sub

    Private Sub BtnTabL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTabL.Click
        If TabPG.TabCount < 1 Then Exit Sub
        If TabPG.SelectedTab IsNot TabPG.TabPages(0) Then
            Dim i As Integer, KUQ As TabPage
            For i = 0 To TabPG.TabCount - 1
                If TabPG.SelectedTab Is TabPG.TabPages(i) Then
                    KUQ = TabPG.TabPages(i - 1)
                    TabPG.TabPages(i - 1) = TabPG.TabPages(i)
                    TabPG.TabPages(i) = KUQ
                    TabPG.SelectTab(i - 1)
                    Exit For
                End If
            Next
        End If
        RTX.Focus()
    End Sub

    Private Sub BtnTabR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTabR.Click
        If TabPG.TabCount < 1 Then Exit Sub
        If TabPG.SelectedTab IsNot TabPG.TabPages(TabPG.TabCount - 1) Then
            Dim i As Integer, KUQ As TabPage
            For i = 0 To TabPG.TabCount - 1
                If TabPG.SelectedTab Is TabPG.TabPages(i) Then
                    KUQ = TabPG.TabPages(i + 1)
                    TabPG.TabPages(i + 1) = TabPG.TabPages(i)
                    TabPG.TabPages(i) = KUQ
                    TabPG.SelectTab(i + 1)
                    Exit For
                End If
            Next
        End If
        RTX.Focus()
    End Sub

    Private Sub BtnCloseTabPG_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCloseTabPG.Move
        BtnTabL.Left = BtnCloseTabPG.Left - 30
        BtnTabR.Left = BtnCloseTabPG.Left - 16
    End Sub

    Private Sub TimerResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerResetToolStripMenuItem.Click
        AutoBakTimer = 0
        BeforeBakTime = 0
        BeforePEyeTime = 0
        LbKT.Text = CTimeT(AutoBakTimer)
    End Sub

    Private Sub RemoveHideSetMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveHideSetMenu.Click
        RemoveHideBodySetItem(True)
    End Sub

    Private Sub MenuAddJe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAddJe.Click
        AddJeNew()
    End Sub

    Private Sub MenuDelJe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDelJe.Click
        If LsJe.Items.Count > 0 Then DeleteLsJe(LsJe.SelectedIndex + 1)
    End Sub

    Private Sub MenuRefreshJe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRefreshJe.Click
        RefreshLsJe()
    End Sub

    Private Sub MenuSel1Je_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSel1Je.Click
        If LsJe.Items.Count > 0 Then JumpLsJe(LsJe.SelectedIndex + 1)
    End Sub

    Private Sub MenuSel2Je_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSel2Je.Click
        If LsJe.Items.Count > 0 Then SelectLsJe(LsJe.SelectedIndex + 1)
    End Sub

    Private Sub MenuSel3Je_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuSel3Je.Click
        If LsJe.Items.Count > 0 Then SelectLsJe(LsJe.SelectedIndex + 1, True)
    End Sub

    Private Sub MenuDelArtJe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuDelArtJe.Click
        Dim RETV As Integer
        RETV = MessageBox.Show(Msg_DeleteArtJe, "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If RETV = DialogResult.OK Then
            If LsJe.Items.Count > 0 Then
                SelectLsJe(LsJe.SelectedIndex + 1, True)
                RTX.SelectedText = ""
                RefreshLsJe()
                LsJe.Focus()
            End If
        End If
    End Sub

    Private Sub FmtChange(Optional ByVal GoChange As Boolean = False, Optional ByVal FmtU8V As Boolean = False)
        If GoChange = True Then
            arrFILE(NowKeyV).FmtU8 = FmtU8V
        End If
        FmtUTF8Menu.Checked = arrFILE(NowKeyV).FmtU8
        FmtDefaultMenu.Checked = Not arrFILE(NowKeyV).FmtU8

    End Sub

    Private Sub FmtUTF8Menu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FmtUTF8Menu.Click
        FmtChange(True, True)
    End Sub

    Private Sub FmtDefaultMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FmtDefaultMenu.Click
        FmtChange(True, False)
    End Sub
End Class
