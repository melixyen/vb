Public Class FormSearch

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.Owner.Name.ToString = "FormEdit" Then
            FormEdit.FindReplace(TxFind.Text, False, TxReplace.Text)
            FormEdit.Focus()
        End If
        AddLastFR()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Me.Owner.Name.ToString = "FormEdit" Then
            FormEdit.FindReplace(TxFind.Text, True, TxReplace.Text)
            FormEdit.Focus()
        End If
        AddLastFR()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Me.Owner.Name.ToString = "FormEdit" Then
            FormEdit.FindReplace(TxFind.Text, True, TxReplace.Text, True)
            FormEdit.Focus()
        End If
        AddLastFR()
    End Sub

    Private Sub FormSearch_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        AddListCbFR()
        CkUp.Checked = FindUp
        TxFind.Text = FindWord
        TxReplace.Text = ReplaceWord
        TxFind.Focus()

    End Sub

    Private Sub FormSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Left = Screen.PrimaryScreen.WorkingArea.Right - Me.Width - 20
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TxReplace.Text = TxFind.Text
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TxFind.Text = TxReplace.Text
    End Sub

    Private Sub CkUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CkUp.CheckedChanged
        FindUp = CkUp.Checked
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If Me.Owner.Name.ToString = "FormEdit" Then
            If FormEdit.RTX.SelectionLength > 0 Then
                FormEdit.ProcBoxProc(TxReplace.Text, True)
                FormEdit.Focus()
            End If
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub TxFind_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxFind.KeyDown
        If e.KeyCode = Keys.Enter Then Button1_Click(sender, e)
    End Sub

    Private Sub TxFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxFind.TextChanged
        FindWord = TxFind.Text
    End Sub


    Private Sub TxReplace_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxReplace.TextChanged
        ReplaceWord = TxReplace.Text
    End Sub

    Private Sub AddLastFR()
        Dim i As Integer, LastFR As Boolean = False
        For i = 0 To LastFind.Length - 1
            If LastFind(i) = TxFind.Text Then LastFR = True
        Next
        If LastFR = False Then
            ReDim Preserve LastFind(LastFind.Length)
            LastFind(i) = TxFind.Text
        End If
        LastFR = False

        For i = 0 To LastReplace.Length - 1
            If LastReplace(i) = TxReplace.Text Then LastFR = True
        Next
        If LastFR = False Then
            ReDim Preserve LastReplace(LastReplace.Length)
            LastReplace(i) = TxReplace.Text
        End If
    End Sub

    Private Sub AddListCbFR()
        Dim i As Integer
        TxFind.Items.Clear()
        TxReplace.Items.Clear()
        For i = LastFind.Length - 1 To 0 Step -1
            TxFind.Items.Add(LastFind(i))
        Next
        For i = LastReplace.Length - 1 To 0 Step -1
            TxReplace.Items.Add(LastReplace(i))
        Next
    End Sub

    Private Sub TxFind_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxFind.SelectedIndexChanged

    End Sub
End Class