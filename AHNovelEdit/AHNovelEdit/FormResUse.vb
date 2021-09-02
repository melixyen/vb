Public Class FormResUse

    Private Sub FormResUse_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.Hide()
    End Sub

    Private Sub FormResUse_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Hide()
        Me.Owner = FormEdit
    End Sub
End Class