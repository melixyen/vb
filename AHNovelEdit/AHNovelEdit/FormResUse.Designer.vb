<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormResUse
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormResUse))
        Me.TxBodySet = New System.Windows.Forms.TextBox
        Me.TxHelp = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'TxBodySet
        '
        Me.TxBodySet.Location = New System.Drawing.Point(1, 2)
        Me.TxBodySet.Multiline = True
        Me.TxBodySet.Name = "TxBodySet"
        Me.TxBodySet.Size = New System.Drawing.Size(19, 19)
        Me.TxBodySet.TabIndex = 0
        Me.TxBodySet.Text = resources.GetString("TxBodySet.Text")
        '
        'TxHelp
        '
        Me.TxHelp.Location = New System.Drawing.Point(26, 2)
        Me.TxHelp.Multiline = True
        Me.TxHelp.Name = "TxHelp"
        Me.TxHelp.Size = New System.Drawing.Size(17, 19)
        Me.TxHelp.TabIndex = 1
        Me.TxHelp.Text = resources.GetString("TxHelp.Text")
        '
        'FormResUse
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(571, 452)
        Me.Controls.Add(Me.TxHelp)
        Me.Controls.Add(Me.TxBodySet)
        Me.Name = "FormResUse"
        Me.Text = "FormResUse"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxBodySet As System.Windows.Forms.TextBox
    Friend WithEvents TxHelp As System.Windows.Forms.TextBox
End Class
