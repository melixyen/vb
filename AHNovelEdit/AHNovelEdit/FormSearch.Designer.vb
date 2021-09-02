<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSearch
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.CkUp = New System.Windows.Forms.CheckBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.TxFind = New System.Windows.Forms.ComboBox
        Me.TxReplace = New System.Windows.Forms.ComboBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "搜尋目標:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 12)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "取代目標:"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button7)
        Me.Panel1.Controls.Add(Me.Button6)
        Me.Panel1.Controls.Add(Me.CkUp)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Location = New System.Drawing.Point(78, 88)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(312, 57)
        Me.Panel1.TabIndex = 4
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(252, 4)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(54, 23)
        Me.Button7.TabIndex = 5
        Me.Button7.Text = "取代(&H)"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(252, 26)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(54, 23)
        Me.Button6.TabIndex = 4
        Me.Button6.Text = "關閉(&X)"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'CkUp
        '
        Me.CkUp.AutoSize = True
        Me.CkUp.Location = New System.Drawing.Point(4, 33)
        Me.CkUp.Name = "CkUp"
        Me.CkUp.Size = New System.Drawing.Size(88, 16)
        Me.CkUp.TabIndex = 3
        Me.CkUp.Text = "向上尋找(&U)"
        Me.CkUp.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(161, 26)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(85, 23)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "全部取代(&A)"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(94, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(152, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "取代這一個並找下一個(&R)"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "找下一個(&F)"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(321, 33)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(19, 21)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "v"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(340, 33)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(19, 21)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "^"
        Me.Button5.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button5.UseVisualStyleBackColor = True
        '
        'TxFind
        '
        Me.TxFind.FormattingEnabled = True
        Me.TxFind.Location = New System.Drawing.Point(78, 9)
        Me.TxFind.Name = "TxFind"
        Me.TxFind.Size = New System.Drawing.Size(281, 20)
        Me.TxFind.TabIndex = 7
        '
        'TxReplace
        '
        Me.TxReplace.FormattingEnabled = True
        Me.TxReplace.Location = New System.Drawing.Point(78, 55)
        Me.TxReplace.Name = "TxReplace"
        Me.TxReplace.Size = New System.Drawing.Size(281, 20)
        Me.TxReplace.TabIndex = 8
        '
        'FormSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 151)
        Me.Controls.Add(Me.TxReplace)
        Me.Controls.Add(Me.TxFind)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FormSearch"
        Me.ShowInTaskbar = False
        Me.Text = "Search"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents CkUp As System.Windows.Forms.CheckBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents TxFind As System.Windows.Forms.ComboBox
    Friend WithEvents TxReplace As System.Windows.Forms.ComboBox
End Class
