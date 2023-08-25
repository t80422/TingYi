<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmTaboo
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.flpMain = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmbType
        '
        Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbType.Font = New System.Drawing.Font("標楷體", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Location = New System.Drawing.Point(12, 12)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(121, 27)
        Me.cmbType.TabIndex = 0
        '
        'flpMain
        '
        Me.flpMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.flpMain.AutoScroll = True
        Me.flpMain.BackColor = System.Drawing.Color.White
        Me.flpMain.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.flpMain.Location = New System.Drawing.Point(12, 44)
        Me.flpMain.Margin = New System.Windows.Forms.Padding(0)
        Me.flpMain.Name = "flpMain"
        Me.flpMain.Size = New System.Drawing.Size(250, 261)
        Me.flpMain.TabIndex = 1
        '
        'btnOK
        '
        Me.btnOK.AutoSize = True
        Me.btnOK.Font = New System.Drawing.Font("標楷體", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnOK.Location = New System.Drawing.Point(185, 12)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 29)
        Me.btnOK.TabIndex = 20
        Me.btnOK.Text = "確定"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'frmTaboo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(272, 314)
        Me.Controls.Add(Me.flpMain)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.btnOK)
        Me.Name = "frmTaboo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "禁忌清單"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbType As ComboBox
    Friend WithEvents flpMain As FlowLayoutPanel
    Friend WithEvents btnOK As Button
End Class
