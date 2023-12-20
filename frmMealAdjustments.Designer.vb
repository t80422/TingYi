<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMealAdjustments
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
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtBreakfastAdj = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtLunchAdj = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDinnerAdj = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnSubtract_break = New System.Windows.Forms.Button()
        Me.btnSubtract_lunch = New System.Windows.Forms.Button()
        Me.btnSubtract_dinner = New System.Windows.Forms.Button()
        Me.btnAdd_break = New System.Windows.Forms.Button()
        Me.btnAdd_lunch = New System.Windows.Forms.Button()
        Me.btnAdd_dinner = New System.Windows.Forms.Button()
        Me.txtDinnerCount = New System.Windows.Forms.TextBox()
        Me.txtLunchCount = New System.Windows.Forms.TextBox()
        Me.txtBreakfastCount = New System.Windows.Forms.TextBox()
        Me.txtDinnerResult = New System.Windows.Forms.TextBox()
        Me.txtLunchResult = New System.Windows.Forms.TextBox()
        Me.txtBreakfastResult = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtBreakfastAdj
        '
        Me.txtBreakfastAdj.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtBreakfastAdj.Location = New System.Drawing.Point(179, 42)
        Me.txtBreakfastAdj.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBreakfastAdj.Name = "txtBreakfastAdj"
        Me.txtBreakfastAdj.ReadOnly = True
        Me.txtBreakfastAdj.Size = New System.Drawing.Size(38, 27)
        Me.txtBreakfastAdj.TabIndex = 3
        Me.txtBreakfastAdj.Text = "0"
        Me.txtBreakfastAdj.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(18, 44)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "早餐"
        '
        'txtLunchAdj
        '
        Me.txtLunchAdj.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtLunchAdj.Location = New System.Drawing.Point(179, 77)
        Me.txtLunchAdj.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLunchAdj.Name = "txtLunchAdj"
        Me.txtLunchAdj.ReadOnly = True
        Me.txtLunchAdj.Size = New System.Drawing.Size(38, 27)
        Me.txtLunchAdj.TabIndex = 5
        Me.txtLunchAdj.Text = "0"
        Me.txtLunchAdj.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label3.Location = New System.Drawing.Point(18, 79)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "午餐"
        '
        'txtDinnerAdj
        '
        Me.txtDinnerAdj.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtDinnerAdj.Location = New System.Drawing.Point(179, 112)
        Me.txtDinnerAdj.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDinnerAdj.Name = "txtDinnerAdj"
        Me.txtDinnerAdj.ReadOnly = True
        Me.txtDinnerAdj.Size = New System.Drawing.Size(38, 27)
        Me.txtDinnerAdj.TabIndex = 7
        Me.txtDinnerAdj.Text = "0"
        Me.txtDinnerAdj.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label4.Location = New System.Drawing.Point(18, 114)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "晚餐"
        '
        'btnSubtract_break
        '
        Me.btnSubtract_break.AutoSize = True
        Me.btnSubtract_break.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnSubtract_break.Location = New System.Drawing.Point(146, 42)
        Me.btnSubtract_break.Name = "btnSubtract_break"
        Me.btnSubtract_break.Size = New System.Drawing.Size(26, 26)
        Me.btnSubtract_break.TabIndex = 8
        Me.btnSubtract_break.Text = "<"
        Me.btnSubtract_break.UseVisualStyleBackColor = True
        '
        'btnSubtract_lunch
        '
        Me.btnSubtract_lunch.AutoSize = True
        Me.btnSubtract_lunch.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnSubtract_lunch.Location = New System.Drawing.Point(146, 76)
        Me.btnSubtract_lunch.Name = "btnSubtract_lunch"
        Me.btnSubtract_lunch.Size = New System.Drawing.Size(26, 26)
        Me.btnSubtract_lunch.TabIndex = 9
        Me.btnSubtract_lunch.Text = "<"
        Me.btnSubtract_lunch.UseVisualStyleBackColor = True
        '
        'btnSubtract_dinner
        '
        Me.btnSubtract_dinner.AutoSize = True
        Me.btnSubtract_dinner.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnSubtract_dinner.Location = New System.Drawing.Point(146, 111)
        Me.btnSubtract_dinner.Name = "btnSubtract_dinner"
        Me.btnSubtract_dinner.Size = New System.Drawing.Size(26, 26)
        Me.btnSubtract_dinner.TabIndex = 10
        Me.btnSubtract_dinner.Text = "<"
        Me.btnSubtract_dinner.UseVisualStyleBackColor = True
        '
        'btnAdd_break
        '
        Me.btnAdd_break.AutoSize = True
        Me.btnAdd_break.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnAdd_break.Location = New System.Drawing.Point(224, 42)
        Me.btnAdd_break.Name = "btnAdd_break"
        Me.btnAdd_break.Size = New System.Drawing.Size(26, 26)
        Me.btnAdd_break.TabIndex = 11
        Me.btnAdd_break.Text = ">"
        Me.btnAdd_break.UseVisualStyleBackColor = True
        '
        'btnAdd_lunch
        '
        Me.btnAdd_lunch.AutoSize = True
        Me.btnAdd_lunch.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnAdd_lunch.Location = New System.Drawing.Point(224, 76)
        Me.btnAdd_lunch.Name = "btnAdd_lunch"
        Me.btnAdd_lunch.Size = New System.Drawing.Size(26, 26)
        Me.btnAdd_lunch.TabIndex = 12
        Me.btnAdd_lunch.Text = ">"
        Me.btnAdd_lunch.UseVisualStyleBackColor = True
        '
        'btnAdd_dinner
        '
        Me.btnAdd_dinner.AutoSize = True
        Me.btnAdd_dinner.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnAdd_dinner.Location = New System.Drawing.Point(224, 111)
        Me.btnAdd_dinner.Name = "btnAdd_dinner"
        Me.btnAdd_dinner.Size = New System.Drawing.Size(26, 26)
        Me.btnAdd_dinner.TabIndex = 13
        Me.btnAdd_dinner.Text = ">"
        Me.btnAdd_dinner.UseVisualStyleBackColor = True
        '
        'txtDinnerCount
        '
        Me.txtDinnerCount.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtDinnerCount.Location = New System.Drawing.Point(67, 111)
        Me.txtDinnerCount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDinnerCount.Name = "txtDinnerCount"
        Me.txtDinnerCount.ReadOnly = True
        Me.txtDinnerCount.Size = New System.Drawing.Size(72, 27)
        Me.txtDinnerCount.TabIndex = 20
        Me.txtDinnerCount.Tag = "ord_dinner"
        Me.txtDinnerCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLunchCount
        '
        Me.txtLunchCount.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtLunchCount.Location = New System.Drawing.Point(67, 76)
        Me.txtLunchCount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLunchCount.Name = "txtLunchCount"
        Me.txtLunchCount.ReadOnly = True
        Me.txtLunchCount.Size = New System.Drawing.Size(72, 27)
        Me.txtLunchCount.TabIndex = 19
        Me.txtLunchCount.Tag = "ord_lunch"
        Me.txtLunchCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBreakfastCount
        '
        Me.txtBreakfastCount.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtBreakfastCount.Location = New System.Drawing.Point(67, 41)
        Me.txtBreakfastCount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBreakfastCount.Name = "txtBreakfastCount"
        Me.txtBreakfastCount.ReadOnly = True
        Me.txtBreakfastCount.Size = New System.Drawing.Size(72, 27)
        Me.txtBreakfastCount.TabIndex = 18
        Me.txtBreakfastCount.Tag = "ord_breakfast"
        Me.txtBreakfastCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDinnerResult
        '
        Me.txtDinnerResult.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtDinnerResult.Location = New System.Drawing.Point(257, 111)
        Me.txtDinnerResult.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDinnerResult.Name = "txtDinnerResult"
        Me.txtDinnerResult.Size = New System.Drawing.Size(72, 27)
        Me.txtDinnerResult.TabIndex = 23
        Me.txtDinnerResult.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtLunchResult
        '
        Me.txtLunchResult.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtLunchResult.Location = New System.Drawing.Point(257, 76)
        Me.txtLunchResult.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLunchResult.Name = "txtLunchResult"
        Me.txtLunchResult.Size = New System.Drawing.Size(72, 27)
        Me.txtLunchResult.TabIndex = 22
        Me.txtLunchResult.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBreakfastResult
        '
        Me.txtBreakfastResult.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtBreakfastResult.Location = New System.Drawing.Point(257, 41)
        Me.txtBreakfastResult.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBreakfastResult.Name = "txtBreakfastResult"
        Me.txtBreakfastResult.Size = New System.Drawing.Size(72, 27)
        Me.txtBreakfastResult.TabIndex = 21
        Me.txtBreakfastResult.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(74, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 16)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "餐點數"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label5.Location = New System.Drawing.Point(176, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 16)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "調整"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label6.Location = New System.Drawing.Point(272, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(41, 16)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "結果"
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnSave.Location = New System.Drawing.Point(101, 161)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(94, 30)
        Me.btnSave.TabIndex = 27
        Me.btnSave.Text = "儲存"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnExit.Location = New System.Drawing.Point(201, 161)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(94, 30)
        Me.btnExit.TabIndex = 28
        Me.btnExit.Text = "取消"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmMealAdjustments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(360, 203)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtDinnerResult)
        Me.Controls.Add(Me.txtLunchResult)
        Me.Controls.Add(Me.txtBreakfastResult)
        Me.Controls.Add(Me.txtDinnerCount)
        Me.Controls.Add(Me.txtLunchCount)
        Me.Controls.Add(Me.txtBreakfastCount)
        Me.Controls.Add(Me.btnAdd_dinner)
        Me.Controls.Add(Me.btnAdd_lunch)
        Me.Controls.Add(Me.btnAdd_break)
        Me.Controls.Add(Me.btnSubtract_dinner)
        Me.Controls.Add(Me.btnSubtract_lunch)
        Me.Controls.Add(Me.btnSubtract_break)
        Me.Controls.Add(Me.txtDinnerAdj)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtLunchAdj)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtBreakfastAdj)
        Me.Controls.Add(Me.Label2)
        Me.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMealAdjustments"
        Me.Text = "配餐調整"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtBreakfastAdj As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtLunchAdj As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtDinnerAdj As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents btnSubtract_break As Button
    Friend WithEvents btnSubtract_lunch As Button
    Friend WithEvents btnSubtract_dinner As Button
    Friend WithEvents btnAdd_break As Button
    Friend WithEvents btnAdd_lunch As Button
    Friend WithEvents btnAdd_dinner As Button
    Friend WithEvents txtDinnerCount As TextBox
    Friend WithEvents txtLunchCount As TextBox
    Friend WithEvents txtBreakfastCount As TextBox
    Friend WithEvents txtDinnerResult As TextBox
    Friend WithEvents txtLunchResult As TextBox
    Friend WithEvents txtBreakfastResult As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents btnSave As Button
    Friend WithEvents btnExit As Button
End Class
