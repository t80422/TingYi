Public Class LoginForm1

    Private Sub OK_Click(sender As Object, e As EventArgs) Handles OK.Click
        Cursor = Cursors.WaitCursor

        Dim sql = "SELECT * FROM employee a " &
                  "LEFT JOIN permissions b ON a.emp_perm_id = b.perm_id " &
                 $"WHERE emp_acct = @emp_acct AND emp_psw = @emp_psw"

        Dim dic As New Dictionary(Of String, Object) From {
            {"emp_acct", txtUsername.Text},
            {"emp_psw", txtPassword.Text}
        }

        Dim rows = SelectTable(sql, dic).Rows

        If rows.Count > 0 Then
            frmMain.Show()

            For Each tp In frmMain.TabControl1.Controls.OfType(Of TabPage).Where(Function(x) Not x.Text = "µn¥X").ToList
                If rows(0)(tp.Tag) = 0 Then
                    tp.Parent = Nothing
                Else
                    tp.Parent = frmMain.TabControl1
                End If
            Next

            frmMain.TabControl1.SelectedIndex = 0
            Hide()

        Else
            MsgBox("±b¸¹±K½X¿ù»~")
        End If

        txtUsername.Clear()
        txtPassword.Clear()

        Cursor = Cursors.Default
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Close()
    End Sub
End Class
