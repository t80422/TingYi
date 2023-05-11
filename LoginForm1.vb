Public Class LoginForm1

    ' TODO: 插入程式碼，利用提供的使用者名稱和密碼執行自訂驗證
    ' (請參閱 https://go.microsoft.com/fwlink/?LinkId=35339)。
    ' 如此便可將自訂主體附加到目前執行緒的主體，如下所示: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' 其中 CustomPrincipal 是用來執行驗證的 IPrincipal 實作。
    ' 接著，My.User 便會傳回封裝在 CustomPrincipal 物件中的識別資訊，
    ' 例如使用者名稱、顯示名稱等。

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If txtUsername.Text = "user" And txtPassword.Text = "user" Then
            frmMain.Show()
            frmMain.TabControl1.SelectedTab = frmMain.tpCustomer
            Me.Hide()
        Else
            MsgBox("帳號密碼錯誤")
        End If

        txtUsername.Clear()
        txtPassword.Clear()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Close()
    End Sub
End Class
