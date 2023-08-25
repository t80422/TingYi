Public Class frmLoadExcel
    Private lst As New List(Of String) From {".", "..", "..."}
    Private i As Integer
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label2.Text = lst(i)
        i += 1
        If i > lst.Count - 1 Then
            i = 0
        End If

    End Sub
End Class