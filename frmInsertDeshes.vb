Public Class frmInsertDeshes
    Public Property Dishes As List(Of String)

    Private Sub frmInsertDeshes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Columns.Add("dish_name", "菜色")
        DataGridView1.Columns.Add("dish_ingredients", "食材")

        With DataGridView1
            .ColumnHeadersDefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
            .DefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(224, 224, 224)
            .EnableHeadersVisualStyles = False
            .ColumnHeadersDefaultCellStyle.BackColor = Color.MediumTurquoise
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .ReadOnly = True
            .AllowUserToResizeColumns = True
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        For Each d In Dishes
            DataGridView1.Rows.Add(d)
        Next

    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Dim cells = DataGridView1.SelectedCells
        If cells.Count = 0 OrElse cells(0).ColumnIndex <> 1 Then Exit Sub
        Dim frm As New frmTaboo
        If frm.ShowDialog = DialogResult.OK Then cells(0).Value = frm.ReturnString
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim dic As New Dictionary(Of String, String)
        Dim table = "dishes"
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("dish_ingredients").Value IsNot Nothing Then
                dic.Add("dish_name", row.Cells("dish_name").Value)
                dic.Add("dish_ingredients", row.Cells("dish_ingredients").Value)
                DeleteData(table, $"dish_name = '{row.Cells("dish_name").Value}'")
                InserTable(table, dic)
            End If
        Next
        Close()
    End Sub
End Class