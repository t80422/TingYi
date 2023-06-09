﻿Public Class frmTaboo
    Private dic As New Dictionary(Of Int32, Boolean)
    Private Sub frmTaboo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '初始化所有選項
        For Each row In frmMain.dtTaboo.AsEnumerable.Select(Function(x) x.Field(Of Int32)("tabo_id"))
            dic.Add(row, False)
        Next
        '初始化分類
        Dim lst As List(Of String) = frmMain.dtTaboo.AsEnumerable.Select(Function(row) row.Field(Of String)("tabo_type")).Distinct.ToList
        With cmbType
            .Items.Add("全部")
            For Each type As String In lst
                .Items.Add(type)
            Next
            .SelectedIndex = 0
        End With
    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        SaveCheck(dic)
        flpMain.Controls.Clear()
        Dim drs As DataRow()
        If sender.text = "全部" Then
            drs = frmMain.dtTaboo.Select()
        Else
            drs = frmMain.dtTaboo.Select($"tabo_type = '{cmbType.Text}'")
        End If
        For Each dr As DataRow In drs
            Dim chk As New CheckBox With {
                .Text = dr.Field(Of String)("tabo_name"),
                .Tag = dr.Field(Of Int32)("tabo_id")
            }
            flpMain.Controls.Add(chk)
        Next
        '將先前勾選的選項勾回去
        For Each chk As CheckBox In flpMain.Controls
            chk.Checked = dic(chk.Tag)
        Next
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        '將已勾選的選項標記至dic
        SaveCheck(dic)
        '將勾選的項目傳至禁忌(txtTaboo) id傳到tag,所以傳道資料庫要從tag抓 文字傳到text
        Dim list As List(Of Int32) = dic.Where(Function(item) item.Value = True).Select(Function(item) item.Key).ToList
        frmMain.txtTaboo.Tag = String.Join(",", list)
        Dim listText As New List(Of String)
        For Each txt As String In list
            Dim name = frmMain.dtTaboo.Select($"tabo_id = '{txt}'").FirstOrDefault.Field(Of String)("tabo_name")
            listText.Add(name)
        Next
        frmMain.txtTaboo.Text = String.Join(",", listText)
        Close()
    End Sub

    '將已勾選的選項紀錄至dic
    Private Sub SaveCheck(dic As Dictionary(Of Int32, Boolean))
        For Each chk As CheckBox In flpMain.Controls
            If chk.Checked Then
                dic(chk.Tag) = True
            Else
                dic(chk.Tag) = False
            End If
        Next
    End Sub
End Class