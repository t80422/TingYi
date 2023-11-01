Public Class frmTaboo
    Private _returnString As String
    Private dic As New Dictionary(Of Integer, Boolean)
    Private dt As DataTable

    Public ReadOnly Property ReturnString() As String
        Get
            Return _returnString
        End Get
    End Property

    Private Sub frmTaboo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dt = SelectTable("SELECT * FROM taboo")

        '初始化所有選項
        For Each row In dt.AsEnumerable.Select(Function(x) x.Field(Of Integer)("tabo_id"))
            dic.Add(row, False)
        Next

        '初始化分類
        Dim rowGroups = SelectTable("SELECT * FROM taboo_group").Rows
        With cmbType
            .Items.Add("全部")
            For Each row As DataRow In rowGroups
                .Items.Add(row("tg_name"))
            Next
            .SelectedIndex = 0
        End With
    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        SaveCheck(dic)
        flpMain.Controls.Clear()
        Dim drs As DataRow()
        If sender.text = "全部" Then
            drs = dt.Select()
        Else
            drs = dt.Select($"tabo_type = '{cmbType.Text}'")
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
        Dim list As List(Of Integer) = dic.Where(Function(item) item.Value = True).Select(Function(item) item.Key).ToList
        Dim listText As New List(Of String)
        For Each txt As String In list
            Dim name = dt.Select($"tabo_id = '{txt}'").FirstOrDefault.Field(Of String)("tabo_name")
            listText.Add(name)
        Next
        _returnString = String.Join(",", listText)
        DialogResult = DialogResult.OK
    End Sub

    '將已勾選的選項紀錄至dic
    Private Sub SaveCheck(dic As Dictionary(Of Integer, Boolean))
        For Each chk As CheckBox In flpMain.Controls
            If chk.Checked Then
                dic(chk.Tag) = True
            Else
                dic(chk.Tag) = False
            End If
        Next
    End Sub
End Class