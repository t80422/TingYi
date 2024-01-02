Module modDistribute
    Public Sub InsertDistribute(day As Date, orderID As Integer, meal As String, dicData As Dictionary(Of String, Object), ByRef count As Integer)
        Dim sDay = day.ToString("yyyy/MM/dd")

        '檢查當日有沒有資料
        Dim dic As New Dictionary(Of String, Object) From {
            {"dist_date", sDay},
            {"dist_ord_id", orderID},
            {"dist_meal", meal}
        }

        '沒資料就新增
        If SelectTable("SELECT * FROM distribute WHERE dist_date = @dist_date AND dist_ord_id = @dist_ord_id AND dist_meal = @dist_meal", dic).Rows.Count = 0 Then
            '取得資料
            dicData.ToList.ForEach(Sub(kvp) dic.Add(kvp.Key, kvp.Value))
            InserTable("distribute", dic)
            count -= 1
        End If
    End Sub

    Public Sub UpdateDistribute(day As Date, orderID As Integer, meal As String, dicData As Dictionary(Of String, Object))
        Dim sDay = day.ToString("yyyy/MM/dd")

        '檢查當日有沒有資料
        Dim dic As New Dictionary(Of String, Object) From {
            {"dist_date", sDay},
            {"dist_ord_id", orderID},
            {"dist_meal", meal}
        }
        Dim dt = SelectTable("SELECT * FROM distribute WHERE dist_date = @dist_date AND dist_ord_id = @dist_ord_id AND dist_meal = @dist_meal", dic)
        If dt.Rows.Count = 0 Then Exit Sub

        Dim id = dt.Rows(0).Field(Of Integer)("dist_id")

        '有資料就更新
        If id > 0 Then
            '取得資料
            dicData.ToList.ForEach(Sub(kvp) dic.Add(kvp.Key, kvp.Value))
            UpdateTable("distribute", dic, $"dist_id = {id}")
        End If
    End Sub

    Public Sub DeleteDistribute(day As Date, orderID As Integer, meal As String)
        Dim sDay = day.ToString("yyyy/MM/dd")

        '取得id
        Dim dic As New Dictionary(Of String, Object) From {
            {"dist_date", sDay},
            {"dist_ord_id", orderID},
            {"dist_meal", meal}
        }
        Dim dt = SelectTable("SELECT * FROM distribute WHERE dist_date = @dist_date AND dist_ord_id = @dist_ord_id AND dist_meal = @dist_meal", dic)

        If dt.Rows.Count = 0 Then Exit Sub

        DeleteData("distribute", $"dist_id = {dt.Rows(0).Field(Of Integer)("dist_id")}")
    End Sub
End Module
