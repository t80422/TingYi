Module modDBServiece
    Public Sub UpdateCustomer(dic As Dictionary(Of String, Object), cusID As Integer)
        UpdateTable("customer", dic, $"cus_id = {cusID}")
    End Sub
End Module