Imports System.Configuration
Imports MySql.Data.MySqlClient

Module modMySQL
    Friend mConn As MySqlConnection

    Friend Sub InitMySQL()
        '設定連線
        Dim myConnectionString As String = ConfigurationManager.AppSettings("myConnectionString").ToString
        mConn = New MySqlConnection(myConnectionString)
    End Sub
    ''' <summary>
    ''' 查詢資料表
    ''' </summary>
    ''' <returns></returns>
    Friend Function SelectFromTable(sSQL As String) As DataTable
        Dim dt As New DataTable()
        Try
            mConn.Open()
            Using cmd As New MySqlCommand(sSQL, mConn)
                Dim adapter As New MySqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        mConn.Close()
        Return dt
    End Function
End Module
