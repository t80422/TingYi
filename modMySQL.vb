Imports System.Configuration
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports MySql.Data.MySqlClient
'MySQL相關
Module modMySQL
    Friend mConn As MySqlConnection
    Private title = "MySQL"

    Sub New()
        '設定連線
        Dim myConnectionString As String = ConfigurationManager.AppSettings("myConnectionString").ToString
        mConn = New MySqlConnection(myConnectionString)
        '測試連線
        Try
            mConn.Open()
        Catch ex As Exception
            '使用3306Port 如果開不起來就是mysql卡住 要到工作管理員結束工作後重開
            MsgBox("未開啟資料庫連線", MsgBoxStyle.Exclamation, "資料庫")
            End
        End Try
        mConn.Close()

    End Sub

    ''' <summary>
    ''' 查詢資料表
    ''' </summary>
    ''' <returns></returns>
    Friend Function SelectTable(sSQL As String) As DataTable
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

    ''' <summary>
    ''' 新增資料至資料表
    ''' </summary>
    ''' <param name="sTable"></param>
    ''' <param name="dicData">key:ColumnName</param>
    ''' <returns></returns>
    ''' 1.參數不能是control、object,遇到cmb需要回傳selectValue的會有問題
    Public Function InserTable(sTable As String, dicData As Dictionary(Of String, String)) As Boolean
        Dim result As Boolean
        Dim cmd As New MySqlCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Keys.Select(Function(key) $"@{key}"))})", mConn)
        Try
            mConn.Open()
            For Each kvp In dicData
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value))
            Next
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        mConn.Close()
        Return result
    End Function
    ''' <summary>
    ''' 新增資料至資料表
    ''' </summary>
    ''' <param name="sTable"></param>
    ''' <param name="dicData">key:欄位名稱 value:控制項</param>
    ''' <returns></returns>
    ''' 對應使用Linq語法抓取容器內的控制項會遇到DateTimePicker的值是value
    <Obsolete("遇到cmb需要回傳selectValue的會有問題")>
    Public Function InserTable(sTable As String, dicData As Dictionary(Of String, Control)) As Boolean
        Dim result As Boolean
        Dim cmd As New MySqlCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Keys.Select(Function(key) $"@{key}"))})", mConn)
        Try
            mConn.Open()
            For Each kvp In dicData
                Dim value = If(TypeOf kvp.Value Is DateTimePicker, DirectCast(kvp.Value, DateTimePicker).Value, kvp.Value.Text)
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(value))
            Next
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message, Title:=title)
        End Try
        mConn.Close()
        Return result
    End Function

    ''' <summary>
    ''' 更新表格
    ''' </summary>
    ''' <param name="table">表格名稱 (試試使用另一個多載)</param>
    ''' <param name="dicFields">更新對象集合</param>
    ''' <param name="condition">Where</param>
    Public Function UpdateTable(table As String, dicFields As Dictionary(Of String, String), condition As String) As Boolean
        Dim result As Boolean = False

        Try
            mConn.Open()
            Dim sql = $"UPDATE {table} SET "
            Dim lst As New List(Of String)

            For Each kvp In dicFields
                lst.Add($"{kvp.Key} = @{kvp.Key}")
            Next

            sql += String.Join(",", lst) + $" WHERE {condition}"
            Dim cmd As New MySqlCommand(sql, mConn)

            For Each kvp In dicFields
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(kvp.Value))
            Next

            If cmd.ExecuteNonQuery() > 0 Then
                result = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, Title:=MethodBase.GetCurrentMethod.Name)
        Finally
            mConn.Close()
        End Try

        Return result
    End Function
    Public Function UpdateTable(table As String, dicFields As Dictionary(Of String, Control), condition As String) As Boolean
        Dim result As Boolean = False

        Try
            mConn.Open()
            Dim sql = $"UPDATE {table} SET "
            Dim lst As New List(Of String)

            For Each kvp In dicFields
                lst.Add($"{kvp.Key} = @{kvp.Key}")
            Next

            sql += String.Join(",", lst) + $" WHERE {condition}"
            Dim cmd As New MySqlCommand(sql, mConn)

            For Each kvp In dicFields
                Dim value = If(TypeOf kvp.Value Is DateTimePicker, DirectCast(kvp.Value, DateTimePicker).Value, kvp.Value.Text)
                cmd.Parameters.AddWithValue($"@{kvp.Key}", Trim(value))
            Next

            If cmd.ExecuteNonQuery() > 0 Then
                result = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, Title:=MethodBase.GetCurrentMethod.Name)
        Finally
            mConn.Close()
        End Try

        Return result
    End Function
    ''' <summary>
    ''' MySQL Delete
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="sWhere">條件</param>
    ''' <returns></returns>
    Public Function DeleteData(sTable As String, sWhere As String) As Boolean
        Dim rowsAffected As Integer
        Dim cmd As New MySqlCommand($"DELETE FROM {sTable} WHERE {sWhere}", mConn)
        Try
            mConn.Open()
            rowsAffected = cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message, Title:="警告")
        End Try
        mConn.Close()
        Return rowsAffected > 0
    End Function

    ''' <summary>
    ''' 取得SQL語句中的表格名稱
    ''' </summary>
    ''' <param name="query"></param>
    ''' <returns></returns>
    Public Function GetTableNamesFromQuery(query As String) As List(Of String)
        Dim tableNames As New List(Of String)

        ' 使用正則表達式搜尋 FROM 和 JOIN 子句中的表名
        Dim regex As New Regex("(?:FROM|JOIN)\s+(\w+)", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = regex.Matches(query)

        ' 迭代匹配的結果，並將表名加入列表
        For Each match As Match In matches
            Dim tableName As String = match.Groups(1).Value
            tableNames.Add(tableName)
        Next

        Return tableNames
    End Function
End Module
