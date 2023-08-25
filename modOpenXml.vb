Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module modOpenXml
    ''' <summary>
    ''' 取得工作表的ID
    ''' </summary>
    ''' <param name="exl"></param>
    ''' <param name="sheetName"></param>
    ''' <returns></returns>
    Public Function GetSheetId(exl As SpreadsheetDocument, sheetName As String) As String
        Return exl.WorkbookPart.Workbook.Sheets.Elements(Of Sheet).Where(Function(s) s.Name = sheetName).Select(Function(s) s.Id.Value).First
    End Function

    ''' <summary>
    ''' 取得指定儲存格的值
    ''' </summary>
    ''' <param name="cellAddress"></param>
    ''' <param name="wbPart"></param>
    ''' <param name="sd"></param>
    ''' <returns></returns>
    Public Function GetCellValue(cellAddress As String, wbPart As WorkbookPart, sd As SheetData) As String
        Dim cellValue As String = String.Empty
        Dim cell As Cell = sd.Descendants(Of Cell)().FirstOrDefault(Function(c) c.CellReference.Value = cellAddress)
        If cell IsNot Nothing Then
            If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                Dim sstPart = wbPart.SharedStringTablePart
                cellValue = sstPart.SharedStringTable.ElementAt(cell.CellValue.InnerText).InnerText
            Else
                cellValue = cell.CellValue.InnerText
            End If
        End If
        Return cellValue
    End Function

    ''' <summary>
    ''' 修改儲存格內容
    ''' </summary>
    ''' <param name="ws"></param>
    ''' <param name="cellAddress"></param>
    ''' <param name="newText"></param>
    Public Sub SetCellValue(ws As Worksheet, cellAddress As String, newText As String, sstPart As SharedStringTablePart)
        Dim sd = ws.GetFirstChild(Of SheetData)()
        Dim cell As Cell = GetOrCreateCell(cellAddress, sd)
        Dim sharedStringIndex As Integer = InsertSharedStringItem(newText, sstPart)
        cell.CellValue = New CellValue(sharedStringIndex.ToString())
        cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
        ws.Save()
    End Sub

    Private Function InsertSharedStringItem(text As String, sharedStringTablePart As SharedStringTablePart) As Integer
        If sharedStringTablePart.SharedStringTable Is Nothing Then
            sharedStringTablePart.SharedStringTable = New SharedStringTable()
        End If
        Dim sharedStringIndex As Integer = 0
        For Each item As SharedStringItem In sharedStringTablePart.SharedStringTable.Elements(Of SharedStringItem)()
            If item.InnerText = text Then
                Return sharedStringIndex
            End If
            sharedStringIndex += 1
        Next
        sharedStringTablePart.SharedStringTable.AppendChild(New SharedStringItem(New Text(text)))
        sharedStringTablePart.SharedStringTable.Save()
        Return sharedStringIndex
    End Function

    Private Function GetOrCreateCell(cellAddress As String, sd As SheetData) As Cell
        Dim cell As Cell = sd.Descendants(Of Cell)().FirstOrDefault(Function(c) c.CellReference.Value = cellAddress)
        If cell Is Nothing Then
            Dim column As String = GetColumnName(cellAddress)
            Dim rowNumber As UInteger = GetRowNumber(cellAddress)

            Dim row As Row = sd.Elements(Of Row)().FirstOrDefault(Function(r) r.RowIndex.Value = rowNumber)
            If row Is Nothing Then
                row = New Row() With {.RowIndex = rowNumber}
                sd.Append(row)
            End If

            cell = New Cell() With {.CellReference = cellAddress}
            row.InsertAt(cell, GetColumnIndex(column))
        End If

        Return cell
    End Function

    Private Function GetColumnName(cellAddress As String) As String
        Return Regex.Replace(cellAddress, "[\d]", String.Empty)
    End Function

    Private Function GetRowNumber(cellAddress As String) As UInteger
        Dim rowNumber As String = Regex.Match(cellAddress, "\d+").Value
        Return UInteger.Parse(rowNumber)
    End Function

    Private Function GetColumnIndex(columnName As String) As UInteger
        Dim baseValue As Integer = Asc("A") - 1
        Dim columnIndex As Integer = 0

        For i As Integer = columnName.Length - 1 To 0 Step -1
            Dim characterValue As Integer = Asc(columnName(i)) - baseValue
            columnIndex += characterValue * System.Math.Pow(26, columnName.Length - 1 - i)
        Next

        Return columnIndex
    End Function

End Module
