Imports System.Text
Imports Microsoft.Office.Interop
Imports MySql.Data.MySqlClient

Public Class frmMain
    Friend dtTaboo As DataTable
    Private tempDistDay As Button '配餐管理月曆所選日期暫存
    'todo 1.登入頁的 大底圖，可以 800*600 px，或是左邊這區塊 320*360 px檔案格式為JPG
    '2.舊會員資料使用EXCEL轉入
    '9.菜單管理可以產生，美工要貼lineat的每週餐點內容
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedTab.Name = "TP_Logout" Then
            If MessageBox.Show("確定要登出嗎??", "登出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.OK Then
                Hide()
                LoginForm1.Show()
            Else
                TabControl1.SelectedTab = tpCustomer
            End If
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '自定義索引標籤、文字顏色
        TabControl1.DrawMode = DrawMode.OwnerDrawFixed
        tcCustomer.DrawMode = DrawMode.OwnerDrawFixed

        InitMySQL()
        InitDataGrid()
        InitProductGroup()
        InitTabooType()
        InitPosition()
        InitSales()

        '初始化收款方式
        cmbMonType.Items.Add("全款")
        cmbMonType.Items.Add("訂金")
        '初始化禁忌清單
        dtTaboo = SelectFromTable("SELECT * FROM taboo")
        '初始化配餐管理 月曆
        txtDistCalendar.Text = DateTime.Now.ToString("Y")
        '初始化菜單版本
        With cmbProdVers_menu
            Dim arr() As String = {"A", "B", "C", "D"}
            .DataSource = arr
            .SelectedIndex = -1
        End With
        'todo 未完成區-----
        TP_Report.Parent = Nothing
    End Sub

    Private Sub frmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        LoginForm1.Close()
    End Sub

    '自定義索引標籤、文字顏色
    Private Sub TabControl_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl1.DrawItem, tcCustomer.DrawItem
        Dim tabControl As TabControl = DirectCast(sender, TabControl)
        Dim tab As TabPage = tabControl.TabPages(e.Index)

        ' 檢查當前索引標籤是否為選中狀態
        Dim isSelected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected

        ' 繪製索引標籤的背景
        Dim backColor As System.Drawing.Color = If(isSelected, System.Drawing.Color.CornflowerBlue, System.Drawing.Color.WhiteSmoke)
        e.Graphics.FillRectangle(New SolidBrush(backColor), e.Bounds)

        ' 繪製索引標籤的文字
        Dim text As String = tab.Text
        Dim textColor As System.Drawing.Color = If(isSelected, System.Drawing.Color.White, System.Drawing.Color.Black)
        Dim font As System.Drawing.Font = tabControl.Font
        e.Graphics.DrawString(text, font, New SolidBrush(textColor), e.Bounds.Location)
    End Sub

    ''' <summary>
    ''' 初始化商品ComboBox
    ''' </summary>
    Private Sub InitProductGroup()
        '初始化商品群組
        Dim col As New Collection From {
        cmbProdGrp_product,
        cmbProdGrp_order
        }
        For i As Short = 1 To col.Count
            With col(i)
                .DataSource = SelectFromTable("SELECT * FROM product_group")
                .DisplayMember = "prod_grp_name"
                .ValueMember = "prod_grp_id"
                .SelectedIndex = -1
            End With
        Next
        '初始化商品
        With cmbProdName_menu
            .DataSource = SelectFromTable("SELECT * FROM product WHERE prod_type = '套餐'")
            .DisplayMember = "prod_name"
            .ValueMember = "prod_id"
            .SelectedIndex = -1
        End With
        '初始化商品管理的商品分類
        Dim items() As String = {"套餐", "單點"}
        cmbProdType.Items.AddRange(items)
    End Sub

    ''' <summary>
    ''' 初始化禁忌分類
    ''' </summary>
    Private Sub InitTabooType()
        With cmbTaboClass
            .DataSource = SelectFromTable("SELECT DISTINCT tabo_type FROM taboo")
            .DisplayMember = "tabo_type"
            .ValueMember = "tabo_type"
            .SelectedIndex = -1
        End With
    End Sub

    ''' <summary>
    ''' 初始化職位ComboBox
    ''' </summary>
    Private Sub InitPosition()
        Dim col As New Collection From {
            cmbPosition_perm,
            cmbPosition_emp
        }
        For i As Short = 1 To col.Count
            With col(i)
                .DataSource = SelectFromTable("SELECT * FROM permissions")
                .DisplayMember = "perm_name"
                .ValueMember = "perm_id"
                .SelectedIndex = -1
            End With
        Next
    End Sub

    ''' <summary>
    ''' 初始化業務人員
    ''' </summary>
    Private Sub InitSales()
        With cmbSales
            .DataSource = SelectFromTable("SELECT a.emp_name,a.emp_id FROM employee a LEFT JOIN permissions b ON a.emp_perm_id=b.perm_id WHERE perm_name = '業務'")
            .DisplayMember = "emp_name"
            .ValueMember = "emp_id"
            .SelectedIndex = -1
        End With
    End Sub

    ''' <summary>
    '''初始化DataGrid欄位
    ''' </summary>
    Private Sub InitDataGrid()
        '客戶管理
        DataToDgv(SelectFromTable(sqlCustomer), "customer", dgvCustomer)
        '商品群組管理
        DataToDgv(SelectFromTable(sqlProductGroup), "product_group", dgvProdgroup)
        '商品管理        
        DataToDgv(SelectFromTable(sqlProduct), "product,product_group", dgvProduct)
        '禁忌管理
        DataToDgv(SelectFromTable(sqlTaboo), "taboo", dgvTaboo)
        '訂單管理
        DataToDgv(SelectFromTable(sqlOrder), "customer,orders,product", dgvOrder)
        '財務管理       
        DataToDgv(SelectFromTable(sqlMoney), "customer,orders,money", dgvMoney)
        '權限管理
        DataToDgv(SelectFromTable(sqlPermision), "permissions", dgvPermissions)
        '員工管理        
        DataToDgv(SelectFromTable(sqlEmployee), "permissions,employee", dgvEmployee)
        '配餐管理        
        DataToDgv(SelectFromTable(sqlDistribute), "distribute,orders,customer,product", dgvDistribute)
        '菜單管理
        DataToDgv(SelectFromTable(sqlMenu), "menu,product", dgvMenu)
    End Sub

    '''' <summary>
    '''' xml取得儲存格資訊
    '''' </summary>
    '''' <returns></returns>
    'Private Function GetCell() As Dictionary(Of String, String)
    '    Dim dic As New Dictionary(Of String, String)
    '    Dim exl = SpreadsheetDocument.Open("D:\WorkWork\挺益\20220126_菜單新格式\最新A版叫貨.xlsx", False)
    '    Dim workbookPart As WorkbookPart = exl.WorkbookPart
    '    Dim targetSheet = workbookPart.Workbook.Sheets.Elements(Of Sheet).FirstOrDefault(Function(x) x.Name = "菜單總表")
    '    Dim worksheetPart As WorksheetPart = workbookPart.GetPartById(targetSheet.Id)
    '    Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()
    '    Dim a As CellWhere
    '    a.Column = "D"
    '    a.RowStart = 3
    '    a.RowEnd = 36
    '    Dim cellName As New List(Of CellWhere) From {
    '        a
    '    }

    '    For Each lst In cellName
    '        For i As Integer = lst.RowStart To lst.RowEnd
    '            If sheetData IsNot Nothing Then
    '                Dim cellRef As String = lst.Column + i.ToString
    '                Dim cell As Cell = sheetData.Descendants(Of Cell)().FirstOrDefault(Function(x) x.CellReference.Value = cellRef)

    '                If cell IsNot Nothing Then
    '                    ' 取得 A1 儲存格的值
    '                    Dim sharedStringTablePart As SharedStringTablePart = workbookPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
    '                    Dim cellValue As String = GetCellValue(cell, sharedStringTablePart)
    '                    dic.Add(cellRef, cellValue)
    '                Else
    '                    Console.WriteLine($"找不到{cellName}儲存格")
    '                End If
    '            Else
    '                Console.WriteLine("找不到工作表資料")
    '                Exit For
    '            End If

    '        Next
    '    Next
    '    Return dic
    'End Function

    '''' <summary>
    '''' xml取出儲存格文字
    '''' </summary>
    '''' <param name="cell"></param>
    '''' <param name="sharedStringTablePart"></param>
    '''' <returns></returns>
    'Private Function GetCellValue(cell As Cell, sharedStringTablePart As SharedStringTablePart) As String
    '    Dim cellValue As String = cell.InnerText

    '    If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
    '        Dim sharedStringIndex As Integer = Integer.Parse(cellValue)
    '        Dim sharedStringItem As SharedStringItem = sharedStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(sharedStringIndex)
    '        cellValue = sharedStringItem.Text.Text
    '    End If

    '    Return cellValue
    'End Function

    '客戶管理-dgv點擊
    Private Sub dgvCustomer_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvCustomer.CellMouseClick
        'todo 參考訂單管理來簡化
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            ClearTabPage(tpConsult_cus)
            ClearTabPage(tpBasic_cus)
            Dim row = dgv.SelectedRows(0)
            Dim list As List(Of String)
            Dim rdo As RadioButton
            Dim id = row.Cells("cus_id").Value.ToString
            '用data table把所有資料放上去
            Dim rowCus = SelectFromTable($"SELECT * FROM customer WHERE cus_id = '{id}'").Rows(0)
            txtCusID.Text = rowCus("cus_id").ToString
            txtCusName_cus.Text = rowCus("cus_name").ToString
            txtBirthday.Text = If(rowCus.IsNull("cus_birthday"), "", Convert.ToDateTime(rowCus("cus_birthday")))
            txtTelephone.Text = rowCus("cus_telephone").ToString
            txtPhone_cus.Text = rowCus("cus_phone").ToString
            txtJob.Text = rowCus("cus_job").ToString
            txtLineID.Text = rowCus("cus_line").ToString
            txtAddress.Text = rowCus("cus_address").ToString

            If Not rowCus.IsNull("cus_gender") Then
                If rowCus("cus_gender") = "男" Then
                    rdoMan.Checked = True
                Else
                    rdoFemale.Checked = True
                End If
            End If

            If Not rowCus.IsNull("cus_marriage") Then
                If rowCus("cus_marriage") = "未婚" Then
                    rdounmarried.Checked = True
                Else
                    rdoMarried.Checked = True
                    txtSpouse.Text = rowCus("cus_spouse").ToString
                    txtChildren.Text = rowCus("cus_children").ToString
                End If
            End If

            For Each rdo In grpAcad_Qual.Controls.OfType(Of RadioButton)
                rdo.Checked = Equals(rdo.Text, rowCus("cus_acad_qual"))
            Next

            If Not rowCus.IsNull("cus_kind") Then
                list = Split(rowCus("cus_kind"), ",").ToList
                For Each check In grpKind.Controls.OfType(Of CheckBox)
                    If list.Contains(check.Text) Then
                        check.Checked = True
                        If check.Text = "術後餐" Then
                            txtKindElse.Text = rowCus("cus_kind_else").ToString
                        End If
                    End If
                Next
                list.Clear()
            End If

            If Not rowCus.IsNull("cus_get_msg") Then
                list = Split(rowCus("cus_get_msg"), ",").ToList
                For Each check In grpGetMsg.Controls.OfType(Of CheckBox)
                    If list.Contains(check.Text) Then
                        check.Checked = True
                        If check.Text = "其他" Then
                            txtGetMsgElse.Text = rowCus("cus_getmsg_else")
                        End If
                    End If
                Next
                list.Clear()
            End If

            'txtDueDate.Text = rowCus("cus_due_date")
            txtDueDate.Text = If(rowCus.IsNull("cus_due_date"), "", Convert.ToDateTime(rowCus("cus_due_date")))
            txtHospital.Text = rowCus("cus_hospital")
            txtManyChild.Text = rowCus("cus_many_child").ToString
            txtConfLoca.Text = rowCus("cus_conf_loca")
            txtConfDay.Text = rowCus("cus_conf_day").ToString
            txtConfBuy.Text = rowCus("cus_conf_buy").ToString
            txtHeight.Text = rowCus("cus_height").ToString
            txtBornWeight.Text = rowCus("cus_born_weight").ToString
            txtWeight.Text = rowCus("cus_weight").ToString

            If Not rowCus.IsNull("cus_disease") Then
                list = Split(rowCus("cus_disease"), ",").ToList
                For Each check In grpDisease.Controls.OfType(Of CheckBox)
                    If list.Contains(check.Text) Then
                        check.Checked = True
                        If check.Text = "其他" Then
                            txtDisease.Text = rowCus("cus_disease_else")
                        End If
                    End If
                Next
                list.Clear()
            End If

            If Not rowCus.IsNull("cus_tabo_id") Then
                list = Split(rowCus("cus_tabo_id"), ",").ToList
                txtTaboo.Tag = String.Join(",", list)
                Dim listText As New List(Of String)
                For Each txt As String In list
                    Dim name = dtTaboo.Select($"tabo_id = '{txt}'").FirstOrDefault.Field(Of String)("tabo_name")
                    listText.Add(name)
                Next
                txtTaboo.Text = String.Join(",", listText)
                list.Clear()
            End If

            txtMealAdj.Text = rowCus("cus_meal_adj").ToString
            txtDietPerp.Text = rowCus("cus_diet_prep").ToString
            txtNutrCons.Text = rowCus("cus_nutr_cons").ToString
            txtMemo_cus.Text = rowCus("cus_memo").ToString

            '顯示歷史訂單
            Dim sql = $"SELECT a.ord_id,a.ord_date,b.cus_name,b.cus_phone FROM orders a LEFT JOIN customer b ON a.ord_cus_id = b.cus_id LEFT JOIN product c ON a.ord_prod_id = c.prod_id LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id WHERE b.cus_id = '{txtCusID.Text}' ORDER BY a.ord_date DESC"
            DataToDgv(SelectFromTable(sql), "customer,orders", dgvOrder_cus)
        End If
    End Sub

    '客戶管理-新增
    Private Sub btnCusInsert_Click(sender As Object, e As EventArgs) Handles btnCusInsert.Click
        Cursor = Cursors.WaitCursor
        Dim table = "customer"
        If Not CheckCustomerData() Then GoTo Finish

        '檢查重複資料
        Dim dic As New Dictionary(Of String, String) From {
            {"cus_name", txtCusName_cus.Text},
            {"cus_phone", txtPhone_cus.Text}
        }
        Dim enu = dic.Select(Function(x) $"{x.Key} = '{x.Value}'")
        Dim dt As DataTable = SelectFromTable($"SELECT cus_id, cus_name, cus_gender, cus_phone FROM {table} WHERE {String.Join(" AND ", enu)}")
        If dt.Rows.Count > 0 Then
            MsgBox("重複資料")
            '列出重複的資料
            DataToDgv(dt, table, dgvCustomer)
            GoTo Finish
        End If
        InserData(table, Bind_TableTextBox(table))

        btnCusCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '客戶管理-修改
    Private Sub btnCusModify_Click(sender As Object, e As EventArgs) Handles btnCusModify.Click
        Cursor = Cursors.WaitCursor
        Dim table = "customer"
        If Not CheckCustomerData() Then GoTo Finish
        UpdateData(table, Bind_TableTextBox(table), $"cus_id = '{txtCusID.Text}'")

        btnCusCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    ''' <summary>
    ''' 檢查Customer即將上傳的內容是否有誤
    ''' </summary>
    ''' <returns>True:正確 False:錯誤</returns>
    Private Function CheckCustomerData() As Boolean
        '去txt頭尾空白
        tpBasic_cus.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Text = Trim(txt.Text))
        tpConsult_cus.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '檢查必填欄位
        If String.IsNullOrWhiteSpace(txtCusName_cus.Text) Then
            MsgBox("姓名 不能空白")
            txtCusName_cus.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtPhone_cus.Text) Then
            MsgBox("手機 不能空白")
            txtPhone_cus.Focus()
            Return False
        End If

        '檢查日期格式
        If Not String.IsNullOrWhiteSpace(txtBirthday.Text) Then
            Dim day As DateTime
            If Not DateTime.TryParse(txtBirthday.Text, day) Then
                MsgBox("生日日期格式錯誤")
                txtBirthday.Focus()
                Return False
            End If
        End If
        If Not String.IsNullOrWhiteSpace(txtDueDate.Text) Then
            Dim day As DateTime
            If Not DateTime.TryParse(txtDueDate.Text, day) Then
                MsgBox("預產期 日期格式錯誤")
                tcCustomer.SelectedTab = tpConsult_cus
                txtDueDate.Focus()
                Return False
            End If
        End If

        '檢查數字
        Dim dic As New Dictionary(Of TextBox, String) From {
            {txtChildren, "子女"},
            {txtManyChild, "第幾胎"},
            {txtConfDay, "月子天數"},
            {txtConfBuy, "欲購買月子餐天數"},
            {txtHeight, "身高"},
            {txtBornWeight, "產前體重"},
            {txtWeight, "目前體重"}
        }
        For Each txt In dic.Keys
            If Not String.IsNullOrEmpty(txtPrice.Text) AndAlso Not IsNumeric(txtPrice.Text) Then
                Dim tp As TabPage = If(TypeOf txtPrice.Parent Is TabPage, txtPrice.Parent, txtPrice.Parent.Parent)
                tcCustomer.SelectedTab = tp
                MsgBox($"{dic(txtPrice)} 請輸入數字")
                txtPrice.Focus()
                Return False
            End If
        Next
        Return True
    End Function

    '客戶管理-刪除
    Private Sub btnCusDelete_Click(sender As Object, e As EventArgs) Handles btnCusDelete.Click
        '檢查是否選擇對象
        Dim id = txtCusID.Text
        If String.IsNullOrEmpty(id) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim table = "customer"
        DeleteData(table, $"cus_id = '{id}'")
        MsgBox("刪除成功")

        btnCusCancel.PerformClick()
    End Sub

    '客戶管理-取消
    Private Sub btnCusCancel_Click(sender As Object, e As EventArgs) Handles btnCusCancel.Click
        DataToDgv(SelectFromTable(sqlCustomer), "customer", dgvCustomer)
        ClearTabPage(tpBasic_cus)
        ClearTabPage(tpConsult_cus)
    End Sub

    '客戶管理-查詢
    Private Sub btnCusQuery_Click(sender As Object, e As EventArgs) Handles btnCusQuery.Click
        Cursor = Cursors.WaitCursor
        Dim sql = sqlCustomer + $" WHERE cus_name LIKE '%{txtCusQuery.Text}%' or cus_phone LIKE '%{txtCusQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), "customer", dgvCustomer)
        ClearTabPage(tpBasic_cus)
        ClearTabPage(tpConsult_cus)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '訂單管理-dgv點擊
    Private Sub dgvOrder_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvOrder.CellMouseClick
        ClearTabPage(tpOrder)

        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count < 0 Then Exit Sub

        Dim row = dgv.SelectedRows(0)
        Dim colName As String
        Dim rowData = SelectFromTable($"SELECT * FROM orders a LEFT JOIN customer b ON a.ord_cus_id = b.cus_id LEFT JOIN product c ON a.ord_prod_id = c.prod_id LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id LEFT JOIN employee e ON a.ord_emp_id=e.emp_id WHERE ord_id = '{row.Cells("ord_id").Value}'").Rows(0)
        For Each ctrl As Windows.Forms.Control In dgv.Parent.Controls
            colName = ctrl.Tag 'TextBox的Tag對應表格的名稱
            If TypeOf ctrl Is TextBox Then
                If Not String.IsNullOrEmpty(colName) Then ctrl.Text = rowData(colName).ToString

            ElseIf TypeOf ctrl Is DateTimePicker Then
                If Not String.IsNullOrEmpty(colName) Then
                    Dim dtp = CType(ctrl, DateTimePicker)
                    dtp.Value = rowData(colName)
                End If

            ElseIf TypeOf ctrl Is ComboBox Then
                If Not String.IsNullOrEmpty(colName) And String.IsNullOrEmpty(rowData(colName).ToString) = False Then
                    Dim cmb = CType(ctrl, ComboBox)
                    cmb.SelectedIndex = cmb.FindStringExact(rowData(colName))
                End If
            End If
        Next

        If Not rowData.IsNull("ord_eat_type") Then
            If rowData("ord_eat_type") = "葷" Then
                rdoMeat.Checked = True
            Else
                rdoVegetarian.Checked = True
            End If
        End If

        If rowData("ord_breakfast").ToString > 0 Then chkBreak_order.Checked = True
        If rowData("ord_lunch").ToString > 0 Then chkLunch_order.Checked = True
        If rowData("ord_dinner").ToString > 0 Then chkDinner_order.Checked = True

        '計算未收帳款
        Dim dt = SelectFromTable($"SELECT mon_income FROM money WHERE mon_ord_id = '{txtOrdID_order.Text}'")
        If dt.Rows.Count = 0 Then
            txtUnpay.Text = txtTotalPrice.Text
            Exit Sub
        End If
        Dim list = dt.AsEnumerable.Select(Function(x) x.Field(Of Int32)("mon_income")).ToList
        Dim sum As Int32
        For Each money In list
            sum += money
        Next
        txtUnpay.Text = txtTotalPrice.Text - sum
    End Sub

    '訂單管理-新增
    Private Sub btnOrdInsert_Click(sender As Object, e As EventArgs) Handles btnOrdInsert.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        dtOrdDate.Value = Now
        If Not CheckOrderData() Then GoTo Finish
        If MsgBox("請確認金額是否正確", MsgBoxStyle.YesNo) = MsgBoxResult.No Then GoTo Finish
        Dim table = "orders"

        '更新customer特定欄位
        Dim dic As New Dictionary(Of String, String) From {
            {"cus_email", txtEmail.Text},
            {"cus_emer_cont", txtEmerCont.Text},
            {"cus_emer_phone", txtEmerPhone.Text}
        }
        Dim dt = SelectFromTable($"SELECT cus_id FROM customer WHERE cus_name = '{txtCusName_order.Text}' AND cus_phone = '{txtPhone_order.Text}'")
        Dim rowCusID As String
        If dt.Rows.Count > 0 Then
            rowCusID = dt.Rows(0)("cus_id").ToString
        Else
            MsgBox("找不到客戶資料")
            GoTo Finish
        End If

        UpdateData("customer", dic, $"cus_id = '{rowCusID}'")

        InserData(table, Bind_TableTextBox(table))

        btnOrdCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '訂單管理-修改
    Private Sub btnOrdModify_Click(sender As Object, e As EventArgs) Handles btnOrdModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If Not CheckOrderData() Then GoTo Finish

        '更新customer特定欄位
        Dim dic As New Dictionary(Of String, String) From {
            {"cus_email", txtEmail.Text},
            {"cus_emer_cont", txtEmerCont.Text},
            {"cus_emer_phone", txtEmerPhone.Text}
        }
        Dim dt = SelectFromTable($"SELECT cus_id FROM customer WHERE cus_name = '{txtCusName_order.Text}' AND cus_phone = '{txtPhone_order.Text}'")
        Dim rowCusID As String
        If dt.Rows.Count > 0 Then
            rowCusID = dt.Rows(0)("cus_id").ToString
        Else
            MsgBox("找不到客戶資料")
            GoTo Finish
        End If
        UpdateData("customer", dic, $"cus_id = '{rowCusID}'")
        UpdateData(sTable, Bind_TableTextBox(sTable), $"ord_id  = '{txtOrdID_order.Text}'")

        btnOrdCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '訂單管理-刪除
    Private Sub btnOrdDelete_Click(sender As Object, e As EventArgs) Handles btnOrdDelete.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = txtOrdID_order
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"ord_id  = '{id.Text}'") Then
            MsgBox("刪除成功")
            btnOrdCancel.PerformClick()
        End If
    End Sub

    '訂單管理-取消
    Private Sub btnOrdCancel_Click(sender As Object, e As EventArgs) Handles btnOrdCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent

        '列出所有資料
        DataToDgv(SelectFromTable(sqlOrder), "customer,orders", dgvOrder)
        ClearTabPage(tp)
        InitSales()
    End Sub

    '訂單管理-查詢
    Private Sub btnOrderQuery_Click(sender As Object, e As EventArgs) Handles btnOrderQuery.Click
        Cursor = Cursors.WaitCursor
        Dim sql = sqlOrder + $" WHERE b.cus_name LIKE '%{txtOrdQuery.Text}%' OR b.cus_phone LIKE '%{txtOrdQuery.Text}%' ORDER BY a.ord_date DESC"
        DataToDgv(SelectFromTable(sql), "customer,orders", dgvOrder)
        ClearTabPage(tpOrder)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    ''' <summary>
    ''' 檢查Orders即將上傳的內容是否有誤
    ''' </summary>
    ''' <returns>True:正確 False:錯誤</returns>
    Private Function CheckOrderData() As Boolean
        '去txt頭尾空白
        tpOrder.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '檢查必填欄位
        If String.IsNullOrWhiteSpace(txtCusName_order.Text) Then
            MsgBox("姓名 不能空白")
            txtCusName_order.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtPhone_order.Text) Then
            MsgBox("手機 不能空白")
            txtPhone_order.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(cmbProdGrp_order.Text) Then
            MsgBox("請選擇 商品群組")
            cmbProdGrp_order.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(cmbProdName_order.Text) Then
            MsgBox("請選擇 商品名稱")
            cmbProdName_order.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtCount.Text) Then
            MsgBox("數量(天) 不能空白")
            txtCount.Focus()
            Return False
        End If

        '檢查數字
        Dim dic As New Dictionary(Of TextBox, String) From {
            {txtCount, "數量"},
            {txtTotalPrice, "金額"},
            {txtDiscount, "折扣金額"},
            {txtTaste, "試吃費"},
            {txtTableware, "押餐具費"}
        }
        For Each txt In dic.Keys
            If Not String.IsNullOrEmpty(txtPrice.Text) AndAlso Not IsNumeric(txtPrice.Text) Then
                Dim tp As TabPage = If(TypeOf txtPrice.Parent Is TabPage, txtPrice.Parent, txtPrice.Parent.Parent)
                MsgBox($"{dic(txtPrice)} 請輸入數字")
                txtPrice.Focus()
                Return False
            End If
        Next

        Return True
    End Function

    '訂單管理-商品群組選擇後過濾商品名稱
    Private Sub cmbProdGrp_order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProdGrp_order.SelectedIndexChanged
        If cmbProdGrp_order.SelectedIndex < 0 Then Exit Sub
        Dim prodGrpID = cmbProdGrp_order.SelectedValue.ToString
        With cmbProdName_order
            .DataSource = SelectFromTable($"SELECT * FROM product WHERE prod_prod_grp_id = '{prodGrpID}'")
            .DisplayMember = "prod_name"
            .ValueMember = "prod_id"
            .SelectedIndex = -1
        End With
    End Sub

    '訂單管理-金額有關的TextBox離開焦點後
    Private Sub txtMoney_Leave(sender As Object, e As EventArgs) Handles txtCount.Leave, txtPrice.Leave, txtDiscount.Leave, txtTaste.Leave, txtTableware.Leave
        MoneyCalculate()
    End Sub

    ''' <summary>
    ''' 金額計算
    ''' </summary>
    Private Sub MoneyCalculate()
        '算式:單價*數量(天)-折讓金額+試吃費+押餐具費
        If String.IsNullOrWhiteSpace(txtPrice.Text) Or String.IsNullOrWhiteSpace(txtCount.Text) Then Exit Sub
        Dim price As Int32 = If(String.IsNullOrWhiteSpace(txtPrice.Text), 0, txtPrice.Text)
        Dim count As Int32 = If(String.IsNullOrWhiteSpace(txtCount.Text), 0, txtCount.Text)
        Dim discount As Int32 = If(String.IsNullOrWhiteSpace(txtDiscount.Text), 0, txtDiscount.Text)
        Dim taste As Int32 = If(String.IsNullOrWhiteSpace(txtTaste.Text), 0, txtTaste.Text)
        Dim tableware As Int32 = If(String.IsNullOrWhiteSpace(txtTableware.Text), 0, txtTableware.Text)
        txtTotalPrice.Text = price * count - discount + taste + tableware
    End Sub

    '訂單管理-商品名稱
    Private Sub cmbProdName_order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProdName_order.SelectedIndexChanged
        If cmbProdName_order.SelectedIndex = -1 Then Exit Sub
        Dim selectedRow As DataRowView = DirectCast(cmbProdName_order.SelectedItem, DataRowView)
        Dim prodId As Integer = selectedRow("prod_id")
        Dim row = SelectFromTable($"SELECT prod_type, prod_price, prod_meal FROM product WHERE prod_id = {prodId}").Rows(0)
        '如果是套餐,顯示餐種供客製勾選
        If row("prod_type") = "套餐" Then
            grpMeal_order.Enabled = True
        Else
            grpMeal_order.Enabled = False
        End If
        '顯示商品價格
        txtPrice.Text = row("prod_price")

        Dim chks As IEnumerable(Of CheckBox) = grpMeal_order.Controls.OfType(Of CheckBox)()
        Dim chk As CheckBox
        '初始化checkbox
        For Each chk In chks
            chk.Checked = False
        Next
        For Each name As String In row("prod_meal").Split(",")
            chk = chks.FirstOrDefault(Function(x) x.Text = name)
            If chk IsNot Nothing Then
                chk.Checked = True
            End If
        Next
    End Sub

    '訂單管理-午餐地址-同上
    Private Sub chkLunchAddr_CheckedChanged(sender As Object, e As EventArgs) Handles chkLunchAddr.Click
        If chkLunchAddr.Checked Then
            txtAddrLunch.Text = txtAddrBreak.Text
        Else
            txtAddrLunch.Text = ""
        End If
    End Sub

    '訂單管理-晚餐地址-同上
    Private Sub chkDinnerAddr_CheckedChanged(sender As Object, e As EventArgs) Handles chkDinnerAddr.Click
        If chkDinnerAddr.Checked Then
            txtAddrDinner.Text = txtAddrLunch.Text
        Else
            txtAddrDinner.Text = ""
        End If
    End Sub

    '員工管理-dgv點擊
    Private Sub dgvEmployee_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvEmployee.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            Dim row = dgv.SelectedRows(0)
            Dim colName As String
            For Each ctrl As Windows.Forms.Control In dgv.Parent.Controls
                'TextBox的Tag對應表格的備註
                If TypeOf ctrl Is TextBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        ctrl.Text = row.Cells(colName).Value.ToString()
                    End If

                ElseIf TypeOf ctrl Is ComboBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        Dim cmb = CType(ctrl, ComboBox)
                        cmb.SelectedIndex = cmb.FindStringExact(row.Cells(colName).Value.ToString)
                    End If
                End If
            Next
        End If
    End Sub

    '員工管理-新增
    Private Sub btnEmpInsert_Click(sender As Object, e As EventArgs) Handles btnEmpInsert.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡
        If Not CheckInsert(sTable, tp) Then GoTo Finish
        If Not InputBox("確認密碼").Equals(txtPsw.Text) Then
            MsgBox("輸入的密碼與先前不同,請再確認")
            GoTo Finish
        End If
        InserData(sTable, Bind_TableTextBox(sTable))
        btnEmpCancel.PerformClick()
        InitSales()
Finish:
        Cursor = Cursors.Default
    End Sub

    '員工管理-修改
    Private Sub btnEmpModify_Click(sender As Object, e As EventArgs) Handles btnEmpModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"emp_id  = '{txtEmpID.Text}'")
        btnEmpCancel.PerformClick()
        InitSales()
Finish:
        Cursor = Cursors.Default
    End Sub

    '員工管理-刪除
    Private Sub btnEmpDelete_Click(sender As Object, e As EventArgs) Handles btnEmpDelete.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "員工編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"emp_id  = '{id.Text}'") Then
            MsgBox("刪除成功")

            btnEmpCancel.PerformClick()
            InitSales()
        End If
    End Sub

    '員工管理-取消
    Private Sub btnEmpCancel_Click(sender As Object, e As EventArgs) Handles btnEmpCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '列出所有表格資料        
        DataToDgv(SelectFromTable(sqlEmployee), "permissions,employee", dgvEmployee)
        ClearTabPage(tp)
    End Sub

    '員工管理-查詢
    Private Sub btnEmpQuery_Click(sender As Object, e As EventArgs) Handles btnEmpQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = tp.Tag.ToString

        Dim Sql = sqlEmployee + $" WHERE emp_name LIKE '%{txtEmpQuery.Text}%' OR emp_phone LIKE '%{txtEmpQuery.Text}%'"
        DataToDgv(SelectFromTable(Sql), "permissions,employee", dgvEmployee)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '配餐管理-dgv點擊
    Private Sub dgvDistribute_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDistribute.CellMouseClick
        ClearTabPage(tpDistribute)

        Dim dgv As DataGridView = dgvDistribute
        If dgv.SelectedRows.Count < 0 Then Exit Sub
        '點dgv後將對象資料傳至各控制項
        Dim dgvRow = dgv.SelectedRows(0)
        Dim colName As String
        Dim rowData = SelectFromTable($"SELECT a.ord_id,b.cus_name,b.cus_phone,c.prod_name,a.ord_delivery,a.ord_breakfast,a.ord_lunch,a.ord_dinner FROM orders a LEFT JOIN customer b ON a.ord_cus_id=b.cus_id LEFT JOIN product c ON a.ord_prod_id=c.prod_id WHERE a.ord_id = '{dgvRow.Cells("ord_id").Value}'").Rows(0)
        For Each ctrl As Windows.Forms.Control In dgv.Parent.Controls
            colName = ctrl.Tag 'TextBox的Tag對應表格欄位名稱
            If TypeOf ctrl Is TextBox Then
                If Not String.IsNullOrEmpty(colName) Then ctrl.Text = rowData(colName).ToString

            ElseIf TypeOf ctrl Is DateTimePicker Then
                If Not String.IsNullOrEmpty(colName) Then
                    Dim dtp = CType(ctrl, DateTimePicker)
                    dtp.Value = rowData(colName)
                End If
            End If
        Next
        '刷新月曆
        SetCalender()
        SetCalenderData()

        CountNotConfigured()
    End Sub

    ''' <summary>
    ''' 設定所選訂單的月曆資料
    ''' </summary>
    Private Sub SetCalenderData()
        If txtOrdID_dist.Text = "" Then Exit Sub
        Dim d As Date = Date.Parse(txtDistCalendar.Text)
        '以當前月曆月份搜尋訂單配餐
        Dim dt = SelectFromTable($"SELECT * FROM distribute WHERE YEAR(dist_date) = {d.Year} AND MONTH(dist_date) = {d.Month} AND dist_ord_id = {txtOrdID_dist.Text}")
        For Each row As DataRow In dt.Rows
            '配餐日期
            d = row.Field(Of Date)("dist_date")
            '取得配餐日的panel
            Dim pnl = tlpCalendar.Controls.OfType(Of Panel).FirstOrDefault(Function(x) CInt(x.Tag) = d.Day)
            '將對應的餐種打包進按鈕的tag並改變顏色
            Dim btn = pnl?.Controls.OfType(Of Button).FirstOrDefault(Function(x) x.Text = row.Field(Of String)("dist_meal"))
            If btn IsNot Nothing Then
                btn.Tag = row
                btn.BackColor = System.Drawing.Color.LightGreen
            End If
        Next
    End Sub

    ''' <summary>
    ''' 計算未配置餐
    ''' </summary>
    Private Sub CountNotConfigured()
        '使用時機 dgv點取 新刪修後
        '計算餐種的數量扣掉未配置餐
        Dim dtOrder = SelectFromTable($"SELECT ord_breakfast, ord_lunch, ord_dinner FROM orders WHERE ord_id = '{txtOrdID_dist.Text}'")
        Dim d As Date = Date.Parse(txtDistCalendar.Text)
        Dim dt = SelectFromTable($"SELECT * FROM distribute WHERE YEAR(dist_date) = {d.Year} AND MONTH(dist_date) = {d.Month} AND dist_ord_id = {txtOrdID_dist.Text}")
        txtBreak.Text = If(dtOrder.Rows(0)("ord_breakfast") > 0, dtOrder.Rows(0)("ord_breakfast") - dt.Select("dist_meal='早'").Count, 0)
        txtLunch.Text = If(dtOrder.Rows(0)("ord_lunch") > 0, dtOrder.Rows(0)("ord_lunch") - dt.Select("dist_meal='午'").Count, 0)
        txtDinner.Text = If(dtOrder.Rows(0)("ord_dinner") > 0, dtOrder.Rows(0)("ord_dinner") - dt.Select("dist_meal='晚'").Count, 0)
    End Sub

    '配餐管理-新增
    Private Sub distInsert_Click(sender As Object, e As EventArgs) Handles distInsert.Click
        If txtOrdID_dist.Text = "" Then Exit Sub

        Cursor = Cursors.WaitCursor
        '若沒勾 改變後續設定 就只新增一天
        Dim count As Integer

        '檢查未配置餐來insert
        Dim checkInsert As Boolean

        If tempDistDay.Text = "早" And txtBreak.Text <> 0 Then
            checkInsert = True
            count = txtBreak.Text
        ElseIf tempDistDay.Text = "午" And txtLunch.Text <> 0 Then
            checkInsert = True
            count = txtLunch.Text
        ElseIf tempDistDay.Text = "晚" And txtDinner.Text <> 0 Then
            checkInsert = True
            count = txtDinner.Text
        End If
        If chkContinue.Checked = False Then count = 1
        If checkInsert Then
            Dim day = tempDistDay.Parent.Controls.OfType(Of Label).FirstOrDefault.Text
            Dim d = Date.Parse(txtDistCalendar.Text + day + "日")
            d = d.AddDays(-1)
            Dim dic As New Dictionary(Of String, String)
            With dic
                .Add("dist_ord_id", txtOrdID_dist.Text)
                .Add("dist_meal", tempDistDay.Text) '早午晚餐
                For i As Integer = count To 1 Step -1
                    d = d.AddDays(1)
                    .Add("dist_date", d) '送餐日期
                    Dim table = "distribute"
                    Dim conditions As List(Of String) = dic.Select(Function(kvp) $"{kvp.Key} = '{kvp.Value}'").ToList()
                    If SelectFromTable($"SELECT * FROM {table} WHERE " + String.Join(" And ", conditions)).Rows.Count > 0 Then
                        MsgBox("重複資料")
                        GoTo Finish
                    End If
                    InserData(table, dic)
                    UpdateData(table, Bind_TableTextBox(table), String.Join(" and ", conditions))
                    dic.Remove("dist_date")
                Next
            End With
            SetCalender()
            SetCalenderData()
            CountNotConfigured()
            MsgBox("新增成功")
        End If
Finish:
        Cursor = Cursors.Default
    End Sub

    '配餐管理-修改
    Private Sub btnDistModify_Click(sender As Object, e As EventArgs) Handles btnDistModify.Click
        If txtOrdID_dist.Text = "" Then Exit Sub

        Cursor = Cursors.WaitCursor
        Dim table = "distribute"
        '抓出所選的天
        If tempDistDay.BackColor = System.Drawing.Color.LightGreen Then
            Dim day = tempDistDay.Parent.Controls.OfType(Of Label).FirstOrDefault.Text
            Dim d = Date.Parse(txtDistCalendar.Text + day + "日")
            Dim dic As New Dictionary(Of String, String)
            With dic
                .Add("dist_ord_id", txtOrdID_dist.Text)
                .Add("dist_meal", tempDistDay.Text) '早午晚餐
                '若沒勾 改變後續設定 就只新增一天
                Dim count As Integer = 1
                If chkContinue.Checked Then
                    count = SelectFromTable($"SELECT * FROM {table} WHERE dist_ord_id = '{txtOrdID_dist.Text}' AND dist_date >= '{d}'").Rows.Count
                End If
                d = d.AddDays(-1)
                For i As Integer = count To 1 Step -1
                    d = d.AddDays(1)
                    .Add("dist_date", d) '送餐日期

                    Dim conditions As List(Of String) = dic.Select(Function(kvp) $"{kvp.Key} = '{kvp.Value}'").ToList()
                    UpdateData(table, Bind_TableTextBox(table), String.Join(" and ", conditions))
                    dic.Remove("dist_date")
                Next
            End With
        Else
            MsgBox("請選擇已新增的日期")
        End If

        SetCalender()
        SetCalenderData()
        CountNotConfigured()
        MsgBox("修改成功")

Finish:
        Cursor = Cursors.Default
    End Sub

    '配餐管理-刪除
    Private Sub btnDistDel_Click(sender As Object, e As EventArgs) Handles btnDistDel.Click
        If txtOrdID_dist.Text = "" Then Exit Sub

        Cursor = Cursors.WaitCursor
        Dim table = "distribute"
        '抓出所選的天
        If tempDistDay.BackColor = System.Drawing.Color.LightGreen Then
            Dim day = tempDistDay.Parent.Controls.OfType(Of Label).FirstOrDefault.Text
            Dim d = Date.Parse(txtDistCalendar.Text + day + "日")
            Dim dic As New Dictionary(Of String, String)
            With dic
                .Add("dist_ord_id", txtOrdID_dist.Text)
                .Add("dist_meal", tempDistDay.Text) '早午晚餐
                '若沒勾 改變後續設定 就只新增一天
                Dim count As Integer = 1
                If chkContinue.Checked Then
                    count = SelectFromTable($"SELECT * FROM {table} WHERE dist_ord_id = '{txtOrdID_dist.Text}' AND dist_date >= '{d}'").Rows.Count
                End If
                d = d.AddDays(-1)
                For i As Integer = count To 1 Step -1
                    d = d.AddDays(1)
                    .Add("dist_date", d) '送餐日期

                    Dim conditions As List(Of String) = dic.Select(Function(kvp) $"{kvp.Key} = '{kvp.Value}'").ToList()
                    DeleteData(table, String.Join(" and ", conditions))
                    dic.Remove("dist_date")
                Next
            End With
        Else
            MsgBox("請選擇已新增的日期")
        End If

        SetCalender()
        SetCalenderData()
        CountNotConfigured()
        MsgBox("刪除成功")
Finish:
        Cursor = Cursors.Default
    End Sub

    '配餐管理-取消
    Private Sub btnDistCancel_Click(sender As Object, e As EventArgs) Handles btnDistCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '列出所有表格資料        
        DataToDgv(SelectFromTable(sqlDistribute), "distribute,orders,customer,product", dgvDistribute)
        ClearTabPage(tp)
        SetCalender()
    End Sub

    '配餐管理-搜尋
    Private Sub btnDistQuery_Click(sender As Object, e As EventArgs) Handles btnDistQuery.Click
        Dim sql = sqlDistribute + $" WHERE b.cus_name LIKE '%{txtDistQuery.Text}%' OR b.cus_phone LIKE '%{txtDistQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), "distribute,orders,customer,product", dgvDistribute)
    End Sub

    ''' <summary>
    ''' 製作月曆一日的物件
    ''' </summary>
    ''' <param name="i">日</param>
    ''' <returns></returns>
    Private Function DayMaker(i As Int16) As Panel
        '框架
        Dim panel As New Panel With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyle.FixedSingle,
            .Tag = i'存日期供搜尋用
        }
        '日期
        Dim font As New System.Drawing.Font("Arial", 12, FontStyle.Bold)
        Dim point As Point
        point = New Point(0, 0)
        Dim lbl As New Label With {.Text = i, .Parent = panel, .Font = font, .AutoSize = True, .Location = point}
        '早中晚的按鈕式CheckBox
        With panel
            .Controls.Add(Setbtn_Dist(New Point(27, 0), "早"))
            .Controls.Add(Setbtn_Dist(New Point(2, 25), "午"))
            .Controls.Add(Setbtn_Dist(New Point(27, 25), "晚"))
        End With

        Return panel
    End Function

    ''' <summary>
    ''' 設定配餐管理月曆裡的按鈕
    ''' </summary>
    ''' <param name="point">在Panel裡的位置</param>
    ''' <param name="txt">顯示的文字</param>
    ''' <returns></returns>
    Private Function Setbtn_Dist(point As Point, txt As String) As Button
        Dim btn As New Button With
        {
            .Text = txt,
            .AutoSize = False,
            .Location = point,
            .Font = New System.Drawing.Font("標楷體", 10, FontStyle.Bold),
            .Height = 25,
            .Width = 25,
            .TextAlign = ContentAlignment.MiddleCenter
        }
        AddHandler btn.Click, AddressOf DistBtnClick
        Return btn
    End Function

    '配餐管理-月曆日期內按鈕點擊事件
    Private Sub DistBtnClick(sender As Object, e As EventArgs)
        '刷新grp
        For Each grp In tpDistribute.Controls.OfType(Of GroupBox)()
            For Each chk In grp.Controls.OfType(Of CheckBox)()
                chk.Checked = False
            Next
            For Each rdo In grp.Controls.OfType(Of RadioButton)()
                rdo.Checked = False
            Next
        Next

        tempDistDay = sender
        '標示現在選取的日期
        Dim day = tempDistDay.Parent.Controls.OfType(Of Label).FirstOrDefault.Text
        Dim d = Date.Parse(txtDistCalendar.Text + day + "日")
        txtSelectDate.Text = d + " " + tempDistDay.Text

        '將btn.tag裡的物件資料送至各grp
        Dim rowData As DataRow = tempDistDay.Tag
        If rowData Is Nothing Then Exit Sub

        Dim dic As New Dictionary(Of String, String)
        With dic
            .Add("dist_soup", "湯盅")
            .Add("dist_oil", "麻油")
            .Add("dist_wine", "酒")
            .Add("dist_vege", "素")
            .Add("dist_other", "其他")
            .Add("dist_customized", "客製需求")
            .Add("dist_tableware", "餐具")
        End With

        For Each kvp In dic
            Dim grp = tpDistribute.Controls.OfType(Of GroupBox)().FirstOrDefault(Function(x) x.Text = kvp.Value)
            If grp IsNot Nothing Then
                If Not rowData.IsNull(kvp.Key) Then
                    Dim lst As List(Of String) = Split(rowData(kvp.Key), ",").ToList
                    For Each check In grp.Controls.OfType(Of CheckBox)
                        If lst.Contains(check.Text) Then
                            check.Checked = True
                        End If
                    Next
                    For Each rdo In grp.Controls.OfType(Of RadioButton)
                        If lst.Contains(rdo.Text) Then
                            rdo.Checked = True
                        End If
                    Next
                    lst.Clear()
                End If
            End If
        Next
    End Sub

    '對配餐管理的月曆加減月份
    Private Sub BtnMonth_Click(sender As Object, e As EventArgs) Handles btnAddMonth.Click, btnMinusMonth.Click
        If txtDistCalendar.Text = "" Then Exit Sub

        Dim dt As Date = Date.Parse(txtDistCalendar.Text)
        Dim btn As Button = CType(sender, Button)
        Dim newDt As Date

        Select Case btn.Name
            Case "btnAddMonth"
                newDt = dt.AddMonths(1)
            Case "btnMinusMonth"
                newDt = dt.AddMonths(-1)
        End Select

        txtDistCalendar.Text = newDt.ToString("yyyy年MM月")
        SetCalenderData()
    End Sub

    '配餐管理-月曆時間一改變就設置日期內的物件
    Private Sub txtDistCalendar_TextChanged(sender As Object, e As EventArgs) Handles txtDistCalendar.TextChanged
        SetCalender()
    End Sub

    ''' <summary>
    ''' 設置配餐管理月曆
    ''' </summary>
    Private Sub SetCalender()
        tlpCalendar.Visible = False
        tlpCalendar.Controls.Clear()
        Dim i, j As Short
        Dim d = Date.Parse(txtDistCalendar.Text)
        Dim days As Short = Date.DaysInMonth(d.Year, d.Month)
        Dim firstDateOfWeek As Short = New DateTime(d.Year, d.Month, 1).DayOfWeek

        For i = 1 To days
            tlpCalendar.Controls.Add(DayMaker(i), firstDateOfWeek, j)
            firstDateOfWeek += 1
            If firstDateOfWeek = 7 Then
                firstDateOfWeek = 0
                j += 1
            End If
        Next
        tlpCalendar.Visible = True
    End Sub

    '菜單管理-Excel匯入
    Private Sub btnMenuExcel_Click(sender As Object, e As EventArgs) Handles btnMenuExcel.Click
        Cursor = Cursors.WaitCursor
        Dim exl As New Excel.Application With {
            .DisplayAlerts = False
        }
        Dim sheet As Excel.Worksheet = exl.Workbooks.Open("D:\WorkWork\挺益\20220126_菜單新格式\最新A版叫貨.xlsx").Sheets("菜單總表")
        Dim rng As String
        Dim cell As Excel.Range
        Dim value As Object
        Dim txt As String
        Dim menu As Menu
        Dim tempCell As Excel.Range
        Dim backColor As Excel.XlRgbColor
        '版本
        rng = "A2"
        cell = sheet.Range(rng)
        value = cell.Value
        txt = value.ToString()
        Dim ver = txt.Substring(txt.Length - 2, 1)
        '蒐集完丟這裡
        Dim lstMenu As New List(Of Menu)
        '月子早餐
        Dim dicMoonSonBreak As New Dictionary(Of Integer, Integer) From {
            {4, Meal_Detail.主食},
            {5, Meal_Detail.主菜},
            {6, Meal_Detail.半葷素},
            {7, Meal_Detail.青菜西飲},
            {8, Meal_Detail.湯品}
        }
        '月子午餐
        Dim dicMoonSonLunch As New Dictionary(Of Integer, Integer) From {
            {9, Meal_Detail.湯盅清補},
            {10, Meal_Detail.湯盅1期},
            {11, Meal_Detail.湯盅3期},
            {12, Meal_Detail.主食},
            {13, Meal_Detail.主菜},
            {14, Meal_Detail.半葷素},
            {15, Meal_Detail.蔬菜1},
            {16, Meal_Detail.水果},
            {17, Meal_Detail.甜品}
        }
        '月子晚餐
        Dim dicMoonSonDinner As New Dictionary(Of Integer, Integer) From {
            {18, Meal_Detail.湯盅清補},
            {19, Meal_Detail.湯盅1期},
            {20, Meal_Detail.湯盅3期},
            {21, Meal_Detail.主食},
            {22, Meal_Detail.主菜},
            {23, Meal_Detail.半葷素},
            {24, Meal_Detail.蔬菜1},
            {25, Meal_Detail.水果}
        }
        '月子晚點
        Dim dicMoonSonNightSnack As New Dictionary(Of Integer, Integer) From {
            {26, Meal_Detail.湯盅清補},
            {27, Meal_Detail.湯盅1期},
            {28, Meal_Detail.湯盅3期}
        }
        '調理餐
        Dim dicConditioning As New Dictionary(Of Integer, Integer) From {
            {39, Meal_Detail.主食},
            {40, Meal_Detail.主菜},
            {41, Meal_Detail.半葷素},
            {42, Meal_Detail.蔬菜1},
            {43, Meal_Detail.蔬菜2},
            {44, Meal_Detail.湯品},
            {45, Meal_Detail.水果}
        }
        '幸福午餐
        Dim dicHappinessLunch As New Dictionary(Of Integer, Integer) From {
            {50, Meal_Detail.主食},
            {51, Meal_Detail.主菜},
            {52, Meal_Detail.半葷素},
            {53, Meal_Detail.蔬菜1},
            {54, Meal_Detail.湯品}
        }
        '幸福晚餐
        Dim dicHappinessDinner As New Dictionary(Of Integer, Integer) From {
            {56, Meal_Detail.主食},
            {57, Meal_Detail.主菜},
            {58, Meal_Detail.半葷素},
            {59, Meal_Detail.蔬菜1},
            {60, Meal_Detail.湯品}
        }
        '住院早餐
        Dim dicHospitalizedBreak As New Dictionary(Of Integer, Integer) From {
            {65, Meal_Detail.主食},
            {66, Meal_Detail.主菜},
            {67, Meal_Detail.半葷素},
            {68, Meal_Detail.蔬菜1},
            {69, Meal_Detail.湯品},
            {70, Meal_Detail.飲品}
        }
        '住院午餐
        Dim dicHospitalizedLunch As New Dictionary(Of Integer, Integer) From {
            {71, Meal_Detail.主食},
            {72, Meal_Detail.主菜},
            {73, Meal_Detail.半葷素},
            {74, Meal_Detail.蔬菜1},
            {75, Meal_Detail.湯品},
            {76, Meal_Detail.水果},
            {77, Meal_Detail.飲品},
            {78, Meal_Detail.甜湯}
        }
        '住院晚餐
        Dim dicHospitalizedDinner As New Dictionary(Of Integer, Integer) From {
            {79, Meal_Detail.主食},
            {80, Meal_Detail.主菜},
            {81, Meal_Detail.半葷素},
            {82, Meal_Detail.蔬菜1},
            {83, Meal_Detail.湯品},
            {84, Meal_Detail.飲品},
            {85, Meal_Detail.夜點}
        }
        '輕食早餐
        Dim dicLightMealBreak As New Dictionary(Of Integer, Integer) From {
            {89, Meal_Detail.主食},
            {90, Meal_Detail.主菜},
            {91, Meal_Detail.蔬菜1},
            {92, Meal_Detail.蔬菜2},
            {93, Meal_Detail.水果},
            {94, Meal_Detail.飲品}
        }
        '輕食午餐
        Dim dicLightMealLunch As New Dictionary(Of Integer, Integer) From {
            {96, Meal_Detail.主食},
            {97, Meal_Detail.主菜},
            {98, Meal_Detail.蔬菜1},
            {99, Meal_Detail.蔬菜2},
            {100, Meal_Detail.水果},
            {101, Meal_Detail.飲品}
        }
        '輕食晚餐
        Dim dicLightMealDinner As New Dictionary(Of Integer, Integer) From {
            {103, Meal_Detail.主食},
            {104, Meal_Detail.主菜},
            {105, Meal_Detail.蔬菜1},
            {106, Meal_Detail.蔬菜2},
            {107, Meal_Detail.水果},
            {108, Meal_Detail.飲品}
        }
        '術後調理早餐
        Dim dicOperationBreak As New Dictionary(Of Integer, Integer) From {
            {4, Meal_Detail.主食},
            {5, Meal_Detail.主菜},
            {6, Meal_Detail.半葷素},
            {7, Meal_Detail.青菜西飲},
            {8, Meal_Detail.湯品}
        }
        '術後調理午餐
        Dim dicOperationLunch As New Dictionary(Of Integer, Integer) From {
            {11, Meal_Detail.主食},
            {12, Meal_Detail.主菜},
            {13, Meal_Detail.半葷素},
            {14, Meal_Detail.蔬菜1},
            {15, Meal_Detail.水果},
            {9, Meal_Detail.湯盅清補}
        }
        '術後調理晚餐
        Dim dicOperationDinner As New Dictionary(Of Integer, Integer) From {
            {18, Meal_Detail.主食},
            {19, Meal_Detail.主菜},
            {20, Meal_Detail.半葷素},
            {21, Meal_Detail.蔬菜1},
            {22, Meal_Detail.水果},
            {16, Meal_Detail.湯盅清補}
        }
        '素食早餐
        Dim dicVegetarianBreak As New Dictionary(Of Integer, Integer) From {
            {27, Meal_Detail.主食},
            {28, Meal_Detail.主菜},
            {29, Meal_Detail.半葷素},
            {30, Meal_Detail.青菜西飲},
            {31, Meal_Detail.湯品}
        }
        '素食午餐
        Dim dicVegetarianLunch As New Dictionary(Of Integer, Integer) From {
            {32, Meal_Detail.湯盅清補},
            {33, Meal_Detail.湯盅2期},
            {34, Meal_Detail.主食},
            {35, Meal_Detail.主菜},
            {36, Meal_Detail.半葷素},
            {37, Meal_Detail.蔬菜1},
            {38, Meal_Detail.甜品}
        }
        '素食晚餐
        Dim dicVegetarianDinner As New Dictionary(Of Integer, Integer) From {
            {39, Meal_Detail.湯盅清補},
            {40, Meal_Detail.湯盅2期},
            {41, Meal_Detail.主食},
            {42, Meal_Detail.主菜},
            {43, Meal_Detail.半葷素},
            {44, Meal_Detail.蔬菜1},
            {46, Meal_Detail.夜點}
        }
        '素食一般午餐
        Dim dicVegetarianNormalLunch As New Dictionary(Of Integer, Integer) From {
            {53, Meal_Detail.主食},
            {54, Meal_Detail.主菜},
            {55, Meal_Detail.蔬菜1},
            {56, Meal_Detail.蔬菜2},
            {57, Meal_Detail.湯品}
        }
        '素食一般晚餐
        Dim dicVegetarianNormalDinner As New Dictionary(Of Integer, Integer) From {
            {58, Meal_Detail.主食},
            {59, Meal_Detail.主菜},
            {60, Meal_Detail.蔬菜1},
            {61, Meal_Detail.蔬菜2},
            {62, Meal_Detail.湯品}
        }
        For col As Integer = Asc("D") To Asc("J")
            '日期
            rng = Chr(col) + "3"
            cell = sheet.Range(rng)
            value = cell.Value
            txt = value.ToString()
            Dim d = Date.Parse(txt)

            For row As Integer = 4 To 8
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "經典月子餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicMoonSonBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicMoonSonBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 9 To 17
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "經典月子餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicMoonSonLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                '判斷顏色,若是黃色在搜尋菜單下面的來替換
                backColor = cell.Interior.Color
                If backColor.ToString = "rgbYellow" Then
                    rng = Chr(col) + "33"
                    tempCell = sheet.Range(rng)
                    If tempCell.Value.ToString = "" Then
                        value = cell.Value
                    Else
                        value = tempCell.Value
                    End If
                Else
                    value = cell.Value
                End If
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
                tempCell = Nothing

                If row = 10 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    '判斷顏色,若是黃色在搜尋菜單下面的來替換
                    backColor = cell.Interior.Color
                    If backColor.ToString = "rgbYellow" Then
                        rng = Chr(col) + "33"
                        tempCell = sheet.Range(rng)
                        If tempCell.Value.ToString = "" Then
                            value = cell.Value
                        Else
                            value = tempCell.Value
                        End If
                    Else
                        value = cell.Value
                    End If
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 11 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    '判斷顏色,若是黃色在搜尋菜單下面的來替換
                    backColor = cell.Interior.Color
                    If backColor.ToString = "rgbYellow" Then
                        rng = Chr(col) + "33"
                        tempCell = sheet.Range(rng)
                        If tempCell.Value.ToString = "" Then
                            value = cell.Value
                        Else
                            value = tempCell.Value
                        End If
                    Else
                        value = cell.Value
                    End If
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicMoonSonLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                If row = 10 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 11 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If
            Next

            For row As Integer = 18 To 25
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "經典月子餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicMoonSonDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                '判斷顏色,若是黃色在搜尋菜單下面的來替換
                backColor = cell.Interior.Color
                If backColor.ToString = "rgbYellow" Then
                    rng = Chr(col) + "33"
                    tempCell = sheet.Range(rng)
                    If tempCell.Value.ToString = "" Then
                        value = cell.Value
                    Else
                        value = tempCell.Value
                    End If
                Else
                    value = cell.Value
                End If
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
                tempCell = Nothing

                If row = 19 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    '判斷顏色,若是黃色在搜尋菜單下面的來替換
                    backColor = cell.Interior.Color
                    If backColor.ToString = "rgbYellow" Then
                        rng = Chr(col) + "33"
                        tempCell = sheet.Range(rng)
                        If tempCell.Value.ToString = "" Then
                            value = cell.Value
                        Else
                            value = tempCell.Value
                        End If
                    Else
                        value = cell.Value
                    End If
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 20 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    '判斷顏色,若是黃色在搜尋菜單下面的來替換
                    backColor = cell.Interior.Color
                    If backColor.ToString = "rgbYellow" Then
                        rng = Chr(col) + "33"
                        tempCell = sheet.Range(rng)
                        If tempCell.Value.ToString = "" Then
                            value = cell.Value
                        Else
                            value = tempCell.Value
                        End If
                    Else
                        value = cell.Value
                    End If
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicMoonSonDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                If row = 19 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 20 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If
            Next

            For row As Integer = 26 To 28
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "經典月子餐",
                    .Meal = Meal.夜點,
                    .Meal_Detail = dicMoonSonNightSnack(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                If row = 27 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 28 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.夜點,
                    .Meal_Detail = dicMoonSonNightSnack(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                If row = 27 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 28 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If
            Next

            For row As Integer = 39 To 44
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "調理餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicConditioning(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                If value Is Nothing Then Continue For
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next
            '調理餐水果
            menu = New Menu With {
                .Version = ver,
                .MenuDate = d,
                .ProductName = "調理餐",
                .Meal = Meal.午餐,
                .Meal_Detail = dicConditioning(45)
            }
            rng = "D45"
            cell = sheet.Range(rng)
            value = cell.Value
            txt = value.ToString()
            menu.Name = txt
            lstMenu.Add(menu)

            For row As Integer = 50 To 54
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "幸福餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicHappinessLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 56 To 60
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "幸福餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicHappinessDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 65 To 70
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "住院餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicHospitalizedBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 71 To 78
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "住院餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicHospitalizedLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 79 To 85
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "住院餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicHospitalizedDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 89 To 94
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "輕食餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicLightMealBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 96 To 101
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "輕食餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicLightMealLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 103 To 108
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "輕食餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicLightMealDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next
        Next

        For col As Integer = Asc("O") To Asc("U")
            '日期
            rng = Chr(col) + "3"
            cell = sheet.Range(rng)
            value = cell.Value
            txt = value.ToString()
            Dim d = Date.Parse(txt)

            For row As Integer = 4 To 8
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "術後調理餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicOperationBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For Each row In dicOperationLunch.Keys
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "術後調理餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicOperationLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For Each row In dicOperationDinner.Keys
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "術後調理餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicOperationDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 27 To 31
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "素食餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicVegetarianBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 32 To 38
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "素食餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicVegetarianLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                If row = 32 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅1期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 33 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅3期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If
            Next

            For Each row In dicVegetarianDinner.Keys
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "素食餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicVegetarianDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)

                If row = 39 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅1期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If

                If row = 40 Then
                    menu = New Menu With {
                        .Version = ver,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅3期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt = value.ToString()
                    menu.Name = txt
                    lstMenu.Add(menu)
                End If
            Next

            For row As Integer = 53 To 57
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "素食一般餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicVegetarianNormalLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next

            For row As Integer = 58 To 62
                menu = New Menu With {
                    .Version = ver,
                    .MenuDate = d,
                    .ProductName = "素食一般餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicVegetarianNormalDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt = value.ToString()
                menu.Name = txt
                lstMenu.Add(menu)
            Next
        Next

        'insert到table
        For Each m In lstMenu
            Dim table = "menu"
            Dim dic As New Dictionary(Of String, String) From {
                {"me_date", m.MenuDate},
                {"me_version", m.Version},
                {"me_meal_id", m.Meal},
                {"me_meal_detail_id", m.Meal_Detail},
                {"me_name", m.Name}
            }
            '與商品匹配
            Dim dt = SelectFromTable($"SELECT prod_id FROM product WHERE prod_name = '{m.ProductName}'")
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                dic.Add("me_prod_id", row("prod_id").ToString)
                '先刪除後新增避免重複
                DeleteData(table, $"me_date = '{m.MenuDate:yyyy-MM-dd}' AND me_version = '{m.Version}' AND me_meal_id = {m.Meal} AND me_meal_detail_id = {m.Meal_Detail} AND me_prod_id = {row("prod_id")}")
                InserData(table, dic)
            Else
                MsgBox("無 " + m.ProductName + " 商品,請先新增")
                GoTo Finish
            End If
        Next
        DataToDgv(SelectFromTable(sqlMenu), "menu,product", dgvMenu)
        MsgBox("匯入完成")
Finish:
        Cursor = Cursors.Default
    End Sub

    '菜單管理-dgv點擊
    Private Sub dgvMenu_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvMenu.CellMouseClick
        ClearTabPage(tpMenu)

        If dgvMenu.SelectedRows.Count < 0 Then Exit Sub
        '點dgv後將對象資料傳至各控制項
        Dim dgvRow = dgvMenu.SelectedRows(0)
        Dim d As Date = dgvRow.Cells("me_date").Value
        Dim ver As String = dgvRow.Cells("me_version").Value
        Dim prod As Integer = dgvRow.Cells("prod_id").Value
        Dim dataMuenu = SelectFromTable($"SELECT * FROM menu WHERE me_date = '{d}' AND me_version = '{ver}' AND me_prod_id = '{prod}'").Rows
        For Each row As DataRow In dataMuenu
            Dim t As String = CStr(row.Field(Of Integer)("me_meal_id")) + "," + CStr(row.Field(Of Integer)("me_meal_detail_id"))
            For Each txt In tpMenu.Controls.OfType(Of TextBox).Where(Function(x) x.Tag = t)
                txt.Text = row.Field(Of String)("me_name")
            Next
        Next
        cmbProdName_menu.SelectedIndex = cmbProdName_menu.FindStringExact(dgvRow.Cells("prod_name").Value.ToString)
        cmbProdVers_menu.SelectedIndex = cmbProdVers_menu.FindStringExact(ver)
        dtMenu.Value = d
    End Sub

    '菜單管理-新增/修改
    Private Sub btnMunuInsert_Click(sender As Object, e As EventArgs) Handles btnMunuInsert.Click
        Cursor = Cursors.WaitCursor

        For Each txt In tpMenu.Controls.OfType(Of TextBox).Where(Function(x) String.IsNullOrEmpty(x.Text) = False)
            Dim dic As New Dictionary(Of String, String)
            With dic
                .Add("me_date", dtMenu.Value)
                If cmbProdVers_menu.SelectedItem Is Nothing Then
                    MsgBox("請選擇版本")
                    cmbProdVers_menu.Focus()
                    GoTo Finish
                End If
                Dim ver = cmbProdVers_menu.SelectedItem.ToString
                .Add("me_version", ver)
                If cmbProdName_menu.SelectedValue Is Nothing Then
                    MsgBox("請選擇商品")
                    cmbProdName_menu.Focus()
                    GoTo Finish
                End If
                Dim prodID = cmbProdName_menu.SelectedValue.ToString
                .Add("me_prod_id", prodID)
                Dim meal As String() = Split(txt.Tag, ",")
                .Add("me_meal_id", meal(0))
                .Add("me_meal_detail_id", meal(1))
                .Add("me_name", txt.Text)

                '先刪除後新增避免重複
                Dim table = "menu"
                DeleteData(table, $"me_date = '{dtMenu.Value:yyyy-MM-dd}' AND me_version = '{ver}' AND me_meal_id = {meal(0)} AND me_meal_detail_id = {meal(1)} AND me_prod_id = {prodID}")
                InserData(table, dic)
            End With
        Next
        ClearTabPage(tpMenu)
        DataToDgv(SelectFromTable(sqlMenu), "menu,product", dgvMenu)
        MsgBox("新增完成")
Finish:
        Cursor = Cursors.Default
    End Sub

    '菜單管理-刪除
    Private Sub btnMenuDel_Click(sender As Object, e As EventArgs) Handles btnMenuDel.Click
        Cursor = Cursors.WaitCursor

        Dim dic As New Dictionary(Of String, String)
        With dic
            .Add("me_date", dtMenu.Value)
            If cmbProdVers_menu.SelectedItem Is Nothing Then
                MsgBox("請選擇版本")
                cmbProdVers_menu.Focus()
                GoTo Finish
            End If
            Dim ver = cmbProdVers_menu.SelectedItem.ToString
            .Add("me_version", ver)
            If cmbProdName_menu.SelectedValue Is Nothing Then
                MsgBox("請選擇商品")
                cmbProdName_menu.Focus()
                GoTo Finish
            End If
            Dim prodID = cmbProdName_menu.SelectedValue.ToString
            .Add("me_prod_id", prodID)
            Dim table = "menu"
            DeleteData(table, $"me_date = '{dtMenu.Value:yyyy-MM-dd}' AND me_version = '{ver}' AND me_prod_id = {prodID}")
        End With
        ClearTabPage(tpMenu)
        DataToDgv(SelectFromTable(sqlMenu), "menu,product", dgvMenu)
        MsgBox("刪除完成")
Finish:
        Cursor = Cursors.Default
    End Sub

    '菜單管理-取消
    Private Sub btnMenuCancel_Click(sender As Object, e As EventArgs) Handles btnMenuCancel.Click
        ClearTabPage(tpMenu)
        DataToDgv(SelectFromTable(sqlMenu), "menu,product", dgvMenu)
    End Sub

    '菜單管理-搜尋
    Private Sub btnMenuQuery_Click(sender As Object, e As EventArgs) Handles btnMenuQuery.Click
        Dim sql = $"SELECT DISTINCT b.prod_name,a.me_date,a.me_version,b.prod_id FROM menu a LEFT JOIN product b ON a.me_prod_id=b.prod_id WHERE a.me_date = '{dtMenu.Value}'"
        If cmbProdVers_menu.SelectedItem IsNot Nothing Then sql += $" AND me_version = '{cmbProdVers_menu.SelectedItem}'"
        If cmbProdName_menu.SelectedValue IsNot Nothing Then sql += $" AND me_prod_id = '{cmbProdName_menu.SelectedValue}'"
        DataToDgv(SelectFromTable(sql), "menu,product", dgvMenu)
    End Sub

    ''' <summary>
    ''' 檢查Table是否有重複資料
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="sWhere">搜尋條件 xxx='xxx' and...</param>
    ''' <returns>True:有重複;False:沒重複</returns>
    Private Function CheckDataDuplication(sTable As String, sWhere As String, dgv As DataGridView) As Boolean
        Dim bResult As Boolean
        Dim dt As DataTable = SelectFromTable($"SELECT * FROM {sTable} WHERE {sWhere}")
        If dt.Rows.Count > 0 Then
            MsgBox("重複資料")

            '列出重複的資料
            DataToDgv(dt, sTable, dgv)
            bResult = True
        End If
        Return bResult
    End Function

    ''' <summary>
    ''' 綁定Table欄位與TextBox
    ''' </summary>
    Private Function Bind_TableTextBox(sTable As String) As Dictionary(Of String, String)
        Dim dicData As New Dictionary(Of String, String)
        Dim row As DataRow
        Dim rdo As RadioButton
        Dim chk As IEnumerable(Of CheckBox)
        Dim list As New List(Of String)
        Dim check As CheckBox
        With dicData
            Select Case sTable
                Case "customer"
                    .Add("cus_name", txtCusName_cus.Text)

                    .Add("cus_phone", txtPhone_cus.Text)

                    .Add("cus_telephone", txtTelephone.Text)

                    .Add("cus_address", txtAddress.Text)

                    .Add("cus_memo", txtMemo_cus.Text)

                    .Add("cus_line", txtLineID.Text) 'LineID

                    If Not String.IsNullOrWhiteSpace(txtBirthday.Text) Then .Add("cus_birthday", txtBirthday.Text)

                    If Not String.IsNullOrWhiteSpace(txtDueDate.Text) Then .Add("cus_due_date", txtDueDate.Text) '預產期

                    .Add("cus_hospital", txtHospital.Text) '產檢醫院

                    rdo = grpGender.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
                    If rdo IsNot Nothing Then
                        .Add("cus_gender", rdo.Text) '性別
                        rdo = Nothing
                    End If

                    .Add("cus_job", txtJob.Text)

                    rdo = grpMarriage.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
                    If rdo IsNot Nothing Then
                        .Add("cus_marriage", rdo.Text)
                        rdo = Nothing
                    End If

                    If rdoMarried.Checked Then
                        .Add("cus_spouse", txtSpouse.Text) '配偶
                        .Add("cus_children", txtChildren.Text) '子女人數                      
                    End If

                    rdo = grpAcad_Qual.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
                    If rdo IsNot Nothing Then
                        .Add("cus_acad_qual", rdo.Text) '學歷
                        rdo = Nothing
                    End If

                    For Each check In grpKind.Controls.OfType(Of CheckBox)
                        If check.Checked Then
                            If check.Text = "術後餐" Then
                                .Add("cus_kind_else", txtKindElse.Text)
                            End If
                            list.Add(check.Text)
                        End If
                    Next
                    .Add("cus_kind", String.Join(",", list)) '種類
                    list.Clear()

                    For Each check In grpGetMsg.Controls.OfType(Of CheckBox)
                        If check.Checked Then
                            If check.Text = "其他" Then
                                .Add("cus_getmsg_else", txtGetMsgElse.Text)
                            End If
                            list.Add(check.Text)
                        End If
                    Next
                    .Add("cus_get_msg", String.Join(",", list)) '得知媒體
                    list.Clear()

                    .Add("cus_many_child", txtManyChild.Text) '第幾胎

                    .Add("cus_conf_loca", txtConfLoca.Text) '月子地點

                    .Add("cus_conf_day", txtConfDay.Text) '月子天數

                    .Add("cus_conf_buy", txtConfBuy.Text) '欲購買月子餐天數

                    .Add("cus_height", txtHeight.Text)

                    .Add("cus_born_weight", txtBornWeight.Text) '產前體重

                    .Add("cus_weight", txtWeight.Text)

                    For Each check In grpDisease.Controls.OfType(Of CheckBox)
                        If check.Checked Then
                            If check.Text = "其他" Then
                                .Add("cus_disease_else", txtDisease.Text)
                            End If
                            list.Add(check.Text)
                        End If
                    Next
                    .Add("cus_disease", String.Join(",", list)) '疾病史
                    list.Clear()

                    .Add("cus_meal_adj", txtMealAdj.Text) '餐點調整

                    .Add("cus_diet_prep", txtDietPerp.Text) '飲食調配

                    .Add("cus_nutr_cons", txtNutrCons.Text) '營養顧問

                    .Add("cus_tabo_id", txtTaboo.Tag) '禁忌編號

                Case "product_group"
                    .Add("prod_grp_name", txtProdGrpName.Text)

                Case "product"
                    .Add("prod_name", txtProdName.Text)
                    .Add("prod_prod_grp_id", cmbProdGrp_product.SelectedValue)
                    .Add("prod_price", txtProdPrice.Text)
                    .Add("prod_cost", txtProdCost.Text)
                    .Add("prod_type", cmbProdType.Text)

                    '取得勾選的餐種
                    chk = grpMeal.Controls.OfType(Of CheckBox)().Where(Function(x) x.Checked)
                    .Add("prod_meal", String.Join(",", chk.Select(Function(x) x.Text)))
                    .Add("prod_memo", txtProdMemo.Text)

                Case "taboo"
                    .Add("tabo_type", cmbTaboClass.Text)
                    .Add("tabo_name", txtTaboName.Text)

                Case "orders"
                    row = SelectFromTable($"SELECT cus_id FROM customer WHERE cus_name = '{txtCusName_order.Text}' AND cus_phone = '{txtPhone_order.Text}'").Rows(0)
                    .Add("ord_cus_id", row("cus_id")) '客戶編號
                    row = SelectFromTable($"SELECT prod_id FROM product WHERE prod_name = '{cmbProdName_order.Text}'").Rows(0)
                    .Add("ord_prod_id", row("prod_id")) '商品編號
                    .Add("ord_date", dtOrdDate.Value.ToString("d")) '訂單日期
                    .Add("ord_count", txtCount.Text) '數量(天)
                    .Add("ord_price", txtTotalPrice.Text) '金額
                    .Add("ord_discount", txtDiscount.Text) '折讓金額
                    .Add("ord_breakfast", If(chkBreak_order.Checked, txtCount.Text, "0")) '早餐份數
                    .Add("ord_lunch", If(chkLunch_order.Checked, txtCount.Text, "0")) '午餐份數
                    .Add("ord_dinner", If(chkDinner_order.Checked, txtCount.Text, "0")) '晚餐份數
                    .Add("ord_delivery", dtDelivery.Value.ToString("d")) '預計送餐日
                    .Add("ord_memo", txtMemo_order.Text)
                    .Add("ord_deli_hosp", txtDeliHosp.Text) '生產醫院
                    .Add("ord_taste", txtTaste.Text) '試吃費
                    .Add("ord_tableware", txtTableware.Text) '押餐具費
                    .Add("ord_break_addr", txtAddrBreak.Text) '早餐地址
                    .Add("ord_lunch_addr", txtAddrLunch.Text) '午餐地址
                    .Add("ord_dinner_addr", txtAddrDinner.Text) '晚餐地址
                    rdo = grpEatType.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
                    If rdo IsNot Nothing Then
                        .Add("ord_eat_type", rdo.Text) '葷素
                        rdo = Nothing
                    End If
                    .Add("ord_emp_id", cmbSales.SelectedValue.ToString)

                Case "money"
                    row = SelectFromTable($"SELECT cus_id FROM customer WHERE cus_name = '{txtCusName_money.Text}' AND cus_phone = '{txtPhone_money.Text}'").Rows(0)
                    .Add("mon_cus_id", row("cus_id"))
                    .Add("mon_ord_id", txtOrdID_money.Text)
                    .Add("mon_date", dtMonDate.Value.ToString("d"))
                    .Add("mon_type", cmbMonType.Text)
                    .Add("mon_Income", txtMoney.Text)
                    .Add("mon_memo", txtMonMemo.Text)

                Case "permissions"
                    .Add("perm_name", cmbPosition_perm.Text)
                    .Add("perm_customer", IIf(chkCustomer.Checked, "Y", "N"))
                    .Add("perm_product", IIf(chkProduct.Checked, "Y", "N"))
                    .Add("perm_menu", IIf(chkMenu.Checked, "Y", "N"))
                    .Add("perm_order", IIf(chkOrders.Checked, "Y", "N"))
                    .Add("perm_distribute", IIf(chkDistr.Checked, "Y", "N"))
                    .Add("perm_report", IIf(chkReport.Checked, "Y", "N"))
                    .Add("perm_money", IIf(chkMoney.Checked, "Y", "N"))
                    .Add("perm_employee", IIf(chkEmployee.Checked, "Y", "N"))
                    .Add("perm_permissions", IIf(chkPermissions.Checked, "Y", "N"))
                    .Add("perm_taboo", IIf(chkTaboo.Checked, "Y", "N"))
                    .Add("perm_product_group", IIf(chkProdGrp.Checked, "Y", "N"))

                Case "employee"
                    .Add("emp_name", txtEmpName.Text)
                    .Add("emp_phone", txtEmpPhone.Text)
                    .Add("emp_tel", txtEmpTel.Text)
                    .Add("emp_address", txtEmpAddr.Text)
                    .Add("emp_perm_id", cmbPosition_emp.SelectedValue.ToString)
                    .Add("emp_acct", txtAcct.Text)
                    .Add("emp_psw", txtPsw.Text)
                    .Add("emp_memo", txtEmpMemo.Text)

                Case "distribute"
                    Dim txt As String
                    For Each grp In tpDistribute.Controls.OfType(Of GroupBox)
                        Select Case grp.Text
                            Case "湯盅"
                                txt = grp.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_soup", txt)
                            Case "麻油"
                                txt = grp.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_oil", txt)
                            Case "酒"
                                txt = grp.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_wine", txt)
                            Case "素"
                                txt = grp.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_vege", txt)
                            Case "其他"
                                chk = grp.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_other", String.Join(",", chk.Select(Function(x) x.Text)))
                            Case "客製需求"
                                chk = grp.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_customized", String.Join(",", chk.Select(Function(x) x.Text)))
                            Case "餐具"
                                chk = grp.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_tableware", String.Join(",", chk.Select(Function(x) x.Text)))
                        End Select
                    Next

            End Select
        End With
        Return dicData
    End Function

    Private Function InserData(sTable As String, dicData As Dictionary(Of String, String)) As Boolean
        Dim result As Boolean
        Dim cmd As New MySqlCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Values.Select(Function(x) $"'{x}'"))})", mConn)
        Try
            mConn.Open()
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        mConn.Close()
        Return result
    End Function

    ''' <summary>
    ''' 更新表格
    ''' </summary>
    ''' <param name="sTable">表格名稱</param>
    ''' <param name="dicFields">更新對象集合</param>
    ''' <param name="sCondition">Where</param>
    Public Function UpdateData(sTable As String, dicFields As Dictionary(Of String, String), sCondition As String) As Boolean
        Dim result As Boolean
        Try
            mConn.Open()

            '建立 UPDATE SQL 陳述式
            Dim sb As New StringBuilder()
            sb.AppendFormat("UPDATE {0} SET ", sTable)

            '取得更新的欄位名稱與值
            Dim lstFields As New List(Of String)
            For Each kvp As KeyValuePair(Of String, String) In dicFields
                lstFields.Add(String.Format("{0} = @{0}", kvp.Key))
            Next
            sb.Append(String.Join(", ", lstFields))

            '加上 WHERE 條件式
            If Not String.IsNullOrWhiteSpace(sCondition) Then
                sb.AppendFormat(" WHERE {0}", sCondition)
            End If

            '執行 SQL 陳述式
            Dim cmd As New MySqlCommand(sb.ToString(), mConn)
            For Each kvp As KeyValuePair(Of String, String) In dicFields
                cmd.Parameters.AddWithValue(String.Format("@{0}", kvp.Key), kvp.Value)
            Next
            If cmd.ExecuteNonQuery() > 0 Then result = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        mConn.Close()
        Return result
    End Function

    ''' <summary>
    ''' Insert前檢查
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <returns></returns>
    Private Function CheckInsert(sTable As String, tp As TabPage) As Boolean
        Dim bResult As Boolean
        If CheckTextNull(sTable, tp) Then GoTo Finish

        '不可重複的欄位
        Dim dic As New Dictionary(Of String, String)
        With dic
            Select Case sTable
                Case "product_group"
                    .Add("prod_grp_name", txtProdGrpName.Text)
                Case "product"
                    .Add("prod_name", txtProdName.Text)
                Case "taboo"
                    .Add("tabo_name", txtTaboName.Text)
                Case Else
                    GoTo Pass
            End Select
        End With
        Dim lst As List(Of String) = dic.Select(Function(x) $"{x.Key} = '{x.Value}'").ToList
        Dim sWhere = String.Join(" AND ", lst)
        Dim dgv = tp.Controls.OfType(Of DataGridView).FirstOrDefault
        If CheckDataDuplication(sTable, sWhere, dgv) Then GoTo Finish
Pass:
        bResult = True
Finish:
        Return bResult
    End Function

    ''' <summary>
    ''' 去頭尾空白後,檢查必填的欄位
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="tp">TabPage</param>
    ''' <returns>True:是空的;False:有文字</returns>
    Private Function CheckTextNull(sTable As String, tp As TabPage) As Boolean
        '去頭尾空白
        tp.Controls.OfType(Of TextBox).ToList().ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '找出資料表不能為空值的欄位註解名稱
        Dim dt As DataTable = SelectFromTable($"SELECT COLUMN_COMMENT FROM information_schema.columns WHERE table_schema = 'tingyi' AND TABLE_NAME='{sTable}' AND is_nullable = 'NO' AND column_key != 'PRI'")

        '比較與當前控制項.tag是否相符
        For Each ctrl As Windows.Forms.Control In tp.Controls
            Dim row As DataRow = dt.AsEnumerable().FirstOrDefault(Function(x) x("COLUMN_COMMENT").ToString() = ctrl.Tag)
            If row IsNot Nothing Then
                If String.IsNullOrWhiteSpace(ctrl.Text) Then
                    MsgBox(ctrl.Tag + "不能空白")
                    ctrl.Focus()
                    Return True
                End If
            End If
        Next
        Return False
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
    ''' 清空TabPage裡的控制項內容
    ''' </summary>
    ''' <param name="tp"></param>
    Private Sub ClearTabPage(tp As TabPage)
        For Each ctrl As Windows.Forms.Control In tp.Controls
            If TypeOf ctrl Is GroupBox Then
                ClearGroupBox(CType(ctrl, GroupBox))
            ElseIf TypeOf ctrl Is TabControl Then '取得TabControl裡的控制項
                For Each tp1 As TabPage In CType(ctrl, TabControl).Controls
                    ClearTabPage(tp1)
                Next
            End If
            ClearControl(ctrl)
        Next

    End Sub

    '搜尋欄位按下"Enter"即可搜尋
    Private Sub txtQuery_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtProdGrpName.KeyPress, txtProdQuery.KeyPress, txtTaboQuery.KeyPress, txtMonQuery.KeyPress, txtEmpQuery.KeyPress, txtOrdQuery.KeyPress, txtCusQuery.KeyPress, txtDistQuery.KeyPress
        If e.KeyChar = vbCr Then
            Dim btn As Button = CType(sender, TextBox).Parent.Controls.OfType(Of Button).FirstOrDefault(Function(x) x.Text = "查詢")
            btn.PerformClick()
        End If
    End Sub

    '清除GroupBox裡的控制項內容
    Private Sub ClearGroupBox(grp As GroupBox)
        Dim ctrl As Windows.Forms.Control
        For Each ctrl In grp.Controls
            ClearControl(ctrl)
        Next
    End Sub

    '清空控制項內容
    Private Sub ClearControl(ctrl As Windows.Forms.Control)
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is CheckBox Then
            CType(ctrl, CheckBox).Checked = False
        ElseIf TypeOf ctrl Is RadioButton Then
            CType(ctrl, RadioButton).Checked = False
        ElseIf TypeOf ctrl Is ComboBox Then
            CType(ctrl, ComboBox).SelectedIndex = -1
        End If
    End Sub

    ''' <summary>
    ''' 將資料放到DataGridView
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="sTable"></param>
    ''' <param name="dgv"></param>
    Private Sub DataToDgv(dt As DataTable, sTable As String, dgv As DataGridView)
        With dgv
            .DataSource = dt

            '用table欄位的備註將dgv的欄位改名
            Dim conditions As String = String.Join(" or ", sTable.Split(","c).Select(Function(x) $"Table_name = '{x.Trim()}'"))
            Dim TableCol As DataTable = SelectFromTable($"SELECT COLUMN_NAME, COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS  WHERE TABLE_SCHEMA = 'tingyi' AND {conditions}")
            For Each col As DataGridViewColumn In .Columns
                Dim row As DataRow = TableCol.AsEnumerable().FirstOrDefault(Function(x) x("COLUMN_NAME").ToString() = col.Name)
                If row IsNot Nothing Then
                    col.HeaderText = row("COLUMN_COMMENT").ToString()
                End If
            Next
            .AutoResizeColumnHeadersHeight()
        End With
    End Sub

    '商品群組管理-dgv點擊
    Private Sub dgvProdgroup_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvProdgroup.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            Dim row = dgv.SelectedRows(0)
            For Each txt As TextBox In dgv.Parent.Controls.OfType(Of TextBox)()
                'TextBox的Tag對應表格的備註
                Dim colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = txt.Tag)?.Name
                If Not String.IsNullOrEmpty(colName) Then
                    txt.Text = row.Cells(colName).Value.ToString()
                End If
            Next
        End If
    End Sub

    '商品管理-dgv點擊
    Private Sub dgvProduct_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvProduct.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            Dim row = dgv.SelectedRows(0)
            Dim colName As String
            For Each ctrl As Windows.Forms.Control In dgv.Parent.Controls
                'TextBox的Tag對應表格的備註
                If TypeOf ctrl Is TextBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        ctrl.Text = row.Cells(colName).Value.ToString()
                    End If

                ElseIf TypeOf ctrl Is ComboBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        Dim cmb = CType(ctrl, ComboBox)
                        cmb.SelectedIndex = cmb.FindStringExact(row.Cells(colName).Value.ToString)
                    End If

                ElseIf TypeOf ctrl Is GroupBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Text)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        Dim grp = CType(ctrl, GroupBox)
                        Dim names = row.Cells(colName).Value.ToString()
                        Dim chks As IEnumerable(Of CheckBox) = grp.Controls.OfType(Of CheckBox)()
                        Dim chk As CheckBox

                        '初始化checkbox
                        For Each chk In grp.Controls.OfType(Of CheckBox)()
                            chk.Checked = False
                        Next

                        For Each name As String In names.Split(","c)
                            chk = chks.FirstOrDefault(Function(x) x.Text = name)
                            If chk IsNot Nothing Then
                                chk.Checked = True
                            End If
                        Next
                    End If
                End If
            Next
        End If
    End Sub

    '禁忌管理-dgv點擊
    Private Sub dgvTaboo_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvTaboo.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            Dim row = dgv.SelectedRows(0)
            Dim colName As String
            For Each ctrl As Windows.Forms.Control In dgv.Parent.Controls
                'TextBox的Tag對應表格的備註
                If TypeOf ctrl Is TextBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        ctrl.Text = row.Cells(colName).Value.ToString()
                    End If

                ElseIf TypeOf ctrl Is ComboBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        Dim cmb = CType(ctrl, ComboBox)
                        cmb.SelectedIndex = cmb.FindStringExact(row.Cells(colName).Value.ToString)
                    End If

                ElseIf TypeOf ctrl Is GroupBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Text)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        Dim grp = CType(ctrl, GroupBox)
                        Dim names = row.Cells(colName).Value.ToString()
                        Dim chks As IEnumerable(Of CheckBox) = grp.Controls.OfType(Of CheckBox)()
                        Dim chk As CheckBox

                        '初始化checkbox
                        For Each chk In grp.Controls.OfType(Of CheckBox)()
                            chk.Checked = False
                        Next

                        For Each name As String In names.Split(","c)
                            chk = chks.FirstOrDefault(Function(x) x.Text = name)
                            If chk IsNot Nothing Then
                                chk.Checked = True
                            End If
                        Next
                    End If
                End If
            Next
        End If
    End Sub

    '財務管理-dgv點擊
    Private Sub dgvMoney_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvMoney.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            Dim row = dgv.SelectedRows(0)
            Dim colName As String
            For Each ctrl As Windows.Forms.Control In dgv.Parent.Controls
                'TextBox的Tag對應表格的備註
                If TypeOf ctrl Is TextBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        ctrl.Text = row.Cells(colName).Value.ToString()
                    End If

                ElseIf TypeOf ctrl Is ComboBox Then
                    colName = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(x) x.HeaderText = ctrl.Tag)?.Name
                    If Not String.IsNullOrEmpty(colName) Then
                        Dim cmb = CType(ctrl, ComboBox)
                        cmb.SelectedIndex = cmb.FindStringExact(row.Cells(colName).Value.ToString)
                    End If
                End If
            Next
        End If
    End Sub

    '權限管理-dgv點擊
    Private Sub dgvPermissions_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvPermissions.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            '初始化checkbox
            dgv.Parent.Controls.OfType(Of CheckBox).ToList().ForEach(Sub(chk) chk.Checked = False)

            Dim row = dgv.SelectedRows(0)
            With row
                txtPermID.Text = .Cells("perm_id").Value.ToString
                cmbPosition_perm.SelectedIndex = cmbPosition_perm.FindStringExact(.Cells("perm_name").Value.ToString)
                If .Cells("perm_customer").Value.ToString = "Y" Then chkCustomer.Checked = True
                If .Cells("perm_product").Value.ToString = "Y" Then chkProduct.Checked = True
                If .Cells("perm_menu").Value.ToString = "Y" Then chkMenu.Checked = True
                If .Cells("perm_order").Value.ToString = "Y" Then chkOrders.Checked = True
                If .Cells("perm_distribute").Value.ToString = "Y" Then chkDistr.Checked = True
                If .Cells("perm_report").Value.ToString = "Y" Then chkReport.Checked = True
                If .Cells("perm_money").Value.ToString = "Y" Then chkMoney.Checked = True
                If .Cells("perm_employee").Value.ToString = "Y" Then chkEmployee.Checked = True
                If .Cells("perm_permissions").Value.ToString = "Y" Then chkPermissions.Checked = True
                If .Cells("perm_taboo").Value.ToString = "Y" Then chkTaboo.Checked = True
                If .Cells("perm_product_group").Value.ToString = "Y" Then chkProdGrp.Checked = True
            End With

        End If
    End Sub

    '商品群組管理-新增
    Private Sub btnProdGrpInsert_Click(sender As Object, e As EventArgs) Handles btnProdGrpInsert.Click
        Cursor = Cursors.WaitCursor

        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡

        If Not CheckInsert(sTable, tp) Then GoTo Finish

        InserData(sTable, Bind_TableTextBox(sTable))

        '列出所有表格資料
        btnProdGrpCancel.PerformClick()
Finish:
        Me.Cursor = Cursors.Default
    End Sub

    '商品管理-新增
    Private Sub btnProdInsert_Click(sender As Object, e As EventArgs) Handles btnProdInsert.Click
        Cursor = Cursors.WaitCursor

        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡

        If Not CheckInsert(sTable, tp) Then GoTo Finish

        InserData(sTable, Bind_TableTextBox(sTable))

        '列出所有表格資料
        btnProdCancel.PerformClick()
Finish:
        Me.Cursor = Cursors.Default
    End Sub

    '禁忌管理-新增
    Private Sub btnTaboInsert_Click(sender As Object, e As EventArgs) Handles btnTaboInsert.Click
        Cursor = Cursors.WaitCursor

        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡
        If Not CheckInsert(sTable, tp) Then GoTo Finish

        InserData(sTable, Bind_TableTextBox(sTable))

        '列出所有表格資料
        btnTaboCancel.PerformClick()
        InitTabooType()
Finish:
        Cursor = Cursors.Default
    End Sub

    '財務管理-新增
    Private Sub btnMonInsert_Click(sender As Object, e As EventArgs) Handles btnMonInsert.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡
        If Not CheckInsert(sTable, tp) Then GoTo Finish
        InserData(sTable, Bind_TableTextBox(sTable))
        '列出所有表格資料
        btnMonCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '權限管理-新增
    Private Sub btnPermInsert_Click(sender As Object, e As EventArgs) Handles btnPermInsert.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡
        If Not CheckInsert(sTable, tp) Then GoTo Finish
        InserData(sTable, Bind_TableTextBox(sTable))
        '列出所有表格資料
        btnPermCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '商品群組管理-修改
    Private Sub btnProdGrpModify_Click(sender As Object, e As EventArgs) Handles btnProdGrpModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"prod_grp_id = '{txtProdGrpID.Text}'")

        '列出所有資料
        btnProdGrpCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '商品管理-修改
    Private Sub btnProdModify_Click(sender As Object, e As EventArgs) Handles btnProdModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        If UpdateData(sTable, Bind_TableTextBox(sTable), $"prod_id  = '{txtProdID.Text}'") Then MsgBox("修改成功")

        '列出所有資料
        btnProdCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '禁忌管理-修改
    Private Sub btnTaboModify_Click(sender As Object, e As EventArgs) Handles btnTaboModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"tabo_id  = '{txtTaboID.Text}'")
        '列出所有資料
        btnTaboCancel.PerformClick()
        InitTabooType()
Finish:
        Cursor = Cursors.Default
    End Sub

    '財務管理-修改
    Private Sub btnMonModify_Click(sender As Object, e As EventArgs) Handles btnMonModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"mon_id  = '{txtMonID.Text}'")
        '列出所有資料
        btnMonCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '權限管理-修改
    Private Sub btnPermModify_Click(sender As Object, e As EventArgs) Handles btnPermModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"perm_id  = '{txtPermID.Text}'")
        '列出所有資料
        btnPermCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

    '商品群組管理-刪除
    Private Sub btnProdGrpDel_Click(sender As Object, e As EventArgs) Handles btnProdGrpDel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"prod_grp_id = '{id.Text}'") Then
            MsgBox("刪除成功")

            btnProdGrpCancel.PerformClick()
        End If
    End Sub

    '商品管理-刪除
    Private Sub btnProdDelete_Click(sender As Object, e As EventArgs) Handles btnProdDelete.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"prod_id  = '{id.Text}'") Then
            MsgBox("刪除成功")

            '顯示table所有資料
            btnProdCancel.PerformClick()
        End If
    End Sub

    '禁忌管理-刪除
    Private Sub btnTaboDel_Click(sender As Object, e As EventArgs) Handles btnTaboDel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"tabo_id  = '{id.Text}'") Then
            MsgBox("刪除成功")

            '顯示table所有資料
            btnTaboCancel.PerformClick()
            InitTabooType()
        End If
    End Sub

    '財務管理-刪除
    Private Sub btnMonDel_Click(sender As Object, e As EventArgs) Handles btnMonDel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"mon_id  = '{id.Text}'") Then
            MsgBox("刪除成功")

            '顯示table所有資料
            btnMonCancel.PerformClick()
        End If
    End Sub

    '權限管理-刪除
    Private Sub btnPermDel_Click(sender As Object, e As EventArgs) Handles btnPermDel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "權限編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"perm_id  = '{id.Text}'") Then
            MsgBox("刪除成功")

            '顯示table所有資料
            btnPermCancel.PerformClick()
        End If
    End Sub

    '商品群組管理-取消
    Private Sub btnProdGrpCancel_Click(sender As Object, e As EventArgs) Handles btnProdGrpCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        '顯示所有資料
        Dim sTable As String = tp.Tag.ToString
        DataToDgv(SelectFromTable(sqlProductGroup), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        InitProductGroup()
    End Sub

    '商品管理-取消
    Private Sub btnProdCancel_Click(sender As Object, e As EventArgs) Handles btnProdCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        '顯示所有資料
        DataToDgv(SelectFromTable(sqlProduct), "product", dgvProduct)
    End Sub

    '禁忌管理-取消
    Private Sub btnTaboCancel_Click(sender As Object, e As EventArgs) Handles btnTaboCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        '顯示所有資料
        DataToDgv(SelectFromTable(sqlTaboo), "taboo", dgvTaboo)
    End Sub

    '財務管理-取消
    Private Sub btnMonCancel_Click(sender As Object, e As EventArgs) Handles btnMonCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        '顯示所有資料       
        DataToDgv(SelectFromTable(sqlMoney), "customer,orders,money", dgvMoney)
    End Sub

    '權限管理-取消
    Private Sub btnPermCancel_Click(sender As Object, e As EventArgs) Handles btnPermCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '顯示所有資料
        DataToDgv(SelectFromTable(sqlPermision), "permissions", dgvPermissions)
        ClearTabPage(tp)
        InitPosition()
    End Sub

    '商品群組管理-查詢
    Private Sub btnProdGrpQuery_Click(sender As Object, e As EventArgs) Handles btnProdGrpQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = tp.Tag
        Dim sql = sqlProductGroup + $" WHERE prod_grp_name LIKE '%{txtProdGrpName.Text}%'"
        DataToDgv(SelectFromTable(sql), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        ClearTabPage(tp)
        InitProductGroup()
        MsgBox("搜尋完畢")

        Cursor = Cursors.Default
    End Sub

    '商品管理-查詢
    Private Sub btnProdQuery_Click(sender As Object, e As EventArgs) Handles btnProdQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = "product,product_group"
        Dim sql = sqlProduct + $" WHERE a.prod_name LIKE '%{txtProdQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '禁忌管理-查詢
    Private Sub btnTaboQuery_Click(sender As Object, e As EventArgs) Handles btnTaboQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sql = sqlTaboo + $" WHERE tabo_type LIKE '%{txtTaboQuery.Text}%' OR tabo_name LIKE '%{txtTaboQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), "taboo", tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '財務管理-查詢
    Private Sub btnMonQuery_Click(sender As Object, e As EventArgs) Handles btnMonQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim Sql = sqlMoney + $" WHERE b.cus_name LIKE '%{txtMonQuery.Text}%' OR b.cus_phone LIKE '%{txtMonQuery.Text}%'"
        DataToDgv(SelectFromTable(Sql), "customer,money,orders", dgvMoney)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    Private Sub btnTaboo_Click(sender As Object, e As EventArgs) Handles btnTaboo.Click
        frmTaboo.Show()
    End Sub

End Class
