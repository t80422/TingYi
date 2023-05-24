Imports System.Configuration
Imports System.Text
Imports MySql.Data.MySqlClient

Public Class frmMain
    Friend dtTaboo As DataTable
    'todo 1.登入頁的 大底圖，可以 800*600 px，或是左邊這區塊 320*360 px檔案格式為JPG
    '2.舊會員資料使用EXCEL轉入
    '4.合約書
    '7.訂單加上業務人員，員工資料要有業務員身分別
    '8.客戶頁多顯示歷史訂單紀錄
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
        InitProduct()
        InitTabooType()
        InitPosition()
        '初始化收款方式
        cmbMonType.Items.Add("全款")
        cmbMonType.Items.Add("訂金")
        '初始化禁忌清單
        dtTaboo = SelectFromTable("SELECT * FROM taboo")

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
        Dim backColor As Color = If(isSelected, Color.CornflowerBlue, Color.WhiteSmoke)
        e.Graphics.FillRectangle(New SolidBrush(backColor), e.Bounds)

        ' 繪製索引標籤的文字
        Dim text As String = tab.Text
        Dim textColor As Color = If(isSelected, Color.White, Color.Black)
        Dim font As Font = tabControl.Font
        e.Graphics.DrawString(text, font, New SolidBrush(textColor), e.Bounds.Location)
    End Sub

    ''' <summary>
    ''' 初始化商品ComboBox
    ''' </summary>
    Private Sub InitProduct()
        '初始化商品
        With cmbProdName_order
            .DataSource = SelectFromTable("SELECT * FROM product")
            .DisplayMember = "prod_name"
            .ValueMember = "prod_id"
            .SelectedIndex = -1
        End With
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

        '初始化商品


        Dim col As New Collection From {
            cmbPosition_perm,
            cmbPosition_emp
        }
        For i As Short = 1 To col.Count
            With col(i)
                .DataSource = SelectFromTable("SELECT * FROM permissions")
                .DisplayMember = "perm_name"
                .SelectedIndex = -1
            End With
        Next
    End Sub

    ''' <summary>
    '''初始化DataGrid欄位
    ''' </summary>
    Private Sub InitDataGrid()
        Dim sql As String
        '客戶管理
        sql = "SELECT cus_id, cus_name, cus_gender, cus_phone FROM customer"
        DataToDgv(SelectFromTable(sql), "customer", dgvCustomer)
        '商品群組管理
        sql = "SELECT * FROM product_group"
        DataToDgv(SelectFromTable(sql), "product_group", dgvProdgroup)
        '商品管理
        sql = "SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id"
        DataToDgv(SelectFromTable(sql), "product,product_group", dgvProduct)
        '禁忌管理
        sql = "SELECT * FROM taboo"
        DataToDgv(SelectFromTable(sql), "taboo", dgvTaboo)
        '訂單管理
        sql = "SELECT a.ord_id,a.ord_date,b.cus_name,b.cus_phone FROM orders a LEFT JOIN customer b ON a.ord_cus_id = b.cus_id LEFT JOIN product c ON a.ord_prod_id = c.prod_id LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id"
        DataToDgv(SelectFromTable(sql), "customer,orders", dgvOrder)
        '財務管理
        sql = "SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id"
        DataToDgv(SelectFromTable(sql), "customer,orders,money", dgvMoney)
        '權限管理
        sql = "SELECT * FROM permissions"
        DataToDgv(SelectFromTable(sql), "permissions", dgvPermissions)
        '員工管理
        sql = "SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_pos_id = b.perm_id"
        DataToDgv(SelectFromTable(sql), "permissions,employee", dgvEmployee)

        '菜單管理
        With dgvMenu
            .Columns.Add("", "編號")
            .Columns.Add("", "版本")
            .Columns.Add("", "日期")
            .Columns.Add("", "商品名稱")
            .Columns.Add("", "早餐-主食")
            .Columns.Add("", "早餐-主菜")
            .Columns.Add("", "早餐-半葷")
            .Columns.Add("", "早餐-青菜/西飲")
            .Columns.Add("", "早餐-湯品")
            .Columns.Add("", "早餐-飲品")
            .Columns.Add("", "午餐-湯盅")
            .Columns.Add("", "午餐-湯盅(1,2期)")
            .Columns.Add("", "午餐-湯盅(3,4期)")
            .Columns.Add("", "午餐-主食")
            .Columns.Add("", "午餐-主菜")
            .Columns.Add("", "午餐-半葷")
            .Columns.Add("", "午餐-青菜")
            .Columns.Add("", "午餐-水果")
            .Columns.Add("", "午餐-甜品")
            .Columns.Add("", "午餐-飲品")
            .Columns.Add("", "晚餐-湯盅")
            .Columns.Add("", "晚餐-湯盅(1,2期)")
            .Columns.Add("", "晚餐-湯盅(3,4期)")
            .Columns.Add("", "晚餐-主食")
            .Columns.Add("", "晚餐-主菜")
            .Columns.Add("", "晚餐-半葷")
            .Columns.Add("", "晚餐-青菜")
            .Columns.Add("", "晚餐-水果")
            .Columns.Add("", "晚餐-飲品")
            .Columns.Add("", "晚點-湯盅")
            .Columns.Add("", "晚點-湯盅(1,2期)")
            .Columns.Add("", "晚點-湯盅(3,4期)")
            .Rows.Add("1", "B", "2023-01-23", "經典月子餐", "黃金小米粥", "泰式沙嗲烤豬", "燻雞香拌雲耳", "蒜香龍鬚菜", "黃芪鮮雞湯", "", "枸杞排骨湯", "枸杞排骨湯", "杜仲燉排骨", "傳香地瓜飯", "蒜蓉海大蝦", "茶油杏菇爆炒腰子", "玉米高麗菜", "柳丁",
                      "紅糖大麥粥", "", "錦蔬鮮魚湯", "錦蔬鮮魚湯", "何首烏鮮魚湯", "枸杞養生飯(茶油)", "醬燒煨豬膝", "塔香肉絲海龍", "吻魚白杏菜", "黃奇果", "", "玉竹鮮雞湯", "干貝鮮雞湯", "八珍干貝鮮雞湯")

            .Rows.Add("2", "B", "2023-01-03", "溫馨月子餐", "照燒梅花三明治", "田園烤白筍", "起司煎蛋", "養生芝麻飲", "青木瓜燉魚湯", "", "枸杞排骨湯", "何首烏排骨湯", "何首烏排骨湯", "芝麻糙米飯", "檸檬香煎海魚",
                      "黃耆炒雞肉", "薑絲蔭醬過貓", "柳丁", "桂圓銀耳甜湯", "", "玉竹鮮雞湯", "紅棗玉竹鮮雞湯", "黨蔘鮮雞湯", "養生紫米飯", "秘製紅酒牛腩", "翡翠鮮菇蒸雙鮮", "腐乳高麗菜", "百香果", "", "棗香龍尾湯", "棗香龍尾湯",
                      "龍尾虎豆燉紅棗")

            .Rows.Add("3", "C", "2023-03-11", "幸福餐", "", "", "", "", "", "", "北蟲草花鮮雞湯", "", "", "香甜栗子飯", "南方澳帶魚捲(烤)", "塔香杏鮑菇", "鮮菇白杏", "", "", "味噌魚頭湯", "", "", "", "養生五穀飯",
                      "磨菇豬小排", "茶香紅棗雞", "蒜香青江菜")

            .Rows.Add("4", "D", "2023-01-19", "住院餐", "田園時蔬雞肉粥", "椒塩烤鮑菇", "茄汁肉丸", "香菇高麗菜", "黃耆片鮮魚湯", "觀音串", "無花果排骨湯", "", "", "茶香珍菇飯", "梅子燒雞", "清炒香蔥魚栁", "金銀蛋莧菜",
                      "四季水果", "紅糖燕麥粥", "杜仲茶", "玉竹鮮雞湯", "紅藜高纖飯", "粉蒸排骨(不要豆鼓)", "美人腿炒雞(茶香)", "吻魚炒青江菜", "", "通乳茶", "北菇燉魚湯")

        End With

        txtProdName_menu.Text = "經典月子餐"
        cmbProdVers_menu.Text = "B"
        dtMenu.Value = "2023-01-23"
        txtBraSta.Text = "黃金小米粥"
        txtBlaMain.Text = "泰式沙嗲烤豬"
        txtBlaHM.Text = "燻雞香拌雲耳"
        txtBlaVag.Text = "蒜香龍鬚菜"
        txtBlaSoup.Text = "黃芪鮮雞湯"
        txtBlaDri.Text = ""
        txtLunSoup.Text = "枸杞排骨湯"
        txtLun1.Text = "枸杞排骨湯"
        txtLun3.Text = "杜仲燉排骨"
        txtLunSta.Text = "傳香地瓜飯"
        txtLunMain.Text = "蒜蓉海大蝦"
        txtLunHM.Text = "茶油杏菇爆炒腰子"
        txtLunVag.Text = "玉米高麗菜"
        txtLunFru.Text = "柳丁"
        txtLunDess.Text = "紅糖大麥粥"
        txtLunDri.Text = ""
        txtDinSoup.Text = "錦蔬鮮魚湯"
        txtDin1.Text = "錦蔬鮮魚湯"
        txtDin3.Text = "何首烏鮮魚湯"
        txtDinSta.Text = "枸杞養生飯(茶油)"
        txtDinMain.Text = "醬燒煨豬膝"
        txtDinHM.Text = "塔香肉絲海龍"
        txtDinVag.Text = "吻魚白杏菜"
        txtDinFru.Text = "黃奇果"
        txtDinDri.Text = ""
        txtNSSoup.Text = "玉竹鮮雞湯"
        txtNS1.Text = "干貝鮮雞湯"
        txtNS3.Text = "八珍干貝鮮雞湯"

        '配餐管理
        txtCusName_dist.Text = "陳小姐"
        txtPhone_dist.Text = "0918-123123"
    End Sub

    Private Sub btnMenuExcel_Click(sender As Object, e As EventArgs) Handles btnMenuExcel.Click
        'Dim excelApp As New Excel.Application

        'Dim workbook As Excel.Workbook = excelApp.Workbooks.Open("C:\ExcelFile.xlsx")
        'Dim worksheet As Excel.Worksheet = workbook.Sheets("Sheet1")

        'Dim cellValue As String = worksheet.Cells(1, 1).Value

        'workbook.Close(False)
        'Marshal.ReleaseComObject(workbook)
        'Marshal.ReleaseComObject(excelApp)

    End Sub

    Private Sub cmdProdName_order_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbProdName_order.SelectedValueChanged
        '更新商品分類
        '若商品分類是套餐則顯示"三餐"

    End Sub

    '客戶管理-dgv點擊
    Private Sub dgvCustomer_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvCustomer.CellMouseClick
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
        End If
    End Sub

    '客戶管理-新增
    Private Sub btnCusInsert_Click(sender As Object, e As EventArgs) Handles btnCusInsert.Click
        Cursor = Cursors.WaitCursor

        Dim table = "customer"
        If Not CheckCustomerData(table) Then GoTo Finish

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
        If Not CheckCustomerData(table) Then GoTo Finish
        UpdateData(table, Bind_TableTextBox(table), $"cus_id = '{txtCusID.Text}'")

        btnCusCancel.PerformClick()
Finish:
        Cursor = Cursors.Default
    End Sub

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
        Dim table = "customer"
        Dim sql = $"SELECT cus_id, cus_name, cus_gender, cus_phone FROM {table}"
        DataToDgv(SelectFromTable(sql), table, dgvCustomer)
        ClearTabPage(tpBasic_cus)
        ClearTabPage(tpConsult_cus)
    End Sub

    '客戶管理-查詢
    Private Sub btnCusQuery_Click(sender As Object, e As EventArgs) Handles btnCusQuery.Click
        Cursor = Cursors.WaitCursor
        Dim table = "customer"
        Dim sql = $"SELECT cus_id, cus_name, cus_gender, cus_phone FROM {table} WHERE cus_name LIKE '%{txtCusQuery.Text}%' or cus_phone LIKE '%{txtCusQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), table, dgvCustomer)
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
        Dim rowData = SelectFromTable($"SELECT * FROM orders a LEFT JOIN customer b ON a.ord_cus_id = b.cus_id LEFT JOIN product c ON a.ord_prod_id = c.prod_id LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id WHERE ord_id = '{row.Cells("ord_id").Value}'").Rows(0)
        For Each ctrl As Control In dgv.Parent.Controls
            colName = ctrl.Tag 'TextBox的Tag對應表格的名稱
            If TypeOf ctrl Is TextBox Then
                If Not String.IsNullOrEmpty(colName) Then ctrl.Text = rowData(colName).ToString

            ElseIf TypeOf ctrl Is DateTimePicker Then
                If Not String.IsNullOrEmpty(colName) Then
                    Dim dtp = CType(ctrl, DateTimePicker)
                    dtp.Value = rowData(colName)
                End If

            ElseIf TypeOf ctrl Is ComboBox Then
                If Not String.IsNullOrEmpty(colName) Then
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
            txtUnpay.Text = txtPrice_order.Text
            Exit Sub
        End If
        Dim dr As DataRow = dt.Rows(0)
        txtUnpay.Text = txtPrice_order.Text - dr("mon_income").ToString
        'todo 要再驗算 如果有兩筆財務的話
    End Sub

    ''' <summary>
    ''' 檢查Customer即將上傳的內容是否有誤
    ''' </summary>
    ''' <returns>True:正確 False:錯誤</returns>
    Private Function CheckCustomerData(table As String) As Boolean
        '去txt頭尾空白
        tpBasic_cus.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Text = Trim(txt.Text))
        tpConsult_cus.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '檢查必填欄位
        If String.IsNullOrWhiteSpace(txtCusName_cus.Text) Then
            MsgBox("姓名不能空白")
            txtCusName_cus.Focus()
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtPhone_cus.Text) Then
            MsgBox("手機不能空白")
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
                MsgBox("預產期日期格式錯誤")
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
            If Not String.IsNullOrEmpty(txt.Text) AndAlso Not IsNumeric(txt.Text) Then
                Dim tp As TabPage = If(TypeOf txt.Parent Is TabPage, txt.Parent, txt.Parent.Parent)
                tcCustomer.SelectedTab = tp
                MsgBox($"{dic(txt)} 請輸入數字")
                txt.Focus()
                Return False
            End If
        Next

        Return True
    End Function

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
                    .Add("ord_cus_id", row("cus_id"))
                    row = SelectFromTable($"SELECT prod_id FROM product WHERE prod_name = '{cmbProdName_order.Text}'").Rows(0)
                    .Add("ord_prod_id", row("prod_id"))
                    .Add("ord_date", dtOrdDate.Value.ToString("d"))
                    .Add("ord_count", txtCount.Text)
                    .Add("ord_price", txtPrice_order.Text)
                    .Add("ord_discount", txtDiscount.Text)
                    .Add("ord_breakfast", IIf(chkBreak_order.Checked, txtCount.Text, "0"))
                    .Add("ord_lunch", IIf(chkLunch_order.Checked, txtCount.Text, "0"))
                    .Add("ord_dinner", IIf(chkDinner_order.Checked, txtCount.Text, "0"))
                    .Add("ord_delivery", dtDelivery.Value.ToString("d"))
                    .Add("ord_memo", txtMemo_order.Text)

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
                    row = SelectFromTable($"SELECT perm_id FROM permissions WHERE perm_name = '{cmbPosition_emp.Text}'").Rows(0)
                    .Add("emp_pos_id", row("perm_id"))
                    .Add("emp_acct", txtAcct.Text)
                    .Add("emp_psw", txtPsw.Text)
                    .Add("emp_memo", txtEmpMemo.Text)
            End Select
        End With
        Return dicData
    End Function

    Private Sub InserData(sTable As String, dicData As Dictionary(Of String, String))
        Dim cmd As New MySqlCommand($"INSERT INTO {sTable} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Values.Select(Function(x) $"'{x}'"))})", mConn)
        Try
            mConn.Open()
            If cmd.ExecuteNonQuery() > 0 Then MsgBox("新增成功")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        mConn.Close()
        btnCusCancel.PerformClick()
    End Sub

    ''' <summary>
    ''' 更新表格
    ''' </summary>
    ''' <param name="sTable">表格名稱</param>
    ''' <param name="dicFields">更新對象集合</param>
    ''' <param name="sCondition">Where</param>
    Public Sub UpdateData(sTable As String, dicFields As Dictionary(Of String, String), sCondition As String)
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
            If cmd.ExecuteNonQuery() > 0 Then
                MsgBox("修改成功")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        mConn.Close()
    End Sub

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
        For Each ctrl As Control In tp.Controls
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
    Private Sub txtQuery_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtProdGrpName.KeyPress, txtProdQuery.KeyPress, txtTaboQuery.KeyPress, txtMonQuery.KeyPress, txtEmpQuery.KeyPress, txtOrdQuery.KeyPress, txtCusQuery.KeyPress
        If e.KeyChar = vbCr Then
            Dim btn As Button = CType(sender, TextBox).Parent.Controls.OfType(Of Button).FirstOrDefault(Function(x) x.Text = "查詢")
            btn.PerformClick()
        End If
    End Sub

    '清除GroupBox裡的控制項內容
    Private Sub ClearGroupBox(grp As GroupBox)
        Dim ctrl As Control
        For Each ctrl In grp.Controls
            ClearControl(ctrl)
        Next
    End Sub

    '清空控制項內容
    Private Sub ClearControl(ctrl As Control)
        If (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is ComboBox) Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is CheckBox Then
            Dim chk As CheckBox = CType(ctrl, CheckBox)
            chk.Checked = False
        ElseIf TypeOf ctrl Is RadioButton Then
            Dim rdo As RadioButton = CType(ctrl, RadioButton)
            rdo.Checked = False
        End If
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
    End Sub

    '更改月曆時間,有訂單就找訂單月份,沒訂單就用現在月份
    Private Sub btnDistQuery_Click(sender As Object, e As EventArgs) Handles btnDistQuery.Click
        tlpCalendar.Visible = False
        '程式碼
        tlpCalendar.Visible = True
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
            For Each ctrl As Control In dgv.Parent.Controls
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
            For Each ctrl As Control In dgv.Parent.Controls
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
            For Each ctrl As Control In dgv.Parent.Controls
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

    '員工管理-dgv點擊
    Private Sub dgvEmployee_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvEmployee.CellMouseClick
        Dim dgv = CType(sender, DataGridView)
        If dgv.SelectedRows.Count > 0 Then
            Dim row = dgv.SelectedRows(0)
            Dim colName As String
            For Each ctrl As Control In dgv.Parent.Controls
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

    '商品群組管理-新增
    Private Sub btnProdGrpInsert_Click(sender As Object, e As EventArgs) Handles btnProdGrpInsert.Click
        Cursor = Cursors.WaitCursor

        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡

        If Not CheckInsert(sTable, tp) Then GoTo Finish

        InserData(sTable, Bind_TableTextBox(sTable))

        '列出所有表格資料
        DataToDgv(SelectFromTable($"SELECT * FROM {sTable}"), sTable, dgvProdgroup)
        ClearTabPage(tp)
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
        DataToDgv(SelectFromTable("SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id;"), "product,product_group", dgvProduct)
        ClearTabPage(tp)

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
        DataToDgv(SelectFromTable("SELECT * FROM taboo"), "taboo", dgvTaboo)
        ClearTabPage(tp)
        InitTabooType()
Finish:
        Cursor = Cursors.Default
    End Sub

    '訂單管理-新增
    Private Sub btnOrdInsert_Click(sender As Object, e As EventArgs) Handles btnOrdInsert.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString 'Table名稱寫在TabPage的Tag裡
        If Not CheckInsert(sTable, tp) Then GoTo Finish
        InserData(sTable, Bind_TableTextBox(sTable))
        '列出所有表格資料
        Dim sql = "SELECT ord.ord_id, cus.cus_name, cus.cus_phone, ord.ord_date, prd.prod_name, pg.prod_grp_name, ord.ord_count, ord.ord_delivery, ord.ord_price, ord.ord_discount, ord.ord_breakfast, ord.ord_lunch,ord. ord_dinner, ord.ord_memo FROM `orders` ord LEFT JOIN customer cus ON ord.ord_cus_id=cus.cus_id LEFT JOIN product prd ON ord.ord_prod_id=prd.prod_id LEFT JOIN product_group pg ON prd.prod_prod_grp_id=pg.prod_grp_id"
        DataToDgv(SelectFromTable(sql), "customer,product,product_group,orders", dgvOrder)
        ClearTabPage(tp)
        InitProduct()
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
        Dim Sql = "SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_Income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id"
        DataToDgv(SelectFromTable(Sql), "customer,orders,money", dgvMoney)
        ClearTabPage(tp)
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
        Dim Sql = "SELECT * FROM permissions"
        DataToDgv(SelectFromTable(Sql), "permissions", dgvPermissions)
        ClearTabPage(tp)
        InitPosition()
Finish:
        Cursor = Cursors.Default
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
        '列出所有表格資料
        Dim Sql = "SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_pos_id = b.perm_id"
        DataToDgv(SelectFromTable(Sql), "permissions,employee", dgvEmployee)
        ClearTabPage(tp)
Finish:
        Cursor = Cursors.Default
    End Sub

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
                Case "customer"
                    '.Add("cus_name", txtCusName_cus.Text)
                    '.Add("cus_phone", txtPhone_cus.Text)
                    'If Not String.IsNullOrWhiteSpace(txtBirthday.Text) Then
                    '    Dim day As DateTime
                    '    If Not DateTime.TryParse(txtBirthday.Text, day) Then
                    '        MsgBox("生日日期格式錯誤")
                    '        txtBirthday.Focus()
                    '        GoTo Finish
                    '    End If
                    'End If
                    'If Not String.IsNullOrWhiteSpace(txtDueDate.Text) Then
                    '    Dim day As DateTime
                    '    If Not DateTime.TryParse(txtDueDate.Text, day) Then
                    '        MsgBox("預產期日期格式錯誤")
                    '        tcCustomer.SelectedTab = tpConsult_cus
                    '        txtDueDate.Focus()
                    '        GoTo Finish
                    '    End If
                    'End If
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
    ''' 去頭尾空白後,檢查不能空值的欄位
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="tp">TabPage</param>
    ''' <returns>True:是空的;False:有文字</returns>
    Private Function CheckTextNull(sTable As String, tp As TabPage) As Boolean
        '去頭尾空白
        tp.Controls.OfType(Of TextBox).ToList().ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '找出資料表不能為空值的欄位 
        Dim dt As DataTable = SelectFromTable($"SELECT COLUMN_COMMENT FROM information_schema.columns WHERE table_schema = 'tingyi' AND TABLE_NAME='{sTable}' AND is_nullable = 'NO' AND column_key != 'PRI'")

        '比較當前父控制項裡的txtbox.tag是否相符
        For Each txt As TextBox In tp.Controls.OfType(Of TextBox)()
            Dim row As DataRow = dt.AsEnumerable().FirstOrDefault(Function(x) x("COLUMN_COMMENT").ToString() = txt.Tag)
            If row IsNot Nothing Then
                If String.IsNullOrWhiteSpace(txt.Text) Then
                    MsgBox(txt.Tag + "不能空白")
                    txt.Focus()
                    Return True
                End If
            End If
        Next
        'todo combobox也要檢查
        Return False
    End Function

    '商品群組管理-修改
    Private Sub btnProdGrpModify_Click(sender As Object, e As EventArgs) Handles btnProdGrpModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"prod_grp_id = '{txtProdGrpID.Text}'")

        '列出所有資料
        DataToDgv(SelectFromTable($"SELECT * FROM {sTable}"), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        ClearTabPage(tp)
        InitProduct()
Finish:
        Cursor = Cursors.Default
    End Sub

    '商品管理-修改
    Private Sub btnProdModify_Click(sender As Object, e As EventArgs) Handles btnProdModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"prod_id  = '{txtProdID.Text}'")
        '列出所有資料
        DataToDgv(SelectFromTable("SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id;"), "product,product_group", dgvProduct)
        ClearTabPage(tp)
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
        DataToDgv(SelectFromTable("SELECT * FROM taboo"), "taboo", dgvTaboo)
        ClearTabPage(tp)
        InitTabooType()
Finish:
        Cursor = Cursors.Default
    End Sub

    '訂單管理-修改
    Private Sub btnOrdModify_Click(sender As Object, e As EventArgs) Handles btnOrdModify.Click
        Cursor = Cursors.WaitCursor
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        If CheckTextNull(sTable, tp) Then GoTo Finish
        UpdateData(sTable, Bind_TableTextBox(sTable), $"ord_id  = '{txtOrdID_order.Text}'")
        '列出所有資料
        Dim sql = "SELECT ord.ord_id, cus.cus_name, cus.cus_phone, ord.ord_date, prd.prod_name, pg.prod_grp_name, ord.ord_count, ord.ord_delivery, ord.ord_price, ord.ord_discount, ord.ord_breakfast, ord.ord_lunch,ord. ord_dinner, ord.ord_memo FROM `orders` ord LEFT JOIN customer cus ON ord.ord_cus_id=cus.cus_id LEFT JOIN product prd ON ord.ord_prod_id=prd.prod_id LEFT JOIN product_group pg ON prd.prod_prod_grp_id=pg.prod_grp_id"
        DataToDgv(SelectFromTable(sql), "customer,product,product_group,orders", dgvOrder)
        ClearTabPage(tp)
        InitProduct()
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
        Dim Sql = "SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_Income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id"
        DataToDgv(SelectFromTable(Sql), "customer,orders,money", dgvMoney)
        ClearTabPage(tp)
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
        Dim Sql = "SELECT * FROM permissions"
        DataToDgv(SelectFromTable(Sql), "permissions", dgvPermissions)
        ClearTabPage(tp)
        InitPosition()
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
        '列出所有資料
        Dim Sql = "SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_pos_id = b.perm_id"
        DataToDgv(SelectFromTable(Sql), "permissions,employee", dgvEmployee)
        ClearTabPage(tp)
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

            '顯示table所有資料
            DataToDgv(SelectFromTable($"SELECT * FROM {sTable}"), sTable, dgvProdgroup)
            ClearTabPage(tp)
            InitProduct()
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
            DataToDgv(SelectFromTable("SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id;"), "product,product_group", dgvProduct)
            ClearTabPage(tp)
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
            DataToDgv(SelectFromTable("SELECT * FROM taboo"), "taboo", dgvTaboo)
            ClearTabPage(tp)
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
            Dim Sql = "SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_Income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id"
            DataToDgv(SelectFromTable(Sql), "customer,orders,money", dgvMoney)
            ClearTabPage(tp)
        End If
    End Sub

    '訂單管理-刪除
    Private Sub btnOrdDelete_Click(sender As Object, e As EventArgs) Handles btnOrdDelete.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag = "訂單編號")
        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        Dim sTable As String = tp.Tag
        If DeleteData(sTable, $"ord_id  = '{id.Text}'") Then
            MsgBox("刪除成功")

            '顯示table所有資料
            Dim sql = "SELECT ord.ord_id, cus.cus_name, cus.cus_phone, ord.ord_date, prd.prod_name, pg.prod_grp_name, ord.ord_count, ord.ord_delivery, ord.ord_price, ord.ord_discount, ord.ord_breakfast, ord.ord_lunch,ord. ord_dinner, ord.ord_memo FROM `orders` ord LEFT JOIN customer cus ON ord.ord_cus_id=cus.cus_id LEFT JOIN product prd ON ord.ord_prod_id=prd.prod_id LEFT JOIN product_group pg ON prd.prod_prod_grp_id=pg.prod_grp_id"
            DataToDgv(SelectFromTable(sql), "customer,product,product_group,orders", dgvOrder)
            ClearTabPage(tp)
            InitProduct()
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
            Dim Sql = "SELECT * FROM permissions"
            DataToDgv(SelectFromTable(Sql), "permissions", dgvPermissions)
            ClearTabPage(tp)
            InitPosition()
        End If
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

            '顯示table所有資料
            Dim Sql = "SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_pos_id = b.perm_id"
            DataToDgv(SelectFromTable(Sql), "permissions,employee", dgvEmployee)
            ClearTabPage(tp)
        End If
    End Sub

    '商品群組管理-取消
    Private Sub btnProdGrpCancel_Click(sender As Object, e As EventArgs) Handles btnProdGrpCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        '顯示所有資料
        Dim sTable As String = tp.Tag.ToString
        DataToDgv(SelectFromTable($"SELECT * FROM {sTable}"), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        InitProduct()
    End Sub

    '商品管理-取消
    Private Sub btnProdCancel_Click(sender As Object, e As EventArgs) Handles btnProdCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        Dim sql = "SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id"
        '顯示所有資料
        DataToDgv(SelectFromTable(sql), "product", dgvProduct)
    End Sub

    '禁忌管理-取消
    Private Sub btnTaboCancel_Click(sender As Object, e As EventArgs) Handles btnTaboCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        Dim sql = "SELECT * FROM taboo"
        '顯示所有資料
        DataToDgv(SelectFromTable(sql), "taboo", dgvTaboo)
    End Sub

    '訂單管理-取消
    Private Sub btnOrdCancel_Click(sender As Object, e As EventArgs) Handles btnOrdCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        InitProduct()
        '顯示所有資料
        Dim sql = "SELECT ord.ord_id, cus.cus_name, cus.cus_phone, ord.ord_date, prd.prod_name, pg.prod_grp_name, ord.ord_count, ord.ord_delivery, ord.ord_price, ord.ord_discount, ord.ord_breakfast, ord.ord_lunch,ord. ord_dinner, ord.ord_memo FROM `orders` ord LEFT JOIN customer cus ON ord.ord_cus_id=cus.cus_id LEFT JOIN product prd ON ord.ord_prod_id=prd.prod_id LEFT JOIN product_group pg ON prd.prod_prod_grp_id=pg.prod_grp_id"
        DataToDgv(SelectFromTable(sql), "customer,product,product_group,orders", dgvOrder)
    End Sub

    '財務管理-取消
    Private Sub btnMonCancel_Click(sender As Object, e As EventArgs) Handles btnMonCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        InitProduct()
        '顯示所有資料
        Dim Sql = "SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_Income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id"
        DataToDgv(SelectFromTable(Sql), "customer,orders,money", dgvMoney)
    End Sub

    '權限管理-取消
    Private Sub btnPermCancel_Click(sender As Object, e As EventArgs) Handles btnPermCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '顯示所有資料
        Dim Sql = "SELECT * FROM permissions"
        DataToDgv(SelectFromTable(Sql), "permissions", dgvPermissions)
        ClearTabPage(tp)
        InitPosition()
    End Sub

    '員工管理-取消
    Private Sub btnEmpCancel_Click(sender As Object, e As EventArgs) Handles btnEmpCancel.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        '顯示所有資料
        Dim Sql = "SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_pos_id = b.perm_id"
        DataToDgv(SelectFromTable(Sql), "permissions,employee", dgvEmployee)
        ClearTabPage(tp)
    End Sub

    '商品群組管理-查詢
    Private Sub btnProdGrpQuery_Click(sender As Object, e As EventArgs) Handles btnProdGrpQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = tp.Tag
        Dim sql = $"SELECT * FROM {sTable} WHERE prod_grp_name LIKE '%{txtProdGrpName.Text}%'"
        DataToDgv(SelectFromTable(sql), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        ClearTabPage(tp)
        InitProduct()
        MsgBox("搜尋完畢")

        Cursor = Cursors.Default
    End Sub

    '商品管理-查詢
    Private Sub btnProdQuery_Click(sender As Object, e As EventArgs) Handles btnProdQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = "product,product_group"
        Dim sql = $"SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id = b.prod_grp_id WHERE a.prod_name LIKE '%{txtProdQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '禁忌管理-查詢
    Private Sub btnTaboQuery_Click(sender As Object, e As EventArgs) Handles btnTaboQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = "taboo"
        Dim sql = $"SELECT * FROM {sTable} WHERE tabo_type LIKE '%{txtTaboQuery.Text}%' OR tabo_name LIKE '%{txtTaboQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), sTable, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '訂單管理-查詢
    Private Sub btnOrderQuery_Click(sender As Object, e As EventArgs) Handles btnOrderQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = "taboo"
        Dim sql = $"SELECT ord.ord_id, cus.cus_name, cus.cus_phone, ord.ord_date, prd.prod_name, pg.prod_grp_name, ord.ord_count, ord.ord_delivery, ord.ord_price, ord.ord_discount, ord.ord_breakfast, ord.ord_lunch,ord. ord_dinner, ord.ord_memo FROM `orders` ord LEFT JOIN customer cus ON ord.ord_cus_id=cus.cus_id LEFT JOIN product prd ON ord.ord_prod_id=prd.prod_id LEFT JOIN product_group pg ON prd.prod_prod_grp_id=pg.prod_grp_id WHERE cus.cus_name LIKE '%{txtOrdQuery.Text}%' OR cus.cus_phone LIKE '%{txtOrdQuery.Text}%'"
        DataToDgv(SelectFromTable(sql), "customer,product,product_group,orders", dgvOrder)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '財務管理-查詢
    Private Sub btnMonQuery_Click(sender As Object, e As EventArgs) Handles btnMonQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = "taboo"

        Dim Sql = $"SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_Income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id WHERE b.cus_name LIKE '%{txtMonQuery.Text}%' OR b.cus_phone LIKE '%{txtMonQuery.Text}%'"
        DataToDgv(SelectFromTable(Sql), "customer,money,orders", dgvMoney)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '員工管理-查詢
    Private Sub btnEmpQuery_Click(sender As Object, e As EventArgs) Handles btnEmpQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = tp.Tag.ToString

        Dim Sql = $"SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_pos_id = b.perm_id WHERE emp_name LIKE '%{txtEmpQuery.Text}%' OR emp_phone LIKE '%{txtEmpQuery.Text}%'"
        DataToDgv(SelectFromTable(Sql), "permissions,employee", dgvEmployee)
        MsgBox("搜尋完畢")
        Cursor = Cursors.Default
    End Sub

    '訂單管理,選擇商品與數量後自動計算價格
    Private Sub txtCount_Leave(sender As Object, e As EventArgs) Handles txtCount.Leave
        If String.IsNullOrWhiteSpace(txtPrice_order.Text) Or String.IsNullOrWhiteSpace(txtCount.Text) Then Exit Sub
        txtPrice_order.Text = CInt(txtPrice_order.Text) * CInt(txtCount.Text)
    End Sub

    Private Sub cmbProdName_order_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbProdName_order.SelectionChangeCommitted
        Dim cmb As ComboBox = CType(sender, ComboBox)
        'todo 群組要跟商品名稱對應
        If cmb.SelectedIndex = -1 Then Exit Sub
        Dim selectedRow As DataRowView = DirectCast(cmb.SelectedItem, DataRowView)
        Dim prodId As Integer = CInt(selectedRow("prod_id"))
        Dim row = SelectFromTable($"SELECT prod_type, prod_price, prod_meal FROM product WHERE prod_id = {prodId}").Rows(0)
        '如果是套餐,顯示餐種供客製勾選
        If row("prod_type") = "套餐" Then
            grpMeal_order.Enabled = True
        Else
            grpMeal_order.Enabled = False
        End If
        '顯示商品價格
        txtPrice_order.Text = row("prod_price")

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

    Private Sub chkLunchAddr_CheckedChanged(sender As Object, e As EventArgs)
        If chkLunchAddr.Checked Then txtAddrLunch.Text = txtAddress.Text
    End Sub

    Private Sub chkDinnerAddr_CheckedChanged(sender As Object, e As EventArgs)
        If chkDinnerAddr.Checked Then txtAddrDinner.Text = txtAddress.Text
    End Sub

    Private Sub btnTaboo_Click(sender As Object, e As EventArgs)
        frmTaboo.Show()
    End Sub

    Private Sub btnTaboo_Click_1(sender As Object, e As EventArgs) Handles btnTaboo.Click
        frmTaboo.Show()
    End Sub

End Class
