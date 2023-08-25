Module modUtility
    '客戶管理
    Friend sqlCustomer As String = "SELECT cus_id, cus_name, cus_gender, cus_phone FROM customer"
    '系統設定-商品群組管理
    Friend sqlProductGroup As String = "SELECT * FROM product_group"
    '商品管理
    Friend sqlProduct As String = "SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id"
    '禁忌管理
    Friend sqlTaboo As String = "SELECT * FROM taboo"
    '訂單管理
    Friend sqlOrder As String = "SELECT a.ord_id, a.ord_date, b.cus_name, b.cus_phone, c.prod_name FROM orders a LEFT JOIN customer b ON a.ord_cus_id = b.cus_id LEFT JOIN product c ON a.ord_prod_id=c.prod_id"
    '財務管理
    Friend sqlMoney As String = "SELECT a.mon_id, c.cus_name, c.cus_phone, a.mon_ord_id, a.mon_date, a.mon_type, a.mon_income, a.mon_memo " +
                                "FROM money a " +
                                "LEFT JOIN orders b On a.mon_ord_id = b.ord_id " +
                                "LEFT JOIN customer c On b.ord_cus_id = c.cus_id"
    '權限管理
    Friend sqlPermision As String = "Select * FROM permissions"
    '員工管理
    Friend sqlEmployee As String = "Select a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo " +
                                   "FROM employee a " +
                                   "LEFT JOIN permissions b On a.emp_perm_id = b.perm_id"
    '配餐管理
    Friend sqlDistribute As String = "Select a.ord_id, a.ord_date, b.cus_name, b.cus_phone, c.prod_name, a.ord_memo" +
                                    " FROM orders a" +
                                    " LEFT JOIN customer b On a.ord_cus_id=b.cus_id" +
                                    " LEFT JOIN product c On a.ord_prod_id=c.prod_id" +
                                    " ORDER BY ord_date DESC"
    '菜單管理
    Friend sqlMenu As String = "Select DISTINCT b.prod_name,a.me_date,a.me_version,b.prod_id FROM menu a LEFT JOIN product b On a.me_prod_id=b.prod_id LIMIT 100"
    '配餐參數管理
    Friend sqlDistributeSystem As String = "Select * FROM distribute_system"

    '資料庫改變這也要改
    Public Enum Meal
        早餐 = 1
        午餐 = 2
        晚餐 = 3
        夜點 = 4
    End Enum

    '資料庫改變這也要改
    Public Enum Meal_Detail
        主食 = 1
        主菜 = 2
        半葷素 = 3
        青菜西飲 = 4
        湯品 = 5
        湯盅清補 = 6
        湯盅1期 = 7
        湯盅2期 = 8
        蔬菜1 = 9
        甜品 = 10
        水果 = 11
        飲品 = 12
        甜湯 = 13
        夜點 = 14
        湯盅3期 = 15
        湯盅4期 = 17
        蔬菜2 = 19

    End Enum

    Public Class Menu
        Public Property MenuDate As Date
        Public Property Version As String
        Public Property ProductName As String
        Public Property Meal As Integer
        Public Property Meal_Detail As Integer
        Public Property Name As String
    End Class

    '''' <summary>
    '''' 取得各產品的菜色在Excel菜單Cell的位置對應到Table的編號
    '''' </summary>
    '''' <param name="pmDetail">產品菜色名稱</param>
    '''' <returns></returns>
    'Public Function GetProductMealDetail(pmDetail As ProductMealDetail) As Dictionary(Of Integer, Integer)
    '    Dim dic As New Dictionary(Of Integer, Integer)
    '    With dic
    '        Select Case pmDetail
    '            Case 0 '月子早餐
    '                .Add(4, Meal_Detail.主食)
    '                .Add(5, Meal_Detail.主菜)
    '                .Add(6, Meal_Detail.半葷素)
    '                .Add(7, Meal_Detail.青菜西飲)
    '                .Add(8, Meal_Detail.湯品)
    '            Case 1 '月子午餐
    '                .Add(9, Meal_Detail.湯盅清補)
    '                .Add(10, Meal_Detail.湯盅1期)
    '                .Add(11, Meal_Detail.湯盅3期)
    '                .Add(12, Meal_Detail.主食)
    '                .Add(13, Meal_Detail.主菜)
    '                .Add(14, Meal_Detail.半葷素)
    '                .Add(15, Meal_Detail.蔬菜1)
    '                .Add(16, Meal_Detail.水果)
    '                .Add(17, Meal_Detail.甜品)
    '            Case 2 '月子晚餐
    '                .Add(18, Meal_Detail.湯盅清補)
    '                .Add(19, Meal_Detail.湯盅1期)
    '                .Add(20, Meal_Detail.湯盅3期)
    '                .Add(21, Meal_Detail.主食)
    '                .Add(22, Meal_Detail.主菜)
    '                .Add(23, Meal_Detail.半葷素)
    '                .Add(24, Meal_Detail.蔬菜1)
    '                .Add(25, Meal_Detail.水果)
    '            Case 3 '月子晚點
    '                .Add(26, Meal_Detail.湯盅清補)
    '                .Add(27, Meal_Detail.湯盅1期)
    '                .Add(28, Meal_Detail.湯盅3期)
    '            Case 4 '調理餐
    '                .Add(39, Meal_Detail.主食)
    '                .Add(40, Meal_Detail.主菜)
    '                .Add(41, Meal_Detail.半葷素)
    '                .Add(42, Meal_Detail.蔬菜1)
    '                .Add(43, Meal_Detail.蔬菜2)
    '                .Add(44, Meal_Detail.湯品)
    '                .Add(45, Meal_Detail.水果)
    '            Case 5 '幸福午餐
    '                .Add(50, Meal_Detail.主食)
    '                .Add(51, Meal_Detail.主菜)
    '                .Add(52, Meal_Detail.半葷素)
    '                .Add(53, Meal_Detail.蔬菜1)
    '                .Add(54, Meal_Detail.湯品)
    '            Case 6 '幸福晚餐
    '                .Add(56, Meal_Detail.主食)
    '                .Add(57, Meal_Detail.主菜)
    '                .Add(58, Meal_Detail.半葷素)
    '                .Add(59, Meal_Detail.蔬菜1)
    '                .Add(60, Meal_Detail.湯品)
    '            Case 7 '住院早餐
    '                .Add(65, Meal_Detail.主食)
    '                .Add(66, Meal_Detail.主菜)
    '                .Add(67, Meal_Detail.半葷素)
    '                .Add(68, Meal_Detail.蔬菜1)
    '                .Add(69, Meal_Detail.湯品)
    '                .Add(70, Meal_Detail.飲品)
    '            Case 8 '住院午餐
    '                .Add(71, Meal_Detail.主食)
    '                .Add(72, Meal_Detail.主菜)
    '                .Add(73, Meal_Detail.半葷素)
    '                .Add(74, Meal_Detail.蔬菜1)
    '                .Add(75, Meal_Detail.湯品)
    '                .Add(76, Meal_Detail.水果)
    '                .Add(77, Meal_Detail.飲品)
    '                .Add(78, Meal_Detail.甜湯)
    '            Case 9 '住院晚餐
    '                .Add(79, Meal_Detail.主食)
    '                .Add(80, Meal_Detail.主菜)
    '                .Add(81, Meal_Detail.半葷素)
    '                .Add(82, Meal_Detail.蔬菜1)
    '                .Add(83, Meal_Detail.湯品)
    '                .Add(84, Meal_Detail.飲品)
    '                .Add(85, Meal_Detail.夜點)
    '            Case 10 '輕食早餐
    '                .Add(89, Meal_Detail.主食)
    '                .Add(90, Meal_Detail.主菜)
    '                .Add(91, Meal_Detail.蔬菜1)
    '                .Add(92, Meal_Detail.蔬菜2)
    '                .Add(93, Meal_Detail.水果)
    '                .Add(94, Meal_Detail.飲品)
    '            Case 11 '輕食午餐
    '                .Add(96, Meal_Detail.主食)
    '                .Add(97, Meal_Detail.主菜)
    '                .Add(98, Meal_Detail.蔬菜1)
    '                .Add(99, Meal_Detail.蔬菜2)
    '                .Add(100, Meal_Detail.水果)
    '                .Add(101, Meal_Detail.飲品)
    '            Case 12 '輕食晚餐
    '                .Add(103, Meal_Detail.主食)
    '                .Add(104, Meal_Detail.主菜)
    '                .Add(105, Meal_Detail.蔬菜1)
    '                .Add(106, Meal_Detail.蔬菜2)
    '                .Add(107, Meal_Detail.水果)
    '                .Add(108, Meal_Detail.飲品)
    '            Case 13 '術後調理早餐
    '                .Add(4, Meal_Detail.主食)
    '                .Add(5, Meal_Detail.主菜)
    '                .Add(6, Meal_Detail.半葷素)
    '                .Add(7, Meal_Detail.青菜西飲)
    '                .Add(8, Meal_Detail.湯品)
    '            Case 14 '術後調理午餐
    '                .Add(11, Meal_Detail.主食)
    '                .Add(12, Meal_Detail.主菜)
    '                .Add(13, Meal_Detail.半葷素)
    '                .Add(14, Meal_Detail.蔬菜1)
    '                .Add(15, Meal_Detail.水果)
    '                .Add(9, Meal_Detail.湯盅清補)
    '            Case 15 '術後調理晚餐
    '                .Add(18, Meal_Detail.主食)
    '                .Add(19, Meal_Detail.主菜)
    '                .Add(20, Meal_Detail.半葷素)
    '                .Add(21, Meal_Detail.蔬菜1)
    '                .Add(22, Meal_Detail.水果)
    '                .Add(16, Meal_Detail.湯盅清補)
    '            Case 16 '素食早餐
    '                .Add(27, Meal_Detail.主食)
    '                .Add(28, Meal_Detail.主菜)
    '                .Add(29, Meal_Detail.半葷素)
    '                .Add(30, Meal_Detail.青菜西飲)
    '                .Add(31, Meal_Detail.湯品)
    '            Case 17 '素食午餐
    '                .Add(32, Meal_Detail.湯盅清補)
    '                .Add(33, Meal_Detail.湯盅2期)
    '                .Add(34, Meal_Detail.主食)
    '                .Add(35, Meal_Detail.主菜)
    '                .Add(36, Meal_Detail.半葷素)
    '                .Add(37, Meal_Detail.蔬菜1)
    '                .Add(38, Meal_Detail.甜品)
    '            Case 18 '素食晚餐
    '                .Add(39, Meal_Detail.湯盅清補)
    '                .Add(40, Meal_Detail.湯盅2期)
    '                .Add(41, Meal_Detail.主食)
    '                .Add(42, Meal_Detail.主菜)
    '                .Add(43, Meal_Detail.半葷素)
    '                .Add(44, Meal_Detail.蔬菜1)
    '                .Add(46, Meal_Detail.夜點)
    '            Case 19 '素食一般午餐
    '                .Add(53, Meal_Detail.主食)
    '                .Add(54, Meal_Detail.主菜)
    '                .Add(55, Meal_Detail.蔬菜1)
    '                .Add(56, Meal_Detail.蔬菜2)
    '                .Add(57, Meal_Detail.湯品)
    '            Case 20 '素食一般晚餐
    '                .Add(58, Meal_Detail.主食)
    '                .Add(59, Meal_Detail.主菜)
    '                .Add(60, Meal_Detail.蔬菜1)
    '                .Add(61, Meal_Detail.蔬菜2)
    '                .Add(62, Meal_Detail.湯品)
    '        End Select
    '    End With
    '    Return dic
    'End Function

    'Public Enum ProductMealDetail
    '    月子早餐
    '    月子午餐
    '    月子晚餐
    '    月子晚點
    '    調理餐
    '    幸福午餐
    '    幸福晚餐
    '    住院早餐
    '    住院午餐
    '    住院晚餐
    '    輕食早餐
    '    輕食午餐
    '    輕食晚餐
    '    術後調理早餐
    '    術後調理午餐
    '    術後調理晚餐
    '    素食早餐
    '    素食午餐
    '    素食晚餐
    '    素食一般午餐
    '    素食一般晚餐
    'End Enum

    ''' <summary>
    ''' 清空指定控制項內其他控制項
    ''' </summary>
    ''' <param name="ctrls">控制項的集合</param>
    Public Sub ClearControls(ctrls As Control, Optional exception As List(Of String) = Nothing)
        For Each ctrl As Control In ctrls.Controls
            If exception IsNot Nothing Then
                If exception.Contains(ctrls.Name) Or exception.Contains(ctrls.Text) Then Continue For
            End If
            If TypeOf ctrl Is GroupBox Then
                Dim grp = CType(ctrl, GroupBox)
                ClearControls(grp)
            ElseIf TypeOf ctrl Is TabControl Then
                For Each tp As TabPage In CType(ctrl, TabControl).Controls
                    ClearControls(ctrls)
                Next
            End If
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
            ElseIf TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Checked = False
            ElseIf TypeOf ctrl Is RadioButton Then
                CType(ctrl, RadioButton).Checked = False
            ElseIf TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).SelectedIndex = -1
            End If
        Next
    End Sub

    ''' <summary>
    ''' 清空TabPage裡的控制項內容
    ''' </summary>
    ''' <param name="tp"></param>
    <Obsolete("舊版 更新成 ClearControls")>
    Public Sub ClearTabPage(tp As TabPage)
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
    ''' 將取得的資料傳至各控制項(控制項的Tag必須寫上表格欄位名稱)
    ''' </summary>
    ''' <param name="ctrls">父容器</param>
    ''' <param name="row"></param>
    Public Sub GetDataToControls(ctrls As Control, row As Object)
        For Each ctrl In ctrls.Controls.Cast(Of Control).Where(Function(c) Not String.IsNullOrEmpty(c.Tag))
            Dim value = GetCellData(row, ctrl.Tag.ToString)
            Select Case ctrl.GetType.Name
                Case "TextBox"
                    ctrl.Text = value
                Case "DateTimePicker"
                    Dim dtp As DateTimePicker = ctrl
                    dtp.Value = value
                Case "ComboBox"
                    Dim cmb As ComboBox = ctrl
                    cmb.SelectedIndex = cmb.FindStringExact(value)
                Case "GroupBox"
                    Dim grp As GroupBox = ctrl
                    For Each c In grp.Controls
                        If TypeOf c Is CheckBox Then
                            Dim chk As CheckBox = c
                            Dim b As Boolean
                            If Boolean.TryParse(value, b) Then
                                chk.Checked = value
                            Else
                                chk.Checked = value.Contains(chk.Text)
                            End If
                        ElseIf TypeOf c Is RadioButton Then
                            Dim rdo As RadioButton = c
                            rdo.Checked = rdo.Text = value
                        End If
                    Next
                    GetDataToControls(ctrl, row)
                Case "CheckBox"
                    Dim chk As CheckBox = ctrl
                    If Boolean.Parse(value) Then
                        chk.Checked = value
                    Else
                        chk.Checked = value.Contains(chk.Text)
                    End If
                Case Else
            End Select
        Next
    End Sub

    Private Function GetCellData(row As Object, colName As String) As String
        Select Case row.GetType.Name
            Case "DataRow"
                Dim r As DataRow = row
                Return r(colName).ToString
            Case "DataGridViewRow"
                Dim r As DataGridViewRow = row
                Return r.Cells(colName).Value.ToString
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>
    ''' 去頭尾空白後,檢查必填的欄位
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <param name="txts">TextBox的集合</param>
    ''' <returns>True:是空的;False:有文字</returns>
    <Obsolete("這是舊版,改成 ChkRequiredCol")>
    Public Function CheckTextNull(sTable As String, txts As List(Of TextBox)) As Boolean
        '去頭尾空白
        txts.ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '找出資料表不能為空值的欄位註解名稱
        Dim dt As DataTable = SelectTable($"Select COLUMN_COMMENT FROM information_schema.columns WHERE table_schema = 'tingyi' AND TABLE_NAME='{sTable}' AND is_nullable = 'NO' AND column_key != 'PRI'")
        '比較與當前控制項.tag是否相符
        For Each ctrl As Windows.Forms.Control In txts
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
    ''' 檢查必填欄位
    ''' </summary>
    ''' <param name="ctrl">比對這個控制項裡的子控制項的tag</param>
    ''' <param name="required">填入key:Table欄位 value:中文名稱</param>
    ''' <returns></returns>
    <Obsolete("直接指明哪個控制項必填就好,不需要無謂的搜尋控制項")>
    Public Function CheckRequiredCol(ctrl As Control, required As Dictionary(Of String, String)) As Boolean
        For Each txt In ctrl.Controls.OfType(Of TextBox)().Where(Function(x) required.Keys.Contains(If(x.Tag, "")) AndAlso String.IsNullOrWhiteSpace(x.Text))
            MsgBox(required(txt.Tag) + " 不能空白")
            txt.Focus()
            Return False
        Next
        For Each txt In ctrl.Controls.OfType(Of ComboBox)().Where(Function(x) required.Keys.Contains(If(x.Tag, "")) AndAlso String.IsNullOrWhiteSpace(x.Text))
            MsgBox(required(txt.Tag) + " 不能空白")
            txt.Focus()
            Return False
        Next
        Return True
    End Function
    ''' <summary>
    ''' 檢查必填欄位
    ''' </summary>
    ''' <param name="required">填入key:欄位名稱 value:控制項</param>
    ''' <returns></returns>
    Public Function CheckRequiredCol(required As Dictionary(Of String, Object)) As Boolean
        For Each kvp In required
            If String.IsNullOrWhiteSpace(kvp.Value.Text) Then
                MsgBox(kvp.Key + " 不能空白")
                kvp.Value.Focus()
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' Insert前檢查
    ''' </summary>
    ''' <param name="sTable">資料表</param>
    ''' <returns></returns>
    Public Function CheckInsert(sTable As String, tp As TabPage) As Boolean
        Dim bResult As Boolean
        If CheckTextNull(sTable, tp.Controls.OfType(Of TextBox).ToList) Then GoTo Finish

        '不可重複的欄位
        Dim dic As New Dictionary(Of String, String)
        With dic
            Select Case sTable
                Case "product_group"
                    .Add("prod_grp_name", frmMain.txtName_prod_grp.Text)
                Case "product"
                    .Add("prod_name", frmMain.txtProdName.Text)
                Case "taboo"
                    .Add("tabo_name", frmMain.txtTaboName.Text)
                Case Else
                    GoTo Pass
            End Select
        End With
        Dim lst As List(Of String) = dic.Select(Function(x) $"{x.Key} = '{x.Value}'").ToList
        Dim sWhere = String.Join(" AND ", lst)
        Dim dgv = tp.Controls.OfType(Of DataGridView).FirstOrDefault
        'If CheckDuplication(sTable, sWhere, dgv) Then GoTo Finish
Pass:
        bResult = True
Finish:
        Return bResult
    End Function

    ''' <summary>
    ''' 檢查是否重複新增
    ''' </summary>
    ''' <param name="selectFrom">SQL前半段</param>
    ''' <param name="list">條件,輸入控制項會自動取得Tag(欄位名稱),Text(值)</param>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    Public Function CheckDuplication(selectFrom As String, list As List(Of Object), dgv As DataGridView) As Boolean
        Dim sql = selectFrom + $" WHERE {String.Join(" AND ", list.Select(Function(x) $"{x.tag} = '{x.text}'"))}"
        If SelectTable(sql).Rows.Count > 0 Then
            MsgBox("重複資料")
            '列出重複的資料
            DataToDgv(sql, dgv)
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 將資料放到DataGridView
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="table"></param>
    ''' <param name="dgv"></param>
    <Obsolete("舊版 請用 DataToDgv(sql As String, dgv As DataGridView)")>
    Public Sub DataToDgv(dt As DataTable, table As String, dgv As DataGridView)
        With dgv
            .DataSource = dt
            '用table欄位的備註將dgv的欄位改名
            Dim conditions As String = String.Join(" Or ", table.Split(","c).Select(Function(x) $"Table_name = '{x.Trim()}'"))
            Dim tableCol As DataTable = SelectTable($"SELECT COLUMN_NAME, COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS  WHERE TABLE_SCHEMA = 'tingyi' AND {conditions}")
            For Each col As DataGridViewColumn In .Columns
                Dim row As DataRow = tableCol.AsEnumerable().FirstOrDefault(Function(x) x("COLUMN_NAME").ToString() = col.Name)
                If row IsNot Nothing Then
                    col.HeaderText = row("COLUMN_COMMENT").ToString()
                End If
            Next
            .AutoResizeColumnHeadersHeight()
        End With
    End Sub
    ''' <summary>
    ''' 將資料放到DataGridView
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="table"></param>
    ''' <param name="dgv"></param>
    Public Sub DataToDgv(sql As String, dgv As DataGridView)
        With dgv
            .DataSource = SelectTable(sql)
            Dim lstTableNames = GetTableNamesFromQuery(sql)
            '條件式
            Dim conditions = String.Join(" OR ", lstTableNames.Select(Function(x) $"Table_name = '{x}'"))
            '用table欄位的備註將dgv的欄位改名
            Dim tableCol = SelectTable($"SELECT COLUMN_NAME, COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'tingyi' AND {conditions}")
            For Each col As DataGridViewColumn In .Columns
                Dim row = tableCol.AsEnumerable().FirstOrDefault(Function(x) x("COLUMN_NAME").ToString() = col.Name)
                If row IsNot Nothing Then
                    col.HeaderText = row("COLUMN_COMMENT").ToString()
                End If
            Next
            .AutoResizeColumnHeadersHeight()
        End With
    End Sub

    ''' <summary>
    ''' 清除鍵共用功能
    ''' </summary>
    ''' <param name="btn"></param>
    Public Sub BtnCancel(btn As Button, sql As String, dgv As DataGridView)
        Dim tp As TabPage = btn.Parent
        ClearControls(tp)
        DataToDgv(sql, dgv)
    End Sub

    ''' <summary>
    ''' 新增鍵共用功能(Table名稱要存在TabPage.Tag)
    ''' </summary>
    ''' <param name="btn">"新增"按鈕</param>
    ''' <param name="id">用來判斷是不是新資料(有insert過就會有id)</param>
    ''' <param name="required">必填欄位 key:欄位中文名稱 value:TextBox</param>
    ''' <returns></returns>
    Public Function BtnInsert(btn As Button, id As TextBox, Optional required As Dictionary(Of String, Object) = Nothing) As Boolean
        '判斷是否可以新增
        If Not String.IsNullOrEmpty(id.Text) Then Return False

        If required IsNot Nothing Then
            If Not CheckRequiredCol(required) Then Return False
        End If
        Dim tp As TabPage = btn.Parent
        Dim table = tp.Tag.ToString
        If Not InserTable(table, frmMain.BindData(table)) Then Return False
        '刷新
        tp.Controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
        Return True
    End Function

    ''' <summary>
    ''' 修改鍵共用功能(Table名稱要存在TabPage.Tag)
    ''' </summary>
    ''' <param name="btn"></param>
    ''' <param name="id">用來判斷是不是新資料(有insert過就會有id)</param>
    ''' <param name="condition">條件式(xxx=xxx)</param>
    ''' <param name="required">必填欄位 key:欄位中文名稱 value:TextBox</param>
    ''' <returns></returns>
    Public Function BtnModify(btn As Button, id As Control, condition As String, Optional required As Dictionary(Of String, Object) = Nothing)
        '判斷是否可以修改
        If String.IsNullOrEmpty(id.Text) Then Return False

        If required IsNot Nothing Then
            If Not CheckRequiredCol(required) Then Return False
        End If
        '該TabPage裡的TextBox文字去頭尾空白
        Dim tp As TabPage = btn.Parent
        tp.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        Dim table = tp.Tag.ToString
        If Not UpdateTable(table, frmMain.BindData(table), condition) Then Return False
        '刷新
        tp.Controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
        Return True
    End Function

    ''' <summary>
    ''' 刪除鍵共用功能(Table名稱要存在TabPage.Tag)
    ''' </summary>
    ''' <param name="btn"></param>
    ''' <param name="id">用來判斷是不是新資料(有insert過就會有id)</param>
    ''' <param name="condition">條件式(xxx=xxx)</param>
    ''' <returns></returns>
    Public Function BtnDelete(btn As Button, id As Control, condition As String) As Boolean
        '判斷是否可以刪除
        If String.IsNullOrEmpty(id.Text) Then Return False

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Return False
        Dim tp = btn.Parent
        If Not DeleteData(tp.Tag, condition) Then Return False
        tp.Controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
        Return True
    End Function

    ''' <summary>
    ''' DataGridVeiw CellMouseClick共用功能
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <returns></returns>
    Public Function DGVCellMouseClick(dgv As DataGridView) As Boolean
        If dgv.SelectedRows.Count <> 1 Then Return False
        Dim tp = dgv.Parent
        ClearControls(tp)
        Dim row = dgv.SelectedRows(0)
        GetDataToControls(tp, row)
        Return True
    End Function

    ''' <summary>
    ''' 設定DataGridView的樣式屬性
    ''' </summary>
    ''' <param name="ctrl">父容器</param>
    Public Sub SetDataGridViewStyle(ctrl As Control)
        For Each dgv In GetControlInParent(Of DataGridView)(ctrl)
            With dgv
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .ColumnHeadersDefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
                .DefaultCellStyle.Font = New Font("標楷體", 12, FontStyle.Bold)
                .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(224, 224, 224)
                .EnableHeadersVisualStyles = False
                .ColumnHeadersDefaultCellStyle.BackColor = Color.MediumTurquoise
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .ReadOnly = True
                .AllowUserToResizeColumns = True
            End With
        Next
    End Sub

    ''' <summary>
    ''' 取得指定控制項內所有的目標控制項
    ''' </summary>
    ''' <typeparam name="T">目標控制項</typeparam>
    ''' <param name="parent">父控制項</param>
    ''' <returns></returns>
    Public Function GetControlInParent(Of T As Control)(parent As Control) As List(Of T)
        Dim lst As New List(Of T)
        If parent.Controls.Count > 0 Then
            For Each ctrl In parent.Controls
                If TypeOf ctrl Is T Then lst.Add(ctrl)
                lst.AddRange(GetControlInParent(Of T)(ctrl))
            Next
        End If
        Return lst
    End Function
End Module
