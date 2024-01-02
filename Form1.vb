Imports System.IO
Imports System.Windows
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop
Imports Path = System.IO.Path
Imports CheckBox = System.Windows.Forms.CheckBox
Imports Control = System.Windows.Forms.Control

Public Class frmMain
    Private tempDistDay As Button '配餐管理月曆所選日期暫存
    Private dragRowIndex As Integer ' 儲存被拖曳的行的索引
    Private rowBackColor As System.Drawing.Color '紀錄被拖曳行的背景顏色

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedTab.Name = "TP_Logout" Then
            If MessageBox.Show("確定要登出嗎??", "登出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.OK Then
                Close()
            Else
                TabControl1.SelectedTab = tpCustomer
            End If
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '自定義索引標籤、文字顏色
        TabControl1.DrawMode = DrawMode.OwnerDrawFixed
        tcCustomer.DrawMode = DrawMode.OwnerDrawFixed
        tcSystem.DrawMode = DrawMode.OwnerDrawFixed

        SetDataGridViewStyle(Me)

        InitDataGrid()
        InitcmbProductGroup()
        InitSales()
        InitPosition()
        InitDistribute()
        InitcmbProduct()

        '初始化各TabPage
        btnCancel_dish_Click(btnCancel_dish, e)
        btnCancel_money_ord_Click(btnCancel_money_ord, EventArgs.Empty)
        btnCancel_taboo_group_Click(btnCancel_taboo_group, EventArgs.Empty)
        btnPermCancel_Click(btnPermCancel, EventArgs.Empty)

        ''初始化配餐管理 月曆
        txtDistCalendar.Text = Date.Now.ToString("Y")

        ''初始化菜單版本
        With cmbProdVers_menu
            Dim arr() As String = {"A", "B", "C", "D"}
            .DataSource = arr
            .SelectedIndex = -1
        End With

    End Sub

    ''' <summary>
    ''' 初始化職位ComboBox
    ''' </summary>
    Private Sub InitPosition()
        With cmbPosition_emp
            .DataSource = SelectTable("SELECT perm_name, perm_id FROM permissions").Copy
            .DisplayMember = "perm_name"
            .ValueMember = "perm_id"
            .SelectedIndex = -1
        End With
    End Sub

    Private Sub frmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        LoginForm1.Show()
    End Sub

    '自定義TabPage索引標籤、文字顏色
    Private Sub TabControl_DrawItem_MediumVioletRed(sender As Object, e As DrawItemEventArgs) Handles TabControl1.DrawItem, tcCustomer.DrawItem, tcSystem.DrawItem
        Dim tabControl As TabControl = DirectCast(sender, TabControl)
        Dim tab As TabPage = tabControl.TabPages(e.Index)

        ' 檢查當前索引標籤是否為選中狀態
        Dim isSelected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected

        ' 繪製索引標籤的背景
        Dim backColor As System.Drawing.Color = If(isSelected, System.Drawing.Color.ForestGreen, System.Drawing.Color.WhiteSmoke)
        e.Graphics.FillRectangle(New SolidBrush(backColor), e.Bounds)

        ' 繪製索引標籤的文字
        Dim text As String = tab.Text
        Dim textColor As System.Drawing.Color = If(isSelected, System.Drawing.Color.White, System.Drawing.Color.Black)
        Dim font As System.Drawing.Font = tabControl.Font
        e.Graphics.DrawString(text, font, New SolidBrush(textColor), e.Bounds.Location)
    End Sub

    Private Sub TabControl_DrawItem_GreenYellow(sender As Object, e As DrawItemEventArgs) Handles tcCustomer.DrawItem, tcSystem.DrawItem
        Dim tabControl As TabControl = DirectCast(sender, TabControl)
        Dim tab As TabPage = tabControl.TabPages(e.Index)

        ' 檢查當前索引標籤是否為選中狀態
        Dim isSelected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected

        ' 繪製索引標籤的背景
        Dim backColor As System.Drawing.Color = If(isSelected, System.Drawing.Color.GreenYellow, System.Drawing.Color.WhiteSmoke)
        e.Graphics.FillRectangle(New SolidBrush(backColor), e.Bounds)

        ' 繪製索引標籤的文字
        Dim text As String = tab.Text
        'Dim textColor As System.Drawing.Color = If(isSelected, System.Drawing.Color.White, System.Drawing.Color.Black)
        Dim textColor As System.Drawing.Color = System.Drawing.Color.Black
        Dim font As System.Drawing.Font = tabControl.Font
        e.Graphics.DrawString(text, font, New SolidBrush(textColor), e.Bounds.Location)
    End Sub

    ''' <summary>
    ''' 初始化商品群組的ComboBox
    ''' </summary>
    Private Sub InitcmbProductGroup()
        Dim list As New List(Of ComboBox) From {cmbProdGrp_order, cmbProdGrp_product}

        For Each cmb In list
            With cmb
                .DataSource = SelectTable("SELECT prod_grp_name, prod_grp_id FROM product_group").Copy
                .DisplayMember = "prod_grp_name"
                .ValueMember = "prod_grp_id"
                .SelectedIndex = -1
            End With
        Next
    End Sub

    Private Sub InitcmbProduct()
        '初始化商品
        With cmbProdName_menu
            .DataSource = SelectTable("SELECT * FROM product WHERE prod_type = '套餐'").Copy
            .DisplayMember = "prod_name"
            .ValueMember = "prod_id"
            .SelectedIndex = -1
        End With
    End Sub

    ''' <summary>
    ''' 初始化業務人員
    ''' </summary>
    Private Sub InitSales()
        With cmbSales
            .DataSource = SelectTable("SELECT a.emp_name,a.emp_id FROM employee a LEFT JOIN permissions b ON a.emp_perm_id=b.perm_id WHERE perm_name = '業務'")
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
        DataToDgv(sqlCustomer, dgvCustomer)

        '系統設定-商品群組管理
        DataToDgv(sqlProductGroup, dgvProdgroup)

        '商品管理        
        DataToDgv(sqlProduct, dgvProduct)

        '訂單管理
        DataToDgv(SelectTable(sqlOrder), "customer,orders,product", dgvOrder)

        '權限管理
        DataToDgv(SelectTable(sqlPermision), "permissions", dgvPermissions)

        '員工管理        
        DataToDgv(SelectTable(sqlEmployee), "permissions,employee", dgvEmployee)

        '配餐管理        
        DataToDgv(SelectTable(sqlDistribute), "distribute,orders,customer,product", dgvDist)

        '菜單管理
        DataToDgv(SelectTable(sqlMenu), "menu,product", dgvMenu)

        '配餐系統管理
        DataToDgv(SelectTable(sqlDistributeSystem), "distribute_system", dgvDistSys)
        dgvDist.Columns("ord_memo").Visible = False

        '報表管理-送餐報表
        ' 添加拖曳事件處理函式
        AddHandler dgvDelivery_new.MouseDown, AddressOf dgv_MouseDown
        AddHandler dgvDelivery_new.DragOver, AddressOf dgv1_DragOver
        AddHandler dgvDelivery_new.DragDrop, AddressOf dgv_DragDrop
        dgvDelivery_new.Columns("路線").ReadOnly = False
    End Sub

    ' 開始拖曳時觸發
    Private Sub dgv_MouseDown(sender As Object, e As MouseEventArgs)
        Dim hit As DataGridView.HitTestInfo = dgvDelivery_new.HitTest(e.X, e.Y)

        If hit.RowIndex >= 0 AndAlso e.Button = MouseButtons.Left Then
            dragRowIndex = hit.RowIndex
            rowBackColor = dgvDelivery_new.Rows(dragRowIndex).DefaultCellStyle.BackColor
            dgvDelivery_new.DoDragDrop(dgvDelivery_new.Rows(dragRowIndex), DragDropEffects.Move)
        End If

    End Sub

    ' 拖曳時觸發
    Private Sub dgv1_DragOver(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.Move
    End Sub

    ' 放下時觸發
    Private Sub dgv_DragDrop(sender As Object, e As DragEventArgs)
        Dim targetDgv As DataGridView = DirectCast(sender, DataGridView)
        Dim dropPoint As System.Drawing.Point = targetDgv.PointToClient(New System.Drawing.Point(e.X, e.Y))
        Dim hit As DataGridView.HitTestInfo = targetDgv.HitTest(dropPoint.X, dropPoint.Y)

        If hit.RowIndex >= 0 Then
            ' 移動行到目標位置
            Dim targetRowIndex As Integer = hit.RowIndex
            Dim draggedRow As DataGridViewRow = TryCast(e.Data.GetData(GetType(DataGridViewRow)), DataGridViewRow)

            If Not IsNothing(draggedRow) Then
                ' 在目標位置插入拖曳的行
                targetDgv.Rows.Insert(targetRowIndex, draggedRow.Clone())
                For Each cell As DataGridViewCell In draggedRow.Cells
                    targetDgv.Rows(targetRowIndex).Cells(cell.ColumnIndex).Value = cell.Value
                Next

                ' 移除原來位置的行
                If targetRowIndex > dragRowIndex Then
                    dgvDelivery_new.Rows.RemoveAt(dragRowIndex)
                Else
                    dgvDelivery_new.Rows.RemoveAt(dragRowIndex + 1)
                End If

                ' 在拖放完成後還原原始行的背景顏色
                targetDgv.Rows(targetRowIndex).DefaultCellStyle.BackColor = rowBackColor
            End If
        End If
    End Sub

    ''' <summary>
    ''' 動態初始化配餐選項
    ''' </summary>
    Private Sub InitDistribute()
        flpDist.Visible = False

        Dim dt = SelectTable("SELECT * FROM distribute_system")

        For Each grp In flpDist.Controls.OfType(Of GroupBox)
            Dim flp = grp.Controls.OfType(Of FlowLayoutPanel).FirstOrDefault

            If flp IsNot Nothing Then flp.Controls.Clear()

            Dim row = dt.Select($"dist_sys_name = '{grp.Text}'")
            Dim options = row.First.Field(Of String)("dist_sys_option")
            Dim type = row.First.Field(Of String)("dist_sys_type")

            For Each txt In Split(options, ",")
                If txt = "" Then Exit For

                If type = "單選" Then
                    grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.Add(New RadioButton With {.Text = txt, .AutoSize = True})

                ElseIf type = "多選" Then
                    Dim chk As New Forms.CheckBox With {.Text = txt, .AutoSize = True}

                    If txt = "最後一餐" Then
                        AddHandler chk.CheckedChanged, AddressOf LastMeal
                    End If

                    grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.Add(chk)
                End If
            Next
        Next

        flpDist.Visible = True
    End Sub

    ''' <summary>
    ''' 勾選最後餐,則一併勾選免洗餐具
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LastMeal(sender, e)
        Dim check = CType(sender, CheckBox)
        flpDist.Controls.OfType(Of GroupBox).Where(Function(g) g.Text = "餐具").FirstOrDefault.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of Forms.CheckBox).Where(Function(chk) chk.Text = "免洗餐具").First.Checked = check.Checked
    End Sub

    '財務管理-取消
    Private Sub btnCancel_money_Click(sender As Object, e As EventArgs) Handles btnCancel_money.Click
        ClearControls(tpMoney, {"grpOrder_money"}.ToList())
        DataToDgv($"SELECT * FROM money WHERE mon_ord_id = '{txtOrderID_money.Text}'", dgvMoney)
    End Sub

    '財務管理-訂單資訊-dgv點擊
    Private Sub dgvOrder_money_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvOrder_money.CellMouseClick
        ClearControls(tpMoney)
        DGVCellMouseClick(dgvOrder_money)
        DataToDgv($"SELECT * FROM money WHERE mon_ord_id = '{txtOrderID_money.Text}'", dgvMoney)
        txtSum.Text = Int(txtPrice_money.Text) + Int(txtTableware_money.Text) + Int(txtTaste_money.Text) - Int(txtDiscount_money.Text) + Int(txtFreight.Text) - Int(txtSpread.Text)
        UpdateNoPay()

    End Sub

    Private Sub UpdateNoPay()
        '計算未繳金額
        Dim totalPay As Integer '已繳金額
        If dgvMoney.RowCount = 0 Then
            totalPay = 0
        Else
            For Each row As DataGridViewRow In dgvMoney.Rows
                totalPay += row.Cells("mon_income").Value
            Next
        End If

        txtNoPay.Text = Int(txtSum.Text) - totalPay
    End Sub

    '財務管理-新增
    Private Sub btnMonInsert_Click(sender As Object, e As EventArgs) Handles btnInsert_money.Click
        Dim btn As Button = sender
        Dim dic = CheckMoney(btn)
        If dic Is Nothing OrElse dic.Count = 0 Then Exit Sub

        Dim tp = btn.Parent
        If InserTable("money", dic) Then
            '刷新
            tp.Controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("新增成功")
            UpdateNoPay()
        End If
    End Sub

    Private Function CheckMoney(sender As Button) As Dictionary(Of String, Object)
        Dim dicReq As New Dictionary(Of String, Object) From {
            {"收款金額", txtMoney},
            {"收款類型", cmbMonType},
            {"付款類型", cmbPayType}
         }
        If Not CheckRequiredCol(dicReq) Then Return Nothing

        Dim tp As TabPage = sender.Parent

        Dim dic As New Dictionary(Of String, Object)
        dic = tp.Controls.OfType(Of TextBox).Where(Function(ctrl) ctrl.Tag IsNot Nothing AndAlso Not ctrl.Tag.ToString.Contains("id") AndAlso
                                                             Not String.IsNullOrWhiteSpace(ctrl.Text)).
            ToDictionary(Function(ctrl) ctrl.Tag.ToString, Function(ctrl) CType(ctrl.Text, Object))
        dic.Add("mon_type", cmbMonType.Text)
        dic.Add("mon_pay_type", cmbPayType.Text)
        dic.Add("mon_date", dtMonDate.Value.ToString("yyyy/MM/dd"))
        dic.Add("mon_ord_id", txtOrderID_money.Text)

        Return dic
    End Function

    '財務管理-dgv點擊
    Private Sub dgvMoney_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvMoney.CellMouseClick
        Dim dgv As DataGridView = sender
        If dgv.SelectedRows.Count <> 1 Then Exit Sub
        Dim tp = dgv.Parent
        Dim list = New List(Of String) From {grpOrder_money.Text}
        ClearControls(tp, list)
        GetDataToControls(tp, dgv.SelectedRows(0))
    End Sub

    '財務管理-修改
    Private Sub btnMonModify_Click(sender As Object, e As EventArgs) Handles btnModify_money.Click
        If String.IsNullOrEmpty(txtID_money.Text) Then
            MsgBox("請先選擇對象")
            Exit Sub
        End If

        Dim tp = sender.Parent
        Dim dic = CheckMoney(sender)

        If dic.Count = 0 Then Exit Sub

        If UpdateTable("money", dic, $"{txtID_money.Tag} = '{txtID_money.Text}'") Then
            btnCancel_money.PerformClick()
            MsgBox("修改成功")
            UpdateNoPay()
        End If
    End Sub

    '財務管理-刪除
    Private Sub btnMonDel_Click(sender As Object, e As EventArgs) Handles btnDel_money.Click
        '判斷是否可以刪除
        If String.IsNullOrEmpty(txtID_money.Text) Then
            MsgBox("請選擇對象")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        If DeleteData("money", $"{txtID_money.Tag} = '{txtID_money.Text}'") Then
            btnCancel_money.PerformClick()
            MsgBox("刪除成功")
            UpdateNoPay()
        End If
    End Sub

    '財務管理-查詢
    Private Sub btnMonQuery_Click(sender As Object, e As EventArgs) Handles btnQuery_money.Click
        Dim btn As Button = sender
        ClearControls(btn)
        Dim Sql = sqlMoney & $" WHERE c.cus_name LIKE '%{txtQuery_money.Text}%' OR c.cus_phone LIKE '%{txtQuery_money.Text}%'"
        DataToDgv(Sql, dgvOrder_money)
    End Sub

    '財務管理-訂單資訊-訂單編號-連動資料
    Private Sub txtOrderID_money_TextChanged(sender As Object, e As EventArgs) Handles txtOrderID_money.TextChanged
        If String.IsNullOrEmpty(txtOrderID_money.Text) Then Exit Sub

        DataToControl_Order(txtOrderID_money.Text)
    End Sub

    Private Sub btnModify_order_money_Click(sender As Object, e As EventArgs) Handles btnModify_order_money.Click
        Dim dicOrder = grpOrder_money.Controls.OfType(Of TextBox).Where(Function(x) Not String.IsNullOrEmpty(x.Tag) AndAlso x.Tag.ToString.Contains("ord_")).
            ToDictionary(Of String, Object)(Function(x) x.Tag, Function(x) x.Text)

        If UpdateTable("orders", dicOrder, $"{txtOrderID_money.Tag} = '{txtOrderID_money.Text}'") Then
            btnCancel_money_ord.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    '訂單管理-數量-計算總金額
    Private Sub txtTotalPrice_TextChanged(sender As Object, e As EventArgs) Handles txtCount.TextChanged, txtPrice.TextChanged
        If String.IsNullOrWhiteSpace(txtCount.Text) OrElse String.IsNullOrWhiteSpace(txtPrice.Text) Then Exit Sub
        Dim price As Integer = txtPrice.Text
        Dim count As Integer = txtCount.Text
        txtTotalPrice.Text = price * count
    End Sub

    '訂單管理-取消
    Private Sub btnCancel_order_Click(sender As Object, e As EventArgs) Handles btnCancel_order.Click
        Dim btn As Button = sender
        ClearControls(btn.Parent)
        DataToDgv(sqlOrder, dgvOrder)
        InitSales()
        InitcmbProductGroup()
    End Sub

    '訂單管理-新增
    Private Sub btnOrdInsert_Click(sender As Object, e As EventArgs) Handles btnOrdInsert.Click
        If Not String.IsNullOrEmpty(txtOrdID_order.Text) Then
            MsgBox("請先按取消再新增")
            Exit Sub
        End If

        Dim btn As Button = sender
        Dim dic = CheckOrder(btn)
        If dic Is Nothing OrElse dic.Count = 0 Then Exit Sub

        Dim tp = btn.Parent
        If InserTable("orders", dic) Then
            '刷新
            tp.Controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            btnDistCancel_Click(btnCancel_dist, EventArgs.Empty)
            btnCancel_money_Click(btnCancel_money, EventArgs.Empty)
            MsgBox("新增成功")
        End If
    End Sub

    '訂單管理-修改
    Private Sub btnOrdModify_Click(sender As Object, e As EventArgs) Handles btnOrdModify.Click
        If String.IsNullOrEmpty(txtOrdID_order.Text) Then
            MsgBox("請先選擇對象")
            Exit Sub
        End If

        Dim btn As Button = sender
        Dim tp = btn.Parent
        Dim table = "orders"
        Dim dic = CheckOrder(btn)

        If dic Is Nothing OrElse dic.Count = 0 Then Exit Sub

        If UpdateTable(table, dic, $"{txtOrdID_order.Tag} = '{txtOrdID_order.Text}'") Then
            tp.Controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            btnDistCancel_Click(btnCancel_dist, EventArgs.Empty)
            btnCancel_money_Click(btnCancel_money, EventArgs.Empty)
            MsgBox("修改成功")
        End If
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

        Dim sTable As String = "orders"
        If DeleteData(sTable, $"ord_id  = '{id.Text}'") Then
            btnCancel_order.PerformClick()
            btnDistCancel_Click(btnCancel_dist, EventArgs.Empty)
            btnCancel_money_Click(btnCancel_money, EventArgs.Empty)
            MsgBox("刪除成功")
        End If
    End Sub

    ''' <summary>
    ''' 檢查Orders必填欄位,取得上傳資料
    ''' </summary>
    ''' <returns>True:正確 False:錯誤</returns>
    Private Function CheckOrder(btn As Button) As Dictionary(Of String, Object)
        Dim dicData As New Dictionary(Of String, Object)
        Dim dicReq As New Dictionary(Of String, Object) From {
            {"商品群組", cmbProdGrp_order},
            {"商品名稱", cmbProdName_order},
            {"數量", txtCount}
        }
        If Not CheckRequiredCol(dicReq) Then Return dicData

        '取得客戶
        Dim cusID As Integer

        If String.IsNullOrEmpty(txtCusID_ord.Text) Then
            Dim dic As New Dictionary(Of String, Object) From {
                {"cus_phone", txtPhone_order.Text}
            }

            Try
                cusID = SelectTable($"SELECT cus_id FROM customer WHERE cus_phone = @cus_phone", dic).Rows(0).Field(Of Integer)("cus_id")
            Catch ex As Exception
                MsgBox("查無此客戶")
                Return Nothing
            End Try
        Else
            cusID = txtCusID_ord.Text

        End If

#Region "取得資料"

        With dicData
            Dim ctrlOrder = tpOrder.Controls

            .Add("ord_cus_id", cusID) '客戶編號
            .Add("ord_prod_id", cmbProdName_order.SelectedValue) '商品編號
            ctrlOrder.OfType(Of DateTimePicker).Where(Function(dtp) String.IsNullOrEmpty(dtp.Tag) = False).ToList.ForEach(Sub(d) .Add(d.Tag, d.Value.ToString))
            ctrlOrder.OfType(Of TextBox) _
                .Where(Function(x) Not String.IsNullOrEmpty(x.Tag) AndAlso Not String.IsNullOrEmpty(x.Text) AndAlso x.Tag.Contains("ord_")).ToList.ForEach(Sub(x) .Add(x.Tag, x.Text))

            Dim rdo = grpEatType.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)

            If rdo IsNot Nothing Then .Add(grpEatType.Tag, rdo.Text) '葷素
            If cmbSales.SelectedIndex <> -1 Then .Add("ord_emp_id", cmbSales.SelectedValue) '業務人員

            '餐種數量
            .Add("ord_breakfast", txtCount.Text)
            .Add("ord_lunch", txtCount.Text)
            .Add("ord_dinner", txtCount.Text)

        End With

#End Region

        Return dicData

    End Function

    '訂單管理-dgv點擊
    Private Sub dgvOrder_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvOrder.CellMouseClick
        Dim dgv As DataGridView = sender
        Dim row = dgv.SelectedRows(0)
        DataToControl_Order(row.Cells("ord_id").Value)
    End Sub

    ''' <summary>
    ''' 將資料傳至訂單管理個控制項
    ''' </summary>
    ''' <param name="ordID"></param>
    Private Sub DataToControl_Order(ordID As String)
        ClearControls(tpOrder)

        Dim rowData = SelectTable("SELECT * " &
                                  "FROM orders a " &
                                  "LEFT JOIN customer b ON a.ord_cus_id = b.cus_id " &
                                  "LEFT JOIN product c ON a.ord_prod_id = c.prod_id " &
                                  "LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id " &
                                  "LEFT JOIN employee e ON a.ord_emp_id=e.emp_id " &
                                 $"WHERE ord_id = '{ordID}'").Rows(0)

        GetDataToControls(tpOrder, rowData)

        ' 計算未收帳款
        Dim dt = SelectTable($"SELECT mon_income FROM money WHERE mon_ord_id = '{ordID}'")

        If dt.Rows.Count = 0 Then
            txtUnpay.Text = txtTotalPrice.Text
        Else
            Dim totalIncome = dt.AsEnumerable().Sum(Function(row) row.Field(Of Integer)("mon_income"))
            txtUnpay.Text = (CDec(txtTotalPrice.Text) - totalIncome).ToString()
        End If
    End Sub


    '訂單管理-訂單編號-連動資料
    Private Sub txtOrdID_order_TextChanged(sender As Object, e As EventArgs) Handles txtOrdID_order.TextChanged
        If String.IsNullOrEmpty(txtOrdID_order.Text) OrElse Not IsNumeric(txtOrdID_order.Text) Then Exit Sub

        Dim dic As New Dictionary(Of String, Object) From {{"ord_id", txtOrdID_order.Text}}
        Dim sql = "SELECT * FROM orders WHERE ord_id = @ord_id"
        DataToControl_Customer(SelectTable(sql, dic).Rows(0).Field(Of Integer)("ord_cus_id"))

        DataToControl_Distribute(txtOrdID_order.Text)
    End Sub

    '配餐管理-新增
    Private Sub distInsert_Click(sender As Object, e As EventArgs) Handles btnDistInsert.Click
        Dim d As Date
        Dim count As Integer

        '早餐有勾選
        If chkBreakfast_dist.Checked Then
            d = dtpStart_dist.Value
            count = txtBreak.Text

            '跑回圈直到結束日期或未配置餐數為0
            Do While (d.Date <= dtpEnd_dist.Value.Date And count > 0)
                InsertDistribute(d, txtOrdID_dist.Text, chkBreakfast_dist.Text, GetDistributeData, count)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        '午餐有勾選
        If chkLunch_dist.Checked Then
            d = dtpStart_dist.Value
            count = txtLunch.Text

            '跑回圈直到結束日期或未配置餐數為0
            Do While (d.Date <= dtpEnd_dist.Value.Date And count > 0)
                InsertDistribute(d, txtOrdID_dist.Text, chkLunch_dist.Text, GetDistributeData, count)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        If chkDinner_dist.Checked Then
            d = dtpStart_dist.Value
            count = txtDinner.Text

            '跑回圈直到結束日期或未配置餐數為0
            Do While (d.Date <= dtpEnd_dist.Value.Date And count > 0)
                InsertDistribute(d, txtOrdID_dist.Text, chkDinner_dist.Text, GetDistributeData, count)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        UpdateCustomerTaboo()
        SetCalender()
        SetCalenderData()
        CountNotConfigured()
        MsgBox("新增成功")
    End Sub

    '配餐管理-修改
    Private Sub btnDistModify_Click(sender As Object, e As EventArgs) Handles btnDistModify.Click
        Dim d As Date

        '早餐有勾選
        If chkBreakfast_dist.Checked Then
            d = dtpStart_dist.Value

            '跑回圈直到結束日期
            Do While d.Date <= dtpEnd_dist.Value.Date
                UpdateDistribute(d, txtOrdID_dist.Text, chkBreakfast_dist.Text, GetDistributeData)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        '午餐有勾選
        If chkLunch_dist.Checked Then
            d = dtpStart_dist.Value

            '跑回圈直到結束日期
            Do While d.Date <= dtpEnd_dist.Value.Date
                UpdateDistribute(d, txtOrdID_dist.Text, chkLunch_dist.Text, GetDistributeData)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        If chkDinner_dist.Checked Then
            d = dtpStart_dist.Value

            '跑回圈直到結束日期
            Do While d.Date <= dtpEnd_dist.Value.Date
                UpdateDistribute(d, txtOrdID_dist.Text, chkDinner_dist.Text, GetDistributeData)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        UpdateCustomerTaboo()
        SetCalender()
        SetCalenderData()
        CountNotConfigured()
        MsgBox("修改成功")
    End Sub

    '配餐管理-刪除
    Private Sub btnDistDel_Click(sender As Object, e As EventArgs) Handles btnDistDel.Click
        Dim d As Date

        '早餐有勾選
        If chkBreakfast_dist.Checked Then
            d = dtpStart_dist.Value

            '跑回圈直到結束日期
            Do While d.Date <= dtpEnd_dist.Value.Date
                DeleteDistribute(d, txtOrdID_dist.Text, chkBreakfast_dist.Text)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        '午餐有勾選
        If chkLunch_dist.Checked Then
            d = dtpStart_dist.Value

            '跑回圈直到結束日期
            Do While d.Date <= dtpEnd_dist.Value.Date
                DeleteDistribute(d, txtOrdID_dist.Text, chkLunch_dist.Text)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        If chkDinner_dist.Checked Then
            d = dtpStart_dist.Value

            '跑回圈直到結束日期
            Do While d.Date <= dtpEnd_dist.Value.Date
                DeleteDistribute(d, txtOrdID_dist.Text, chkDinner_dist.Text)
                '前進一天
                d = d.AddDays(1)
            Loop
        End If

        SetCalender()
        SetCalenderData()
        CountNotConfigured()
        MsgBox("刪除成功")
    End Sub

    Private Sub UpdateCustomerTaboo()
        '禁忌編號
        If Not String.IsNullOrEmpty(txtTaboo_dist.Text) Then
            Dim lst As New List(Of String)

            For Each n In Split(txtTaboo_dist.Text, ",")
                Dim dicTaboo As New Dictionary(Of String, Object) From {{"tabo_name", n}}
                lst.Add(SelectTable("SELECT * FROM taboo WHERE tabo_name = @tabo_name", dicTaboo).Rows(0).Field(Of Integer)("tabo_id"))
            Next

            Dim dic = New Dictionary(Of String, Object) From {{"cus_tabo_id", String.Join(",", lst)}}
            UpdateCustomer(dic, txtCusID_dist.Text)
        End If
    End Sub

    Private Function GetDistributeData() As Dictionary(Of String, Object)
        Dim result = New Dictionary(Of String, Object)

        With result
            For Each grp In flpDist.Controls.OfType(Of GroupBox)
                Dim flp = grp.Controls.OfType(Of FlowLayoutPanel).FirstOrDefault
                Dim flpControls As Forms.Control.ControlCollection = Nothing

                If flp IsNot Nothing Then flpControls = flp.Controls

                Select Case grp.Text
                    Case "湯盅"
                        .Add("dist_soup", flpControls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault())
                    Case "麻油"
                        .Add("dist_oil", flpControls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault())
                    Case "酒"
                        .Add("dist_wine", flpControls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault())
                    Case "素"
                        .Add("dist_vege", flpControls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault())
                    Case "其他"
                        .Add("dist_other", String.Join(",", flpControls.OfType(Of Forms.CheckBox).Where(Function(x) x.Checked = True).Select(Function(x) x.Text)))
                    Case "客製需求"
                        .Add("dist_customized", String.Join(",", flpControls.OfType(Of Forms.CheckBox).Where(Function(x) x.Checked = True).Select(Function(x) x.Text)))
                    Case "餐具"
                        .Add("dist_tableware", String.Join(",", flpControls.OfType(Of Forms.CheckBox).Where(Function(x) x.Checked = True).Select(Function(x) x.Text)))
                    Case "飲品需求"
                        .Add("dist_drink", String.Join(",", flpControls.OfType(Of Forms.CheckBox).Where(Function(x) x.Checked = True).Select(Function(x) x.Text)))
                End Select
            Next

            '取得訂單預設地址
            Dim dicAddress As New Dictionary(Of String, Object) From {{"ord_id", txtOrdID_dist.Text}}
            .Add("dist_address", SelectTable("SELECT ord_break_addr FROM orders WHERE ord_id = @ord_id", dicAddress).Rows(0).Field(Of String)("ord_break_addr"))

            .Add("dist_memo", dgvDist.SelectedRows(0).Cells("ord_memo").Value.ToString)
            .Add("dist_memo2", txtMemo2_dist.Text)
        End With

        Return result
    End Function

    '配餐管理-訂單編號-連動資料
    Private Sub txtOrdID_dist_TextChanged(sender As Object, e As EventArgs) Handles txtOrdID_dist.TextChanged
        If String.IsNullOrEmpty(txtOrdID_dist.Text) OrElse txtOrdID_dist.Text = txtOrdID_order.Text Then Exit Sub
        DataToControl_Order(txtOrdID_dist.Text)
    End Sub

    '客戶管理-dgv點擊
    Private Sub dgvCustomer_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvCustomer.CellMouseClick
        Dim dgv As DataGridView = sender
        Dim cusID As Integer = dgv.SelectedRows(0).Cells("cus_id").Value
        Dim obj As Object = Nothing

        DataToControl_Customer(cusID)

        If Not String.IsNullOrEmpty(txtCusName_cus.Text) Then
            txtOrdQuery.Text = txtCusName_cus.Text
            txtDistQuery.Text = txtCusName_cus.Text
            btnOrderQuery_Click(obj, EventArgs.Empty)
            btnDistQuery_Click(obj, EventArgs.Empty)
        End If
    End Sub

    ''' <summary>
    ''' 將資料傳至客戶管理各控制項
    ''' </summary>
    ''' <param name="cusID"></param>
    Private Sub DataToControl_Customer(cusID As Integer)
        ClearControls(tpConsult_cus)
        ClearControls(tpBasic_cus)

        Dim dic = New Dictionary(Of String, Object) From {{"cus_id", cusID}}
        Dim rowCus = SelectTable("SELECT * FROM customer WHERE cus_id = @cus_id", dic).Rows(0)

        GetDataToControls(tpBasic_cus, rowCus)
        GetDataToControls(tpConsult_cus, rowCus)

        '禁忌
        txtTaboo_cus.Text = GetTabooNames(rowCus("cus_tabo_id").ToString())

        '坐月子地點
        grpConfLoca.Controls.OfType(Of TextBox)().ToList().ForEach(Sub(x)
                                                                       x.Text = If(IsDBNull(rowCus(x.Tag)), String.Empty, rowCus(x.Tag).ToString())
                                                                   End Sub)


        '顯示歷史訂單
        Dim sql = "SELECT a.ord_id,a.ord_date,b.cus_name,b.cus_phone" &
             " FROM orders a" &
             " LEFT JOIN customer b ON a.ord_cus_id = b.cus_id" &
             " LEFT JOIN product c ON a.ord_prod_id = c.prod_id" &
             " LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id" &
            $" WHERE b.cus_id = '{cusID}'" &
             " ORDER BY a.ord_date DESC"

        DataToDgv(sql, dgvOrder_cus)
    End Sub

    '客戶管理-新增
    Private Sub btnCusInsert_Click(sender As Object, e As EventArgs) Handles btnCusInsert.Click
        If Not String.IsNullOrEmpty(txtCusID.Text) Then
            MsgBox("請先按 取消 清空畫面再新增")
            Exit Sub
        End If

        Dim dic = CheckCustomer()

        Dim list As New List(Of Object) From {txtCusName_cus, txtPhone_cus}
        If Not CheckDuplication(sqlCustomer, list, dgvCustomer) Then Exit Sub

        If dic.Count = 0 Then Exit Sub
        If InserTable("customer", dic) Then
            btnCusCancel.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '客戶管理-修改
    Private Sub btnCusModify_Click(sender As Object, e As EventArgs) Handles btnCusModify.Click
        If String.IsNullOrEmpty(txtCusID.Text) Then
            MsgBox("請先選擇修改對象")
            Exit Sub
        End If

        Dim dic = CheckCustomer()
        If dic.Count = 0 Then Exit Sub
        If UpdateTable("customer", dic, $"cus_id = '{txtCusID.Text}'") Then
            btnCusCancel.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    Private Function CheckCustomer() As Dictionary(Of String, Object)
        Dim dic As New Dictionary(Of String, Object)

        Dim dicReq As New Dictionary(Of String, Object) From {
            {"姓名", txtCusName_cus},
            {"手機", txtPhone_cus}
        }

        If Not CheckRequiredCol(dicReq) Then Return dic

        '檢查日期格式
        Dim d As Date
        Dim dicDate As New Dictionary(Of String, TextBox) From {
            {"生日", txtBirthday},
            {"預產期", txtDueDate}
        }

        For Each kvp In dicDate
            If Not String.IsNullOrEmpty(kvp.Value.Text) And Not Date.TryParse(kvp.Value.Text, d) Then
                MsgBox(kvp.Key + " 日期格式錯誤")
                kvp.Value.Focus()
                Return dic
            End If
        Next

        '檢查數字
        Dim dicNumber As New Dictionary(Of String, TextBox) From {
            {"子女人數", txtChildren},
            {"第幾胎", txtManyChild},
            {"月子天數", txtConfDay},
            {"欲購買月子餐天數", txtConfBuy},
            {"身高", txtHeight},
            {"產前體重", txtBornWeight},
            {"目前體重", txtWeight}
        }

        If Not String.IsNullOrEmpty(txtChildren.Text) And Not IsNumeric(txtChildren.Text) Then
            MsgBox("子女人數 不為數字")
            txtChildren.Focus()
            Return dic
        End If

        With dic
            tpBasic_cus.Controls.OfType(Of TextBox).Where(Function(txt) Not String.IsNullOrEmpty(txt.Tag)).ToList.
                ForEach(Sub(t) .Add(t.Tag, t.Text))

            tpConsult_cus.Controls.OfType(Of TextBox).Where(Function(txt) Not String.IsNullOrEmpty(txt.Tag)).ToList.
                ForEach(Sub(t) .Add(t.Tag, t.Text))

            Dim rdo = grpGender.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
            If rdo IsNot Nothing Then .Add("cus_gender", rdo.Text) '性別

            rdo = grpMarriage.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
            If rdo IsNot Nothing Then .Add("cus_marriage", rdo.Text) '家庭成員

            If rdoMarried.Checked Then
                .Add("cus_spouse", txtSpouse.Text) '配偶
                .Add("cus_children", txtChildren.Text) '子女人數                      
            End If

            rdo = grpAcad_Qual.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(x) x.Checked)
            If rdo IsNot Nothing Then .Add("cus_acad_qual", rdo.Text) '學歷

            Dim list As New List(Of String)

            For Each check In grpKind.Controls.OfType(Of Forms.CheckBox)
                If check.Checked Then
                    If check.Text = "術後餐" Then
                        .Add("cus_kind_else", txtKindElse.Text)
                    End If
                    list.Add(check.Text)
                End If
            Next
            .Add("cus_kind", String.Join(",", list)) '種類
            list.Clear()

            For Each check In grpGetMsg.Controls.OfType(Of Forms.CheckBox)
                If check.Checked Then
                    If check.Text = "其他" Then
                        .Add("cus_getmsg_else", txtGetMsgElse.Text)
                    End If
                    list.Add(check.Text)
                End If
            Next
            .Add("cus_get_msg", String.Join(",", list)) '得知媒體
            list.Clear()

            For Each check In grpDisease.Controls.OfType(Of Forms.CheckBox)
                If check.Checked Then
                    If check.Text = "其他" Then
                        .Add("cus_disease_else", txtDisease.Text)
                    End If
                    list.Add(check.Text)
                End If
            Next
            .Add("cus_disease", String.Join(",", list)) '疾病史
            list.Clear()

            '禁忌編號
            If String.IsNullOrEmpty(txtTaboo_cus.Text) = False Then
                Dim lst As New List(Of String)

                For Each n In Split(txtTaboo_cus.Text, ",")
                    Dim dicTaboo As New Dictionary(Of String, Object) From {{"tabo_name", n}}
                    lst.Add(SelectTable("SELECT * FROM taboo WHERE tabo_name = @tabo_name", dicTaboo).Rows(0).Field(Of Integer)("tabo_id"))
                Next

                .Add("cus_tabo_id", String.Join(",", lst))
            End If

            '坐月子地點
            grpConfLoca.Controls.OfType(Of TextBox).ToList.ForEach(Sub(x) .Add(x.Tag, x.Text))

        End With

        Return dic
    End Function

    '客戶管理-取消
    Private Sub btnCusCancel_Click(sender As Object, e As EventArgs) Handles btnCusCancel.Click
        Cursor = Cursors.WaitCursor

        DataToDgv(sqlCustomer, dgvCustomer)
        ClearControls(tpBasic_cus)
        ClearControls(tpConsult_cus)

        Cursor = Cursors.Default
    End Sub

    '系統設定-禁忌食物管理-禁忌群組-取消
    Private Sub btnCancel_taboo_group_Click(sender As Object, e As EventArgs) Handles btnCancel_taboo_group.Click
        DataToDgv("SELECT * FROM taboo_group", dgvTabooGroup)
        ClearControls(tpTaboo)
        dgvTaboo.DataSource = Nothing
    End Sub

    '系統設定-禁忌食物管理-禁忌群組-新增
    Private Sub btnInsert_taboo_group_Click(sender As Object, e As EventArgs) Handles btnInsert_taboo_group.Click
        Dim btn As Button = sender
        Dim controls = btn.Parent.Controls

        If Not String.IsNullOrEmpty(controls.OfType(Of TextBox).First(Function(txt) txt.Tag.ToString.Contains("id")).Text) Then
            MsgBox("請先按 取消 清空畫面再新增")
            Exit Sub
        End If

        Dim dic = CheckTabooGroup()
        If dic.Count = 0 Then Exit Sub

        Dim list As New List(Of Object) From {txtTabooGroupName}
        If Not CheckDuplication("SELECT * FROM taboo_group", list, dgvTabooGroup) Then Exit Sub

        If InserTable("taboo_group", dic) Then
            controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '系統設定-禁忌食物管理-禁忌群組-修改
    Private Sub btnModify_taboo_group_Click(sender As Object, e As EventArgs) Handles btnModify_taboo_group.Click
        Dim btn As Button = sender
        Dim controls = btn.Parent.Controls
        Dim txtID As TextBox = controls.OfType(Of TextBox).First(Function(txt) txt.Tag.ToString.Contains("id"))

        If String.IsNullOrEmpty(txtID.Text) Then
            MsgBox("請先選擇修改對象")
            Exit Sub
        End If

        Dim dic = CheckTabooGroup()

        If dic.Count = 0 Then Exit Sub

        If UpdateTable("taboo_group", dic, $"{txtID.Tag} = '{txtID.Text}'") Then
            controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    Private Function CheckTabooGroup() As Dictionary(Of String, Object)
        Dim dic As New Dictionary(Of String, Object)

        Dim dicReq As New Dictionary(Of String, Object) From {
            {"群組名稱", txtTabooGroupName}
        }
        If Not CheckRequiredCol(dicReq) Then Return dic

        grpTabooGroup.Controls.OfType(Of TextBox).Where(Function(txt) Not String.IsNullOrEmpty(txt.Tag) AndAlso Not String.IsNullOrWhiteSpace(txt.Text)).ToList.
                ForEach(Sub(t) dic.Add(t.Tag, t.Text))

        Return dic
    End Function

    '系統設定-禁忌食物管理-禁忌群組-刪除
    Private Sub btnDelete_taboo_group_Click(sender As Object, e As EventArgs) Handles btnDelete_taboo_group.Click
        Dim btn As Button = sender
        Dim controls = btn.Parent.Controls
        Dim txtID As TextBox = controls.OfType(Of TextBox).First(Function(txt) txt.Tag.ToString.Contains("id"))

        If String.IsNullOrEmpty(txtID.Text) Then
            MsgBox("請先選擇刪除對象")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        If DeleteData("taboo_group", $"{txtID.Tag} = '{txtID.Text}'") Then
            DeleteData("taboo", $"tabo_tg_id = {txtID.Text}")
            controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("刪除成功")
        End If

    End Sub

    '系統設定-禁忌食物管理-禁忌群組-dgv點擊
    Private Sub dgvTabooGroup_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvTabooGroup.CellMouseClick
        Dim dgv As DataGridView = sender
        Dim row = dgv.SelectedRows(0)
        GetDataToControls(grpTabooGroup, row)
        DataToDgv($"SELECT tabo_id, tabo_name FROM taboo WHERE tabo_tg_id = '{row.Cells("tg_id").Value}'", dgvTaboo)
    End Sub

    '系統設定-禁忌管理-新增
    Private Sub btnTaboInsert_Click(sender As Object, e As EventArgs) Handles btnTaboInsert.Click
        Dim btn As Button = sender
        Dim controls = btn.Parent.Controls

        If Not String.IsNullOrEmpty(controls.OfType(Of TextBox).First(Function(txt) txt.Tag.ToString.Contains("id")).Text) Then
            MsgBox("請先按 取消 清空畫面再新增")
            Exit Sub
        End If

        Dim dic = CheckTaboo()
        If dic.Count = 0 Then Exit Sub

        Dim list As New List(Of Object) From {txtTaboName}
        If Not CheckDuplication("SELECT tabo_id, tabo_name FROM taboo", list, dgvTaboo) Then Exit Sub

        If InserTable("taboo", dic) Then
            controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    '系統設定-禁忌管理-修改
    Private Sub btnTaboModify_Click(sender As Object, e As EventArgs) Handles btnTaboModify.Click
        Dim btn As Button = sender
        Dim controls = btn.Parent.Controls
        Dim txtID As TextBox = controls.OfType(Of TextBox).First(Function(txt) txt.Tag.ToString.Contains("id"))

        If String.IsNullOrEmpty(txtID.Text) Then
            MsgBox("請先選擇修改對象")
            Exit Sub
        End If

        Dim dic = CheckTaboo()

        If dic.Count = 0 Then Exit Sub

        If UpdateTable("taboo", dic, $"{txtID.Tag} = '{txtID.Text}'") Then
            controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    Private Function CheckTaboo() As Dictionary(Of String, Object)
        Dim dic As New Dictionary(Of String, Object)
        Dim tgID = txtID_taboo_group.Text

        If String.IsNullOrEmpty(tgID) Then
            MsgBox("請先選擇群組")
            Return dic
        End If

        Dim dicReq As New Dictionary(Of String, Object) From {{"食材名稱", txtTaboName}}
        If Not CheckRequiredCol(dicReq) Then Return dic

        tpTaboo.Controls.OfType(Of TextBox).Where(Function(txt) Not String.IsNullOrEmpty(txt.Tag) AndAlso Not String.IsNullOrWhiteSpace(txt.Text)).ToList.
                ForEach(Sub(t) dic.Add(t.Tag, t.Text))
        dic.Add("tabo_tg_id", tgID)
        Return dic
    End Function

    '系統設定-禁忌管理-刪除
    Private Sub btnTaboDel_Click(sender As Object, e As EventArgs) Handles btnTaboDel.Click
        Dim btn As Button = sender
        Dim tp As TabPage = sender.Parent
        Dim controls = btn.Parent.Controls
        '取得編號
        Dim id As TextBox = tp.Controls.OfType(Of TextBox)().FirstOrDefault(Function(x) x.Tag.ToString.Contains("id"))

        If String.IsNullOrEmpty(id.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If

        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub

        If DeleteData(tp.Tag, $"tabo_id  = '{id.Text}'") Then
            controls.OfType(Of Button).First(Function(b) b.Text = "取  消").PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '系統設定-禁忌管理-查詢
    Private Sub btnTaboQuery_Click(sender As Object, e As EventArgs) Handles btnTaboQuery.Click
        Dim tp = CType(sender, Button).Parent
        Dim sql = $"SELECT tabo_id, tabo_name FROM taboo WHERE tabo_name LIKE '%{txtTaboQuery.Text}%'"
        DataToDgv(sql, tp.Controls.OfType(Of DataGridView).FirstOrDefault)

        grpTabooGroup.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Clear())
        tpTaboo.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Clear())
    End Sub

    '系統設定-禁忌管理-dgv點擊
    Private Sub dgvTaboo_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvTaboo.CellMouseClick
        Dim dgv As DataGridView = sender
        Dim parent = dgv.Parent
        Dim row = dgv.SelectedRows(0)
        GetDataToControls(parent, row)

        '搜尋後點擊要回推食材的群組編號

        Dim grpRow = SelectTable($"SELECT b.tg_id, b.tg_name FROM taboo a LEFT JOIN taboo_group b ON a.tabo_tg_id = b.tg_id WHERE a.tabo_id = '{txtTaboID.Text}'").Rows(0)
        txtID_taboo_group.Text = grpRow("tg_id")
        txtTabooGroupName.Text = grpRow("tg_name")

    End Sub

    '系統設定-禁忌管理-取消
    Private Sub btnCancel_taboo_Click(sender As Object, e As EventArgs) Handles btnCancel_taboo.Click
        tpTaboo.Controls.OfType(Of TextBox).ToList.ForEach(Sub(txt) txt.Clear())
        DataToDgv($"SELECT tabo_id, tabo_name FROM taboo WHERE tabo_tg_id = '{txtID_taboo_group.Text}'", dgvTaboo)
    End Sub

    '取消-權限管理
    Private Sub btnPermCancel_Click(sender As Object, e As EventArgs) Handles btnPermCancel.Click
        BtnCancel(sender, sqlPermision, dgvPermissions)
        InitPosition()
    End Sub

    '取消-員工管理
    Private Sub btnEmpCancel_Click(sender As Object, e As EventArgs) Handles btnEmpCancel.Click
        BtnCancel(sender, sqlEmployee, dgvEmployee)
    End Sub

    'mark by v0.13-----
    ''清除-送餐管理
    'Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear_drive.Click
    '    dgvDrive.DataSource = Nothing
    '    For Each grp In grpQuickSet.Controls.OfType(Of GroupBox)
    '        grp.Controls.OfType(Of ComboBox).ToList.ForEach(Sub(cmb) cmb.DataSource = Nothing)
    '    Next
    '    InitDriver()
    '    dgvDrive.ReadOnly = False
    'End Sub
    '-----

    '配餐管理-取消
    Private Sub btnDistCancel_Click(sender As Object, e As EventArgs) Handles btnCancel_dist.Click
        BtnCancel(sender, sqlDistribute, dgvDist)
        SetCalender()
        lblBreak_dist.BackColor = System.Drawing.Color.White
        lblLunch_dist.BackColor = System.Drawing.Color.White
        lblDinner_dist.BackColor = System.Drawing.Color.White
        InitDistribute()
        btnDistInsert.Enabled = False
        btnDistModify.Enabled = False
        btnDistDel.Enabled = False
    End Sub

    '取消-系統設定-商品群組管理
    Private Sub btnProdGrpCancel_Click(sender As Object, e As EventArgs) Handles btnCancel_prod_grp.Click
        BtnCancel(sender, sqlProductGroup, dgvProdgroup)
    End Sub

    '取消-菜單管理
    Private Sub btnMenuCancel_Click(sender As Object, e As EventArgs) Handles btnCancel_menu.Click
        ClearControls(tpMenu)
        DataToDgv(sqlMenu, dgvMenu)
        InitcmbProduct()
    End Sub

    '新增-系統設定-權限管理
    Private Sub btnPermInsert_Click(sender As Object, e As EventArgs) Handles btnInsert_perm.Click

        Dim dicReq As New Dictionary(Of String, Object) From {{"職位", txtPermName}}
        If BtnInsert(sender, txtId_perm, dicReq) Then MsgBox("新增成功")
    End Sub

    '新增-員工管理
    Private Sub btnEmpInsert_Click(sender As Object, e As EventArgs) Handles btnInsert_emp.Click
        Dim dicReq As New Dictionary(Of String, Object) From {
            {"姓名", txtName_emp},
            {"手機", txtPhone_emp},
            {"帳號", txtAcct},
            {"密碼", txtPsw},
            {"職位", cmbPosition_emp}
        }
        If BtnInsert(sender, txtId_emp, dicReq) Then MsgBox("新增成功")
    End Sub

    '新增-系統設定-商品群組管理
    Private Sub btnProdGrpInsert_Click(sender As Object, e As EventArgs) Handles btnInsert_prod_grp.Click
        Dim required As New Dictionary(Of String, Object) From {{"名稱", txtName_prod_grp}}
        If BtnInsert(sender, txtId_prod_grp, required) Then MsgBox("新增成功")
        InitcmbProductGroup()
    End Sub

    '新增-商品管理
    Private Sub btnProdInsert_Click(sender As Object, e As EventArgs) Handles btnProdInsert.Click
        Dim table = "product"
        Dim dic As New Dictionary(Of String, Object) From {
            {"商品名稱", txtProdName},
            {"售價", txtProdPrice},
            {"商品群組", cmbProdGrp_product},
            {"商品分類", cmbProdType}
        }
        If Not CheckRequiredCol(dic) Then Exit Sub
        If Not grpMeal.Controls.OfType(Of CheckBox).Any(Function(chk) chk.Checked) Then
            MsgBox("請勾選餐種")
            Exit Sub
        End If
        If InserTable(table, BindData(table)) Then
            btnCancel_prod.PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    'dgv點擊-系統設定-權限管理
    Private Sub dgvPermissions_CellMouseClick(sender As DataGridView, e As DataGridViewCellMouseEventArgs) Handles dgvPermissions.CellMouseClick
        Dim dgv As DataGridView = sender
        ClearControls(tpPermissions)
        Dim row = sender.SelectedRows(0)
        GetDataToControls(tpPermissions, row)
        GetDataToControls(grpPosition, row)
    End Sub

    'dgv點擊-員工管理
    Private Sub dgvEmployee_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvEmployee.CellMouseClick
        ClearControls(tpEmployee)
        GetDataToControls(tpEmployee, sender.SelectedRows(0))
    End Sub

    'dgv點擊-配餐管理
    Private Sub dgvDistribute_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDist.CellMouseClick
        '點dgv後將對象資料傳至各控制項
        Dim dgvRow = dgvDist.SelectedRows(0)
        DataToControl_Distribute(dgvRow.Cells("ord_id").Value)
    End Sub

    Private Sub DataToControl_Distribute(ordID As Integer)
        ClearControls(tpDistribute)

        '初始化目前選取早午晚餐的燈號
        lblBreak_dist.BackColor = System.Drawing.Color.White
        lblLunch_dist.BackColor = System.Drawing.Color.White
        lblDinner_dist.BackColor = System.Drawing.Color.White

        InitDistribute()

        Dim dic = New Dictionary(Of String, Object) From {{"ord_id", ordID}}
        Dim sql = "SELECT *" &
                    " FROM orders a" &
                    " LEFT JOIN customer b ON a.ord_cus_id=b.cus_id" &
                    " LEFT JOIN product c ON a.ord_prod_id=c.prod_id" &
                    " LEFT JOIN distribute d ON a.ord_id=d.dist_ord_id" &
                    " WHERE a.ord_id = @ord_id" &
                    " ORDER BY dist_date"
        Dim rowData = SelectTable(sql, dic).Rows(0)

        GetDataToControls(dgvDist.Parent, rowData)

        '禁忌
        txtTaboo_dist.Text = GetTabooNames(rowData("cus_tabo_id").ToString())

        '設定最近訂餐日期到月曆日期
        If Not IsDBNull(rowData("dist_date")) Then txtDistCalendar.Text = Date.Parse(rowData("dist_date")).ToString("Y")

        '刷新月曆
        SetCalender()
        SetCalenderData()
        CountNotConfigured()
        btnDistInsert.Enabled = True
        btnDistModify.Enabled = True
        btnDistDel.Enabled = True
    End Sub

    'dgv點擊-系統設定-配餐參數管理
    Private Sub dgvDistSys_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDistSys.CellMouseClick
        ClearControls(tpDistSys)
        GetDataToControls(tpDistSys, dgvDistSys.SelectedRows(0))
    End Sub

    'dgv點擊-商品群組管理
    Private Sub dgvProdgroup_CellMouseClick(dgv As DataGridView, e As DataGridViewCellMouseEventArgs) Handles dgvProdgroup.CellMouseClick
        DGVCellMouseClick(dgv)
    End Sub

    'dgv點擊-商品管理
    Private Sub dgvProduct_CellMouseClick(sender As DataGridView, e As DataGridViewCellMouseEventArgs) Handles dgvProduct.CellMouseClick
        ClearControls(tpProduct)
        GetDataToControls(tpProduct, sender.SelectedRows(0))
    End Sub

    '修改-系統設定-權限管理
    Private Sub btnPermModify_Click(sender As Object, e As EventArgs) Handles btnPermModify.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = tp.Tag.ToString
        Dim dicReq As New Dictionary(Of String, String) From {{"perm_name", "職位"}}
        If Not CheckRequiredCol(tp, dicReq) Then Exit Sub
        If Not UpdateTable(sTable, BindData(sTable), $"perm_id  = '{txtId_perm.Text}'") Then Exit Sub
        btnPermCancel.PerformClick()
        MsgBox("修改成功")
    End Sub

    '修改-員工管理
    Private Sub btnEmpModify_Click(sender As Object, e As EventArgs) Handles btnEmpModify.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim sTable = dgvEmployee.Tag
        Dim dicReq As New Dictionary(Of String, String) From {
            {"emp_name", "姓名"},
            {"emp_phone", "手機"},
            {"emp_acct", "帳號"},
            {"emp_psw", "密碼"},
            {"emp_perm_id", "職位"}
        }
        If Not CheckRequiredCol(tp, dicReq) Then Exit Sub
        If Not UpdateTable(sTable, BindData(sTable), $"emp_id  = '{txtId_emp.Text}'") Then Exit Sub
        btnEmpCancel.PerformClick()
        InitSales()
        MsgBox("修改成功")
    End Sub

    '修改-系統設定-商品群組管理
    Private Sub btnProdGrpModify_Click(btn As Object, e As EventArgs) Handles btnModify_prod_grp.Click
        Dim required As New Dictionary(Of String, Object) From {{"名稱", txtName_prod_grp}}
        If BtnModify(btn, txtId_prod_grp, txtId_prod_grp.Tag.ToString + " = " + txtId_prod_grp.Text, required) Then MsgBox("修改成功")
        InitcmbProductGroup()
    End Sub

    '修改-商品管理
    Private Sub btnProdModify_Click(sender As Object, e As EventArgs) Handles btnProdModify.Click
        Dim table = "product"
        Dim dic As New Dictionary(Of String, Object) From {
            {"商品名稱", txtProdName},
            {"售價", txtProdPrice},
            {"商品群組", cmbProdGrp_product},
            {"商品分類", cmbProdType}
        }
        If Not CheckRequiredCol(dic) Then Exit Sub
        If Not grpMeal.Controls.OfType(Of Forms.CheckBox).Any(Function(chk) chk.Checked) Then
            MsgBox("請勾選餐種")
            Exit Sub
        End If
        If UpdateTable(table, BindData(table), $"prod_id  = '{txtProdID.Text}'") Then
            '列出所有資料
            btnCancel_prod.PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    ''' <summary>
    ''' 繫結Table欄位與TextBox
    ''' </summary>
    Public Function BindData(sTable As String) As Dictionary(Of String, Object)
        Dim dicData As New Dictionary(Of String, Object)
        Dim chk As IEnumerable(Of Forms.CheckBox)
        Dim list As New List(Of String)

        With dicData
            Select Case sTable
                Case "product_group"
                    For Each ctrl In tpProdGroup.Controls.OfType(Of Forms.Control).Where(Function(control) Not String.IsNullOrEmpty(control.Tag) AndAlso Not String.IsNullOrEmpty(control.Text))
                        .Add(ctrl.Tag, ctrl.Text)
                    Next

                Case "product"
                    .Add("prod_name", txtProdName.Text)
                    .Add("prod_prod_grp_id", cmbProdGrp_product.SelectedValue)
                    .Add("prod_price", txtProdPrice.Text)
                    .Add("prod_cost", txtProdCost.Text)
                    .Add("prod_type", cmbProdType.Text)

                    '取得勾選的餐種
                    chk = grpMeal.Controls.OfType(Of Forms.CheckBox)().Where(Function(x) x.Checked)
                    .Add("prod_meal", String.Join(",", chk.Select(Function(x) x.Text)))
                    .Add("prod_memo", txtProdMemo.Text)

                Case "taboo"
                    '.Add("tabo_type", cmbTaboClass.Text)
                    .Add("tabo_name", txtTaboName.Text)

                Case "permissions"
                    .Add(txtPermName.Tag, txtPermName.Text)
                    tpPermissions.Controls.OfType(Of CheckBox).ToList.ForEach(Sub(x) .Add(x.Tag.ToString, If(x.Checked, "1", "0")))
                    grpPosition.Controls.OfType(Of CheckBox).ToList.ForEach(Sub(x) .Add(x.Tag.ToString, If(x.Checked, "1", "0")))

                Case "employee"
                    For Each txt In tpEmployee.Controls.OfType(Of TextBox).Where(Function(t) t.Tag IsNot Nothing AndAlso Not String.IsNullOrEmpty(t.Text))
                        .Add(txt.Tag, txt.Text)
                    Next
                    .Add("emp_perm_id", cmbPosition_emp.SelectedValue)

                Case "distribute"
                    .Add("dist_ord_id", txtOrdID_dist.Text)
                    .Add("dist_meal", tempDistDay.Text) '早午晚餐

                    Dim txt As String
                    For Each grp In flpDist.Controls.OfType(Of GroupBox)
                        Select Case grp.Text
                            Case "湯盅"
                                txt = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_soup", txt)
                            Case "麻油"
                                txt = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_oil", txt)
                            Case "酒"
                                txt = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_wine", txt)
                            Case "素"
                                txt = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of RadioButton).Where(Function(x) x.Checked = True).Select(Function(x) x.Text).FirstOrDefault()
                                .Add("dist_vege", txt)
                            Case "其他"
                                chk = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_other", String.Join(",", chk.Select(Function(x) x.Text)))
                            Case "客製需求"
                                chk = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_customized", String.Join(",", chk.Select(Function(x) x.Text)))
                            Case "餐具"
                                chk = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_tableware", String.Join(",", chk.Select(Function(x) x.Text)))
                            Case "飲品需求"
                                chk = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls.OfType(Of CheckBox).Where(Function(x) x.Checked = True)
                                .Add("dist_drink", String.Join(",", chk.Select(Function(x) x.Text)))
                        End Select
                    Next

                    '更新預設地址
                    Dim rowAddr As DataRow = Nothing
                    Dim city As String = ""
                    Dim area As String = ""
                    Dim address As String = ""

                    Select Case tempDistDay.Text
                        Case "早"
                            '取得訂單預設地址
                            rowAddr = SelectTable($"SELECT ord_break_city, ord_break_area, ord_break_addr FROM orders WHERE ord_id = {txtOrdID_dist.Text}").Rows(0)
                            city = If(rowAddr.Field(Of String)("ord_break_city"), "")
                            area = If(rowAddr.Field(Of String)("ord_break_area"), "")
                            address = If(rowAddr.Field(Of String)("ord_break_addr"), "")

                        Case "午"
                            '取得訂單預設地址
                            rowAddr = SelectTable($"SELECT ord_lunch_ctiy, ord_lunch_area, ord_lunch_addr FROM orders WHERE ord_id = {txtOrdID_dist.Text}").Rows(0)
                            city = If(rowAddr.Field(Of String)("ord_lunch_ctiy"), "")
                            area = If(rowAddr.Field(Of String)("ord_lunch_area"), "")
                            address = If(rowAddr.Field(Of String)("ord_lunch_addr"), "")

                        Case "晚"
                            '取得訂單預設地址
                            rowAddr = SelectTable($"SELECT ord_dinner_city, ord_dinner_area, ord_dinner_addr FROM orders WHERE ord_id = {txtOrdID_dist.Text}").Rows(0)
                            city = If(rowAddr.Field(Of String)("ord_dinner_city"), "")
                            area = If(rowAddr.Field(Of String)("ord_dinner_area"), "")
                            address = If(rowAddr.Field(Of String)("ord_dinner_addr"), "")
                    End Select

                    .Add("dist_city", city)
                    .Add("dist_area", area)
                    .Add("dist_address", address)

                Case "distribute_system"
                    .Add("dist_sys_type", cmbType_dist_sys.Text)
                    .Add("dist_sys_option", txtOption.Text)
            End Select
        End With

        Return dicData
    End Function

    '刪除-客戶管理
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
        btnCusCancel.PerformClick()
        MsgBox("刪除成功")
    End Sub

    '刪除-系統設定-權限管理
    Private Sub btnPermDel_Click(sender As Button, e As EventArgs) Handles btnPermDel.Click
        If BtnDelete(sender, txtId_perm, txtId_perm.Tag.ToString + " = " + txtId_perm.Text) Then MsgBox("刪除成功")
    End Sub

    '刪除-員工管理
    Private Sub btnEmpDelete_Click(sender As Object, e As EventArgs) Handles btnEmpDelete.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        If String.IsNullOrEmpty(txtId_emp.Text) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        Dim sTable As String = "employee"
        If DeleteData(sTable, $"emp_id  = '{txtId_emp.Text}'") Then
            btnEmpCancel.PerformClick()
            InitSales()
            'InitDriver()
            MsgBox("刪除成功")
        End If
    End Sub

    '刪除--系統設定-商品群組管理
    Private Sub btnProdGrpDel_Click(sender As Object, e As EventArgs) Handles btnDel_prod_grp.Click
        If BtnDelete(sender, txtId_prod_grp, txtId_prod_grp.Tag.ToString + " = " + txtId_prod_grp.Text) Then MsgBox("刪除成功")
        InitcmbProductGroup()
    End Sub

    '刪除-商品管理
    Private Sub btnProdDelete_Click(sender As Object, e As EventArgs) Handles btnProdDelete.Click
        '檢查是否選擇對象
        Dim id = txtProdID.Text
        If String.IsNullOrEmpty(id) Then
            MsgBox("請選擇刪除對象", Title:="提醒")
            Exit Sub
        End If
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        If DeleteData("product", $"prod_id = '{id}'") Then
            btnCancel_prod.PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '刪除-菜單管理
    Private Sub btnMenuDel_Click(sender As Object, e As EventArgs) Handles btnMenuDel.Click
        Dim dic As New Dictionary(Of String, String)
        With dic
            .Add("me_date", dtMenu.Value)
            If cmbProdVers_menu.SelectedItem Is Nothing Then
                MsgBox("請選擇版本")
                cmbProdVers_menu.Focus()
                Exit Sub
            End If
            Dim ver = cmbProdVers_menu.SelectedItem.ToString
            .Add("me_version", ver)
            If cmbProdName_menu.SelectedValue Is Nothing Then
                MsgBox("請選擇商品")
                cmbProdName_menu.Focus()
                Exit Sub
            End If
            Dim prodID = cmbProdName_menu.SelectedValue.ToString
            .Add("me_prod_id", prodID)
            Dim table = "menu"
            If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
            If DeleteData(table, $"me_date = '{dtMenu.Value:yyyy-MM-dd}' AND me_version = '{ver}' AND me_prod_id = {prodID}") Then
                btnCancel_menu.PerformClick()
                MsgBox("刪除完成")
            End If
        End With

    End Sub

    '查詢-客戶管理
    Private Sub btnCusQuery_Click(sender As Object, e As EventArgs) Handles btnCusQuery.Click
        Dim sql = sqlCustomer + $" WHERE cus_name LIKE '%{txtCusQuery.Text}%' or cus_phone LIKE '%{txtCusQuery.Text}%'"
        DataToDgv(sql, dgvCustomer)
        ClearControls(tpBasic_cus)
        ClearControls(tpConsult_cus)
    End Sub

    '查詢-配餐管理
    Private Sub btnDistQuery_Click(sender As Object, e As EventArgs) Handles btnDistQuery.Click
        Dim indexOrderBy = sqlDistribute.IndexOf("ORDER BY")
        Dim sql = sqlDistribute.Insert(indexOrderBy, $" WHERE b.cus_name Like '%{txtDistQuery.Text}%' OR b.cus_phone LIKE '%{txtDistQuery.Text}%' ")

        DataToDgv(sql, dgvDist)
    End Sub

    '查詢-員工管理
    Private Sub btnEmpQuery_Click(sender As Object, e As EventArgs) Handles btnEmpQuery.Click
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = tp.Tag.ToString

        Dim Sql = sqlEmployee + $" WHERE emp_name LIKE '%{txtEmpQuery.Text}%' OR emp_phone LIKE '%{txtEmpQuery.Text}%'"
        DataToDgv(SelectTable(Sql), "permissions,employee", dgvEmployee)
        MsgBox("搜尋完畢")
    End Sub

    '查詢-系統設定-商品群組管理
    Private Sub btnQuery_prod_grp_Click(btn As Object, e As EventArgs) Handles btnQuery_prod_grp.Click
        Dim Sql = sqlProductGroup + $" WHERE prod_grp_name LIKE '%{txtQuery_prod_grp.Text}%' OR prod_grp_aka LIKE '%{txtQuery_prod_grp.Text}%'"
        DataToDgv(Sql, dgvProdgroup)
    End Sub

    '搜尋欄位按下"Enter"即可搜尋
    Private Sub txtQuery_KeyPress(txt As TextBox, e As KeyPressEventArgs) Handles txtName_prod_grp.KeyPress, txtProdQuery.KeyPress, txtTaboQuery.KeyPress, txtQuery_money.KeyPress, txtEmpQuery.KeyPress, txtOrdQuery.KeyPress, txtCusQuery.KeyPress, txtDistQuery.KeyPress, txtQuery_prod_grp.KeyPress
        If e.KeyChar = vbCr Then
            Dim btn As Button = txt.Parent.Controls.OfType(Of Button).FirstOrDefault(Function(x) x.Text = "查詢")
            btn.PerformClick()
        End If
    End Sub

    '訂單管理-查詢
    Private Sub btnOrderQuery_Click(sender As Object, e As EventArgs) Handles btnOrderQuery.Click
        Cursor = Cursors.WaitCursor

        Dim sql = sqlOrder & $" WHERE b.cus_name Like '%{txtOrdQuery.Text}%' OR b.cus_phone LIKE '%{txtOrdQuery.Text}%' ORDER BY a.ord_date DESC"
        Dim exception As New List(Of String) From {"txtOrdQuery"}

        DataToDgv(sql, dgvOrder)
        ClearControls(tpOrder, exception)

        Cursor = Cursors.Default
    End Sub

    '訂單管理-商品群組選擇後過濾商品名稱
    Private Sub cmbProdGrp_order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProdGrp_order.SelectedIndexChanged
        If cmbProdGrp_order.SelectedIndex < 0 Then Exit Sub
        Dim prodGrpID = cmbProdGrp_order.SelectedValue.ToString
        With cmbProdName_order
            .DataSource = SelectTable($"SELECT * FROM product WHERE prod_prod_grp_id = '{prodGrpID}'")
            .DisplayMember = "prod_name"
            .ValueMember = "prod_id"
            .SelectedIndex = -1
        End With
    End Sub

    '訂單管理-商品名稱
    Private Sub cmbProdName_order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProdName_order.SelectedIndexChanged
        If cmbProdName_order.SelectedIndex = -1 Then
            txtPrice.Clear()
            ClearControls(grpMeal_order)
            Exit Sub
        End If
        Dim selectedRow As DataRowView = cmbProdName_order.SelectedItem
        Dim prodId As Integer = selectedRow("prod_id")
        Dim row = SelectTable($"SELECT prod_type, prod_price, prod_meal FROM product WHERE prod_id = {prodId}").Rows(0)
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
            txtAddrLunch.Clear()
        End If
    End Sub

    '訂單管理-晚餐地址-同上
    Private Sub chkDinnerAddr_CheckedChanged(sender As Object, e As EventArgs) Handles chkDinnerAddr.Click
        If chkDinnerAddr.Checked Then
            txtAddrDinner.Text = txtAddrLunch.Text
        Else
            txtAddrDinner.Clear()
        End If
    End Sub

    ''' <summary>
    ''' 設定所選訂單的月曆資料
    ''' </summary>
    Private Sub SetCalenderData()
        If txtOrdID_dist.Text = "" Then Exit Sub

        Dim d As Date = Date.Parse(txtDistCalendar.Text)

        '以當前月曆月份搜尋訂單配餐
        'Dim dt = SelectTable($"SELECT * FROM distribute WHERE YEAR(dist_date) = {d.Year} AND MONTH(dist_date) = {d.Month} AND dist_ord_id = {txtOrdID_dist.Text}")
        Dim sql = "SELECT * FROM product_group a " &
                  "LEFT JOIN product b ON b.prod_prod_grp_id = a.prod_grp_id " &
                  "LEFT JOIN orders c ON c.ord_prod_id = b.prod_id " &
                  "LEFT JOIN distribute d ON d.dist_ord_id = c.ord_id " &
                  "LEFT JOIN customer e ON e.cus_id = c.ord_cus_id " &
                  $"WHERE YEAR(dist_date) = @year AND MONTH(dist_date) = @month AND dist_ord_id = @dist_ord_id"

        Dim dic = New Dictionary(Of String, Object) From {
                {"dist_ord_id", txtOrdID_dist.Text},
                {"year", d.Year},
                {"month", d.Month}
            }

        Dim dt = SelectTable(sql, dic)
        Dim toolTip As New ToolTip

        For Each row As DataRow In dt.Rows
            '配餐日期
            d = row.Field(Of Date)("dist_date")

            '取得配餐日的panel
            Dim pnl = tlpCalendar.Controls.OfType(Of Panel).FirstOrDefault(Function(x) CInt(x.Tag) = d.Day)

            '將對應的餐種打包進按鈕的tag並改變顏色
            Dim btn = pnl?.Controls.OfType(Of Button).FirstOrDefault(Function(x) x.Text = row.Field(Of String)("dist_meal"))

            If btn IsNot Nothing Then
                btn.Tag = row
                Select Case btn.Text
                    Case "早"
                        btn.BackColor = System.Drawing.Color.LightGreen
                    Case "午"
                        btn.BackColor = System.Drawing.Color.Yellow
                    Case "晚"
                        btn.BackColor = System.Drawing.Color.LightCoral
                End Select
            End If

            'TipTool顯示
            Dim taboos = row("cus_tabo_id")
            Dim caption As String = $"床:{row("prod_grp_aka")} 電話:{row("cus_phone")}" & vbCrLf &
                                    $"地址:{row("cus_address")}" & vbCrLf &
                                    $"備註:{row("ord_memo")}" & vbCrLf &
                                    $"湯盅:{row("dist_soup")}" & vbCrLf &
                                    $"麻油:{row("dist_oil")}" & vbCrLf &
                                    $"酒:{row("dist_wine")}" & vbCrLf &
                                    $"客製需求:{row("dist_customized")}" & vbCrLf &
                                    $"禁忌:{GetTabooNames(row("cus_tabo_id").ToString)}" & vbCrLf &
                                    $"飲品需求:{row("dist_drink")}"

            toolTip.SetToolTip(btn, caption)
        Next
    End Sub

    ''' <summary>
    ''' 計算未配置餐
    ''' </summary>
    Private Sub CountNotConfigured()
        Dim dic As New Dictionary(Of String, Object) From {{"ord_id", txtOrdID_dist.Text}}
        Dim dtOrder = SelectTable("SELECT ord_breakfast, ord_lunch, ord_dinner FROM orders WHERE ord_id = @ord_id", dic)
        Dim dtDist = SelectTable("SELECT dist_meal FROM distribute WHERE dist_ord_id = @ord_id", dic)

        txtBreak.Text = If(dtOrder.Rows(0)("ord_breakfast") > 0, dtOrder.Rows(0)("ord_breakfast") - dtDist.Select("dist_meal='早'").Count, 0)
        txtLunch.Text = If(dtOrder.Rows(0)("ord_lunch") > 0, dtOrder.Rows(0)("ord_lunch") - dtDist.Select("dist_meal='午'").Count, 0)
        txtDinner.Text = If(dtOrder.Rows(0)("ord_dinner") > 0, dtOrder.Rows(0)("ord_dinner") - dtDist.Select("dist_meal='晚'").Count, 0)
    End Sub

    ''' <summary>
    ''' 製作月曆一日的物件
    ''' </summary>
    ''' <param name="i">日</param>
    ''' <returns></returns>
    Private Function DayMaker(i As Short) As Panel
        '框架
        Dim panel As New Panel With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyle.FixedSingle,
            .Tag = i
        }

        '日期
        Dim font As New System.Drawing.Font("Arial", 12, FontStyle.Bold)
        Dim point As System.Drawing.Point
        point = New System.Drawing.Point(0, 0)
        Dim lbl As New Label With {.Text = i, .Parent = panel, .Font = font, .AutoSize = True, .Location = point}

        '早中晚的按鈕式CheckBox
        With panel
            .Controls.Add(Setbtn_Dist(New System.Drawing.Point(27, 0), "早"))
            .Controls.Add(Setbtn_Dist(New System.Drawing.Point(2, 25), "午"))
            .Controls.Add(Setbtn_Dist(New System.Drawing.Point(27, 25), "晚"))
        End With

        Return panel
    End Function

    ''' <summary>
    ''' 設定配餐管理月曆裡的按鈕
    ''' </summary>
    ''' <param name="point">在Panel裡的位置</param>
    ''' <param name="txt">顯示的文字</param>
    ''' <returns></returns>
    Private Function Setbtn_Dist(point As System.Drawing.Point, txt As String) As Button
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
        If String.IsNullOrEmpty(txtOrdID_dist.Text) Then Exit Sub

        '刷新grp,除了未配置餐
        InitDistribute()

        tempDistDay = sender
        '標示現在選取的日期
        Dim day = tempDistDay.Parent.Controls.OfType(Of Label).FirstOrDefault.Text
        txtSelectDate.Text = Date.Parse(txtDistCalendar.Text + day + "日")
        Select Case tempDistDay.Text
            Case "早"
                lblBreak_dist.BackColor = System.Drawing.Color.LightGreen
                lblLunch_dist.BackColor = System.Drawing.Color.White
                lblDinner_dist.BackColor = System.Drawing.Color.White
            Case "午"
                lblLunch_dist.BackColor = System.Drawing.Color.Yellow
                lblBreak_dist.BackColor = System.Drawing.Color.White
                lblDinner_dist.BackColor = System.Drawing.Color.White
            Case "晚"
                lblDinner_dist.BackColor = System.Drawing.Color.LightCoral
                lblLunch_dist.BackColor = System.Drawing.Color.White
                lblBreak_dist.BackColor = System.Drawing.Color.White
        End Select

        '將btn.tag裡的物件資料送至各grp
        Dim rowData As DataRow = tempDistDay.Tag
        If rowData Is Nothing Then Exit Sub

        'grp名稱與欄位繫結
        Dim dic As New Dictionary(Of String, String)
        With dic
            .Add("dist_soup", "湯盅")
            .Add("dist_oil", "麻油")
            .Add("dist_wine", "酒")
            .Add("dist_vege", "素")
            .Add("dist_other", "其他")
            .Add("dist_customized", "客製需求")
            .Add("dist_tableware", "餐具")
            .Add("dist_drink", "飲品需求")
        End With

        For Each kvp In dic
            Dim grp = flpDist.Controls.OfType(Of GroupBox)().FirstOrDefault(Function(x) x.Text = kvp.Value)
            If Not rowData.IsNull(kvp.Key) Then
                Dim lst As List(Of String) = Split(rowData(kvp.Key), ",").ToList
                Dim flpCtrls = grp.Controls.OfType(Of FlowLayoutPanel).First.Controls

                For Each check In flpCtrls.OfType(Of CheckBox)
                    If lst.Contains(check.Text) Then
                        check.Checked = True
                    End If
                Next

                For Each rdo In flpCtrls.OfType(Of RadioButton)
                    If lst.Contains(rdo.Text) Then
                        rdo.Checked = True
                    End If
                Next

                lst.Clear()
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

        Dim lstMenu1 As New List(Of Menu) '蒐集完丟這裡
        If Not ReadExcel(lstMenu1) Then GoTo Finish

        '找出所有菜單有但菜色沒有的新菜色
        Dim dt = SelectTable("SELECT dish_name FROM dishes")
        Dim lstNewDishes As New List(Of String)

        For Each str As String In lstMenu1.Select(Function(x) x.Name).Distinct
            If dt.Select($"dish_name = '{str}'").Count = 0 AndAlso Not lstNewDishes.Contains(str) Then lstNewDishes.Add(str)
        Next

        '彈出視窗讓使用者快速新增
        Dim frm As New frmInsertDeshes With {.Dishes = lstNewDishes}
        frm.ShowDialog()

Finish:
        Cursor = Cursors.Default
    End Sub

    Private Function ReadExcel(ByRef lstMenu1 As List(Of Menu)) As Boolean
        Dim result As Boolean = False
        '讀取檔案
        Dim path As String = ""
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            path = OpenFileDialog1.FileName
        Else
            GoTo Finish
        End If

        '顯示請稍後
        frmLoadExcel.Show()

        Dim exl As New Excel.Application With {.DisplayAlerts = False}
        Dim sheet As Excel.Worksheet = exl.Workbooks.Open(path).Sheets("菜單總表")
        Dim rng As String
        Dim cell As Excel.Range
        Dim value As Object
        Dim txt1 As String
        Dim menu1 As Menu
        Dim tempCell As Excel.Range
        Dim backColor As Excel.XlRgbColor
        '版本
        rng = "A2"
        cell = sheet.Range(rng)
        value = cell.Value
        txt1 = value.ToString().Trim
        Dim ver2 = txt1.Substring(txt1.Length - 2, 1)
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
            Application.DoEvents()
            '日期
            rng = Chr(col) + "3"
            cell = sheet.Range(rng)
            value = cell.Value
            txt1 = value.ToString().Trim
            Dim d = Date.Parse(txt1)

            For row As Integer = 4 To 8
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "經典月子餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicMoonSonBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicMoonSonBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 9 To 17
                menu1 = New Menu With {
                    .Version = ver2,
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
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
                tempCell = Nothing

                If row = 10 Then
                    menu1 = New Menu With {
                        .Version = ver2,
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
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 11 Then
                    menu1 = New Menu With {
                        .Version = ver2,
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
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicMoonSonLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                If row = 10 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 11 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If
            Next

            For row As Integer = 18 To 25
                menu1 = New Menu With {
                    .Version = ver2,
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
                    rng = Chr(col) + "34"
                    tempCell = sheet.Range(rng)
                    If tempCell.Value.ToString = "" Then
                        value = cell.Value
                    Else
                        value = tempCell.Value
                    End If
                Else
                    value = cell.Value
                End If
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
                tempCell = Nothing

                If row = 19 Then
                    menu1 = New Menu With {
                        .Version = ver2,
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
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 20 Then
                    menu1 = New Menu With {
                        .Version = ver2,
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
                        rng = Chr(col) + "34"
                        tempCell = sheet.Range(rng)
                        If tempCell.Value.ToString = "" Then
                            value = cell.Value
                        Else
                            value = tempCell.Value
                        End If
                    Else
                        value = cell.Value
                    End If
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicMoonSonDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                If row = 19 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 20 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If
            Next

            For row As Integer = 26 To 28
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "經典月子餐",
                    .Meal = Meal.夜點,
                    .Meal_Detail = dicMoonSonNightSnack(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                If row = 27 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 28 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "經典月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "溫馨月子餐",
                    .Meal = Meal.夜點,
                    .Meal_Detail = dicMoonSonNightSnack(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                If row = 27 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅2期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 28 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "溫馨月子餐",
                        .Meal = Meal.夜點,
                        .Meal_Detail = Meal_Detail.湯盅4期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim.Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If
            Next

            For row As Integer = 39 To 44
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "調理餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicConditioning(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                If value Is Nothing Then Continue For
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next
            '調理餐水果
            menu1 = New Menu With {
                .Version = ver2,
                .MenuDate = d,
                .ProductName = "調理餐",
                .Meal = Meal.午餐,
                .Meal_Detail = dicConditioning(45)
            }
            rng = "D45"
            cell = sheet.Range(rng)
            value = cell.Value
            txt1 = value.ToString().Trim.Trim
            menu1.Name = txt1
            lstMenu1.Add(menu1)

            For row As Integer = 50 To 54
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "幸福餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicHappinessLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 56 To 60
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "幸福餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicHappinessDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim.Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 65 To 70
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "住院餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicHospitalizedBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 71 To 78
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "住院餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicHospitalizedLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 79 To 85
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "住院餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicHospitalizedDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 89 To 94
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "輕食餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicLightMealBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 96 To 101
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "輕食餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicLightMealLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 103 To 108
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "輕食餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicLightMealDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next
        Next

        For col As Integer = Asc("O") To Asc("U")
            Application.DoEvents()
            '日期
            rng = Chr(col) + "3"
            cell = sheet.Range(rng)
            value = cell.Value
            txt1 = value.ToString().Trim
            Dim d = Date.Parse(txt1)

            For row As Integer = 4 To 8
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "術後調理餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicOperationBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For Each row In dicOperationLunch.Keys
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "術後調理餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicOperationLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For Each row In dicOperationDinner.Keys
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "術後調理餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicOperationDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 27 To 31
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "素食餐",
                    .Meal = Meal.早餐,
                    .Meal_Detail = dicVegetarianBreak(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 32 To 38
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "素食餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicVegetarianLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                If row = 32 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅1期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 33 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.午餐,
                        .Meal_Detail = Meal_Detail.湯盅3期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If
            Next

            For Each row In dicVegetarianDinner.Keys
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "素食餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicVegetarianDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)

                If row = 39 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅1期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If

                If row = 40 Then
                    menu1 = New Menu With {
                        .Version = ver2,
                        .MenuDate = d,
                        .ProductName = "素食餐",
                        .Meal = Meal.晚餐,
                        .Meal_Detail = Meal_Detail.湯盅3期
                    }
                    rng = Chr(col) + row.ToString
                    cell = sheet.Range(rng)
                    value = cell.Value
                    txt1 = value.ToString().Trim
                    menu1.Name = txt1
                    lstMenu1.Add(menu1)
                End If
            Next

            For row As Integer = 53 To 57
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "素食一般餐",
                    .Meal = Meal.午餐,
                    .Meal_Detail = dicVegetarianNormalLunch(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next

            For row As Integer = 58 To 62
                menu1 = New Menu With {
                    .Version = ver2,
                    .MenuDate = d,
                    .ProductName = "素食一般餐",
                    .Meal = Meal.晚餐,
                    .Meal_Detail = dicVegetarianNormalDinner(row)
                }
                rng = Chr(col) + row.ToString
                cell = sheet.Range(rng)
                value = cell.Value
                txt1 = value.ToString().Trim
                menu1.Name = txt1
                lstMenu1.Add(menu1)
            Next
        Next

        'insert到table
        For Each m In lstMenu1
            Dim table = "menu"
            Dim dic As New Dictionary(Of String, Object) From {
                {"me_date", m.MenuDate},
                {"me_version", m.Version},
                {"me_meal_id", m.Meal},
                {"me_meal_detail_id", m.Meal_Detail},
                {"me_name", m.Name}
            }
            '與商品匹配
            Dim dt = SelectTable($"SELECT prod_id FROM product WHERE prod_name = '{m.ProductName}'")
            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)
                dic.Add("me_prod_id", row("prod_id").ToString)
                '先刪除後新增避免重複
                DeleteData(table, $"me_date = '{m.MenuDate:yyyy-MM-dd}' AND me_version = '{m.Version}' AND me_meal_id = {m.Meal} AND me_meal_detail_id = {m.Meal_Detail} AND me_prod_id = {row("prod_id")}")
                InserTable(table, dic)
            Else
                MsgBox("無 " + m.ProductName + " 商品,請先新增")
                GoTo Finish
            End If
        Next
        DataToDgv(sqlMenu, dgvMenu)
        result = True
        MsgBox("匯入完成")
Finish:
        frmLoadExcel.Close()
        Return result
    End Function

    '菜單管理-dgv點擊
    Private Sub dgvMenu_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvMenu.CellMouseClick
        ClearControls(tpMenu)

        '點dgv後將對象資料傳至各控制項
        Dim dgvRow = dgvMenu.SelectedRows(0)
        Dim d As Date = dgvRow.Cells("me_date").Value
        Dim ver As String = dgvRow.Cells("me_version").Value
        Dim prod As Integer = dgvRow.Cells("prod_id").Value
        Dim dic As New Dictionary(Of String, Object) From {
            {"me_date", d.ToString("yyyy-MM-dd")},
            {"me_version", ver},
            {"me_prod_id", prod}
        }
        Dim dataMuenu = SelectTable("SELECT * FROM menu WHERE me_date = @me_date AND me_version = @me_version AND me_prod_id = @me_prod_id", dic).Rows

        For Each row As DataRow In dataMuenu
            Dim t As String = CStr(row.Field(Of Integer)("me_meal_id")) + "," + CStr(row.Field(Of Integer)("me_meal_detail_id"))

            For Each txt In tpMenu.Controls.OfType(Of TextBox).Where(Function(x) x.Tag = t)
                txt.Text = row.Field(Of String)("me_name")
                dic.Clear()
                dic.Add("dish_name", txt.Text)

                Dim dt = SelectTable("SELECT dish_ingredients FROM dishes WHERE dish_name = @dish_name", dic)

                If dt.Rows.Count > 0 Then
                    Dim dishes = dt.Rows(0).Field(Of String)("dish_ingredients")
                    Dim toolTip As New ToolTip
                    toolTip.SetToolTip(txt, dishes)
                End If
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
            Dim dic As New Dictionary(Of String, Object)
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
                InserTable(table, dic)
            End With
        Next
        ClearControls(tpMenu)
        DataToDgv(sqlMenu, dgvMenu)
        MsgBox("新增/修改完成")
Finish:
        Cursor = Cursors.Default
    End Sub

    '菜單管理-搜尋
    Private Sub btnMenuQuery_Click(sender As Object, e As EventArgs) Handles btnMenuQuery.Click
        Dim sql = $"SELECT DISTINCT b.prod_name,a.me_date,a.me_version,b.prod_id FROM menu a LEFT JOIN product b ON a.me_prod_id=b.prod_id WHERE a.me_date = '{dtMenu.Value}'"
        If cmbProdVers_menu.SelectedItem IsNot Nothing Then sql += $" AND me_version = '{cmbProdVers_menu.SelectedItem}'"
        If cmbProdName_menu.SelectedValue IsNot Nothing Then sql += $" AND me_prod_id = '{cmbProdName_menu.SelectedValue}'"
        DataToDgv(SelectTable(sql), "menu,product", dgvMenu)
    End Sub

    '配餐系統管理-修改
    Private Sub btnModify_dist_sys_Click(sender As Object, e As EventArgs) Handles btnModify_dist_sys.Click
        Cursor = Cursors.WaitCursor
        UpdateTable("distribute_system", BindData("distribute_system"), $"dist_sys_id  = '{txtId_dist_sys.Text}'")
        btnCancel_dist_sys.PerformClick()
        MsgBox("修改成功")
Finish:
        Cursor = Cursors.Default
    End Sub

    '配餐系統管理-取消
    Private Sub btnCancel_dist_sys_Click(sender As Object, e As EventArgs) Handles btnCancel_dist_sys.Click
        '列出所有表格資料        
        DataToDgv(SelectTable(sqlDistributeSystem), "distribute_system", dgvDistSys)
        ClearTabPage(tpDistSys)
    End Sub

    '商品管理-取消
    Private Sub btnProdCancel_Click(sender As Object, e As EventArgs) Handles btnCancel_prod.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        ClearTabPage(tp)
        '顯示所有資料
        DataToDgv(SelectTable(sqlProduct), "product", dgvProduct)
    End Sub

    '商品管理-查詢
    Private Sub btnProdQuery_Click(sender As Object, e As EventArgs) Handles btnProdQuery.Click
        Cursor = Cursors.WaitCursor
        Dim tp = CType(sender, Button).Parent
        Dim sTable As String = "product,product_group"
        Dim sql = sqlProduct + $" WHERE a.prod_name LIKE '%{txtProdQuery.Text}%'"
        DataToDgv(sql, tp.Controls.OfType(Of DataGridView).FirstOrDefault)
        Cursor = Cursors.Default
    End Sub

    Private Sub btnTaboo_Click(sender As Object, e As EventArgs) Handles btnTaboo_cus.Click, btnTaboo_dist.Click
        Dim frm As New frmTaboo
        Dim btn As Button = sender
        Dim name = btn.Name.Replace("btn", "")
        Dim txt = Controls.Find("txt" & name, True).First

        If frm.ShowDialog = DialogResult.OK Then
            txt.Text = frm.ReturnString
        End If
    End Sub

    '報表管理-早寫單
    Private Sub btnDistBreak_report_Click(sender As Object, e As EventArgs) Handles btnDistBreak_report.Click
        DistReport("早")
    End Sub

    '報表管理-午寫單
    Private Sub btnDistLunch_report_Click(sender As Object, e As EventArgs) Handles btnDistLunch_report.Click
        DistReport("午")
    End Sub

    '報表管理-晚寫單
    Private Sub btnDistDinner_report_Click(sender As Object, e As EventArgs) Handles btnDistDinner_report.Click
        DistReport("晚")
    End Sub

    '報表管理-送餐報表-建立報表
    Private Sub btnDelivery_Click(sender As Object, e As EventArgs) Handles btnDelivery.Click
        Dim day = dtpReport.Value.ToString("yyyy-MM-dd")
        Dim meal = grpMeal_delivery.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text

        '取得明日要送的客戶
        Dim dic As New Dictionary(Of String, Object) From {
            {"dist_date", day},
            {"dist_meal", meal}
        }

        Dim newDelivery = SelectTable(
            "SELECT '' AS no, c.cus_name, c.cus_phone, a.dist_address, a.dist_memo, '' AS line, a.dist_id, a.dist_date, b.ord_id " &
            "FROM distribute a " &
            "LEFT JOIN orders b ON a.dist_ord_id = b.ord_id " &
            "LEFT JOIN customer c ON b.ord_cus_id = c.cus_id " &
            "WHERE a.dist_date = @dist_date " &
            "AND a.dist_meal = @dist_meal", dic)

        If newDelivery.Rows.Count = 0 Then
            MsgBox($"{day} 無送餐資訊")
            Exit Sub
        End If

        grpDelivery.Text = $"編輯 {day} {meal}餐 送餐報表"

        Dim delivery As New DataTable

        '找出上一次的報表
        Dim lastReport = SelectTable("SELECT * FROM delivery WHERE del_date < @dist_date ORDER BY del_date DESC", dic)

        If lastReport.Rows.Count > 0 Then
            day = lastReport.Rows(0)("del_date")
            dic.Clear()
            dic("dist_date") = day

            Dim sql = "SELECT b.del_dtl_no, e.cus_name, e.cus_phone, c.dist_address, c.dist_memo, b.del_dtl_line, c.dist_id, c.dist_date, d.ord_id " &
                "FROM delivery a " &
                "LEFT JOIN delivery_detail b ON a.del_id = b.del_dtl_del_id " &
                "LEFT JOIN distribute c ON b.del_dtl_dist_id = c.dist_id " &
                "LEFT JOIN orders d ON c.dist_ord_id = d.ord_id " &
                "LEFT JOIN customer e ON d.ord_cus_id = e.cus_id " &
                "WHERE a.del_date = @dist_date " &
                "AND a.del_meal = @dist_meal"

            '讀取前天的報表
            delivery = SelectTable(sql, dic)
        End If

        '刷新dgv
        dgvDelivery_new.Rows.Clear()
        dgvDelivery_old.Rows.Clear()

        If delivery.Rows.Count > 0 Then
            delivery.Rows.Cast(Of DataRow).ToList.ForEach(Sub(row) dgvDelivery_old.Rows.Add(row.ItemArray))

            '取兩者的交集
            Dim intersectionQuery = From delRow As DataRow In delivery.AsEnumerable
                                    Join newRow As DataRow In newDelivery.AsEnumerable
                                    On delRow("ord_id") Equals newRow("ord_id")
                                    Select
                                        no = delRow("del_dtl_no"),
                                        name = newRow("cus_name"),
                                        phone = newRow("cus_phone"),
                                        address = newRow("dist_address"),
                                        memo = newRow("dist_memo"),
                                        distID = newRow("dist_id"),
                                        distDate = newRow("dist_date"),
                                        ordID = newRow("ord_id"),
                                        line = delRow("del_dtl_line")
                                    Order By no

            '交集的資料放到dgv
            For Each row In intersectionQuery
                Dim rowIndex = dgvDelivery_new.Rows.Add(row.no, row.name, row.phone, row.address, row.memo, row.line, row.distID, row.distDate, row.ordID)
                dgvDelivery_new.Rows(rowIndex).DefaultCellStyle.BackColor = System.Drawing.Color.Orange
            Next
        End If

        '取今日新增的客戶
        For Each row As DataRow In newDelivery.Rows
            Dim id As Integer = Int(row("ord_id"))
            If dgvDelivery_new.Rows.Count = 0 OrElse Not dgvDelivery_new.Rows.Cast(Of DataGridViewRow).Any(Function(r) Int(r.Cells("ord_id").Value) = id) Then
                dgvDelivery_new.Rows.Add(row.ItemArray)
            End If
        Next
    End Sub

    '報表管理-送餐報表-儲存
    Private Sub btnSave_delivery_Click(sender As Object, e As EventArgs) Handles btnSave_delivery.Click
        Cursor = Cursors.WaitCursor

        Dim rows = dgvDelivery_new.Rows

        If rows.Count = 0 Then
            MsgBox("沒有資料需要儲存", MsgBoxStyle.Information, "儲存")
            GoTo Finish
        End If

        Dim d As String = dtpReport.Value.ToString("yyyy-MM-dd")
        Dim meal = grpMeal_delivery.Controls.OfType(Of RadioButton).First(Function(rdo) rdo.Checked).Text

        '提醒覆蓋
        If SelectTable($"SELECT * FROM delivery WHERE del_date = '{d}' AND del_meal = '{meal}'").Rows.Count > 0 Then
            If MsgBox($"{d} {meal}餐 已有資料,是否要覆蓋?", vbYesNo, "儲存") = vbNo Then Exit Sub
        End If

        If SelectTable($"SELECT * FROM delivery WHERE del_date = '{d}' AND del_meal = '{meal}'").Rows.Count = 0 Then
            Dim dicDel = New Dictionary(Of String, Object) From {
                {"del_date", d},
                {"del_meal", meal}
            }

            If Not InserTable("delivery", dicDel) Then GoTo Finish
        End If

        Dim delID As String = SelectTable($"SELECT * FROM delivery WHERE del_date = '{d}' AND del_meal = '{meal}'").Rows(0)("del_id")
        DeleteData("delivery_detail", $"del_dtl_del_id = '{delID}'")

        For i As Integer = 1 To rows.Count
            Dim row = rows(i - 1)
            Dim dicDelDtl = New Dictionary(Of String, Object) From {
                {"del_dtl_del_id", delID},
                {"del_dtl_no", i},
                {"del_dtl_dist_id", row.Cells("配餐編號").Value},
                {"del_dtl_line", row.Cells("路線").Value}
            }

            If Not InserTable("delivery_detail", dicDelDtl) Then GoTo Finish
        Next

        dgvDelivery_new.Rows.Clear()
        MsgBox("儲存成功")

Finish:
        Cursor = Cursors.Default
    End Sub

    '報表管理-送餐報表-匯出報表
    Private Sub btnExport_report_Click(sender As Object, e As EventArgs) Handles btnExport_report.Click
        Dim bytes As Byte()
        Dim day = dtpReport.Value.ToString("yyyy-MM-dd")

        Using ms = New MemoryStream
            bytes = File.ReadAllBytes(Path.Combine(Application.StartupPath, "Report", "送餐.xlsx"))
            ms.Write(bytes, 0, bytes.Length)

            Using exl = SpreadsheetDocument.Open(ms, True)
                Dim wbPart = exl.WorkbookPart
                Dim sstPart = wbPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault
                Dim lst = New List(Of String)

                lst.AddRange({"早", "午", "晚"})

                '寫入每餐排的路線
                For Each meal In lst
                    Dim wsPart As WorksheetPart = wbPart.GetPartById(GetSheetId(exl, meal))
                    Dim ws = wsPart.Worksheet
                    Dim sd = ws.GetFirstChild(Of SheetData)

                    '寫入日期
                    SetCellValue(ws, "A1", day + $" {meal} 送餐報表", sstPart)

                    Dim dic As New Dictionary(Of String, Object) From {
                        {"del_date", day},
                        {"del_meal", meal}
                    }
                    Dim rows = SelectTable("SELECT * " &
                                           "FROM delivery a " &
                                           "LEFT JOIN delivery_detail b ON a.del_id = b.del_dtl_del_id " &
                                           "LEFT JOIN distribute c ON b.del_dtl_dist_id = c.dist_id " &
                                           "LEFT JOIN orders d ON c.dist_ord_id = d.ord_id " &
                                           "LEFT JOIN customer e ON d.ord_cus_id = e.cus_id " &
                                           "LEFT JOIN product f ON d.ord_prod_id = f.prod_id " &
                                           "LEFT JOIN product_group g ON f.prod_prod_grp_id	= g.prod_grp_id " &
                                           "WHERE a.del_date = @del_date " &
                                           "AND a.del_meal = @del_meal " &
                                           "ORDER BY b.del_dtl_no", dic).Rows

                    For i As Integer = 0 To rows.Count - 1
                        '編號
                        SetCellValue(ws, "A" + (i + 3).ToString, IIf(IsDBNull(rows(i)("del_dtl_no")), "", rows(i)("del_dtl_no")), sstPart)
                        '姓名
                        SetCellValue(ws, "B" + (i + 3).ToString, rows(i)("cus_name"), sstPart)
                        '餐飲種類
                        SetCellValue(ws, "D" + (i + 3).ToString, rows(i)("prod_grp_name"), sstPart)
                        '電話
                        SetCellValue(ws, "E" + (i + 3).ToString, rows(i)("cus_phone"), sstPart)
                        '送餐地址
                        SetCellValue(ws, "F" + (i + 3).ToString, rows(i)("dist_address"), sstPart)
                        '送餐注意事項
                        SetCellValue(ws, "G" + (i + 3).ToString, rows(i)("dist_memo"), sstPart)
                        '路線
                        SetCellValue(ws, "H" + (i + 3).ToString, If(IsDBNull(rows(i)("del_dtl_line")), "", rows(i)("del_dtl_line")), sstPart)
                    Next
                Next

                exl.Save()
            End Using

            bytes = ms.ToArray
        End Using

        Dim saveFileDialog As New SaveFileDialog With {
            .Filter = "Excel檔 (*.xlsx)|*.xlsx",
            .Title = "選取存檔位置",
            .FileName = day & "送餐.xlsx"
        }

        If saveFileDialog.ShowDialog = DialogResult.OK Then
            Try
                File.WriteAllBytes(saveFileDialog.FileName, bytes)
            Catch ex As Exception
                MsgBox(ex.Message, Title:=Reflection.MethodBase.GetCurrentMethod.Name)
                Exit Sub
            End Try

            MsgBox("報表建立成功!")
        End If
    End Sub

    '報表管理-送餐報表-餐種-點擊餐種搜尋
    Private Sub SerchDelivery(sender As Object, e As EventArgs) Handles rdoBreakfast_delivery.CheckedChanged, rdoLunch_delivery.CheckedChanged, rdoDinner_delivery.CheckedChanged, dtpReport.ValueChanged
        Dim rdo As RadioButton

        If TypeOf sender Is RadioButton Then
            rdo = sender

            If rdo.Checked = False OrElse String.IsNullOrEmpty(rdo.Text) Then Exit Sub
        Else
            rdo = grpMeal_delivery.Controls.OfType(Of RadioButton).FirstOrDefault(Function(r) r.Checked)

            If rdo Is Nothing Then Exit Sub
        End If

        Dim day = dtpReport.Value.ToString("yyyy-MM-dd")
        Dim meal = rdo.Text
        Dim dic As New Dictionary(Of String, Object) From {
            {"del_date", day},
            {"del_meal", meal}
        }
        Dim delivery = SelectTable(
            "SELECT b.del_dtl_no, e.cus_name, e.cus_phone, c.dist_address, c.dist_memo, b.del_dtl_line, c.dist_id, c.dist_date, d.ord_id " &
            "FROM delivery a " &
            "LEFT JOIN delivery_detail b ON a.del_id = b.del_dtl_del_id " &
            "LEFT JOIN distribute c ON b.del_dtl_dist_id = c.dist_id " &
            "LEFT JOIN orders d ON c.dist_ord_id = d.ord_id " &
            "LEFT JOIN customer e ON d.ord_cus_id = e.cus_id " &
            "WHERE a.del_date = @del_date " &
            "AND a.del_meal = @del_meal " &
            "ORDER BY b.del_dtl_no", dic)

        dgvDelivery_new.Rows.Clear()
        delivery.Rows.Cast(Of DataRow).ToList.ForEach(Sub(row) dgvDelivery_new.Rows.Add(row.ItemArray))
        grpDelivery.Text = $"{day} {meal}餐 送餐報表"
    End Sub

    ''' <summary>
    ''' 生成打餐報表
    ''' </summary>
    ''' <param name="meal">餐種</param>
    Private Sub DistReport(meal As String)
        Dim sheetName As String
        Dim sourceFileName As String

        Select Case meal
            Case "早"
                sourceFileName = "隔天早寫單"
                sheetName = "隔天早大報表"
            Case "午"
                sourceFileName = "午寫單"
                sheetName = "午大報表"
            Case "晚"
                sourceFileName = "晚寫單"
                sheetName = "晚大報表"
            Case Else
                Exit Sub
        End Select

        Dim bytes As Byte()

        Using ms = New MemoryStream
            bytes = File.ReadAllBytes(Application.StartupPath + $"\Report\{sourceFileName}.xlsx")
            ms.Write(bytes, 0, bytes.Length)

            Using exl = SpreadsheetDocument.Open(ms, True)
                Dim wbPart = exl.WorkbookPart
                Dim sstPart As SharedStringTablePart = wbPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                Dim wsPart As WorksheetPart = wbPart.GetPartById(GetSheetId(exl, sheetName))
                Dim ws = wsPart.Worksheet
                Dim sd = ws.GetFirstChild(Of SheetData)
                Dim day = dtpReport.Value.ToString("d")

                '寫入日期
                SetCellValue(ws, "A1", day + $" {meal}餐 月子餐打餐報表", sstPart)

                '找出當日所有配餐(禁忌要另外解析)
                Dim dic As New Dictionary(Of String, Object) From {
                    {"dist_date", day},
                    {"dist_meal", meal}
                }
                Dim rows = SelectTable("SELECT d.prod_grp_aka, e.cus_name, a.dist_customized, a.dist_drink, e.cus_tabo_id, a.dist_other, e.cus_plate " +
                                       "FROM distribute a " +
                                       "LEFT JOIN orders b ON a.dist_ord_id = b.ord_id " +
                                       "LEFT JOIN product c ON b.ord_prod_id = c.prod_id " +
                                       "LEFT JOIN product_group d ON c.prod_prod_grp_id = d.prod_grp_id " +
                                       "LEFT JOIN customer e ON e.cus_id = b.ord_cus_id " +
                                       "LEFT JOIN taboo f ON f.tabo_id = e.cus_tabo_id " +
                                       "WHERE dist_date = @dist_date " +
                                       "AND dist_meal = @dist_meal " +
                                       "ORDER BY d.prod_grp_aka DESC", dic).Rows

                For i As Integer = 0 To rows.Count - 1
                    '編號
                    SetCellValue(ws, "A" + (i + 3).ToString, i + 1, sstPart)
                    '產品簡稱
                    SetCellValue(ws, "B" + (i + 3).ToString, rows(i)("prod_grp_aka") + rows(i)("cus_plate"), sstPart)
                    '客戶姓名
                    SetCellValue(ws, "C" + (i + 3).ToString, rows(i)("cus_name"), sstPart)
                    '加減
                    SetCellValue(ws, "D" + (i + 3).ToString, rows(i)("dist_customized"), sstPart)
                    '飲品需求
                    SetCellValue(ws, "E" + (i + 3).ToString, rows(i)("dist_drink"), sstPart)
                    '禁忌
                    SetCellValue(ws, "F" + (i + 3).ToString, GetTaboo(If(IsDBNull(rows(i)("cus_tabo_id")), "", rows(i)("cus_tabo_id"))), sstPart)
                    '備註
                    SetCellValue(ws, "G" + (i + 3).ToString, rows(i)("dist_other"), sstPart)
                Next

                '重新計算公式
                Dim wb = wbPart.Workbook
                Dim cp = wb.CalculationProperties

                cp.ForceFullCalculation = True
                wb.Save()
                exl.Save()
            End Using

            bytes = ms.ToArray()
        End Using

        '另存新檔
        SaveFileDialog1.FileName = dtpReport.Value.ToString("yyyy.MM.dd") + $"{meal}寫單.xlsx"

        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            Try
                File.WriteAllBytes(SaveFileDialog1.FileName, bytes)
            Catch ex As Exception
                MsgBox(ex.Message, Title:=Reflection.MethodBase.GetCurrentMethod.Name)
                Exit Sub
            End Try
            MsgBox("報表建立成功!")
        End If
    End Sub

    ''' <summary>
    ''' 取得禁忌名稱
    ''' </summary>
    ''' <param name="str">customer 裡的 cus_tabo_id</param>
    ''' <returns></returns>
    Private Function GetTaboo(str As String) As String
        If str = "" Then Return str
        Dim lstSource = str.Split(",").ToList
        Dim lstValue As New List(Of String)
        For Each c In lstSource
            lstValue.Add(SelectTable($"SELECT tabo_name FROM taboo WHERE tabo_id = {c}").Rows(0)("tabo_name").ToString)
        Next
        Return String.Join(",", lstValue)
    End Function

    Private Sub btnIngredients_Click(sender As Object, e As EventArgs) Handles btnIngredients.Click
        Dim frm As New frmTaboo
        If frm.ShowDialog = DialogResult.OK Then txtingredients.Text = frm.ReturnString
    End Sub

    '菜品管理-取消
    Private Sub btnCancel_dish_Click(sender As Object, e As EventArgs) Handles btnCancel_dish.Click
        DataToDgv("SELECT * FROM dishes", dgvDishes)
        ClearControls(tpDishes)
    End Sub

    '菜品管理-新增
    Private Sub btnInsert_dish_Click(sender As Object, e As EventArgs) Handles btnInsert_dish.Click
        Dim tp As TabPage = CType(sender, Button).Parent
        Dim dic As Dictionary(Of String, Object) = CheckDishes(sender)
        If dic Is Nothing Then Exit Sub
        If InserTable(tp.Tag, dic) Then
            tp.Controls.OfType(Of Button).First(Function(btn) btn.Text = "取  消").PerformClick()
            MsgBox("新增成功")
        End If
    End Sub

    Private Function CheckDishes(sender As Button)
        Dim dicReq As New Dictionary(Of String, Object) From {
             {"菜名", txtDishes},
             {"食材", txtingredients}
         }
        If Not CheckRequiredCol(dicReq) Then Return Nothing

        Dim tp As TabPage = sender.Parent
        Dim dic As New Dictionary(Of String, String)
        dic = tp.Controls.OfType(Of Forms.Control).Where(Function(ctrl) Not String.IsNullOrEmpty(ctrl.Tag) AndAlso ctrl.Tag <> "comp_id" AndAlso Not String.IsNullOrWhiteSpace(ctrl.Text)).
            ToDictionary(Function(ctrl) ctrl.Tag.ToString, Function(ctrl) ctrl.Text)
        Return dic
    End Function

    '菜品管理-dgv點擊
    Private Sub dgvDishes_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvDishes.CellMouseClick
        Dim tp As TabPage = sender.Parent
        ClearControls(tp)
        GetDataToControls(tp, sender.SelectedRows(0))
    End Sub

    '菜品管理-修改
    Private Sub btnModify_dish_Click(sender As Object, e As EventArgs) Handles btnModify_dish.Click
        Dim tp As TabPage = sender.Parent
        Dim dic As Dictionary(Of String, Object) = CheckDishes(sender)
        If dic Is Nothing Then Exit Sub
        If UpdateTable(tp.Tag, dic, $"{txtDishes.Tag} = '{txtDishes.Text}'") Then
            tp.Controls.OfType(Of Button).First(Function(btn) btn.Text = "取  消").PerformClick()
            MsgBox("修改成功")
        End If
    End Sub

    '菜品管理-刪除
    Private Sub btnDelete_dish_Click(sender As Object, e As EventArgs) Handles btnDelete_dish.Click
        Dim tp As TabPage = sender.Parent
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.No Then Exit Sub
        If DeleteData(tp.Tag, $"{txtDishes.Tag} = '{txtDishes.Text}'") Then
            tp.Controls.OfType(Of Button).First(Function(btn) btn.Text = "取  消").PerformClick()
            MsgBox("刪除成功")
        End If
    End Sub

    '菜品管理-查詢
    Private Sub btnQuery_dish_Click(sender As Object, e As EventArgs) Handles btnQuery_dish.Click
        Dim tp As TabPage = sender.Parent
        DataToDgv($"SELECT * FROM {tp.Tag} WHERE dish_name LIKE '%{txtDishes.Text}%' ", tp.Controls.OfType(Of DataGridView).First)
    End Sub

    '菜品管理-查詢未登入食材的菜名
    Private Sub btnNoIngredients_Click(sender As Object, e As EventArgs) Handles btnNoIngredients.Click
        Dim tp As TabPage = sender.Parent
        DataToDgv($"SELECT * FROM {tp.Tag} WHERE dish_ingredients = '' ", tp.Controls.OfType(Of DataGridView).First)
    End Sub

    Private Sub btnMealAdj_Click(sender As Object, e As EventArgs) Handles btnMealAdj.Click
        If txtOrdID_dist.Text = "" Then
            MsgBox("請選擇對象")
            Exit Sub
        End If

        Using form As New frmMealAdjustments With {.OrderID = txtOrdID_dist.Text}
            form.ShowDialog()
            CountNotConfigured()
        End Using

    End Sub

    Private Sub btnCancel_money_ord_Click(sender As Object, e As EventArgs) Handles btnCancel_money_ord.Click
        ClearControls(tpMoney)
        DataToDgv(sqlMoney, dgvOrder_money)
        dgvMoney.DataSource = Nothing
    End Sub

    Private Sub btnAddNewOrder_Click(sender As Object, e As EventArgs) Handles btnAddNewOrder.Click
        btnCancel_order.PerformClick()
        txtCusName_order.Text = txtCusName_cus.Text
        txtPhone_order.Text = txtPhone_cus.Text
        txtAddrBreak.Text = txtAddressOne_cus.Text
        txtCusID_ord.Text = txtCusID.Text
        TabControl1.SelectedTab = tpOrder
    End Sub

    Private Sub txtCount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCount.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' 取消此次鍵盤操作，以阻止非數字的輸入
            e.Handled = True
        End If
    End Sub
End Class
