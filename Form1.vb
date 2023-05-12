Imports System.Configuration
Imports System.IO
Imports Microsoft.Office.Interop
Imports MySql.Data.MySqlClient

Public Class frmMain
    Dim conn As MySqlConnection
    Dim mAdapter As MySqlDataAdapter
    'Dim dt As New System.Data.DataTable()

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedTab.Name = "TP_Logout" Then
            If MessageBox.Show("確定要登出嗎??", "登出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.OK Then
                Me.Hide()
                LoginForm1.Show()
            Else
                Me.TabControl1.SelectedTab = tpCustomer
            End If
        End If
    End Sub

    'Sub ReadDataGridWidth(dgv As String)
    '    Dim myObject As Object

    '    Select Case dgv
    '        Case "DGV_Customer"
    '            myObject = Me.DGV_Customer
    '        Case "DGV_Product"
    '            myObject = Me.DGV_Product
    '        Case "DGV_Staff"
    '            myObject = Me.DGV_Staff
    '        Case "DGV_Order"
    '            myObject = Me.DGV_Order
    '        Case "DGV_Taboo"
    '            myObject = Me.DGV_Taboo
    '        Case "DGV_TabooClass"
    '            myObject = Me.DGV_TabooClass
    '        Case "DGV_ProductClass"
    '            myObject = Me.DGV_ProductClass
    '        Case "DGV_Parameter"
    '            myObject = Me.DGV_Parameter
    '    End Select
    '    With myObject
    '        Dim tmpWidth As String
    '        Dim objStreamReader As StreamReader
    '        Try
    '            objStreamReader = New StreamReader(dgv + ".set", False)
    '            tmpWidth = objStreamReader.ReadLine()
    '            objStreamReader.Close()
    '            Dim tmpW() = Split(tmpWidth, ",")
    '            For j = 1 To .ColumnCount
    '                .Columns(j - 1).Width = tmpW(j - 1)
    '            Next
    '        Catch ex As Exception

    '        End Try
    '    End With
    'End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '設定連線
        Dim myConnectionString As String = ConfigurationManager.AppSettings("myConnectionString").ToString
        conn = New MySqlConnection(myConnectionString)

        InitDataGrid()

        '初始化商品管理的商品分類
        Dim items() As String = {"套餐", "單點"}
        cmbProdType_product.Items.AddRange(items)
    End Sub

    Private Sub btnPermAdd_Click(sender As Object, e As EventArgs) Handles btnPermAdd.Click
        'Dim oCtrl As Control
        'For Each oCtrl In tpPermissions.Controls
        '    If oCtrl === CheckBox Then

        '    End If
        'Next
        'Dim sSQL As String

        'sSQL = "INSERT INTO PERMISSION ()"
    End Sub


    'Private Sub Setup_retrieve()
    '    'DGV_Parameter.Rows.Clear()
    '    'SQL STMT
    '    Dim sql As String = "SELECT * FROM sys_para"
    '    cmd = New MySqlCommand(sql, conn)
    '    'OPEN CON,RETRIEVE,FILL,DGVIEW
    '    Try
    '        conn.Open()
    '        adapter = New MySqlDataAdapter(cmd)
    '        Dim mydt As New System.Data.DataTable()
    '        adapter.Fill(mydt)
    '        'FILL DGVIEW
    '        For Each row In mydt.Rows
    '            Setup_Populate(row(0), row(1), row(2), row(3))
    '        Next
    '        conn.Close()
    '        'CLEAR DT
    '        mydt.Rows.Clear()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '        conn.Close()
    '    End Try

    'End Sub

    'Private Sub Setup_Populate(sp_sn As String, sp_name As String, sp_type As String, sp_option As String)
    '    Dim row As String() = New String() {sp_sn, sp_name, sp_type, sp_option}

    '    'ADD ROW TO ROWS COLLEC
    '    'DGV_Parameter.Rows.Add(row)
    'End Sub

    'Private Sub DGV_Setup_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs)
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Sub SaveDataGridWidth(dgv As String)
    '    Dim myObject As Object

    '    Select Case dgv
    '        Case "DGV_Customer"
    '            myObject = Me.DGV_Customer
    '        Case "DGV_Product"
    '            myObject = Me.DGV_Product
    '        Case "DGV_Staff"
    '            myObject = Me.DGV_Staff
    '        Case "DGV_Order"
    '            myObject = Me.DGV_Order
    '        Case "DGV_Taboo"
    '            myObject = Me.DGV_Taboo
    '        Case "DGV_TabooClass"
    '            myObject = Me.DGV_TabooClass
    '        Case "DGV_ProductClass"
    '            myObject = Me.DGV_ProductClass
    '        Case "DGV_Parameter"
    '            myObject = Me.DGV_Parameter

    '    End Select
    '    With myObject
    '        Dim tmpWidth As String
    '        tmpWidth = .Columns(0).Width.ToString
    '        For j = 2 To .ColumnCount
    '            tmpWidth = tmpWidth + "," + .Columns(j - 1).Width.ToString
    '        Next
    '        Dim objStreamWriter As StreamWriter
    '        objStreamWriter = New StreamWriter(dgv + ".set", False)
    '        objStreamWriter.WriteLine(tmpWidth)
    '        objStreamWriter.Close()
    '    End With
    'End Sub

    'Private Sub DGV_Customer_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_Customer.ColumnWidthChanged
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_Product_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_Product.ColumnWidthChanged
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_Staff_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_Staff.ColumnWidthChanged
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_Order_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_Order.ColumnWidthChanged
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_Parameter_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs)
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_ProductClass_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs)
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_TabooClass_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_TabooClass.ColumnWidthChanged
    '    SaveDataGridWidth(sender.name)
    'End Sub

    'Private Sub DGV_Taboo_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs)
    '    SaveDataGridWidth(sender.name)
    'End Sub


    '初始化DataGrid欄位
    Private Sub InitDataGrid()
        '客戶管理
        'With dgvCustomer
        '    .Columns.Add("", "編號")
        '    .Columns.Add("", "姓名")
        '    .Columns.Add("", "電話")
        '    .Columns.Add("", "手機")
        '    .Columns.Add("", "公司電話")
        '    .Columns.Add("", "地址")
        '    .Columns.Add("", "早餐送餐地址")
        '    .Columns.Add("", "午餐送餐地址")
        '    .Columns.Add("", "晚餐送餐地址")
        '    .Columns.Add("", "床號")
        '    .Columns.Add("", "備註")
        '    .AutoResizeColumnHeadersHeight()
        'End With

        'todo 顯示所有客戶資料
        '將table資料塞到dgv
        Dim cmd As New MySqlCommand("SELECT * FROM customer", conn)
        Dim dtData As New DataTable()
        conn.Open()
        Dim adapter As New MySqlDataAdapter(cmd)
        adapter.Fill(dtData)
        With dgvCustomer
            .DataSource = dtData
            .AutoResizeColumnHeadersHeight()
        End With

        '用table欄位的備註將dgv的欄位改名
        cmd = New MySqlCommand("SELECT COLUMN_NAME, COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'tingyi' AND TABLE_NAME = 'customer'", conn)
        adapter = New MySqlDataAdapter(cmd)
        Dim dtColName As New DataTable()
        adapter.Fill(dtColName)
        For Each col As DataGridViewColumn In dgvCustomer.Columns
            Dim row As DataRow = dtColName.AsEnumerable().FirstOrDefault(Function(dr) dr("COLUMN_NAME").ToString() = col.Name)
            If row IsNot Nothing Then
                col.HeaderText = row("COLUMN_COMMENT").ToString()
            End If
        Next
        conn.Close()

        '商品管理
        With dgProduct
            .Columns.Add("", "編號")
            .Columns.Add("", "商品群組")
            .Columns.Add("", "商品分類")
            .Columns.Add("", "品名")
            .Columns.Add("", "餐種")
            .Columns.Add("", "售價")
            .Columns.Add("", "成本")
            .Columns.Add("", "備註")
            .Rows.Add("1", "月子餐", "套餐", "經典月子餐", "早,午,晚", "3200", "1700")
            .Rows.Add("2", "月子餐_優惠", "套餐", "經典月子餐_優惠", "早,午,晚", "3990", "3000", "7日以上")
            .Rows.Add("3", "調養餐", "套餐", "小產調養餐", "早,午,晚", "2900", "1400")
            .Rows.Add("4", "月子餐", "單點", "經典月子早餐", "早", "1000", "500")
            .Rows.Add("5", "月子餐", "單點", "經典月子午餐", "午", "1100", "600")
            .Rows.Add("6", "月子餐", "單點", "經典月子晚餐", "晚", "1100", "600")
            .Rows.Add("7", "調養餐", "單點", "小產調養早餐", "早", "900", "400")
            .Rows.Add("8", "調養餐", "單點", "小產調養午餐", "午", "1000", "500")
            .Rows.Add("9", "調養餐", "單點", "小產調養晚餐", "晚", "1000", "500")
            .Rows.Add("10", "月子餐_優惠", "單點", "經典月子早餐_優惠", "晚", "1000", "500")
        End With

        txtProdName_product.Text = "經典月子餐"
        cmbProdType_product.Text = "套餐"
        cmbProdGroup_product.Text = "月子餐"
        txtProdPrice_product.Text = "3200"
        txtProdCost_product.Text = "1700"
        chkBleak_product.Checked = True
        chkLunch_product.Checked = True
        chkDinner_product.Checked = True

        '菜單管理
        With DGV_Menu
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

        '訂單管理
        With dgOrder
            .Columns.Add("", "訂單編號")
            .Columns.Add("", "客戶姓名")
            .Columns.Add("", "手機")
            .Columns.Add("", "訂單日期")
            .Columns.Add("", "商品名稱")
            .Columns.Add("", "早餐")
            .Columns.Add("", "午餐")
            .Columns.Add("", "晚餐")
            .Columns.Add("", "折讓金額")
            .Columns.Add("", "金額")
            .Columns.Add("", "預計送餐日")
            .Columns.Add("", "禁忌食物")
            .Columns.Add("", "備註")

            .Rows.Add("1", "陳小姐", "0918-123123", "2023/5/5", "小產調養餐", "10", "10", "10", "0", "29000", "2023/5/11", "蝦,花生")
            .Rows.Add("2", "李小姐", "0918-456456", "2023/5/6", "經典月子餐", "10", "10", "0", "0", "21000", "2023/5/21")
            .Rows.Add("3", "王太太", "0918-852852", "2023/5/7", "月子早餐", "2688", "1")
        End With

        txtCusName_order.Text = "陳小姐"
        cmdProdName_order.Text = "月子餐30日"
        txtPrice_order.Text = "57000"
        txtPhone_order.Text = "0918-123123"
        txtTaboo.Text = "蝦,花生"
        txtCount.Text = "90"

        '配餐管理
        txtCusName_dist.Text = "陳小姐"
        txtPhone_dist.Text = "0918-123123"

        '財務管理
        With dgMoney
            .Columns.Add("", "編號")
            .Columns.Add("", "日期")
            .Columns.Add("", "客戶姓名")
            .Columns.Add("", "客戶手機")
            .Columns.Add("", "訂單編號")
            .Columns.Add("", "商品名稱")
            .Columns.Add("", "收款金額")
            .Columns.Add("", "收款類型")
            .Columns.Add("", "收款說明")
            .Rows.Add("1", "2023-03-01", "陳小姐", "0918-123123", "1", "月子餐30日", "10000", "訂金", "123")
            .Rows.Add("2", "2023-03-05", "3", "王太太", "月子早餐", "2688", "全款")
        End With

        txtCusName_money.Text = "陳小姐"
        txtPhone_money.Text = "0918-123123"
        txtOrdID_money.Text = "1"
        dtMoney.Value = "2023-03-01"
        txtMoney.Text = "10000"
        txtMonType.Text = "訂金"
        txtMonMemo.Text = "123"

        '員工管理
        With DGV_Staff
            .Columns.Add("", "編號")
            .Columns.Add("", "姓名")
            .Columns.Add("", "電話")
            .Columns.Add("", "手機")
            .Columns.Add("", "地址")
            .Columns.Add("", "帳號")
            .Columns.Add("", "職位")
            .Columns.Add("", "備註")
            .Rows.Add("1", "小陳", "05-1111111", "0900-123123", "嘉義縣大林鎮中山路1號", "user1")
            .Rows.Add("2", "小李", "05-2222222", "0900-456456", "嘉義縣東區世賢路二段567號", "user2", "廚師")
            .Rows.Add("3", "老王", "05-3333333", "0900-852852", "雲林縣斗六市大學路52號", "user3")
            .Rows.Add("4", "小張", "05-5555555", "0900-147147", "嘉義縣太保市市政路23號", "user4")
            .Rows.Add("5", "小高", "05-6666666", "0900-369369", "雲林縣虎尾鎮中正路100號", "user5")
        End With

        txtEmpName_emp.Text = "小李"
        txtEmpTel.Text = "05-2222222"
        txtEmpPhone_Emp.Text = "0900-456456"
        txtEmpAddr.Text = "嘉義縣東區世賢路二段567號"
        txtEmpMemo.Text = ""
        txtEmpAcct.Text = "user2"
        txtPsw.Text = "********"
        txtPswCheck.Text = "********"
        cmbEmpPos_emp.Text = "系統管理員"

        '禁忌食物管理
        With dgTaboo
            .Columns.Add("", "編號")
            .Columns.Add("", "分類")
            .Columns.Add("", "名稱")
            .Rows.Add("1", "雞", "雞屁股")
            .Rows.Add("2", "豬", "豬舌頭")
            .Rows.Add("3", "魚", "魚眼睛")
        End With

        cmbTaboClass.Text = "雞"
        txtTaboName.Text = "雞屁股"

        '權限管理
        With dgPermission
            .Columns.Add("", "編號")
            .Columns.Add("", "職位")
            .Columns.Add("", "客戶管理")
            .Columns.Add("", "商品管理")
            .Columns.Add("", "菜單管理")
            .Columns.Add("", "訂單管理")
            .Columns.Add("", "報表管理")
            .Columns.Add("", "財務管理")
            .Columns.Add("", "員工管理")
            .Columns.Add("", "權限管理")
            .Columns.Add("", "配餐管理")
            .Rows.Add("1", "系統管理員", "Y", "Y", "Y", "Y", "Y", "Y", "Y", "Y", "Y")
            .Rows.Add("2", "新人", "N", "N", "N", "N", "N", "N", "N", "N", "Y")
        End With

        cmbPosition.Text = "新人"
        chkCustomer.Checked = True
        chkItem.Checked = True
        chkItem_detail.Checked = True
        chkOrders.Checked = True
        chkReport.Checked = True
        chkForbid.Checked = True
        chkEmployee.Checked = True
        chkPermission.Checked = True
        chkDistr.Checked = True
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles btnDistDel.Click
        MsgBox("是否往後延一餐?", vbYesNo)
    End Sub

    Private Sub btnTaboo_Click_1(sender As Object, e As EventArgs) Handles btnTaboo.Click
        frmTaboo.Show()
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles btnDistModify.Click
        MsgBox("是否更改後續配置?", vbYesNo)
    End Sub

    Private Sub btnCusInsert_Click(sender As Object, e As EventArgs) Handles btnCusInsert.Click
        '檢查不能空值的欄位
        Dim txts As New Collection
        txts.Add(txtCusName_cus)
        txts.Add(txtPhone_cus)
        If Empty_Textbox(txts) Then Exit Sub

        '去頭尾空白
        btnCusInsert.Parent.Controls.OfType(Of TextBox).ToList().ForEach(Sub(txt) txt.Text = Trim(txt.Text))

        '塞資料
        Dim dicData As New Dictionary(Of String, String)

        With dicData
            .Add("cus_name", txtCusName_cus.Text) '客戶姓名
            .Add("cus_phone", txtPhone_cus.Text) '客戶手機
            .Add("cus_tel_home", txtTelHome.Text) '客戶住家電話
            .Add("cus_tel_comp", txtTelComp.Text) '客戶公司電話
            .Add("cus_addr_home", txtAddrHome.Text) '客戶住家地址
            .Add("cus_addr_break", txtAddrBreak.Text) '早餐地址
            .Add("cus_addr_lunch", txtAddrLunch.Text) '午餐地址
            .Add("cus_addr_dinner", txtAddrDinner.Text) '晚餐地址
            .Add("cus_bed", txtBed.Text) '床號
            .Add("cus_memo", txtMemo_cus.Text)
        End With
        InserTable("customer", dicData)
    End Sub
    '檢查textbox是否為空
    Private Function Empty_Textbox(ctrls As Collection) As Boolean
        For Each ctrl As TextBox In ctrls
            If String.IsNullOrWhiteSpace(ctrl.Text) Then
                MsgBox(ctrl.Tag + "不能空白")
                ctrl.Focus()
                Return True
            End If
        Next

        Return False
    End Function

    Private Sub InserTable(sTableName As String, dicData As Dictionary(Of String, String))
        Dim cmd As New MySqlCommand($"INSERT INTO {sTableName} ({String.Join(",", dicData.Keys)}) VALUES ({String.Join(",", dicData.Values.Select(Function(x) $"'{x}'"))})", conn)
        Try
            conn.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                MsgBox("新增成功")
                btnCusCancel.PerformClick()
                'todo 重新搜尋新增目標
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conn.Close()
    End Sub

    Private Sub btnCusModify_Click(sender As Object, e As EventArgs) Handles btnCusModify.Click
        MsgBox("修改成功")
    End Sub

    Private Sub btnCusDelete_Click(sender As Object, e As EventArgs) Handles btnCusDelete.Click
        If MsgBox("確定要刪除?", vbYesNo, "警告") = MsgBoxResult.Yes Then
            MsgBox("刪除成功")
        End If

    End Sub

    Private Sub btnProdAdd_Click(sender As Object, e As EventArgs) Handles btnProdAdd.Click
        MsgBox("新增成功")
    End Sub

    Private Sub btnProdModify_Click(sender As Object, e As EventArgs) Handles btnProdModify.Click
        MsgBox("修改成功")
    End Sub

    Private Sub btnProdDelete_Click(sender As Object, e As EventArgs) Handles btnProdDelete.Click
        MsgBox("刪除成功")
    End Sub
    '客戶管理-查詢
    Private Sub btnCusQuery_Click(sender As Object, e As EventArgs) Handles btnCusQuery.Click
        'msSQL = "SELECT * FROM customer"
        'mSQLCmd = New MySqlCommand(msSQL, conn)

        'Try
        '    conn.Open()
        '    mAdapter = New MySqlDataAdapter(mSQLCmd)
        '    Dim dt As New DataTable()
        '    mAdapter.Fill(dt)
        '    Dim row As Object
        '    Dim i As Int16
        '    For Each row In dt.Rows
        '        dgCustomer.Rows.Add(row(i))
        '    Next

        'Catch ex As Exception
        '    Debug.Print(ex.Message)
        'End Try
        'conn.Close()
    End Sub
    '清除鍵,清除畫面
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles btnCusCancel.Click, btnProdCancel.Click, btnMenuCancel.Click, btnOrdCancel.Click, btnMonCancel.Click, btnEmpCancel.Click, btnTaboCancel.Click, btnPermCancel.Click, btnDistCancel.Click

        Dim btn As Button = CType(sender, Button)

        ClearTabPage(btn.Parent)
    End Sub
    '清除TabPage裡的控制項內容
    Private Sub ClearTabPage(tabpage As TabPage)
        Dim ctrl As Control
        For Each ctrl In tabpage.Controls
            If TypeOf ctrl Is GroupBox Then
                Dim grp As GroupBox = CType(ctrl, GroupBox)
                ClearGroupBox(grp)
            ElseIf TypeOf ctrl Is TabControl Then '取得TabControl裡的控制項
                Dim tc As TabControl = CType(ctrl, TabControl)
                Dim tp As TabPage
                For Each tp In tc.Controls
                    ClearTabPage(tp)
                Next
            End If

            ClearControl(ctrl)
        Next
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

    Private Sub btnMenuExcel_Click(sender As Object, e As EventArgs) Handles btnMenuExcel.Click
        'Dim excelApp As New Excel.Application

        'Dim workbook As Excel.Workbook = excelApp.Workbooks.Open("C:\ExcelFile.xlsx")
        'Dim worksheet As Excel.Worksheet = workbook.Sheets("Sheet1")

        'Dim cellValue As String = worksheet.Cells(1, 1).Value

        'workbook.Close(False)
        'Marshal.ReleaseComObject(workbook)
        'Marshal.ReleaseComObject(excelApp)

    End Sub

    Private Sub frmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        LoginForm1.Close()
    End Sub

    Private Sub cmdProdName_order_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmdProdName_order.SelectedValueChanged
        '更新商品分類
        '若商品分類是套餐則顯示"三餐"

    End Sub
End Class
