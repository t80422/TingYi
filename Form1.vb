Imports System.Configuration
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class frmMain
    Dim conn As MySqlConnection
    Dim msSQL As String
    Dim mSQLCmd As MySqlCommand
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
        InitDataGrid()

        '設定連線
        Dim myConnectionString As String = ConfigurationSettings.AppSettings("myConnectionString").ToString
        conn = New MySqlConnection(myConnectionString)


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

    Private Sub InitDataGrid()
        '初始化DataGrid欄位
        '客戶管理
        With dgCustomer
            .Columns.Add("", "編號")
            .Columns.Add("", "姓名")
            .Columns.Add("", "電話")
            .Columns.Add("", "手機")
            .Columns.Add("", "公司電話")
            .Columns.Add("", "地址")
            .Columns.Add("", "早餐送餐地址")
            .Columns.Add("", "午餐送餐地址")
            .Columns.Add("", "晚餐送餐地址")
            .Columns.Add("", "床號")
            .Columns.Add("", "備註")
            .Rows.Add("1", "陳小姐", "05-1234567", "0918-123123", "05-5885888", "嘉義縣大林鎮中山路1號", "嘉義縣大林鎮中正路123號", "嘉義縣大林鎮中山路1號", "嘉義縣大林鎮中山路1號", "B01")
            .Rows.Add("2", "李小姐", "05-2222222", "0918-456456", "05-5456456", "嘉義縣東區世賢路二段567號", "", "嘉義縣東區世賢路二段567號", "嘉義縣東區世賢路二段567號", "", "到達前10分鐘請電話通知")
            .Rows.Add("3", "王太太", "05-3852852", "0918-852852", "05-5852258", "雲林縣斗六市大學路52號", "嘉義縣大林鎮中正路123號", "嘉義縣東區世賢路二段567號", "雲林縣斗六市大學路52號", "A01", "到達前10分鐘請電話通知")
            .Rows.Add("4", "張女士", "05-5147741", "0918-147147", "05-5456741", "嘉義縣太保市市政路23號", "", "嘉義縣太保市市政路23號", "嘉義縣太保市市政路23號")
            .Rows.Add("5", "高小姐", "05-6951159", "0918-369369", "05-5951159", "雲林縣虎尾鎮中正路100號", "", "雲林縣虎尾鎮中正路100號", "雲林縣虎尾鎮中正路100號")
            .AutoResizeColumnHeadersHeight()
        End With

        txtCusName_cus.Text = "李小姐"
        txtTelHom_cus.Text = "05-2222222"
        txtPhone_cus.Text = "0918-456456"
        txtTelCom_cus.Text = "05-5456456"
        txtAddr_cus.Text = "嘉義縣東區世賢路二段567號"
        txtAddrBla_cus.Text = ""
        txtAddrLun_cus.Text = "嘉義縣東區世賢路二段567號"
        txtAddrDin_cus.Text = "嘉義縣東區世賢路二段567號"
        txtBedNo_cus.Text = ""
        txtMemo_cus.Text = "到達前10分鐘請電話通知"

        '商品管理
        With dgProduct
            .Columns.Add("", "編號")
            .Columns.Add("", "群組")
            .Columns.Add("", "商品分類")
            .Columns.Add("", "品名")
            .Columns.Add("", "售價")
            .Columns.Add("", "成本")
            .Columns.Add("", "備註")
            .Rows.Add("1", "月子餐", "套餐", "月子餐30日", "57000", "50000")
            .Rows.Add("2", "月子餐", "套餐", "月子餐21日", "39900", "30000")
            .Rows.Add("3", "調養餐", "套餐", "小產調養餐30日", "56789", "50000")
            .Rows.Add("4", "調養餐", "套餐", "小產調養餐21日", "37800", "30000")
            .Rows.Add("5", "月子餐", "單點", "月子早餐", "2688", "1800")
            .Rows.Add("6", "月子餐", "單點", "月子午餐", "1688", "1200")
            .Rows.Add("7", "調養餐", "單點", "調養早餐", "3780", "3000")
            .Rows.Add("8", "調養餐", "單點", "調養午餐", "5400", "5000")
        End With

        txtProdName_product.Text = "小產調養餐30日"
        cmdProdType_product.Text = "套餐"
        cmbProdGroup_product.Text = "調養餐"
        txtProdPrice_product.Text = "56789"
        txtProdCost_product.Text = "50000"

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
            '.Rows.Add("1", "B", "2023-01-23", "經典月子餐", "黃金小米粥", "泰式沙嗲烤豬", "燻雞香拌雲耳", "蒜香龍鬚菜", "黃芪鮮雞湯", "", "枸杞排骨湯", "枸杞排骨湯", "杜仲燉排骨", "傳香地瓜飯", "蒜蓉海大蝦", "茶油杏菇爆炒腰子", "玉米高麗菜", "柳丁",
            '          "紅糖大麥粥", "", "錦蔬鮮魚湯", "錦蔬鮮魚湯", "何首烏鮮魚湯", "枸杞養生飯(茶油)", "醬燒煨豬膝", "塔香肉絲海龍", "吻魚白杏菜", "黃奇果", "", "玉竹鮮雞湯", "干貝鮮雞湯", "八珍干貝鮮雞湯")

            '.Rows.Add("2", "B", "2023-01-03", "溫馨月子餐", "照燒梅花三明治", "田園烤白筍", "起司煎蛋", "養生芝麻飲", "青木瓜燉魚湯", "", "枸杞排骨湯", "何首烏排骨湯", "何首烏排骨湯", "芝麻糙米飯", "檸檬香煎海魚",
            '          "黃耆炒雞肉", "薑絲蔭醬過貓", "柳丁", "桂圓銀耳甜湯", "", "玉竹鮮雞湯", "紅棗玉竹鮮雞湯", "黨蔘鮮雞湯", "養生紫米飯", "秘製紅酒牛腩", "翡翠鮮菇蒸雙鮮", "腐乳高麗菜", "百香果", "", "棗香龍尾湯", "棗香龍尾湯",
            '          "龍尾虎豆燉紅棗")

            '.Rows.Add("3", "C", "2023-03-11", "幸福餐", "", "", "", "", "", "", "北蟲草花鮮雞湯", "", "", "香甜栗子飯", "南方澳帶魚捲(烤)", "塔香杏鮑菇", "鮮菇白杏", "", "", "味噌魚頭湯", "", "", "", "養生五穀飯",
            '          "磨菇豬小排", "茶香紅棗雞", "蒜香青江菜")

            '.Rows.Add("4", "D", "2023-01-19", "住院餐", "田園時蔬雞肉粥", "椒塩烤鮑菇", "茄汁肉丸", "香菇高麗菜", "黃耆片鮮魚湯", "觀音串", "無花果排骨湯", "", "", "茶香珍菇飯", "梅子燒雞", "清炒香蔥魚栁", "金銀蛋莧菜",
            '          "四季水果", "紅糖燕麥粥", "杜仲茶", "玉竹鮮雞湯", "紅藜高纖飯", "粉蒸排骨(不要豆鼓)", "美人腿炒雞(茶香)", "吻魚炒青江菜", "", "通乳茶", "北菇燉魚湯")

        End With

        'txtProdName_menu.Text = "經典月子餐"
        'cmbProdVers_menu.Text = "B"
        'dtMenu.Value = "2023-01-23"
        'txtBraSta.Text = "黃金小米粥"
        'txtBlaMain.Text = "泰式沙嗲烤豬"
        'txtBlaHM.Text = "燻雞香拌雲耳"
        'txtBlaVag.Text = "蒜香龍鬚菜"
        'txtBlaSoup.Text = "黃芪鮮雞湯"
        'txtBlaDri.Text = ""
        'txtLunSoup.Text = "枸杞排骨湯"
        'txtLun1.Text = "枸杞排骨湯"
        'txtLun3.Text = "杜仲燉排骨"
        'txtLunSta.Text = "傳香地瓜飯"
        'txtLunMain.Text = "蒜蓉海大蝦"
        'txtLunHM.Text = "茶油杏菇爆炒腰子"
        'txtLunVag.Text = "玉米高麗菜"
        'txtLunFru.Text = "柳丁"
        'txtLunDess.Text = "紅糖大麥粥"
        'txtLunDri.Text = ""
        'txtDinSoup.Text = "錦蔬鮮魚湯"
        'txtDin1.Text = "錦蔬鮮魚湯"
        'txtDin3.Text = "何首烏鮮魚湯"
        'txtDinSta.Text = "枸杞養生飯(茶油)"
        'txtDinMain.Text = "醬燒煨豬膝"
        'txtDinHM.Text = "塔香肉絲海龍"
        'txtDinVag.Text = "吻魚白杏菜"
        'txtDinFru.Text = "黃奇果"
        'txtDinDri.Text = ""
        'txtNSSoup.Text = "玉竹鮮雞湯"
        'txtNS1.Text = "干貝鮮雞湯"
        'txtNS3.Text = "八珍干貝鮮雞湯"

        '訂單管理
        With dgOrder
            .Columns.Add("", "訂單編號")
            .Columns.Add("", "客戶姓名")
            .Columns.Add("", "手機")
            .Columns.Add("", "商品名稱")
            .Columns.Add("", "售價")
            .Columns.Add("", "禁忌食物")
            .Columns.Add("", "餐數")
            .Columns.Add("", "預計送餐日")
            .Columns.Add("", "備註")
            '.Rows.Add("1", "陳小姐", "0918-123123", "調養餐30日", "54000", "蝦,花生", "90")
            '.Rows.Add("李小姐", "0918-456456", "孕期餐7日", "36888","21")
            '.Rows.Add("王太太", "0918-852852", "月子早餐", "2688","1")
        End With

        'txtCusName_order.Text = "陳小姐"
        'cmdProdName_order.Text = "月子餐30日"
        'txtPrice_order.Text = "57000"
        'txtPhone_order.Text = "0918-123123"
        'txtTaboo.Text = "蝦,花生"
        'txtCount.Text = "90"

        '配餐管理
        'txtCusName_dist.Text = "陳小姐"
        'txtPhone_dist.Text = "0918-123123"
        'Dim list As New List(Of String) From {"月子餐30日", "調養午餐"}
        'cmbProdName_dist.DataSource = list


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
            '.Rows.Add("1", "2023-03-01", "陳小姐", "0918-123123", "1", "月子餐30日", "10000", "訂金", "123")
            '.Rows.Add("2", "2023-03-05", "3", "王太太", "月子早餐", "2688", "全款")
        End With

        'txtCusName_money.Text = "陳小姐"
        'txtPhone_money.Text = "0918-123123"
        'txtOrdID_money.Text = "1"
        'dtMoney.Value = "2023-03-01"
        'txtMoney.Text = "10000"
        'txtMonType.Text = "訂金"
        'txtMonMemo.Text = "123"

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
            '.Rows.Add("1", "小陳", "05-1111111", "0900-123123", "嘉義縣大林鎮中山路1號", "user1")
            '.Rows.Add("2", "小李", "05-2222222", "0900-456456", "嘉義縣東區世賢路二段567號", "user2", "廚師")
            '.Rows.Add("3", "老王", "05-3333333", "0900-852852", "雲林縣斗六市大學路52號", "user3")
            '.Rows.Add("4", "小張", "05-5555555", "0900-147147", "嘉義縣太保市市政路23號", "user4")
            '.Rows.Add("5", "小高", "05-6666666", "0900-369369", "雲林縣虎尾鎮中正路100號", "user5")
        End With

        'txtEmpName_emp.Text = "小李"
        'txtEmpTel.Text = "05-2222222"
        'txtEmpPhone_Emp.Text = "0900-456456"
        'txtEmpAddr.Text = "嘉義縣東區世賢路二段567號"
        'txtEmpMemo.Text = ""
        'txtEmpAcct.Text = "user2"
        'txtPsw.Text = "********"
        'txtPswCheck.Text = "********"
        'cmbEmpPos_emp.Text = "系統管理員"

        '禁忌食物管理
        With dgTaboo
            .Columns.Add("", "編號")
            .Columns.Add("", "分類")
            .Columns.Add("", "名稱")
            '.Rows.Add("1", "雞", "雞屁股")
            '.Rows.Add("2", "豬", "豬舌頭")
            '.Rows.Add("3", "魚", "魚眼睛")
        End With

        'cmbTaboClass.Text = "雞"
        'txtTaboName.Text = "雞屁股"

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

    Private Sub btn_CusAdd_Click(sender As Object, e As EventArgs) Handles btn_CusAdd.Click
        'msSQL = "INSERT INTO customer (cus_name,cus_tel_hom,cus_tel_com,cus_phone,cus_addr_hom,cus_addr_bla,cus_addr_lun,cus_addr_din,cus_bed,cus_memo)"
        'mSQLCmd = New MySqlCommand(msSQL, conn)

        'With mSQLCmd
        '    .Parameters.AddWithValue("@cus_name", txtName_cus.Text)
        '    .Parameters.AddWithValue("@cus_tel_hom", txtTelHom_cus.Text)
        '    .Parameters.AddWithValue("@cus_tel_com", txtTelCom_cus.Text)
        '    .Parameters.AddWithValue("@cus_phone", txtPhone_cus.Text)
        '    .Parameters.AddWithValue("@cus_addr_hom", txtAddr_cus.Text)
        '    .Parameters.AddWithValue("@cus_addr_bla", txtAddrBla_cus.Text)
        '    .Parameters.AddWithValue("@cus_addr_lun", txtAddrLun_cus.Text)
        '    .Parameters.AddWithValue("@cus_addr_din", txtAddrDin_cus.Text)
        '    .Parameters.AddWithValue("@cus_bed", txtBedNo_cus.Text)
        '    .Parameters.AddWithValue("@cus_memo", txtMemo_cus.Text)
        'End With

        'Try
        '    conn.Open()

        '    If mSQLCmd.ExecuteNonQuery() > 0 Then
        '        MsgBox("新增成功")
        '        '清空textbox
        '    End If
        '    '重新搜尋新增目標
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        'conn.Close()
        MsgBox("新增成功")
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

    Private Sub btnCusQuery_Click(sender As Object, e As EventArgs) Handles btnCusQuery.Click
        msSQL = "SELECT * FROM test"
        mSQLCmd = New MySqlCommand(msSQL, conn)

        Try
            conn.Open()
            mAdapter = New MySqlDataAdapter(mSQLCmd)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub
    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles btnCusCancel.Click, btnProdCancel.Click, btnMenuCancel.Click, btnOrdCancel.Click, btnMonCancel.Click, btnEmpCancel.Click, btnTaboCancel.Click, btnPermCancel.Click, btnDistCancel.Click
        '清除鍵,清除畫面
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
                '思考怎麼清掉tabpage裡的tabcontrol 除行跑跑看是抓到control or page
            ElseIf TypeOf ctrl Is TabPage Then
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
        ElseIf TypeOf ctrl Is DataGridView Then
            Dim dg As DataGridView = CType(ctrl, DataGridView)
            dg.Rows.Clear()
        ElseIf TypeOf ctrl Is CheckBox Then
            Dim chk As CheckBox = CType(ctrl, CheckBox)
            chk.Checked = False
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

    Private Sub btnDistQuery_Click(sender As Object, e As EventArgs) Handles btnDistQuery.Click
        '更改月曆時間,有訂單就找訂單月份,沒訂單就用現在月份
    End Sub
End Class
