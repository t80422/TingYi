Module modUtility
    '客戶管理
    Friend sqlCustomer As String = "SELECT cus_id, cus_name, cus_gender, cus_phone FROM customer"
    '商品群組管理
    Friend sqlProductGroup As String = "SELECT * FROM product_group"
    '商品管理
    Friend sqlProduct As String = "SELECT a.prod_id,a.prod_name,b.prod_grp_name,a.prod_price,a.prod_cost,a.prod_type,a.prod_meal,a.prod_memo FROM product a LEFT JOIN product_group b on a.prod_prod_grp_id=b.prod_grp_id"
    '禁忌管理
    Friend sqlTaboo As String = "SELECT * FROM taboo"
    '訂單管理
    Friend sqlOrder As String = "SELECT a.ord_id, a.ord_date, b.cus_name, b.cus_phone, c.prod_name FROM orders a LEFT JOIN customer b ON a.ord_cus_id = b.cus_id LEFT JOIN product c ON a.ord_prod_id=c.prod_id"
    '財務管理
    Friend sqlMoney As String = "SELECT a.mon_id, b.cus_name, b.cus_phone, c.ord_id, a.mon_date, a.mon_type, a.mon_income, a.mon_memo FROM money a LEFT JOIN customer b ON a.mon_cus_id=b.cus_id LEFT JOIN orders c on a.mon_ord_id=c.ord_id"
    '權限管理
    Friend sqlPermision As String = "SELECT * FROM permissions"
    '員工管理
    Friend sqlEmployee As String = "SELECT a.emp_id, a.emp_name, a.emp_phone, a.emp_tel, a.emp_address, b.perm_name, a.emp_acct, a.emp_psw, a.emp_memo FROM employee a LEFT JOIN permissions b ON a.emp_perm_id = b.perm_id"
    '配餐管理
    Friend sqlDistribute As String = "SELECT b.cus_name,b.cus_phone,a.ord_id,c.prod_name FROM orders a LEFT JOIN customer b ON a.ord_cus_id=b.cus_id LEFT JOIN product c ON a.ord_prod_id=c.prod_id"
    '菜單管理
    Friend sqlMenu As String = "SELECT DISTINCT b.prod_name,a.me_date,a.me_version,b.prod_id FROM menu a LEFT JOIN product b ON a.me_prod_id=b.prod_id LIMIT 100"

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
End Module
