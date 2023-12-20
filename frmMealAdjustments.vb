Public Class frmMealAdjustments
    Private _orderID As Integer

    Public WriteOnly Property OrderID As Integer
        Set
            _orderID = Value
        End Set
    End Property

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Close()
    End Sub

    Private Sub frmMealAdjustments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadOrderData()
    End Sub

    Private Sub LoadOrderData()
        Try
            Dim dic As New Dictionary(Of String, Object) From {
                {"ord_id", _orderID}
            }
            Dim row = SelectTable("SELECT ord_breakfast, ord_lunch, ord_dinner FROM orders WHERE ord_id = @ord_id", dic).Rows(0)

            GetDataToControls(Me, row)
            UpdateMealResult()

        Catch ex As Exception
            MsgBox(ex)
        End Try
    End Sub

    Private Sub UpdateMealResult()
        txtBreakfastResult.Text = (Int(txtBreakfastCount.Text) + Int(txtBreakfastAdj.Text)).ToString()
        txtLunchResult.Text = (Int(txtLunchCount.Text) + Int(txtLunchAdj.Text)).ToString()
        txtDinnerResult.Text = (Int(txtDinnerCount.Text) + Int(txtDinnerAdj.Text)).ToString()
    End Sub

    Private Sub btnSubtract_Click(sender As Object, e As EventArgs) Handles btnSubtract_break.Click, btnSubtract_lunch.Click, btnSubtract_dinner.Click
        Dim txtAdj As TextBox = Nothing

        Select Case True
            Case sender Is btnSubtract_break
                txtAdj = txtBreakfastAdj
            Case sender Is btnSubtract_lunch
                txtAdj = txtLunchAdj
            Case sender Is btnSubtract_dinner
                txtAdj = txtDinnerAdj
        End Select

        If txtAdj IsNot Nothing Then
            txtAdj.Text = (Int(txtAdj.Text) - 1).ToString()
        End If

        UpdateMealResult()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd_break.Click, btnAdd_lunch.Click, btnAdd_dinner.Click
        Dim txtAdj As TextBox = Nothing

        Select Case True
            Case sender Is btnAdd_break
                txtAdj = txtBreakfastAdj
            Case sender Is btnAdd_lunch
                txtAdj = txtLunchAdj
            Case sender Is btnAdd_dinner
                txtAdj = txtDinnerAdj
        End Select

        If txtAdj IsNot Nothing Then
            txtAdj.Text = (Int(txtAdj.Text) + 1).ToString()
        End If

        UpdateMealResult()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim dic As New Dictionary(Of String, Object) From {
            {"ord_breakfast", txtBreakfastResult.Text},
            {"ord_lunch", txtLunchResult.Text},
            {"ord_dinner", txtDinnerResult.Text}
        }

        UpdateTable("orders", dic, $"ord_id = {_orderID}")

        Close()
    End Sub
End Class
