Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Invoice
    Dim c_dataCustomer2 As DataTable
    Private Sub frmView_Invoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtEnter1.ReadOnly = True
        txtEnter1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.ReadOnly = True
        txtDate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtJob.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtJob.ReadOnly = True
        txtService.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtService.ReadOnly = True
        txtNet1.ReadOnly = True
        txtNet1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNext.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtNext.ReadOnly = True
        txtTec1.ReadOnly = True
        txtTec2.Text = True
        txtTP1.ReadOnly = True
        txtCname1.ReadOnly = True
        txtAddress1.ReadOnly = True
        txtCount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount1.ReadOnly = True
        txtMtr1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMtr1.ReadOnly = True
        Call Load_Gride3()

    End Sub

    Function Load_Gride3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_iNVOICE_UNIQ
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 170
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(7).Width = 110


            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            '.DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Search_Invoice()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Dim VALUE As Double
        Dim _sT As String
        Try
            Sql = "SELECT * FROM T08Sales_Header WHERE T08Job_No='" & Trim(txtJob.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtDate.Text = M01.Tables(0).Rows(0)("T08Date")
                txtEnter1.Text = Trim(M01.Tables(0).Rows(0)("T08Invo_No"))
                txtV_No.Text = Trim(M01.Tables(0).Rows(0)("T08V_No"))
                Call Search_Vehicle_No_1()
                txtMtr1.Text = Trim(M01.Tables(0).Rows(0)("T08St_Mtr"))
                txtService.Text = Trim(M01.Tables(0).Rows(0)("T08Service_on"))
                txtNext.Text = Trim(M01.Tables(0).Rows(0)("T08End_mtr"))
            End If

            Sql = "select M10Name from T10Technicion_Comm inner join M10Employee on M10Code=T10Emp where T10INV_No='" & Trim(txtEnter1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                If i = 0 Then
                    txtTec1.Text = Trim(M01.Tables(0).Rows(0)("M10Name"))
                Else
                    txtTec2.Text = Trim(M01.Tables(0).Rows(0)("M10Name"))
                End If
                i = i + 1
            Next
            '================================================================================
            Sql = "select * from T09Sales_Flutter  where T09Inv_No='" & Trim(txtEnter1.Text) & "' and T09Department<>'-'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Department") = Trim(M01.Tables(0).Rows(i)("T09Department"))
                newRow("#Part No") = Trim(M01.Tables(0).Rows(i)("T09Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("T09Item_Name"))
                VALUE = M01.Tables(0).Rows(i)("T09Retail")
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Retail Price") = _sT
                VALUE = M01.Tables(0).Rows(i)("T09Qty")
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Qty") = _sT
                VALUE = M01.Tables(0).Rows(i)("T09Free")
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Free Issue") = _sT
                VALUE = M01.Tables(0).Rows(i)("T09Discount")
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Discount%") = _sT
                VALUE = (CDbl(M01.Tables(0).Rows(i)("T09Retail")) * CDbl(M01.Tables(0).Rows(i)("T09Qty")))
                _sT = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _sT = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
                newRow("Total") = _sT
                ' _TOTAL = VALUE + _TOTAL
                c_dataCustomer2.Rows.Add(newRow)
                i = i + 1
            Next
            txtCount1.Text = UltraGrid2.Rows.Count
            Call Calculation_Net()
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function
    Function Search_Vehicle_No_1() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Vehicle_No_1 = False
            Sql = "SELECT T05Job_No,T05Mtr FROM T05Job_Card WHERE T05Vehi_No='" & Trim(txtV_No.Text) & "' AND T05Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtJob.Text = Trim(M01.Tables(0).Rows(0)("T05Job_No"))
                txtMtr1.Text = Trim(M01.Tables(0).Rows(0)("T05Mtr"))
                ' Call Search_Jobno_1()
            End If
            '=====================================================================
            Sql = "select * from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  and M07V_No='" & Trim(txtV_No.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Vehicle_No_1 = True
                '   _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                'cboBrand.Text = Trim(M01.Tables(0).Rows(0)("M07Brand_Name"))
                'cbov_Type.Text = Trim(M01.Tables(0).Rows(0)("M07Type"))
                txtTP1.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                txtCname1.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtAddress1.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                ' cboCus_Type.Text = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
            End If
            ' Call lOAD_DATA_GRIDE()
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Calculation_Net()
        On Error Resume Next
        Dim Value As Double
        Dim i As Integer
        i = 0
        Value = 0
        For Each uRow As UltraGridRow In UltraGrid2.Rows
            If IsNumeric(UltraGrid2.Rows(i).Cells(7).Text) Then
                Value = Value + CDbl(UltraGrid2.Rows(i).Cells(7).Text)
            End If

            i = i + 1
        Next

        txtNet1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtNet1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function

   
End Class
