
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Marchant
    Dim c_dataCustomer1 As System.Data.DataTable

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_Request_Merch
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 290
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(5).Width = 80
            '.DisplayLayout.Bands(0).Columns(5).AutoEdit = True
            ''   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_Request_M_Maneger
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 290
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110
            '.DisplayLayout.Bands(0).Columns(5).AutoEdit = True
            ''   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub frmView_Marchant_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Trim(strUGroup) = "MERCHENT" Then
            Call Load_Gride()
            Call Load_Data_Merchant()
        ElseIf Trim(strUGroup) = "M/MANEGER" Then
            Call Load_Gride1()
            Call Load_Data_MManeger()
        End If
    End Sub

    Function Load_Data_Merchant()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer
        Dim Value As Double
        Dim _St As String

        Try
            Sql = "select * from T01_TEC_Development_Request inner join M01_TEC_Customer on M01Cus_RefNo=T01Customer_Ref where T01Merchant='" & strDisname & "' and T01Status='DR' "
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                'tmpQty = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = Trim(m01.Tables(0).Rows(i)("T01Req_No_St"))
                newRow("Requested Date") = Month(m01.Tables(0).Rows(i)("T01Req_Date")) & "/" & Microsoft.VisualBasic.Day(m01.Tables(0).Rows(i)("T01Req_Date")) & "/" & Year(m01.Tables(0).Rows(i)("T01Req_Date"))
                newRow("Required Date") = Month(m01.Tables(0).Rows(i)("T01Requied_Date")) & "/" & Microsoft.VisualBasic.Day(m01.Tables(0).Rows(i)("T01Requied_Date")) & "/" & Year(m01.Tables(0).Rows(i)("T01Requied_Date"))
                newRow("Customer Name") = Trim(m01.Tables(0).Rows(i)("M01Cus_Name"))
                Value = Trim(m01.Tables(0).Rows(i)("T01Order_Qty"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Order Qty (mtr)") = _St
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            '  UltraGrid4.DisplayLayout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_MManeger()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim Sql As String
        Dim i As Integer
        Dim Value As Double
        Dim _St As String

        Try
            Sql = "select * from T01_TEC_Development_Request inner join M01_TEC_Customer on M01Cus_RefNo=T01Customer_Ref where  T01Status='DR' order by T01Req_No_St"
            m01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow4 As DataRow In m01.Tables(0).Rows
                'tmpQty = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = Trim(m01.Tables(0).Rows(i)("T01Req_No_St"))
                newRow("Requested Date") = Month(m01.Tables(0).Rows(i)("T01Req_Date")) & "/" & Microsoft.VisualBasic.Day(m01.Tables(0).Rows(i)("T01Req_Date")) & "/" & Year(m01.Tables(0).Rows(i)("T01Req_Date"))
                newRow("Required Date") = Month(m01.Tables(0).Rows(i)("T01Requied_Date")) & "/" & Microsoft.VisualBasic.Day(m01.Tables(0).Rows(i)("T01Requied_Date")) & "/" & Year(m01.Tables(0).Rows(i)("T01Requied_Date"))
                newRow("Customer Name") = Trim(m01.Tables(0).Rows(i)("M01Cus_Name"))
                Value = Trim(m01.Tables(0).Rows(i)("T01Order_Qty"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Order Qty (mtr)") = _St
                newRow("Merchant") = Trim(m01.Tables(0).Rows(i)("T01Merchant"))
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            '  UltraGrid4.DisplayLayout.Bands(0).Columns("Process Root").ValueList = Me.UltraDropDown3
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim i As Integer
        strFab_Req_No = 0

        i = UltraGrid1.ActiveRow.Index
        strFab_Req_No = UltraGrid1.Rows(i).Cells(0).Value
        Me.Close()

        frmDisplay_Merchant.MdiParent = MDIMain
        ' m_ChildFormNumber += 1
        ' frmWinner.Text = "WNF"
        frmDisplay_Merchant.Show()

    End Sub

End Class