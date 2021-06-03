Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Imports DBLotVbnet.modlVar1
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmGrn_T_Acc
    Dim _Acc_Type As String
    Dim _Comcode As String
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim c_dataCustomer4 As DataTable
    Dim _From As Date
    Dim _To As Date

    Dim _Suppcode As String
    Dim _Search_Status As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub frmGrn_T_Acc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride2()
    End Sub


    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Creditors_Statment
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
           

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(6).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Loc_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboAccount
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                ' .Rows.Band.Columns(1).Width = 180


            End With

          

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub BySupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem.Click
        Panel3.Visible = True
        Call Load_Supplier()
        cboAccount.ToggleDropdown()
    End Sub

    Private Sub UltraButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton8.Click
        Call Load_Gride2()
        Call Load_Data()
        _Search_Status = "A1"
        Panel3.Visible = False
        _Suppcode = Trim(cboAccount.Text)
    End Sub

    Function Load_Data()
        Dim Sql As String
        Dim M01 As DataSet

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _dr As Double

        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double
        Try
            Sql = "SELECT * FROM View_GRN_T_ACCOUNT WHERE M09Name='" & Trim(cboAccount.Text) & "' AND T01To_Loc_Code='" & _Comcode & "' ORDER BY T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _dr = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = M01.Tables(0).Rows(i)("D1")
                newRow("Ref No") = Trim(M01.Tables(0).Rows(i)("Type"))
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
              
                Value = M01.Tables(0).Rows(i)("CR")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("CR") = _St

                Value = M01.Tables(0).Rows(i)("gross")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _dr = Value + _dr
                newRow("DR") = _St

                newRow("Chq No") = Trim(M01.Tables(0).Rows(i)("Chq"))
                ' newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("CR") = _St

            Value = _dr
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("DR") = _St

            c_dataCustomer1.Rows.Add(newRow1)


            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            con.CLOSE()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride2()
        Panel3.Visible = False
    End Sub
End Class