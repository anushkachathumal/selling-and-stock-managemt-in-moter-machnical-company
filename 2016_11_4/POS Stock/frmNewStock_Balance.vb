Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmNewStock_Balance
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _ExStatus As Boolean

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [Location] from M04Location "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 320
                '  .Rows.Band.Columns(1).Width = 160


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Name as [Item Name] from M03Item_Master where M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '  .Rows.Band.Columns(1).Width = 160


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet

        Dim i As Integer
        Try
            Sql = "select * from M04Location where  M04Loc_Name='" & Trim(cboLocation.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Loccode = Trim(M01.Tables(0).Rows(0)("M04Loc_Code"))
                Search_Location = True
            End If
            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmNewStock_Balance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Combo()
        Call Load_Item()
        Call Load_GrideStock()
        txtSt_New.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtTotQty.ReadOnly = True
        txtLast_Date.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCurrent.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtNew.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.Text = Today
        txtCurrent.ReadOnly = True
        txtLast_Date.ReadOnly = True
        txtGRN.ReadOnly = True
        txtSales.ReadOnly = True
        txtTransfer.ReadOnly = True
        txtGRN.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSales.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTransfer.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEx_Date.Text = Today

    End Sub

    Function Load_GrideStock()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_ExStock
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function
    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            txtCode.Focus()
        End If
    End Sub

    Function Search_ItemName_1() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim Value As Double

        Dim i As Integer
        Try
            If Search_Location() = True Then
            Else
                MsgBox("Please select the Location", MsgBoxStyle.Information, "Information .....")
                con.close()
                Exit Function
            End If

            Sql = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "'  and M03Com_Code='" & _Comcode & "' and M03Cost_Price='" & txtCost.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboItem.Text = Trim(M01.Tables(0).Rows(0)("M03Item_Name"))
                Search_ItemName_1 = True
            End If

            Value = 0
            'Current Stock
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCurrent.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'LAST OB DATE
            Sql = "select * from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type='OB' and S01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtLast_Date.Text = Month(M01.Tables(0).Rows(0)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("S01Date")) & "/" & Year(M01.Tables(0).Rows(0)("S01Date"))

            End If
            'SALES
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type IN ('DR','HS') and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtSales.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'GRN
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='GRN' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtGRN.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If
            'TRANSFER

            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='TR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtTransfer.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If

            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_ItemName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim Value As Double

        Dim i As Integer
        Try
            If Search_Location() = True Then
            Else
                MsgBox("Please select the Location", MsgBoxStyle.Information, "Information .....")
                con.close()
                Exit Function
            End If

            Sql = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "'  and M03Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboItem.Text = Trim(M01.Tables(0).Rows(0)("M03Item_Name"))
                Search_ItemName = True
            End If

            Value = 0
            'Current Stock
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCurrent.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'LAST OB DATE
            Sql = "select * from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type='OB' and S01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtLast_Date.Text = Month(M01.Tables(0).Rows(0)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("S01Date")) & "/" & Year(M01.Tables(0).Rows(0)("S01Date"))

            End If
            'SALES
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type IN ('DR','HS') and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtSales.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'GRN
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='GRN' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtGRN.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If
            'TRANSFER

            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='TR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtTransfer.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If

            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemName()
            txtCost.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            Call Load_Gride()
            txtFind.Text = ""
            Call Load_Gride()
            OPR5.Visible = True
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR7.Visible = False
            OPR5.Visible = False
        End If
    End Sub

    Function Search_Itemcode_1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim Value As Double
        Dim A As String

        Dim i As Integer
        Try
            Call Search_Location()
            Sql = "select * from M03Item_Master where M03Item_Name='" & cboItem.Text & "' and M03Com_Code='" & _Comcode & "' and M03Cost_Price='" & txtCost.Text & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCode.Text = Trim(M01.Tables(0).Rows(0)("M03Item_Code"))

            End If

            Value = 0
            'Current Stock
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCurrent.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'LAST OB DATE
            Sql = "select * from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type='OB' and S01Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtLast_Date.Text = Month(M01.Tables(0).Rows(0)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("S01Date")) & "/" & Year(M01.Tables(0).Rows(0)("S01Date"))

            End If
            'SALES
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type IN ('DR','HS') and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtSales.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'GRN
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='GRN' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtGRN.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If
            'TRANSFER

            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='TR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtTransfer.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If

            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_Itemcode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim Value As Double
        Dim A As String

        Dim i As Integer
        Try
            Call Search_Location()
            Sql = "select * from M03Item_Master where M03Item_Name='" & cboItem.Text & "' and M03Com_Code='" & _Comcode & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCode.Text = Trim(M01.Tables(0).Rows(0)("M03Item_Code"))

            End If

            Value = 0
            'Current Stock
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCurrent.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'LAST OB DATE
            Sql = "select * from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type='OB' and S01Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtLast_Date.Text = Month(M01.Tables(0).Rows(0)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(0)("S01Date")) & "/" & Year(M01.Tables(0).Rows(0)("S01Date"))

            End If
            'SALES
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type IN ('DR','HS') and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtSales.Text = Trim(M01.Tables(0).Rows(0)("Qty"))

            End If
            'GRN
            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='GRN' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtGRN.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If
            'TRANSFER

            Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & _Loccode & "' and S01Item_Code='" & txtCode.Text & "' and S01Status='A' AND S01Trans_Type ='TR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Loc_Code,S01Com_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtTransfer.Text = Trim(M01.Tables(0).Rows(0)("Qty"))
            End If

            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cboItem_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItem.AfterCloseUp
        Call Search_Itemcode()
    End Sub

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = 13 Then
            UltraGrid1.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            txtCode.Focus()
        End If
    End Sub

 

    Private Sub txtFind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFind.TextChanged
        Call Load_Gride1()
    End Sub

    Function Load_Gride1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & txtFind.Text & "%' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & cboItem.Text & "%' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid1.ActiveRow.Index


        txtCode.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Text
        Search_ItemName()
        OPR5.Visible = False

    End Sub

  
    Function Clear_Text()
        Me.cboItem.Text = ""
        'Me.cboLocation.Text = ""
        Me.txtCode.Text = ""
        Me.txtCurrent.Text = ""
        Me.txtNew.Text = ""
        Me.txtSales.Text = ""
        Me.txtGRN.Text = ""
        Me.txtTransfer.Text = ""
        Me.txtSt_New.Text = ""
        Me.txtTotQty.Text = ""
        Me.txtCost.Text = ""
        Me.txtRetail.Text = ""
        Call Load_GrideStock()
        OPR5.Visible = False
        OPR7.Visible = False
        txtFind.Text = ""
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
        cboLocation.Text = ""
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim i As Integer
        Dim result1 As String
        Dim m01 As DataSet
        Dim M02 As Double

        Try
            If txtCost.Text <> "" Then
            Else
                MsgBox("Please enter the cost price", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                txtCost.Focus()
                Exit Sub
            End If

            If IsNumeric(txtCost.Text) Then
            Else
                MsgBox("Please enter the correct Cost Price", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                txtCost.Focus()
                Exit Sub
            End If

            If txtRetail.Text <> "" Then
            Else
                MsgBox("Please enter the Retail price", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                txtRetail.Focus()
                Exit Sub
            End If

            If IsNumeric(txtRetail.Text) Then
            Else
                MsgBox("Please enter the correct Retail Price", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                txtRetail.Focus()
                Exit Sub
            End If
            If txtNew.Text <> "" Then
                If IsNumeric(txtNew.Text) Then
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    Exit Sub
                End If
            Else
                txtNew.Text = "0"
            End If

            If Search_ItemName() = True Then
            Else
                MsgBox("Please enter the correct Item Code", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                Exit Sub
            End If
            result1 = MsgBox("Are you sure you want to update this stock", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ......")
            If result1 = vbYes Then

                'nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "' and m03Status='A' and M03ExPair='YES'"
                'm01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(m01) Then
                '    If txtTotQty.Text <> "" Then
                '        If CDbl(txtTotQty.Text) = CDbl(txtNew.Text) Then
                '        Else
                '            MsgBox("Please check the New Balance qty", MsgBoxStyle.Information, "Information ......")
                '            OPR7.Visible = True
                '            txtSt_New.Focus()
                '            connection.Close()
                '            Exit Sub
                '        End If
                '    Else
                '        If txtNew.Text > 0 Then


                '            OPR7.Visible = True
                '            txtSt_New.Focus()
                '            connection.Close()
                '            Exit Sub
                '        End If
                '        End If
                'End If

                nvcFieldList1 = "UPDATE S01Stock_Balance set S01Status='I' where S01Item_Code='" & txtCode.Text & "' and S01Loc_Code='" & _Loccode & "' and S01Com_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Com_Code,S01Status)" & _
                                                                 " values('" & _Loccode & "', '" & (Trim(txtCode.Text)) & "','" & txtDate.Text & "','OB','" & txtNew.Text & "','0','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S04Stock_Price set S04Status='I' where S04Item_Code='" & txtCode.Text & "' and s04Location='" & _Loccode & "' and S04Com_Code='" & _Comcode & "' AND S04Cost='" & Trim(txtCost.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S04Stock_Price(s04Location,S04Item_Code,S04Date,S04Tr_Type,S04Qty,S04Com_Code,S04Status,S04Cost,S04Rate)" & _
                                                                 " values('" & _Loccode & "', '" & (Trim(txtCode.Text)) & "','" & txtDate.Text & "','OB','" & txtNew.Text & "','" & _Comcode & "','A','" & txtCost.Text & "' ,'" & txtRetail.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'EXPAIRE STOCK UPDATE
                nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "' and m03Status='A' and M03ExPair='YES' and M03Com_Code='" & _Comcode & "'"
                m01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(m01) Then
                    nvcFieldList1 = "UPDATE S03Ex_Stock set S03Status='I' where S03Item_Code='" & txtCode.Text & "' and S03Loc_Code='" & _Loccode & "' and S03Com_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    i = 0
                    For Each uRow As UltraGridRow In UltraGrid2.Rows
                        nvcFieldList1 = "Insert Into S03Ex_Stock(S03Loc_Code,S03Tr_Type,S03Item_Code,S03Qty,S03Ex_Date,S03Status,S03Com_Code)" & _
                                                               " values('" & _Loccode & "','OB', '" & (Trim(txtCode.Text)) & "','" & UltraGrid2.Rows(i).Cells(4).Text & "','" & UltraGrid2.Rows(i).Cells(2).Text & "','A','" & _Comcode & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        i = i + 1
                    Next
                End If
                'nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "' and m03Status='A' and M03ExPair='YES'"
                'm01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(m01) Then

                'End If
                MsgBox("Stock update successully", MsgBoxStyle.Information, "Information .......")
                transaction.Commit()
            End If
            connection.Close()
            Call Clear_Text()
            txtCode.Focus()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub txtNew_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

   


    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = Keys.Escape Then
            OPR7.Visible = False
            OPR5.Visible = False
        ElseIf e.KeyCode = Keys.F1 Then
            Call Load_Gride()
            txtFind.Text = ""
            Call Load_Gride()
            OPR5.Visible = True
            txtFind.Focus()
        ElseIf e.KeyCode = 13 Then
            txtCost.Focus()
        End If
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        OPR7.Visible = False
    End Sub

    Private Sub txtSt_New_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSt_New.KeyUp
        Dim i As Integer

        Try
            If txtTotQty.Text <> "" Then
            Else
                txtTotQty.Text = "0"
            End If
            If e.KeyCode = 13 Then
                If txtSt_New.Text <> "" Then
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ........")
                    txtSt_New.Focus()
                    Exit Sub
                End If

                If IsNumeric(txtSt_New.Text) Then
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ........")

                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ........")
                    txtSt_New.Focus()
                    Exit Sub
                End If

                i = 0
                For Each uRow As UltraGridRow In UltraGrid2.Rows
                    If Trim(txtCode.Text) = Trim(UltraGrid2.Rows(i).Cells(0).Text) And txtEx_Date.Text = Trim(UltraGrid2.Rows(i).Cells(2).Text) Then
                        UltraGrid2.Rows(i).Cells(3).Value = txtSt_New.Text
                        txtTotQty.Text = CDbl(txtTotQty.Text) + CDbl(txtSt_New.Text)
                        Exit For
                    End If
                    i = i + 1
                Next

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = Trim(txtCode.Text)
                newRow("Item Name") = cboItem.Text
                newRow("Ex Date") = txtEx_Date.Text
                newRow("Current Qty") = "0"
                newRow("New Qty") = txtSt_New.Text
                ' newRow("Free Issue") = txtFree.Text
                c_dataCustomer1.Rows.Add(newRow)
            
                txtTotQty.Text = CDbl(txtTotQty.Text) + txtSt_New.Text
                txtSt_New.Text = ""
                txtSt_New.Focus()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Sub


    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        OPR7.Visible = False
        cmdAdd.Focus()
    End Sub

    
  

    Private Sub txtCost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCost.KeyUp
        On Error Resume Next
        If e.KeyCode = 13 Then
            If IsNumeric(txtCost.Text) Then
                Dim Value As Double
                Value = txtCost.Text
                txtCost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Call Search_ItemName_1()
            End If
            txtRetail.Focus()
        End If
    End Sub

    Private Sub txtRetail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetail.KeyUp
        On Error Resume Next
        If e.KeyCode = 13 Then
            If IsNumeric(txtRetail.Text) Then
                Dim Value As Double
                Value = txtRetail.Text
                txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtNew.Focus()
        End If
    End Sub

    Private Sub txtCost_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCost.ValueChanged

    End Sub
End Class