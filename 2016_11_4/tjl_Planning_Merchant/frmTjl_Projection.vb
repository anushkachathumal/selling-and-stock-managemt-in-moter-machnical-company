Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.Drawing.Point
Public Class frmTjl_Projection
    Dim c_dataCustomer As DataTable
    Dim c_dataCustomer1 As DataTable

    Private Sub frmTjl_Projection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCF.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPrint_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPro_Year.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGreige_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFG.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSales_Year.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYarn_Dye.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtUSD_Kg.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtUSD_Mtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txt_Kg.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txt_Mtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRev.ReadOnly = True
        txtRev.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCF.ReadOnly = True
        Call Load_Quality()
        Call Load_Product()
        Call Load_Product_Stap()
        Call Load_Shade()
        Call Load_Planned()
        Call Load_BIZUNIT()
        Call Load_PO()
        Call Load_Retailer()
        Call Load_Month()
        txtSales_Year.Text = Year(Today)
        txtPro_Year.Text = Year(Today)

        Call Load_Gride()
        Call Load_Gride_Shade()

    End Sub

    Function Load_Quality()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select m22Quality as [##] from M22Tec_Spec group by m22Quality order by m22Quality"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboQuality
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_Product()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M56Des as [##] from M56Tjl_ProductType "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboProduct_Type
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_Product_Stap()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M57Dis as [##] from M57Production_Stap"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboProduct_Stap
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_Shade()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M58Dis as [##] from M58Tjl_Shade"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboShade
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_Planned()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M59Dis as [##] from M59TJL_Planned"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPlanned
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_BIZUNIT()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M60Dis as [##] from M60TJL_BizUnit"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBussiness_Unit
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_PO()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M61Dis as [##] from M61TJL_PO"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_Retailer()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M55Retailer as [##] from M55Tjl_Projection group by M55Retailer"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRetailer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function

    Function Load_Month()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "select M13Name as [##] from M13Month order by M13Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSales_Month
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                '.Rows.Band.Columns(1).Width = 260
            End With

            With cboPro_Month
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                '.Rows.Band.Columns(1).Width = 260
            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try

    End Function
    Function Load_Gride_BIZFm_Qlty()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String

        Try

            i = 0
            _Shade = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _Shade = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next
            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' and M55Biz_Unit in (" & _Shade & "') and  M55Quality='" & Trim(cboFQulty.Text) & "' order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_BIZShade_Qlty()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String
        Dim _BizUnit As String

        Try

            i = 0
            _Shade = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _Shade = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next

            i = 0
            _BizUnit = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(2).Value) = True Then
                    If _BizUnit <> "" Then
                        _BizUnit = _BizUnit & "','" & Trim(UltraGrid2.Rows(i).Cells(3).Value)
                    Else
                        _BizUnit = "'" & Trim(UltraGrid2.Rows(i).Cells(3).Value)
                    End If
                End If
                i = i + 1
            Next

            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' and M55Biz_Unit in (" & _BizUnit & "') and  M55Quality='" & Trim(cboFQulty.Text) & "' and M55Shade in ('" & _Shade & "') order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_BIZShade()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String
        Dim _BizUnit As String

        Try

            i = 0
            _Shade = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _Shade = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next

            i = 0
            _BizUnit = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(2).Value) = True Then
                    If _BizUnit <> "" Then
                        _BizUnit = _BizUnit & "','" & Trim(UltraGrid2.Rows(i).Cells(3).Value)
                    Else
                        _BizUnit = "'" & Trim(UltraGrid2.Rows(i).Cells(3).Value)
                    End If
                End If
                i = i + 1
            Next

            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' and M55Biz_Unit in (" & _BizUnit & "')  and M55Shade in (" & _Shade & "') order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_ShadeFm_Qlty()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String

        Try

            i = 0
            _Shade = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _Shade = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next
            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' and M55Shade in (" & _Shade & "') and  M55Quality='" & Trim(cboFQulty.Text) & "' order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_ShadeFm()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String

        Try

            i = 0
            _Shade = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _Shade = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next
            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' and M55Shade in (" & _Shade & "') order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride_Quality()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String

        Try

          
            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality='" & Trim(cboFQulty.Text) & "' order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_BIZFm()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String

        Try

            i = 0
            _Shade = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _Shade = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next
            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and proMonth_No>='" & Month(Today) & "' and M55Pro_Year>='" & Year(Today) & "' and M55Quality<>'' and M55Biz_Unit in (" & _Shade & "') order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Clear_Text()
        Me.cboQuality.Text = ""
        Me.cboRetailer.Text = ""
        Me.cboPlanned.Text = ""
        Me.cboBussiness_Unit.Text = ""
        Me.cboPO.Text = ""
        Me.cboPro_Month.Text = ""
        Me.cboSales_Month.Text = ""
        Me.cboSales_Stage.Text = ""
        Me.cboShade.Text = ""
        Me.cboCustomer.Text = ""
        Me.cboProduct_Type.Text = ""
        Me.txtUSD_Mtr.Text = ""
        Me.txtUSD_Kg.Text = ""
        Me.txtSales_Year.Text = Year(Today)
        Me.txtPro_Year.Text = Year(Today)
        Me.txtQty.Text = ""
        Me.txtCF.Text = ""
        Me.txtPrint_Cost.Text = ""
        Me.txtGreige_Cost.Text = ""
        Me.txtFG.Text = ""
        Me.txt_Mtr.Text = ""
        Me.txt_Kg.Text = ""
        Me.txtYarn_Dye.Text = ""
        Me.chkFG1.Checked = False
        Me.chkFG2.Checked = False
        Me.chkFG3.Checked = False
        Me.chkFG4.Checked = False
        Me.chkFG5.Checked = False
        Me.chkOS1.Checked = False
        Me.chkOS2.Checked = False
        Me.chkOS3.Checked = False
        Me.chkOS4.Checked = False
        Me.chkOS5.Checked = False

        Call Load_Quality()
        Call Load_Product()
        Call Load_Product_Stap()
        Call Load_Shade()
        Call Load_Planned()
        Call Load_BIZUNIT()
        Call Load_PO()
        Call Load_Retailer()

        cboQuality.ToggleDropdown()

    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
        Call Load_Gride()
    End Sub

    Private Sub chkFG1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFG1.CheckedChanged
        If chkFG1.Checked = True Then
            chkFG2.Checked = False
            chkFG3.Checked = False
            chkFG4.Checked = False
            chkFG5.Checked = False

        End If
    End Sub

    Private Sub chkFG2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFG2.CheckedChanged
        If chkFG2.Checked = True Then
            chkFG1.Checked = False
            chkFG3.Checked = False
            chkFG4.Checked = False
            chkFG5.Checked = False

        End If
    End Sub

    Private Sub chkFG3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFG3.CheckedChanged
        If chkFG3.Checked = True Then
            chkFG2.Checked = False
            chkFG1.Checked = False
            chkFG4.Checked = False
            chkFG5.Checked = False

        End If
    End Sub

    Private Sub chkFG4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFG4.CheckedChanged
        If chkFG4.Checked = True Then
            chkFG2.Checked = False
            chkFG3.Checked = False
            chkFG1.Checked = False
            chkFG5.Checked = False

        End If
    End Sub

    Private Sub chkFG5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFG5.CheckedChanged
        If chkFG5.Checked = True Then
            chkFG2.Checked = False
            chkFG3.Checked = False
            chkFG4.Checked = False
            chkFG1.Checked = False

        End If
    End Sub

    Private Sub chkOS1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOS1.CheckedChanged
        If chkOS1.Checked = True Then
            chkOS2.Checked = False
            chkOS3.Checked = False
            chkOS4.Checked = False
            chkOS5.Checked = False

        End If
    End Sub

    Private Sub chkOS2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOS2.CheckedChanged
        If chkOS2.Checked = True Then
            chkOS1.Checked = False
            chkOS3.Checked = False
            chkOS4.Checked = False
            chkOS5.Checked = False

        End If
    End Sub

    Private Sub chkOS3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOS3.CheckedChanged
        If chkOS3.Checked = True Then
            chkOS2.Checked = False
            chkOS1.Checked = False
            chkOS4.Checked = False
            chkOS5.Checked = False

        End If
    End Sub

    Private Sub chkOS4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOS4.CheckedChanged
        If chkOS4.Checked = True Then
            chkOS2.Checked = False
            chkOS3.Checked = False
            chkOS1.Checked = False
            chkOS5.Checked = False

        End If
    End Sub

    Private Sub chkOS5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOS5.CheckedChanged
        If chkOS5.Checked = True Then
            chkOS2.Checked = False
            chkOS3.Checked = False
            chkOS4.Checked = False
            chkOS1.Checked = False

        End If
    End Sub


    Private Sub cboQuality_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboQuality.KeyUp
        If e.KeyCode = 13 Then
            cboShade.ToggleDropdown()
        End If
    End Sub

    Private Sub cboShade_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboShade.KeyUp
        If e.KeyCode = 13 Then
            cboPlanned.ToggleDropdown()

        End If
    End Sub

    Private Sub cboPlanned_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPlanned.KeyUp
        If e.KeyCode = 13 Then
            cboProduct_Type.ToggleDropdown()
        End If
    End Sub


    Private Sub cboProduct_Type_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboProduct_Type.KeyUp
        If e.KeyCode = 13 Then
            cboProduct_Stap.ToggleDropdown()
        End If
    End Sub

    Private Sub cboProduct_Stap_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboProduct_Stap.KeyUp
        If e.KeyCode = 13 Then
            cboRetailer.ToggleDropdown()
        End If
    End Sub

    Private Sub cboRetailer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboRetailer.KeyUp
        If e.KeyCode = 13 Then
            cboBussiness_Unit.ToggleDropdown()
        End If
    End Sub

  
    Private Sub cboBussiness_Unit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBussiness_Unit.KeyUp
        If e.KeyCode = 13 Then
            cboPO.ToggleDropdown()

        End If
    End Sub

    Private Sub cboPO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPO.KeyUp
        If e.KeyCode = 13 Then
            cboCustomer.ToggleDropdown()
        End If
    End Sub

    Private Sub cboCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCustomer.KeyUp
        If e.KeyCode = 13 Then
            txtPrint_Cost.Focus()
        End If
    End Sub

    Private Sub txtPrint_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPrint_Cost.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtPrint_Cost.Text <> "" Then
                If IsNumeric(txtPrint_Cost.Text) Then
                    Value = txtPrint_Cost.Text
                    txtPrint_Cost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtPrint_Cost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txtGreige_Cost.Focus()
        End If
    End Sub

    Private Sub txtGreige_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGreige_Cost.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtGreige_Cost.Text <> "" Then
                If IsNumeric(txtGreige_Cost.Text) Then
                    Value = txtGreige_Cost.Text
                    txtGreige_Cost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtGreige_Cost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txtFG.Focus()
        End If
    End Sub

    Private Sub txtFG_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFG.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtFG.Text <> "" Then
                If IsNumeric(txtFG.Text) Then
                    Value = txtFG.Text
                    txtFG.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtFG.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txtYarn_Dye.Focus()
        End If
    End Sub

    Private Sub txtYarn_Dye_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYarn_Dye.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtYarn_Dye.Text <> "" Then
                If IsNumeric(txtYarn_Dye.Text) Then
                    Value = txtYarn_Dye.Text
                    txtYarn_Dye.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtYarn_Dye.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            cboSales_Month.ToggleDropdown()

        End If
    End Sub

    Private Sub cboSales_Month_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSales_Month.KeyUp
        If e.KeyCode = 13 Then
            cboSales_Stage.Focus()
        End If
    End Sub

    Private Sub cboSales_Stage_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSales_Stage.KeyUp
        ' Dim Value As Double
        If e.KeyCode = 13 Then
            txtUSD_Mtr.Focus()
        End If
    End Sub

    Private Sub txtUSD_Mtr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUSD_Mtr.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtUSD_Mtr.Text <> "" Then
                If IsNumeric(txtUSD_Mtr.Text) Then
                    Value = txtUSD_Mtr.Text
                    txtUSD_Mtr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtUSD_Mtr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txtUSD_Kg.Focus()
        End If
    End Sub

    Private Sub txtUSD_Kg_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUSD_Kg.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtUSD_Kg.Text <> "" Then
                If IsNumeric(txtUSD_Kg.Text) Then
                    Value = txtUSD_Kg.Text
                    txtUSD_Kg.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtUSD_Kg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txt_Mtr.Focus()
        End If
    End Sub

    Private Sub txt_Mtr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_Mtr.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txt_Mtr.Text <> "" Then
                If IsNumeric(txt_Mtr.Text) Then
                    Value = txt_Mtr.Text
                    txt_Mtr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txt_Mtr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txt_Kg.Focus()
        End If
    End Sub

    Private Sub txt_Kg_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_Kg.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txt_Kg.Text <> "" Then
                If IsNumeric(txt_Kg.Text) Then
                    Value = txt_Kg.Text
                    txt_Kg.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txt_Kg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            cboPro_Month.ToggleDropdown()
        End If
    End Sub

   

    Private Sub cboPro_Month_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPro_Month.KeyUp
        If e.KeyCode = 13 Then
            txtQty.Focus()
        End If
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()
        End If
    End Sub

   
    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        OPRFilter.Visible = False
    End Sub

    Private Sub chkFilter3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFilter3.CheckedChanged
        If chkFG3.Checked = True Then
            cboFQulty.Text = ""
            chkFG1.Checked = False
            chkFG2.Checked = False
            chkFG4.Checked = False
            chkFG5.Checked = False
            chkFilter1.Checked = False
            chkFilter2.Checked = False
            chkFilter4.Checked = False
        End If
    End Sub

    Private Sub chkFL2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFL2.CheckedChanged
        If chkFL2.Checked = True Then
            chkFL3.Checked = False
        End If
    End Sub

    Private Sub chkFL1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFL1.CheckedChanged
        If chkFL3.Checked = True Then
            chkFL2.Checked = False
        End If
    End Sub

    Function Load_Gride_Shade()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = CustomerDataClass.MakeDataTable_ShadeNew()
        UltraGrid2.DataSource = c_dataCustomer
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(1).Width = 80
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 210
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_Gride_Shade_Biz()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_ShadeWithBiz
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 50
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            .DisplayLayout.Bands(0).Columns(2).Header.Caption = "##"
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_Gride_PrMonth_Biz()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_ProMonthWithBiz
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 50
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            .DisplayLayout.Bands(0).Columns(2).Header.Caption = "##"
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_Gride_BizUnit()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = CustomerDataClass.MakeDataTable_BizUnit
        UltraGrid2.DataSource = c_dataCustomer
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(1).Width = 80
            '.DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 210
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_ShadeData()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select M58Dis from M58Tjl_Shade  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer.NewRow
                '_Qty = 0
                newRow("##") = False
                newRow("Shade") = M01.Tables(0).Rows(i)("M58Dis")
                c_dataCustomer.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_BizData_with_Shade()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select M60Dis from M60TJL_BizUnit  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                '_Qty = 0
                newRow("##1") = False
                newRow("##") = False
                newRow("Biz Unit") = M01.Tables(0).Rows(i)("M60Dis")
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            i = 0
            Sql = "select M58Dis from M58Tjl_Shade  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
              
                ' UltraGrid2.Rows(i).Cells(1).Text = Trim(M01.Tables(0).Rows(i)("M58Dis"))
                UltraGrid2.Rows(i).Cells(1).Value = Trim(M01.Tables(0).Rows(i)("M58Dis"))
                i = i + 1
            Next

            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_BizData_with_ProMonth()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer
        Dim _ProMonth As Date
        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select M60Dis from M60TJL_BizUnit  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                '_Qty = 0
                newRow("##1") = False
                newRow("##") = False
                newRow("Biz Unit") = M01.Tables(0).Rows(i)("M60Dis")
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            i = 0
            _ProMonth = Month(Today) & "/1/" & Year(Today)
            For i = 0 To 7
                UltraGrid2.Rows(i).Cells(3).Value = MonthName(Month(_ProMonth)) & "-" & Year(_ProMonth)
                _ProMonth = _ProMonth.AddDays(+32)
            Next
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride_BIZProMonth_Qlty()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String
        Dim _BizUnit As String
        Dim _proMonth As Date
        Dim _year As String

        Try

            i = 0
            _BizUnit = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _BizUnit <> "" Then
                        _BizUnit = _BizUnit & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _BizUnit = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next

            i = 0
            _Shade = ""
            _year = ""
            _proMonth = Month(Today) & "/1/" & Year(Today)
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(2).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & MonthName(Month(_proMonth))
                        _year = _year & "','" & Year(_proMonth)
                    Else
                        _Shade = MonthName(Month(_proMonth))
                        _year = Year(_proMonth)
                    End If
                End If
                _proMonth = _proMonth.AddDays(+32)
                i = i + 1
            Next

            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and M55Pro_Month in (" & _Shade & "') and M55Pro_Year in (" & _year & "') and M55Quality<>'' and M55Biz_Unit in (" & _BizUnit & "') and  M55Quality='" & Trim(cboFQulty.Text) & "' and M55Shade in ('" & _Shade & "') order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride_BIZProMonth()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim _Shade As String
        Dim _BizUnit As String
        Dim _proMonth As Date
        Dim _year As String

        Try

            i = 0
            _BizUnit = ""
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Value) = True Then
                    If _BizUnit <> "" Then
                        _BizUnit = _BizUnit & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    Else
                        _BizUnit = "'" & Trim(UltraGrid2.Rows(i).Cells(1).Value)
                    End If
                End If
                i = i + 1
            Next

            i = 0
            _Shade = ""
            _year = ""
            _proMonth = Month(Today) & "/1/" & Year(Today)
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(2).Value) = True Then
                    If _Shade <> "" Then
                        _Shade = _Shade & "','" & MonthName(Month(_proMonth))
                        _year = _year & "','" & Year(_proMonth)
                    Else
                        _Shade = MonthName(Month(_proMonth))
                        _year = Year(_proMonth)
                    End If
                End If
                _proMonth = _proMonth.AddDays(+32)
                i = i + 1
            Next

            Sql = "SELECT     M55Ref_No AS [Ref No], M55Quality AS [Quality No], M55Shade AS Shade, CONVERT(varchar,CAST(M55CF AS money), 1) AS CF, M55Planed AS Planned, M55Product_Type AS [Product Type], M55Production_Stap AS [Production Step], M55Retailer AS Retailer, M55Biz_Unit AS [Business Unit], M55PO AS PO, M55Customer AS Customer, M55Sales_Month AS [Sales Month], M55Sales_Year AS [Sales Year],CONVERT(varchar,CAST(M55USD_Mtr AS money), 1)  AS [USD Mtr],CONVERT(varchar,CAST(M55USD_Kg AS money), 1)  AS [USD Kg],CONVERT(varchar,CAST(M55Sales_Vol_Mtr AS money), 1)  AS [Sales volume Mtr],M55Pro_Month AS [Production Month], M55Pro_Year AS [Production Year], M55Sales_Stage AS [Sales Stage],CONVERT(varchar,CAST(M55Sales_Vol_Kg AS money), 1)  AS [Sales volume Kg],CONVERT(varchar,CAST(M55Qty AS money), 1)  AS [Production Volume],CONVERT(varchar,CAST(Rev AS money), 1) as [Rev USD],CONVERT(varchar,CAST(M55Print_Cost AS money), 1)  AS [Print Cost],CONVERT(varchar,CAST(M55Gerige_Cost AS money), 1)  AS [Greige Cost],CONVERT(varchar,CAST(M55FG AS money), 1)  AS [FG Cost],CONVERT(varchar,CAST(M55Yarn_Dye AS money), 1)  AS [Yarn Dye Cost],M55User as [Merchant] FROM View_TJLProjection where M55Status='A' and M55Pro_Month in (" & _Shade & "') and M55Pro_Year in (" & _year & "') and M55Quality<>'' and M55Biz_Unit in (" & _BizUnit & "')  and M55Shade in ('" & _Shade & "') order by M55Quality,M55Pro_Year,M55Pro_Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 60
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 60
                .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(3).Width = 60
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(4).Width = 120
                .DisplayLayout.Bands(0).Columns(5).Width = 90
                .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(6).Width = 130
                .DisplayLayout.Bands(0).Columns(7).Width = 130
                .DisplayLayout.Bands(0).Columns(8).Width = 90
                .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(9).Width = 80
                .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(10).Width = 180
                .DisplayLayout.Bands(0).Columns(11).Width = 70
                .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(12).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(13).Width = 90
                .DisplayLayout.Bands(0).Columns(13).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(14).Width = 90
                .DisplayLayout.Bands(0).Columns(14).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(15).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(16).Width = 90
                .DisplayLayout.Bands(0).Columns(16).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(17).Width = 90
                .DisplayLayout.Bands(0).Columns(17).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(18).Width = 190
                .DisplayLayout.Bands(0).Columns(19).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(20).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(21).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(22).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(23).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(24).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                .DisplayLayout.Bands(0).Columns(25).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            End With
            DBEngin.CloseConnection(con)
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.close()
            End If
        End Try
    End Function

    Function Load_BizData()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select M60Dis from M60TJL_BizUnit  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer.NewRow
                '_Qty = 0
                newRow("##") = False
                newRow("Biz Unit") = M01.Tables(0).Rows(i)("M60Dis")
                c_dataCustomer.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If Trim(cboFQulty.Text) <> "" And chkFilter3.Checked = True And chkFilter2.Checked = True Then
            Call Load_Gride_Shade_Biz()
            Call Load_BizData_with_Shade()
        ElseIf chkFilter3.Checked = True And chkFilter2.Checked = True Then
        Call Load_Gride_Shade_Biz()
            Call Load_BizData_with_Shade()
        ElseIf Trim(cboFQulty.Text) <> "" And chkFilter3.Checked = True And chkFL2.Checked = True Then
            Call Load_Gride_PrMonth_Biz()
            Call Load_BizData_with_ProMonth()
        ElseIf chkFilter3.Checked = True And chkFL2.Checked = True Then
            Call Load_Gride_PrMonth_Biz()
            Call Load_BizData_with_ProMonth()
        ElseIf cboFQulty.Text <> "" And chkFilter2.Checked = True Then
            Call Load_Gride_Shade()
            Call Load_ShadeData()
        ElseIf chkFilter2.Checked = True Then
            Call Load_Gride_Shade()
            Call Load_ShadeData()
        ElseIf cboFQulty.Text <> "" And chkFilter3.Checked = True Then
            Call Load_Gride_BizUnit()
            Call Load_BizData()

        ElseIf chkFilter3.Checked = True Then
            Call Load_Gride_BizUnit()
            Call Load_BizData()
        End If
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If Trim(cboFQulty.Text) <> "" And chkFilter3.Checked = True And chkFilter2.Checked = True Then
            Call Load_Gride_BIZShade_Qlty()
        ElseIf chkFilter3.Checked = True And chkFilter2.Checked = True Then
            Call Load_Gride_BIZShade()
        ElseIf cboFQulty.Text <> "" And chkFilter2.Checked = True Then
            Call Load_Gride_ShadeFm_Qlty()
        ElseIf Trim(cboFQulty.Text) <> "" And chkFilter3.Checked = True And chkFL2.Checked = True Then

            Call Load_Gride_BIZProMonth_Qlty()
        ElseIf chkFilter3.Checked = True And chkFL2.Checked = True Then
            Call Load_Gride_BIZProMonth()
        ElseIf cboFQulty.Text <> "" And chkFilter2.Checked = True Then
            Call Load_Gride_ShadeFm_Qlty()
        ElseIf chkFilter2.Checked = True Then
            Call Load_Gride_ShadeFm()
        ElseIf cboFQulty.Text <> "" And chkFilter3.Checked = True Then
            Call Load_Gride_BIZFm_Qlty()
        ElseIf chkFilter3.Checked = True Then
            Call Load_Gride_BIZFm()
        ElseIf cboFQulty.Text <> "" Then
            Call Load_Gride_Quality()
        End If
    End Sub

    Private Sub cmbMin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMin.Click
        With OPR4
            .Height = 97
            .Width = 983
            '.Location.X = 12
            '.Location.Y = 359
            .Location = New Point(12, 401)
        End With

        With UltraGrid1
            .Width = 966
            .Height = 62

            .Location = New Point(6, 29)
        End With
        ' UltraGrid1.Width = 108


        OPR1.Visible = True
        OPR2.Visible = True

        cmdSave.Visible = True
        cmdReset.Visible = True
        cmdExit.Visible = True
        cmdCancel.Visible = True

    End Sub

    Private Sub cmbIncress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbIncress.Click
        With OPR4
            .Height = 446
            .Width = 983
            '.Location.X = 12
            '.Location.Y = 62
            .Location = New Point(12, 49)
        End With

        With UltraGrid1
            .Width = 966
            .Height = 411

            .Location = New Point(6, 29)
        End With

        OPR1.Visible = False
        OPR2.Visible = False
        cmdSave.Visible = False
        cmdReset.Visible = False
        cmdExit.Visible = False
        cmdCancel.Visible = False
        ' UltraGrid1.Width = 208
    End Sub

   
    Private Sub UltraGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.Click
        Dim _Rowcount As Integer

        _Rowcount = UltraGrid1.ActiveRow.Index
        strQuality_Find = UltraGrid1.Rows(_Rowcount).Cells(0).Text
        strFindStatus = False
    End Sub

    

    Function Search_Record()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim value As Double

        Try
            Sql = "select * from M55Tjl_Projection where m55ref_no='" & strQuality_Find & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboQuality.Text = Trim(M01.Tables(0).Rows(0)("M55Quality"))
                cboShade.Text = Trim(M01.Tables(0).Rows(0)("M55Shade"))
                txtCF.Text = Trim(M01.Tables(0).Rows(0)("M55CF"))
                cboPlanned.Text = Trim(M01.Tables(0).Rows(0)("M55Planed"))
                cboProduct_Type.Text = Trim(M01.Tables(0).Rows(0)("M55Product_Type"))
                cboProduct_Stap.Text = Trim(M01.Tables(0).Rows(0)("M55Production_Stap"))
                cboRetailer.Text = Trim(M01.Tables(0).Rows(0)("M55Retailer"))
                ' MsgBox(Trim(M01.Tables(0).Rows(0)("M55Biz_Unit")))
                ' Call Load_BIZUNIT()
                cboBussiness_Unit.Text = Trim(M01.Tables(0).Rows(0)("M55Biz_Unit"))

                cboPO.Text = Trim(M01.Tables(0).Rows(0)("M55PO"))
                cboCustomer.Text = Trim(M01.Tables(0).Rows(0)("M55Customer"))
                value = Trim(M01.Tables(0).Rows(0)("M55Print_Cost"))
                txtPrint_Cost.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPrint_Cost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                value = Trim(M01.Tables(0).Rows(0)("M55Gerige_Cost"))
                txtGreige_Cost.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGreige_Cost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                value = Trim(M01.Tables(0).Rows(0)("M55FG"))
                txtFG.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtFG.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                value = Trim(M01.Tables(0).Rows(0)("M55Yarn_Dye"))
                txtYarn_Dye.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtYarn_Dye.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                If Trim(M01.Tables(0).Rows(0)("M55FG_Supplier")) = "IN" Then
                    chkFG1.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55FG_Supplier")) = "OUT" Then
                    chkFG2.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55FG_Supplier")) = "TJ" Then
                    chkFG3.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55FG_Supplier")) = "OCI" Then
                    chkFG4.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55FG_Supplier")) = "PTL" Then
                    chkFG5.Checked = True
                End If

                If Trim(M01.Tables(0).Rows(0)("M55OS_Greige")) = "Y" Then
                    chkOS1.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55OS_Greige")) = "N" Then
                    chkOS2.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55OS_Greige")) = "Soma" Then
                    chkOS3.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55OS_Greige")) = "Thinx" Then
                    chkOS4.Checked = True
                ElseIf Trim(M01.Tables(0).Rows(0)("M55OS_Greige")) = "Decathlon" Then
                    chkOS5.Checked = True
                End If

                ' MsgBox(Trim(M01.Tables(0).Rows(0)("M55Sales_Month")))
                'cboSales_Month.DropDownStyle = UltraComboStyle.DropDownList
                cboSales_Month.Text = Trim(M01.Tables(0).Rows(0)("M55Sales_Month"))
                cboSales_Month.DropDownStyle = UltraComboStyle.DropDownList
                txtSales_Year.Text = Trim(M01.Tables(0).Rows(0)("M55Sales_Year"))
                cboSales_Stage.Text = Trim(M01.Tables(0).Rows(0)("M55Sales_Stage"))

                value = Trim(M01.Tables(0).Rows(0)("M55USD_Kg"))
                txtUSD_Kg.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUSD_Kg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                value = Trim(M01.Tables(0).Rows(0)("M55USD_Mtr"))
                txtUSD_Mtr.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUSD_Mtr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                value = Trim(M01.Tables(0).Rows(0)("M55Sales_Vol_Kg"))
                txt_Kg.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txt_Kg.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                value = Trim(M01.Tables(0).Rows(0)("M55Sales_Vol_Mtr"))
                txt_Mtr.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txt_Mtr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))
                cboPro_Month.DropDownStyle = UltraComboStyle.DropDown
                cboPro_Month.Text = Trim(M01.Tables(0).Rows(0)("M55Pro_Month"))
                txtPro_Year.Text = Trim(M01.Tables(0).Rows(0)("M55Pro_Year"))

                If IsDBNull((M01.Tables(0).Rows(0)("M55Qty"))) Then
                Else
                    value = (M01.Tables(0).Rows(0)("M55Qty"))
                    txtQty.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtQty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))
                End If

               

            End If
            ' cboSales_Month.DropDownStyle = UltraComboStyle.DropDownList
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Private Sub cboPro_Month_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboPro_Month.InitializeLayout

    End Sub

    Private Sub UltraCheckEditor1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraCheckEditor1.CheckedChanged
        If UltraCheckEditor1.Checked = True Then
            If IsNumeric(txtQty.Text) Then

            End If
        End If
    End Sub
End Class