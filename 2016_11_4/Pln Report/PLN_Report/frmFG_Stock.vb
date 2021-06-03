Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Imports Microsoft.Office.Interop.Excel

Public Class frmFG_Stock
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim c_dataCustomer2 As System.Data.DataTable
    Dim c_dataCustomer3 As System.Data.DataTable
    Dim c_dataCustomer4 As System.Data.DataTable
    Dim c_dataCustomer5 As System.Data.DataTable
    Dim c_dataCustomer6 As System.Data.DataTable
    Dim c_dataCustomer7 As System.Data.DataTable

    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String
    Dim _Location As String
    Dim _Bu As String


    Function Load_Gride_WithRecords()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim _Date As Integer
        Dim diff1 As System.TimeSpan
        Dim date2 As System.DateTime
        Dim date1 As System.DateTime
        Dim _WeekNo As Integer
        Dim vcWharer As String
        Dim _QtyMtr As Double
        Dim _QtyKG As Double

        _WeekNo = DatePart(DateInterval.WeekOfYear, Today)

        _QtyKG = 0
        _QtyMtr = 0

        Try
            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" And txtBU.Text <> "" And Trim(txtLocation.Text) <> "" And cboStatus.Text <> "" Then

            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" And txtBU.Text <> "" And Trim(txtLocation.Text) <> "" Then

            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" And txtBU.Text <> "" Then

            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

            ElseIf Trim(cboStatus.Text) <> "" And Trim(txtLocation.Text) <> "" And Trim(txtBU.Text) <> "" Then
                If Trim(cboStatus.Text) = "Below One Month" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over One Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Two Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Three Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Fore Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Five Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over Six Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over One Year" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                End If


            ElseIf Trim(cboStatus.Text) <> "" And Trim(txtLocation.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                If Trim(cboStatus.Text) = "Below One Month" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over One Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Two Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Three Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Fore Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Five Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over Six Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over One Year" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                End If

            ElseIf Trim(cboStatus.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                If Trim(cboStatus.Text) = "Below One Month" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over One Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Two Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Three Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Fore Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Five Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over Six Months" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over One Year" Then
                    vcWharer = "M08Merchant in ('" & _Merchant & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                End If
            ElseIf Trim(cboStatus.Text) <> "" And Trim(txtBU.Text) <> "" Then
                If Trim(cboStatus.Text) = "Below One Month" Then
                    vcWharer = "M14Name in ('" & _Bu & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over One Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Two Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Three Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Fore Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Five Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over Six Months" Then
                    vcWharer = "M14Name in ('" & _Bu & "')  and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over One Year" Then
                    vcWharer = "M14Name in ('" & _Bu & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                End If

            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(cboStatus.Text) <> "" Then
                If Trim(cboStatus.Text) = "Below One Month" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over One Months" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Two Months" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Three Months" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Fore Months" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Five Months" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over Six Months" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over One Year" Then
                    vcWharer = "M01Cuatomer_Name in ('" & _Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                End If
            ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtLocation.Text) <> "" Then
                vcWharer = " M01Cuatomer_Name in ('" & _Customer & "') and M08Location in ('" & _Location & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

            ElseIf Trim(txtMerchant.Text) <> "" And Trim(txtLocation.Text) <> "" Then
                vcWharer = " M08Merchant in ('" & _Merchant & "') and M08Location in ('" & _Location & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

            ElseIf Trim(txtCustomer.Text) <> "" Then
                vcWharer = " M01Cuatomer_Name in ('" & _Customer & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
            ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtLocation.Text) <> "" Then
                vcWharer = " M08Retailer in ('" & _Department & "') AND  M08Location in ('" & _Location & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
            ElseIf Trim(txtDepartment.Text) <> "" Then
                vcWharer = " M08Retailer in ('" & _Department & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

            ElseIf Trim(txtMerchant.Text) <> "" Then
                vcWharer = " M08Merchant in ('" & _Merchant & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
            ElseIf Trim(txtBU.Text) <> "" Then
                vcWharer = " M14Name in ('" & _Bu & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
            ElseIf Trim(txtLocation.Text) <> "" Then
                vcWharer = " M08Location in ('" & _Location & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))


            ElseIf Trim(cboStatus.Text) <> "" Then
                If Trim(cboStatus.Text) = "Below One Month" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over One Months" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Two Months" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Three Months" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Fore Months" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over Five Months" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                ElseIf Trim(cboStatus.Text) = "Over Six Months" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(cboStatus.Text) = "Over One Year" Then
                    vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                End If
            Else
                Sql = "select * from M08Stock order by M08Meterial"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            End If


            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                Dim Value As Double
                '  If I = 747 Then
                '  MsgBox("")
                '  End If
                ' Value = M01.Tables(0).Rows(I)("M08Qty_Mtr")
                If CInt(M01.Tables(0).Rows(I)("M08Qty_Mtr")) > 0 Then
                    'newRow("Qty(Mtr)") = Value.ToString("#.#", System.Globalization.CultureInfo.InvariantCulture)
                    ' newRow("Qty(Mtr)") = M01.Tables(0).Rows(I)("M08Qty_Mtr")

                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    newRow("Location") = M01.Tables(0).Rows(I)("M08Location")
                    newRow("Material") = M01.Tables(0).Rows(I)("M08Meterial")
                    newRow("Material Description") = M01.Tables(0).Rows(I)("M08Dis")
                    newRow("Retailer") = M01.Tables(0).Rows(I)("M08Retailer")
                    newRow("Merchant") = M01.Tables(0).Rows(I)("M08Merchant")
                    newRow("Sales Order") = M01.Tables(0).Rows(I)("M08Sales_Order")
                    newRow("Line Item") = M01.Tables(0).Rows(I)("M08Line_Item")
                    newRow("Roll No") = M01.Tables(0).Rows(I)("M08RollNo")
                    Value = 0
                    ' MsgBox(M01.Tables(0).Rows(I)("M08Qty_Mtr"))
                    Value = M01.Tables(0).Rows(I)("M08Qty_Mtr")
                    If CInt(Value) > 0 Then
                        newRow("Qty(Mtr)") = Value.ToString("#.#", System.Globalization.CultureInfo.InvariantCulture)
                        ' newRow("Qty(Mtr)") = M01.Tables(0).Rows(I)("M08Qty_Mtr")
                    End If
                    _QtyMtr = _QtyMtr + M01.Tables(0).Rows(I)("M08Qty_Mtr")
                    Value = 0
                    Value = M01.Tables(0).Rows(I)("M08Qty_KG")
                    If Value > 0 Then
                        newRow("Qty(Kg)") = Value.ToString("#.#", System.Globalization.CultureInfo.InvariantCulture)
                    End If
                    'newRow("Qty(Kg)") = M01.Tables(0).Rows(I)("M08Qty_KG")
                    _QtyKG = _QtyKG + M01.Tables(0).Rows(I)("M08Qty_KG")
                    newRow("GRN Date") = Month(M01.Tables(0).Rows(I)("M08TR_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(I)("M08TR_Date")) & "/" & Year(M01.Tables(0).Rows(I)("M08TR_Date"))
                    date2 = M01.Tables(0).Rows(I)("M08TR_Date")
                    date1 = Today
                    diff1 = date1.Subtract(date2)
                    _Date = diff1.Days
                    If _Date <= 30 Then
                        newRow("Ageing") = "Below One Month"
                    ElseIf _Date > 30 And _Date < 60 Then
                        newRow("Ageing") = "Over One Months"
                    ElseIf _Date >= 60 And _Date < 90 Then
                        newRow("Ageing") = "Over Two Months"
                    ElseIf _Date >= 90 And _Date < 120 Then
                        newRow("Ageing") = "Over Three Months"
                    ElseIf _Date >= 120 And _Date < 150 Then
                        newRow("Ageing") = "Over Fore Months"
                    ElseIf _Date >= 150 And _Date < 180 Then
                        newRow("Ageing") = "Over Five Months"
                    ElseIf _Date >= 180 And _Date < 365 Then
                        newRow("Ageing") = "Over Six Months"

                    ElseIf _Date >= 365 Then
                        newRow("Ageing") = "Over One Year"

                    End If

                    Sql = "select * from T09Stock_Comments where  T09Year=" & Year(Today) & " and T09Location='" & M01.Tables(0).Rows(I)("M08Location") & "' and T09Roll_No='" & M01.Tables(0).Rows(I)("M08RollNo") & "' order by T09Date DEsc"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(dsUser) Then
                        If _WeekNo = dsUser.Tables(0).Rows(0)("T09Week") Then
                            newRow("New Update") = dsUser.Tables(0).Rows(0)("T09Comment")
                            newRow("latest update") = dsUser.Tables(0).Rows(0)("T09Comment")
                        Else
                            newRow("latest update") = dsUser.Tables(0).Rows(0)("T09Comment")
                        End If
                        If IsDBNull(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = True Then
                        Else

                            If Year(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = "1900" Then
                            Else
                                newRow("Dedline Date") = dsUser.Tables(0).Rows(0)("T09Ded_Date")

                            End If
                        End If
                    End If
                    c_dataCustomer1.Rows.Add(newRow)
                End If
                I = I + 1

            Next
            If I > 0 Then
                Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                newRow1("Qty(Mtr)") = _QtyMtr.ToString("#.#", System.Globalization.CultureInfo.InvariantCulture)
                newRow1("Qty(Kg)") = _QtyKG.ToString("#.#", System.Globalization.CultureInfo.InvariantCulture)
                c_dataCustomer1.Rows.Add(newRow1)

                Dim _Rcount As Integer
                _Rcount = UltraGrid4.Rows.Count
                UltraGrid4.Rows(_Rcount - 1).Cells(8).Appearance.BackColor = Color.Blue
                UltraGrid4.Rows(_Rcount - 1).Cells(9).Appearance.BackColor = Color.Blue
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MsgBox(I)
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try

    End Function

    Function Load_Status()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M25Dis as [##] from M25Aging order by M25code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboStatus
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 245
                End With
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function


    Private Sub frmFG_Stock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'txtFromDate.Text = Today
        'txtTodate.Text = Today
        Call Load_Status()
        txtBU.ReadOnly = True
        txtCustomer.ReadOnly = True
        txtDepartment.ReadOnly = True
        txtLocation.ReadOnly = True
        txtMerchant.ReadOnly = True


        Call Load_Gride()
        'Call Load_Gride_WithRecords()
        Call Load_GrideCustomer()
        Call Load_GrideRetailer()
        Call Load_GrideMerchant()
        Call Load_GrideLocation()
        Call Load_GrideBU()
    End Sub

    Function Load_GrideCustomer()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer2 = CustomerDataClass.MakeDataTableCheck_Customer
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideRetailer()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer3 = CustomerDataClass.MakeDataTableCheck_Retailer
        UltraGrid3.DataSource = c_dataCustomer3
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideMerchant()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer4 = CustomerDataClass.MakeDataTableCheck_Merch
        UltraGrid1.DataSource = c_dataCustomer4
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideLocation()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer6 = CustomerDataClass.MakeDataTableCheck_Location
        UltraGrid5.DataSource = c_dataCustomer6
        With UltraGrid5
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_FG
        UltraGrid4.DataSource = c_dataCustomer1
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 120
            .DisplayLayout.Bands(0).Columns(3).Width = 60
            .DisplayLayout.Bands(0).Columns(4).Width = 60
            .DisplayLayout.Bands(0).Columns(5).Width = 60
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(11).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideBU()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer5 = CustomerDataClass.MakeDataTableCheck_BU
        UltraGrid6.DataSource = c_dataCustomer5
        With UltraGrid6
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '  .DisplayLayout.Bands(0).Columns(4).Width = 90
            '  .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            Sql = "select M01Cuatomer_Name from M01Sales_Order_SAP inner join M08Stock on M08Sales_Order=M01Sales_Order  and M08Line_Item=M01Line_Item group by M01Cuatomer_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Customer Name") = M01.Tables(0).Rows(I)("M01Cuatomer_Name")

                c_dataCustomer2.Rows.Add(newRow)
                I = I + 1

            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Function Load_Retailer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            Sql = "select M08Retailer from M08Stock  group by M08Retailer"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer3.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Retailer Name") = M01.Tables(0).Rows(I)("M08Retailer")

                c_dataCustomer3.Rows.Add(newRow)
                I = I + 1

            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Function Load_Merch()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            Sql = "select M13Merchant from M13Biz_Unit  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer4.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Merchant") = M01.Tables(0).Rows(I)("M13Merchant")

                c_dataCustomer4.Rows.Add(newRow)
                I = I + 1

            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Function Load_BU()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            Sql = "select M14Name from M14Retailer  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer5.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("BU") = M01.Tables(0).Rows(I)("M14Name")

                c_dataCustomer5.Rows.Add(newRow)
                I = I + 1

            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            Sql = "select M08Location from M08Stock group by M08Location "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer6.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Location") = M01.Tables(0).Rows(I)("M08Location")

                c_dataCustomer6.Rows.Add(newRow)
                I = I + 1

            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Dim i As Integer
        ' UltraGrid1.Visible = False
        'UltraGrid3.Visible = False
        _Customer = ""
        If UltraGrid2.Visible = False Then

            ' Call Load_GrideCustomer()
            Call Load_Customer()
            UltraGrid2.Visible = True
        Else
            txtCustomer.Text = ""
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If UltraGrid2.Rows(i).Cells(0).Value = True Then
                    If Trim(txtCustomer.Text) <> "" Then
                        txtCustomer.Text = txtCustomer.Text & "," & UltraGrid2.Rows(i).Cells(1).Value
                        _Customer = _Customer & "','" & UltraGrid2.Rows(i).Cells(1).Value
                    Else
                        txtCustomer.Text = UltraGrid2.Rows(i).Cells(1).Value
                        _Customer = UltraGrid2.Rows(i).Cells(1).Value


                    End If
                End If
                i = i + 1
            Next
            pln_Customer = _Customer
            UltraGrid2.Visible = False
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Load_Gride()
        Call Load_Gride_WithRecords()
    End Sub

    Private Sub chkCus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCus.CheckedChanged
        If chkCus.Checked = True Then
            chkDep.Checked = False
            chkAge.Checked = False
            chkLocation.Checked = False
            chkMerch.Checked = False
            chkOTD.Checked = False

            Call Load_GrideCustomer()
            Call Load_Customer()

        End If
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim _weekNo As Integer

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Try
            _weekNo = DatePart(DateInterval.WeekOfYear, Today)
            i = 0
            For Each uRow As UltraGridRow In UltraGrid4.Rows
                If Trim(Trim(UltraGrid4.Rows(i).Cells(0).Text)) <> "" Then
                    nvcFieldList1 = "select * from T09Stock_Comments where T09Week=" & _weekNo & " and T09Year=" & Year(Today) & " and T09Location='" & Trim(UltraGrid4.Rows(i).Cells(0).Text) & "' and T09Roll_No='" & Trim(UltraGrid4.Rows(i).Cells(7).Text) & "'"
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(dsUser) Then
                        ' If Trim(Trim(UltraGrid4.Rows(i).Cells(12).Text)) <> "" Then
                        nvcFieldList1 = "update T09Stock_Comments set T09Comment='" & Trim(Trim(UltraGrid4.Rows(i).Cells(12).Text)) & "',T09Date='" & Today & "',T09Ded_Date='" & UltraGrid4.Rows(i).Cells(14).Value & "'  where T09Week=" & _weekNo & " and T09Year=" & Year(Today) & " and T09Location='" & Trim(UltraGrid4.Rows(i).Cells(0).Text) & "' and T09Roll_No='" & Trim(UltraGrid4.Rows(i).Cells(7).Text) & "'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        'End If
                    Else
                        With UltraGrid4

                            'If Trim(Trim(UltraGrid4.Rows(i).Cells(12).Text)) <> "" Then
                            nvcFieldList1 = "Insert Into T09Stock_Comments(T09Location,T09Roll_No,T09Week,T09Year,T09Comment,T09Date,T09Ded_Date)" & _
                                                        " values('" & .Rows(i).Cells(0).Value & "','" & .Rows(i).Cells(7).Value & "','" & _weekNo & "','" & Year(Today) & "','" & .Rows(i).Cells(12).Value & "','" & Today & "','" & .Rows(i).Cells(14).Value & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            'End If
                        End With
                    End If

                End If
                i = i + 1
            Next

            MsgBox("Record Update Successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.Close()
            Call Load_Gride()
            txtBU.Text = ""
            txtCustomer.Text = ""
            txtDepartment.Text = ""
            txtLocation.Text = ""
            txtMerchant.Text = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()

            End If
        End Try
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Dim i As Integer
        ' UltraGrid1.Visible = False
        'UltraGrid3.Visible = False
        _Customer = ""
        If UltraGrid3.Visible = False Then

            ' Call Load_GrideCustomer()
            Call Load_Retailer()
            UltraGrid3.Visible = True
        Else
            txtDepartment.Text = ""
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                If UltraGrid3.Rows(i).Cells(0).Value = True Then
                    If Trim(txtDepartment.Text) <> "" Then
                        txtDepartment.Text = txtDepartment.Text & "," & UltraGrid3.Rows(i).Cells(1).Value
                        _Department = _Department & "','" & UltraGrid3.Rows(i).Cells(1).Value
                    Else
                        txtDepartment.Text = UltraGrid3.Rows(i).Cells(1).Value
                        _Department = UltraGrid3.Rows(i).Cells(1).Value


                    End If
                End If
                i = i + 1
            Next
            pln_Retailer = _Department
            UltraGrid3.Visible = False
        End If
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        Dim i As Integer
        ' UltraGrid1.Visible = False
        'UltraGrid3.Visible = False
        _Merchant = ""
        If UltraGrid1.Visible = False Then

            ' Call Load_GrideCustomer()
            Call Load_Merch()
            UltraGrid1.Visible = True
        Else
            txtMerchant.Text = ""
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If UltraGrid1.Rows(i).Cells(0).Value = True Then
                    If Trim(txtMerchant.Text) <> "" Then
                        txtMerchant.Text = txtMerchant.Text & "," & UltraGrid1.Rows(i).Cells(1).Value
                        _Merchant = _Merchant & "','" & UltraGrid1.Rows(i).Cells(1).Value
                    Else
                        txtMerchant.Text = UltraGrid1.Rows(i).Cells(1).Value
                        _Merchant = UltraGrid1.Rows(i).Cells(1).Value


                    End If
                End If
                i = i + 1
            Next
            pln_Merchnt = _Merchant
            UltraGrid1.Visible = False
        End If
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim i As Integer
        ' UltraGrid1.Visible = False
        'UltraGrid3.Visible = False
        _Location = ""

        If UltraGrid5.Visible = False Then
            Call Load_GrideLocation()
            ' Call Load_GrideCustomer()
            Call Load_Location()
            UltraGrid5.Visible = True
        Else
            txtLocation.Text = ""
            ' MsgBox(UltraGrid5.Rows.Count)
            For Each uRow As UltraGridRow In UltraGrid5.Rows
                If UltraGrid5.Rows(i).Cells(0).Value = True Then
                    If Trim(txtLocation.Text) <> "" Then
                        txtLocation.Text = txtLocation.Text & "," & UltraGrid5.Rows(i).Cells(1).Value
                        _Location = _Location & "','" & UltraGrid5.Rows(i).Cells(1).Value
                    Else
                        txtLocation.Text = UltraGrid5.Rows(i).Cells(1).Value
                        _Location = UltraGrid5.Rows(i).Cells(1).Value


                    End If
                End If
                i = i + 1
            Next
            pln_Location = _Location
            UltraGrid5.Visible = False
        End If
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim i As Integer
        ' UltraGrid1.Visible = False
        'UltraGrid3.Visible = False
        _bu = ""
        If UltraGrid6.Visible = False Then

            ' Call Load_GrideCustomer()
            Call Load_BU()
            UltraGrid6.Visible = True
        Else
            txtBU.Text = ""
            For Each uRow As UltraGridRow In UltraGrid6.Rows
                If UltraGrid6.Rows(i).Cells(0).Value = True Then
                    If Trim(txtBU.Text) <> "" Then
                        txtBU.Text = txtBU.Text & "," & UltraGrid6.Rows(i).Cells(1).Value
                        _Bu = _Bu & "','" & UltraGrid6.Rows(i).Cells(1).Value
                    Else
                        txtBU.Text = UltraGrid6.Rows(i).Cells(1).Value
                        _Bu = UltraGrid6.Rows(i).Cells(1).Value


                    End If
                End If
                i = i + 1
            Next
            pln_BU = _Bu
            UltraGrid6.Visible = False
        End If
    End Sub

    Private Sub UltraButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton7.Click
        Dim i As Integer
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim BatchNo As String
        Dim _Coment As String
        Dim _PossibleDate As String
        Dim _Reason As String
        Dim _LabDip As String

        Try
            Dim x1 As Integer
            strFileName = ConfigurationManager.AppSettings("MtnUpload") + "\LOTs.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                BatchNo = Trim(fields(0))
                _Coment = Trim(fields(1))
                If Trim(BatchNo) = "60811363" Then
                    '  MsgBox("")
                End If
                _PossibleDate = ""
                If IsDate(Trim(fields(2))) Then
                    _PossibleDate = Trim(fields(2))
                End If

                i = 1
                For Each uRow As UltraGridRow In UltraGrid4.Rows
                    'QUARANTINE REASON FOR REPORT TABLE
                    With UltraGrid4
                        If (UltraGrid4.Rows.Count) = i Then
                        Else

                            If Trim(.Rows(i).Cells(7).Text) = Trim(BatchNo) Then
                                UltraGrid4.Rows(i).Cells(12).Value = _Coment
                                If IsDate(_PossibleDate) Then
                                    UltraGrid4.Rows(i).Cells(14).Value = _PossibleDate
                                End If
                                '  UltraGrid4.Rows(i).Cells(15).Value = _Reason
                                ' UltraGrid4.Rows(i).Cells(16).Value = _LabDip
                                Exit For
                            End If
                        End If
                    End With
                    i = i + 1
                Next

                x1 = x1 + 1
            Next

            ' Con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
                'Con.close()

            End If
        End Try
    End Sub

    Private Sub UltraGrid6_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid6.InitializeLayout

    End Sub
End Class