
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
'Imports System.Drawing
'Imports Spire.XlS

Public Class frmOTD
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String
    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M01Sales_Order as [Sales Order] from M01Sales_Order_SAP where M01Sales_Order<>'' group by M01Sales_Order"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboSales_Order
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
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

    Function Load_PO()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M01PO as [PO Number] from M01Sales_Order_SAP group by M01PO"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboPO
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
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


    Function Upload_Fileotdv()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Delay_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double


        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _TollPLS As Integer
        Dim _TollMIN As Integer

        Try
            nvcFieldList1 = "delete from OTD_Records1"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



            strFileName = ConfigurationManager.AppSettings("FilePath") + "\otdv.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 2336 Then
                    '  MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                _Location = Trim(fields(15))
                If Trim(fields(0)) <> "" Then
                    _PO_No = Trim(fields(0))
                    _sales_Order = Trim(fields(1))
                    _LineItem = Trim(fields(2))
                    _Shadule = Trim(fields(3))
                    characterToRemove = "'"

                    _Shadule = (Replace(_Shadule, characterToRemove, ""))
                    _Material = Trim(fields(4))
                    _Material_Dis = Trim(fields(5))
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                    _Customer = Trim(fields(6))
                    _Customer = Microsoft.VisualBasic.Left(_Customer, 11)
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _Customer = (Replace(_Customer, characterToRemove, ""))
                    _Department = Trim(fields(7))
                    If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                        _Department = Microsoft.VisualBasic.Left(_Department, 3)
                    End If
                    _Merchnat = Trim(fields(8))
                    Dim B As String
                    'e
                    B = Microsoft.VisualBasic.Left(Trim(fields(9)), 6)
                    _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(9)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                    '_Del_Date = Trim(fields(9))
                    ' End If
                    _Order_Qty = Trim(fields(10))
                    _Del_Qty = Trim(fields(11))
                    _Delay_Qty = Trim(fields(12))
                    _FGStock = Trim(fields(13))
                    _Balance = Trim(fields(14))
                    _Location = Trim(fields(15))
                    _PRD_Qty = Trim(fields(16))
                    _PRD_OrderQty = Trim(fields(17))
                    _Grg_Qty = Trim(fields(18))
                    _NCComment = Trim(fields(19))

                    _Awaiting = Trim(fields(20))
                    _TollPLS = Trim(fields(22))
                    _TollMIN = Trim(fields(23))
                    _depComm = Trim(fields(25))
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _depComm = (Replace(_depComm, characterToRemove, ""))
                    _Comm2 = Trim(fields(26))
                    _OTDStatus = "Fales"

                    characterToRemove = ";"

                    'MsgBox(Trim(fields(9)))
                    _PO_No = (Replace(_PO_No, characterToRemove, ""))
                    Dim _Week As Integer

                    _Week = DatePart(DateInterval.WeekOfYear, _Del_Date)

                    nvcFieldList1 = "select * from OTD_Records1 where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & Trim(_LineItem) & "' and Del_Date='" & _Del_Date & "' and Location='" & _Location & "' and Prduct_Order='" & _PRD_Qty & "' and PRD_Qty='" & _PRD_OrderQty & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                    Else
                        nvcFieldList1 = "Insert Into OTD_Records1(PO_No,Sales_Order,Line_Item,Shadule,Metrrial,Met_Des,Customer,Department,Merchant,Del_Date,Del_Qty,FG_Stock,Balance,Location,Prduct_Order,PRD_Qty,Greige_Req_Qty,NC_Comment,Awaiting_Grie,Department_Comment,Comment2,Status,Run_Date,OTD_Year,OTD_Week,Order_Qty,Delay_Qty,Tollarance_PLS,Tollarance_MIN)" & _
                                                            " values('" & Trim(_PO_No) & "', '" & Trim(_sales_Order) & "','" & Trim(_LineItem) & "','" & Trim(_Shadule) & "','" & _Material & "','" & Trim(_Material_Dis) & "','" & Trim(_Customer) & "','" & Trim(_Department) & "','" & Trim(_Merchnat) & "','" & Trim(_Del_Date) & "','" & Trim(_Del_Qty) & "','" & Trim(_FGStock) & "','" & Trim(_Balance) & "','" & _Location & "','" & _PRD_Qty & "','" & _PRD_OrderQty & "','" & Trim(_Grg_Qty) & "','" & Trim(_NCComment) & "','" & Trim(_Awaiting) & "','" & Trim(_depComm) & "','" & Trim(_Comm2) & "','" & _OTDStatus & "','" & Today & "','" & Year(_Del_Date) & "','" & _Week & "','" & _Order_Qty & "','" & _Delay_Qty & "','" & _TollPLS & "','" & _TollMIN & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    ' pbCount.Value = pbCount.Value + 1


                    lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                    _PO_No = ""
                    _sales_Order = ""
                    _LineItem = ""
                    '_LineItem = ""
                    _Awaiting = ""
                    _Balance = 0
                    _TollPLS = 0
                    _TollMIN = 0
                    _Grg_Qty = 0
                    _PRD_OrderQty = 0
                    _PRD_Qty = ""
                    _NCComment = ""
                    _Del_Qty = 0
                    _Comm2 = ""
                    _Customer = ""
                    _FGStock = 0
                    _depComm = ""
                    _Material = ""
                    _Material_Dis = ""
                    _Merchnat = ""
                    _Department = ""
                    _Shadule = ""

                End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")

        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11 & "otdv.txt")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Delay_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double


        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _TollPLS As Integer
        Dim _TollMIN As Integer

        Try
            nvcFieldList1 = "delete from OTD_Records"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from OTD_SMS"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from FR_Update"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\otdv.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 2918 Then
                    ' MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                _Location = Trim(fields(15))
                If _Location <> "" Then
                    _PO_No = Trim(fields(0))
                    _sales_Order = Trim(fields(1))
                    _LineItem = Trim(fields(2))
                    _Shadule = Trim(fields(3))
                    characterToRemove = "'"

                    _Shadule = (Replace(_Shadule, characterToRemove, ""))
                    _Material = Trim(fields(4))
                    _Material_Dis = Trim(fields(5))
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                    _Customer = Trim(fields(6))
                    _Customer = Microsoft.VisualBasic.Left(_Customer, 11)
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _Customer = (Replace(_Customer, characterToRemove, ""))
                    _Department = Trim(fields(7))
                    If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                        _Department = Microsoft.VisualBasic.Left(_Department, 3)
                    End If
                    _Merchnat = Trim(fields(8))
                    Dim B As String
                    B = Microsoft.VisualBasic.Left(Trim(fields(9)), 6)
                    _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(9)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                    '_Del_Date = Trim(fields(9))
                    _Order_Qty = Trim(fields(10))
                    _Del_Qty = Trim(fields(11))
                    _Delay_Qty = Trim(fields(12))
                    _FGStock = Trim(fields(13))
                    _Balance = Trim(fields(14))
                    _Location = Trim(fields(15))
                    _PRD_Qty = Trim(fields(16))
                    _PRD_OrderQty = Trim(fields(17))
                    _Grg_Qty = Trim(fields(18))
                    _NCComment = Trim(fields(19))
                    '  _NCComment = " "
                    _Awaiting = Trim(fields(20))
                    _TollPLS = Trim(fields(22))
                    _TollMIN = Trim(fields(23))
                    _depComm = Trim(fields(25))
                    characterToRemove = "'"

                    'MsgBox(Trim(fields(9)))
                    _depComm = (Replace(_depComm, characterToRemove, ""))
                    _Comm2 = Trim(fields(26))
                    _OTDStatus = "Fales"

                    characterToRemove = ";"

                    'MsgBox(Trim(fields(9)))
                    _PO_No = (Replace(_PO_No, characterToRemove, ""))
                    Dim _Week As Integer

                    _Week = DatePart(DateInterval.WeekOfYear, _Del_Date)

                    nvcFieldList1 = "select * from OTD_Records where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & Trim(_LineItem) & "' and Del_Date='" & _Del_Date & "' and Location='" & _Location & "' and Prduct_Order='" & _PRD_Qty & "' and PRD_Qty='" & _PRD_OrderQty & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                    Else
                        nvcFieldList1 = "Insert Into OTD_Records(PO_No,Sales_Order,Line_Item,Shadule,Metrrial,Met_Des,Customer,Department,Merchant,Del_Date,Del_Qty,FG_Stock,Balance,Location,Prduct_Order,PRD_Qty,Greige_Req_Qty,NC_Comment,Awaiting_Grie,Department_Comment,Comment2,Status,Run_Date,OTD_Year,OTD_Week,Order_Qty,Delay_Qty,Tollarance_PLS,Tollarance_MIN)" & _
                                                            " values('" & Trim(_PO_No) & "', '" & Trim(_sales_Order) & "','" & Trim(_LineItem) & "','" & Trim(_Shadule) & "','" & _Material & "','" & Trim(_Material_Dis) & "','" & Trim(_Customer) & "','" & Trim(_Department) & "','" & Trim(_Merchnat) & "','" & Trim(_Del_Date) & "','" & Trim(_Del_Qty) & "','" & Trim(_FGStock) & "','" & Trim(_Balance) & "','" & _Location & "','" & _PRD_Qty & "','" & _PRD_OrderQty & "','" & Trim(_Grg_Qty) & "','" & Trim(_NCComment) & "','" & Trim(_Awaiting) & "','" & Trim(_depComm) & "','" & Trim(_Comm2) & "','" & _OTDStatus & "','" & Today & "','" & Year(_Del_Date) & "','" & _Week & "','" & _Order_Qty & "','" & _Delay_Qty & "','" & _TollPLS & "','" & _TollMIN & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    ' pbCount.Value = pbCount.Value + 1


                    lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                    _PO_No = ""
                    _sales_Order = ""
                    _LineItem = ""
                    '_LineItem = ""
                    _Awaiting = ""
                    _Balance = 0
                    _TollPLS = 0
                    _TollMIN = 0
                    _Grg_Qty = 0
                    _PRD_OrderQty = 0
                    _PRD_Qty = ""
                    _NCComment = ""
                    _Del_Qty = 0
                    _Comm2 = ""
                    _Customer = ""
                    _FGStock = 0
                    _depComm = ""
                    _Material = ""
                    _Material_Dis = ""
                    _Merchnat = ""
                    _Department = ""
                    _Shadule = ""

                End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            Call Edit_OTD_Status()
            Call FR_Update()

        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11 & " on otdv.txt")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Edit_OTD_Status()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As String
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Delay_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double



        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String

        Dim X11 As Integer
        Dim _CusCode As String

        Try
            nvcFieldList1 = "delete from OTD_SMS"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\otdStatus.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 2 Then
                    ' MsgBox("")
                End If
                ' _Location = Trim(fields(15))
                'If _Location <> "" Then

                _sales_Order = Trim(fields(1))
                _LineItem = Trim(fields(3))
                _CusCode = Trim(fields(0))
                _Customer = Trim(fields(16))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Customer = (Replace(_Customer, characterToRemove, " "))

                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(8)), 6)
                _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(8)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))

                ' _Del_Date = Trim(fields(8))


                If Trim(fields(11)) = "1" Then
                    _OTDStatus = "True"
                Else
                    _OTDStatus = "Fales"
                End If
                '_OTDStatus = "Fales"

                characterToRemove = "."

                'MsgBox(Trim(fields(9)))
                '_Del_Date = (Replace(_Del_Date, characterToRemove, "/"))
                Dim A1 As String

                A1 = (Microsoft.VisualBasic.Left(_Del_Date, 5))
                ' _Del_Date = Microsoft.VisualBasic.Right(A1, 2) & "/" & Microsoft.VisualBasic.Left(A1, 2) & "/" & Microsoft.VisualBasic.Right(_Del_Date, 4)
                'Dim oDate As DateTime = Convert.ToDateTime(_Del_Date)
                'MsgBox(oDate.Day & " " & oDate.Month & "  " & oDate.Year)
                _Department = Trim(fields(18))
                If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                    _Department = Microsoft.VisualBasic.Left(_Department, 3)
                Else

                    _Department = Trim(fields(18))
                End If

                _Merchnat = Trim(fields(17))


                'nvcFieldList1 = "select * from OTD_Records where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & Trim(_LineItem) & "' and Del_Date='" & _Del_Date & "' "
                'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(M01) Then
                '    nvcFieldList1 = "update OTD_Records set Cus_Code='" & _CusCode & "',Customer='" & _Customer & "',Status='" & _OTDStatus & "' where Sales_Order='" & Trim(_sales_Order) & "' and Line_Item='" & _LineItem & "' and Del_Date='" & _Del_Date & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                'End If
                ' pbCount.Value = pbCount.Value + 1

                'INSERT OTD_SMS

                nvcFieldList1 = "SELECT * FROM OTD_SMS WHERE Sales_Order='" & Trim(_sales_Order) & "' AND Line_Item='" & _LineItem & "' AND Del_Date='" & _Del_Date & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into OTD_SMS(Sales_Order,Line_Item,Cus_Code,Customer,Del_Date,Status,Department,Merchant)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_CusCode) & "','" & Trim(_Customer) & "','" & _Del_Date & "','" & Trim(_OTDStatus) & "','" & _Department & "','" & _Merchnat & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                _CusCode = ""
                _Customer = ""
                _sales_Order = ""
                _LineItem = ""
                _OTDStatus = ""
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            ' MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")

            'nvcFieldList1 = "select * from OTD_SMS S inner join OTD_Records r on r.Sales_Order=s.Sales_Order and r.Line_Item=s.Line_Item and r.Del_Date=s.Del_Date where s.Status='Fales' and r.Delay_Qty=r.FG_Stock"
            'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            'I = 0
            'For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '    nvcFieldList1 = "update OTD_Records set Status='True' where Sales_Order='" & Trim(M01.Tables(0).Rows(I)("Sales_Order")) & "' and Line_Item='" & Trim(M01.Tables(0).Rows(I)("Line_Item")) & "' and Del_Date='" & M01.Tables(0).Rows(I)("Del_Date") & "'"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '    I = I + 1
            'Next
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function FR_Update()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As String
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Delay_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double


        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String
        Dim _DyeMC As String

        Dim X11 As Integer
        Dim _CusCode As String

        Try
            nvcFieldList1 = "delete from FR_Update"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\FR_otd.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 278 Then
                    '  MsgBox("")
                End If
                ' _Location = Trim(fields(15))
                'If _Location <> "" Then

                If X11 = 0 Then

                Else
                    _sales_Order = Trim(fields(0))
                    _LineItem = Trim(fields(1))

                    _Department = Trim(fields(2))
                    _Del_Date = Trim(fields(3))

                    _DyeMC = Trim(fields(4))

                    '_OTDStatus = "Fales"

                    characterToRemove = "-"
                    'MsgBox(Trim(fields(9)))
                    _Del_Date = (Replace(_Del_Date, characterToRemove, "/"))
                    Dim A1 As String

                    ' A1 = (Microsoft.VisualBasic.Left(_Del_Date, 5))
                    '_Del_Date = Microsoft.VisualBasic.Right(A1, 2) & "/" & Microsoft.VisualBasic.Left(A1, 2) & "/" & Microsoft.VisualBasic.Right(_Del_Date, 4)
                    'Dim oDate As DateTime = Convert.ToDateTime(_Del_Date)
                    'MsgBox(oDate.Day & " " & oDate.Month & "  " & oDate.Year)



                    nvcFieldList1 = "SELECT * FROM FR_Update WHERE Batch_No='" & Trim(_sales_Order) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "update FR_Update set Stock_Code='" & _LineItem & "',Recipy_Status='" & _Department & "',Dye_Pln_Date='" & _Del_Date & "',Dye_Machine='" & _DyeMC & "' where Batch_No='" & _sales_Order & "'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into FR_Update(Batch_No,Stock_Code,Recipy_Status,Dye_Pln_Date,Dye_Machine)" & _
                                                            " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_Department) & "','" & Trim(_Del_Date) & "','" & _DyeMC & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    _CusCode = ""
                    _Customer = ""
                    _sales_Order = ""
                    _Department = ""
                    _LineItem = ""
                    _OTDStatus = ""

                    'cmdEdit.Enabled = True
                End If
                X11 = X11 + 1
            Next
            ' MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Upload_File()
        Call Upload_Fileotdv()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmOTD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today

        ' Call Load_Customer()
        'Call Load_Merchant()

        'Call Load_Department()
        Call Load_Status()
        Call Load_Gride()
        Call Load_Customer()

        Call Load_GrideDep()
        Call Load_Department()

        Call Load_GrideMerch()
        Call Load_Merch()

        txtCustomer.ReadOnly = True
        txtDepartment.ReadOnly = True
        txtMerchant.ReadOnly = True

        Call Load_Combo()
        Call Load_PO()
    End Sub

    Function Load_Status()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select Status as [Status] from OTD_SMS group by Status"
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

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation    
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCheck_Customer
        UltraGrid2.DataSource = c_dataCustomer1
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

    Function Load_GrideDep()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCheck_Dep
        UltraGrid1.DataSource = c_dataCustomer1
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

    Function Load_GrideMerch()
        Dim CustomerDataClass As New DAL_InterLocation
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCheck_Merch
        UltraGrid3.DataSource = c_dataCustomer1
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

    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            If Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                Sql = "select Customer from OTD_SMS where Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')   group by Customer"
            ElseIf Trim(txtDepartment.Text) <> "" Then
                Sql = "select Customer from OTD_SMS where Department in ('" & _Department & "') group by Customer"
            ElseIf Trim(txtMerchant.Text) <> "" Then
                Sql = "select Customer from OTD_SMS where Merchant in ('" & _Merchant & "') group by Customer"
            Else
                Sql = "select Customer from OTD_SMS group by Customer"
            End If

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Customer Name") = M01.Tables(0).Rows(I)("Customer")

                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1

            Next
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

    Function Load_Department()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try
            If Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                Sql = "select Department from OTD_SMS where Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "')   group by Department"
            ElseIf Trim(txtCustomer.Text) <> "" Then
                Sql = "select Department from OTD_SMS where Customer in ('" & _Customer & "') group by Department"
            ElseIf Trim(txtMerchant.Text) <> "" Then
                Sql = "select Department from OTD_SMS where Merchant in ('" & _Merchant & "') group by Department"
            Else
                Sql = "select Department from OTD_SMS group by Department"
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Department") = M01.Tables(0).Rows(I)("Department")

                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1

            Next
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

    Function Load_Merch()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer

        Try

            If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                Sql = "select Merchant from OTD_SMS where Customer in ('" & _Customer & "') and Department in ('" & _Department & "')   group by Merchant"
            ElseIf Trim(txtCustomer.Text) <> "" Then
                Sql = "select Merchant from OTD_SMS where Customer in ('" & _Customer & "') group by Merchant"
            ElseIf Trim(txtDepartment.Text) <> "" Then
                Sql = "select Merchant from OTD_SMS where Department in ('" & _Department & "') group by Merchant"
            Else
                Sql = "select Merchant from OTD_SMS group by Merchant"
            End If


            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0

            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("##") = False
                newRow("Merchant") = M01.Tables(0).Rows(I)("Merchant")

                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1

            Next
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

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Dim i As Integer
        UltraGrid1.Visible = False
        UltraGrid3.Visible = False
        _Customer = ""
        If UltraGrid2.Visible = False Then
            Call Load_Gride()
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
            UltraGrid2.Visible = False
        End If
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Dim i As Integer
        UltraGrid2.Visible = False
        UltraGrid3.Visible = False
        _Department = ""
        If UltraGrid1.Visible = False Then
            Call Load_GrideDep()
            Call Load_Department()
            UltraGrid1.Visible = True
        Else
            txtDepartment.Text = ""
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If UltraGrid1.Rows(i).Cells(0).Value = True Then
                    If txtDepartment.Text <> "" Then
                        txtDepartment.Text = txtDepartment.Text & "," & UltraGrid1.Rows(i).Cells(1).Value
                        _Department = _Department & "','" & UltraGrid1.Rows(i).Cells(1).Value
                    Else
                        txtDepartment.Text = UltraGrid1.Rows(i).Cells(1).Value
                        _Department = UltraGrid1.Rows(i).Cells(1).Value
                    End If
                End If
                i = i + 1
            Next
            UltraGrid1.Visible = False
        End If
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        Dim i As Integer
        UltraGrid2.Visible = False
        UltraGrid1.Visible = False
        _Merchant = ""
        If UltraGrid3.Visible = False Then
            Call Load_GrideMerch()
            Call Load_Merch()
            UltraGrid3.Visible = True
        Else
            txtMerchant.Text = ""
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                If UltraGrid3.Rows(i).Cells(0).Value = True Then
                    If txtMerchant.Text <> "" Then
                        txtMerchant.Text = txtMerchant.Text & "," & UltraGrid3.Rows(i).Cells(1).Value
                        _Merchant = _Merchant & "','" & UltraGrid3.Rows(i).Cells(1).Value
                    Else
                        txtMerchant.Text = UltraGrid3.Rows(i).Cells(1).Value
                        _Merchant = UltraGrid3.Rows(i).Cells(1).Value
                    End If
                End If
                i = i + 1
            Next
            UltraGrid3.Visible = False
        End If
    End Sub

    Private Sub chkCus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCus.CheckedChanged
        If chkCus.Checked = True Then
            UltraButton3.Enabled = True
        Else
            UltraButton3.Enabled = False
        End If
    End Sub

    Private Sub chkDep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDep.CheckedChanged
        If chkDep.Checked = True Then
            UltraButton4.Enabled = True
        Else
            UltraButton4.Enabled = False
        End If
    End Sub

    Private Sub chkMerch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMerch.CheckedChanged
        If chkMerch.Checked = True Then
            UltraButton5.Enabled = True
        Else
            UltraButton5.Enabled = False
        End If
    End Sub

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim tblDye As DataSet

        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        Dim _FirstChr As Integer
        Dim _Possible_Date As Date
        Dim _Last As Integer

        Dim exc As New Application

        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

        exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer
        Dim T04 As DataSet
        Dim n_per As Double
        Dim Y As Integer
        Dim _cOUNT As Integer
        Dim _Fail_Batch As Integer
        Dim _TotalFail As Integer

        _TotalFail = 0
        _Fail_Batch = 0
        Try
            '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
            ' workbooks.Application.Sheets.Add()
            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 12
            worksheet2.Columns("B").ColumnWidth = 8
            worksheet2.Columns("C").ColumnWidth = 8
            worksheet2.Columns("D").ColumnWidth = 60
            worksheet2.Columns("K").ColumnWidth = 20
            worksheet2.Columns("L").ColumnWidth = 20
            worksheet2.Columns("M").ColumnWidth = 20
            worksheet2.Columns("N").ColumnWidth = 18
            worksheet2.Columns("O").ColumnWidth = 25
            worksheet2.Columns("P").ColumnWidth = 20

            worksheet1.Cells(2, 4) = "OTD Details"
            worksheet1.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("d2:e2").Interior.Color = RGB(165, 165, 165)
            worksheet1.Rows(2).Font.size = 8
            worksheet1.Rows(2).Font.name = "Tahoma"
            worksheet2.Rows(2).rowheight = 24.25
            worksheet1.Range("d2:e2").MergeCells = True
            worksheet1.Range("d2:e2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("e2", "e2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e2", "e2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e2", "e2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Cells(3, 4) = "No of Pass Line Items"
            worksheet1.Cells(4, 4) = "No:of fail line items"
            worksheet1.Cells(5, 4) = "Total Line Items"
            worksheet1.Cells(6, 4) = "Actual  OTD %"
            worksheet1.Cells(7, 4) = "Possible OTD %"

            'worksheet1.Rows(3).Font.size = 8
            'worksheet1.Rows(4).Font.size = 8
            'worksheet1.Rows(5).Font.size = 8
            'worksheet1.Rows(6).Font.size = 8
            'worksheet1.Rows(7).Font.size = 8

            i = 4
            Dim _Char As Integer
            _Char = 68
            ' MsgBox(ChrW(_Char))
            For i = 3 To 7
                _Char = 68
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                _Char = _Char + 1

                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            Next
            '-----------------------------------------------------------------
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True'  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True'  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and Customer in ('" & _Customer & "')   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Department in ('" & _Department & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True' and  Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='True'   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                End If
            Else
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') group by Status"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  group by Status"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') group by Status"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') group by Status"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and Customer in ('" & _Customer & "')  group by Status"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and  Department in ('" & _Department & "')  group by Status"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' and  Merchant in ('" & _Merchant & "') group by Status"
                Else

                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Status='True' group by Status"
                End If
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(3, 5) = T01.Tables(0).Rows(0)("Status")
                worksheet1.Cells(3, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            End If

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where Status='Fales' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales'  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where Status='Fales' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales'  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where Status='Fales' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and Customer in ('" & _Customer & "')   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Department in ('" & _Department & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales' and  Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Status='Fales'   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                End If
            Else
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') group by Sales_Order,Line_Item,del_Date"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  group by Sales_Order,Line_Item,del_Date"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') group by Sales_Order,Line_Item,del_Date"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') group by Sales_Order,Line_Item,del_Date"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  group by Sales_Order,Line_Item,del_Date"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Department in ('" & _Department & "')  group by Sales_Order,Line_Item,del_Date"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Merchant in ('" & _Merchant & "') group by Sales_Order,Line_Item,del_Date"
                Else

                    SQL = "select count(status) as status from OTD_Records where del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' group by  Sales_Order,Line_Item,del_Date "
                End If
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(4, 5) = T01.Tables(0).Rows.Count
                worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            End If

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   M01PO='" & Trim(cboPO.Text) & "' and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order,M01PO"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Department in ('" & _Department & "')  and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Merchant in ('" & _Merchant & "') and M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   M01PO='" & Trim(cboPO.Text) & "'  group by Status,M01PO"
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "')   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   Customer in ('" & _Customer & "')   and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Department in ('" & _Department & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where    Merchant in ('" & _Merchant & "')  and M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                Else
                    SQL = "select  count(Status) as Status from OTD_SMS inner join M01Sales_Order_SAP on M01Sales_Order=Sales_Order and Line_Item=M01Line_Item  where   M01Sales_Order='" & Trim(cboSales_Order.Text) & "' group by Status,M01Sales_Order"
                End If
            Else
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') group by Status"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  group by Status"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Merchant in ('" & _Merchant & "') group by Status"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') group by Status"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Customer in ('" & _Customer & "')  group by Status"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Department in ('" & _Department & "')  group by Status"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  count(Status) as Status from OTD_SMS where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and  Merchant in ('" & _Merchant & "') group by Status"
                Else

                    SQL = "select  count(PO_No) as Status from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Delay_Qty<>FG_Stock group by Sales_Order,Line_Item,Del_Date"
                End If
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            Dim _Tot As Double
            _Tot = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _Tot = _Tot + T01.Tables(0).Rows(i)("Status")
                i = i + 1
            Next

            worksheet1.Cells(5, 5) = _Tot
            worksheet1.Cells(5, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(3, 5) = "=e5-e4"
            worksheet1.Cells(3, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(3, 5)

            worksheet1.Cells(6, 5) = "=e3/e5"
            worksheet1.Cells(6, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(6, 5)
            range1.NumberFormat = "0.0%"


            worksheet1.Cells(10, 4) = "Department "
            worksheet1.Cells(10, 5) = "Qty (m)"
            worksheet1.Cells(10, 6) = "No of Batch"
            worksheet1.Cells(10, 7) = "Agreed Qty (m) "
            worksheet1.Cells(10, 8) = "Release Qty (m)"
            worksheet1.Cells(10, 9) = "Week release Qty "
            worksheet1.Cells(10, 10) = "Pending Qty"

            worksheet1.Range("d10:j10").Interior.Color = RGB(165, 165, 165)
            worksheet2.Rows(10).Font.Bold = True
            worksheet1.Rows(10).Font.size = 8
            worksheet1.Rows(10).Font.name = "Tahoma"
            worksheet2.Rows(10).rowheight = 30.25
            worksheet2.Columns("e").ColumnWidth = 13
            worksheet2.Columns("f").ColumnWidth = 15
            worksheet2.Columns("g").ColumnWidth = 15
            worksheet2.Columns("h").ColumnWidth = 18
            worksheet2.Columns("i").ColumnWidth = 18
            worksheet2.Columns("j").ColumnWidth = 15
            worksheet1.Columns("Q").ColumnWidth = 28

            worksheet1.Range("d10:d10").MergeCells = True
            worksheet1.Range("d10:d10").VerticalAlignment = XlVAlign.xlVAlignCenter
            ' worksheet1.Range("c10:c10").Interior.Color = RGB(255, 0, 0)

            worksheet1.Range("e10:e10").MergeCells = True
            worksheet1.Range("e10:e10").VerticalAlignment = XlVAlign.xlVAlignCenter
            'worksheet1.Range("d10:d10").Interior.Color = RGB(217, 151, 149)

            worksheet1.Range("f10:f10").MergeCells = True
            worksheet1.Range("f10:f10").VerticalAlignment = XlVAlign.xlVAlignCenter
            '  worksheet1.Range("e10:e10").Interior.Color = RGB(250, 192, 144)
            worksheet1.Cells(10, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("g10:g10").MergeCells = True
            worksheet1.Range("g10:g10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(10, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("h10:h10").MergeCells = True
            worksheet1.Range("h10:h10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(10, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("i10:i10").MergeCells = True
            worksheet1.Range("i10:i10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(10, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("j10:j10").MergeCells = True
            worksheet1.Range("j10:j10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(10, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Range("j10:j10").MergeCells = True
            'worksheet1.Range("j10:j10").VerticalAlignment = XlVAlign.xlVAlignCenter
            'worksheet1.Cells(10, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

            Dim x As Integer
            i = 10
            _Char = 68
            For x = 1 To 7
                ' _Char = 67
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                _Char = _Char + 1

                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            Next

            worksheet1.Cells(11, 4) = "To Be Planned"
            worksheet1.Cells(12, 4) = "Aw yarn "
            worksheet1.Cells(13, 4) = "Aw Dyed Yarn"
            worksheet1.Cells(14, 4) = "Aw Greige "
            worksheet1.Cells(15, 4) = "To be Batch Card Send "
            worksheet1.Cells(16, 4) = "Aw Recipe "
            worksheet1.Cells(17, 4) = "Aw prepare"

            worksheet1.Cells(18, 4) = "Aw Presetting"
            worksheet1.Cells(19, 4) = "Aw Dyeing "
            worksheet1.Cells(20, 4) = "Aw for 2062 location "
            worksheet1.Cells(21, 4) = "Aw for 2065 location  "
            worksheet1.Cells(22, 4) = "Aw Pigment "
            worksheet1.Cells(23, 4) = "Aw  Internal Shade comment "
            worksheet1.Cells(24, 4) = "Aw  customer Shade comment  "
            worksheet1.Cells(25, 4) = "Held in  N/C due to  dyeing issues "
            worksheet1.Cells(26, 4) = "Held in  N/C due to  Finishing issues"
            worksheet1.Cells(27, 4) = "Aw Finishing "
            worksheet1.Cells(28, 4) = "Aw for 2070 location  "
            worksheet1.Cells(29, 4) = "Aw Quality "
            worksheet1.Cells(30, 4) = "Total "

            For i = 11 To 30
                _Char = 68
                worksheet1.Range("e" & i & ":e" & i).Interior.Color = RGB(217, 151, 149)
                worksheet1.Range("f" & i & ":f" & i).Interior.Color = RGB(250, 192, 144)

                For x = 1 To 7

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1

                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                Next
            Next
            '--------------------------------------------------------------------------------------------------------
            worksheet2.Rows(33).Font.Bold = True
            worksheet1.Rows(33).Font.size = 8
            worksheet1.Rows(33).Font.name = "Tahoma"
            worksheet2.Rows(33).rowheight = 20.25
            'worksheet2.Columns("e").ColumnWidth = 15
            'worksheet2.Columns("e").ColumnWidth = 15
            'worksheet2.Columns("f").ColumnWidth = 18
            'worksheet2.Columns("g").ColumnWidth = 18
            'worksheet2.Columns("h").ColumnWidth = 18
            'worksheet2.Columns("i").ColumnWidth = 15



            worksheet1.Cells(33, 1) = "Sales Order"
            worksheet1.Cells(33, 2) = "Line Item"
            worksheet1.Cells(33, 3) = "Material"
            worksheet1.Cells(33, 4) = "Material Description"
            worksheet1.Cells(33, 5) = "Del.Date"
            worksheet1.Cells(33, 6) = "Order Qty"
            worksheet1.Cells(33, 7) = "Delivered.Qty"
            worksheet1.Cells(33, 8) = "Tobe Delivered Qty"
            worksheet1.Cells(33, 9) = "Batch Qty"
            worksheet1.Cells(33, 10) = "Batch No"
            worksheet1.Cells(33, 11) = "Update on " & Today
            worksheet1.Cells(33, 12) = "Department"
            worksheet1.Cells(33, 13) = "Marchant"
            worksheet1.Cells(33, 14) = "Possible Del Date"
            worksheet1.Cells(33, 15) = "Department Comments"
            worksheet1.Cells(33, 16) = "Department2"
            worksheet1.Cells(33, 17) = "PO Number"

            _Char = 65
            i = 33
            For x = 1 To 17

                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(165, 165, 165)
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).MergeCells = True
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(i, x).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                _Char = _Char + 1

            Next
            i = i + 1

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                Else
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='To Be planned' and r.PO_No='" & Trim(cboPO.Text) & "'"
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Location='To Be planned' and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Location='To Be planned' and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Location='To Be planned' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and s.Customer in ('" & _Customer & "')"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Department in ('" & _Department & "')"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned' and r.Merchant in ('" & _Merchant & "')"
                Else
                    SQL = "select r.Tollarance_MIN,r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location='To Be planned'"
                End If
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            If isValidDataset(T01) Then
                worksheet1.Cells(i, 1) = "To Be Planned"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next
                '-------------------------------------------------------------------------------------------------
                'Dim _Last As Integer

                _Last = 0
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    Dim _OrderQty As Double

                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    _OrderQty = 0
                    SQL = "select * from M01Sales_Order_SAP where M01Sales_Order='" & Trim(T01.Tables(0).Rows(Y)("Sales_Order")) & "' and M01Line_Item='" & Trim(T01.Tables(0).Rows(Y)("Line_Item")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        _OrderQty = dsUser.Tables(0).Rows(0)("M01SO_Qty")
                    End If
                    If (_OrderQty * T01.Tables(0).Rows(Y)("Tollarance_MIN")) / 100 >= T01.Tables(0).Rows(Y)("PRD_Qty") Then

                    Else

                        Dim _sTATUS As Boolean
                        _sTATUS = False
                        For x = 34 To i
                            'check if the cell value matches the search string.
                            'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                            If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                                _sTATUS = True
                            End If

                        Next
                        If _sTATUS = False Then
                            _TotalFail = _TotalFail + 1
                        End If
                        worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                        worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                        worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                        worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                        worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                        '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                        ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                        worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                        range1 = worksheet1.Cells(i, 6)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                        '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                        worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                        range1 = worksheet1.Cells(i, 7)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                        worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                        range1 = worksheet1.Cells(i, 8)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                        '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                        range1 = worksheet1.Cells(i, 9)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                        worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                        worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                        worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                        worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                        worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                        worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                        worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        'Dim _sTATUS As Boolean
                        'For x = 34 To i
                        '    'check if the cell value matches the search string.
                        '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                        '        _sTATUS = True
                        '    End If

                        'Next
                        'If _sTATUS = False Then
                        '    _TotalFail = _TotalFail + 1
                        'End If

                        _Last = _Last + 1
                        _Char = 65
                        For x = 1 To 17
                            ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            _Char = _Char + 1
                        Next

                        i = i + 1
                    End If
                    Y = Y + 1

                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E11").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(11, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(11, 6) = _Last
                worksheet1.Cells(11, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '-----------------------------------------------------------------------------------------------------
            'Aw yarn 
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and  r.Department in ('" & _Department & "')and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else

                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and  r.Department in ('" & _Department & "')and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                Else

                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Aw yarn' and  r.Department in ('" & _Department & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn' and r.Merchant in ('" & _Merchant & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else

                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw yarn'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and s.Customer in ('" & _Customer & "') "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and  r.Department in ('" & _Department & "') "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn' and r.Merchant in ('" & _Merchant & "')"
                Else

                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw yarn'"
                End If
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            If isValidDataset(T01) Then
                i = i + 1
                worksheet1.Cells(i, 1) = "Aw yarn "
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next
                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                '   Dim _Last As Integer

                _Last = 0
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'Dim _sTATUS As Boolean
                    'For x = 34 To i
                    '    'check if the cell value matches the search string.
                    '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                    '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                    '        _sTATUS = True
                    '    End If

                    'Next
                    'If _sTATUS = False Then
                    '    _TotalFail = _TotalFail + 1
                    'End If

                    _Last = 0
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E12").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(12, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(12, 6) = _Last
                worksheet1.Cells(12, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '--------------------------------------------------------------------------------------------------------
            'Aw Dyed Yarn

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Aw Dyed Yarn'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Aw Dyed Yarn' and  r.Department in ('" & _Department & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Aw Dyed Yarn'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn' and s.Customer in ('" & _Customer & "') "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn' and  r.Department in ('" & _Department & "') "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn' and r.Merchant in ('" & _Merchant & "')"
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Aw Dyed Yarn'"
                End If
            End If
            ' Dim _Last As Integer
            _Last = 0
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            If isValidDataset(T01) Then
                i = i + 1
                worksheet1.Cells(i, 1) = "Aw Dyed Yarn"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next
                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"
                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If
                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'Dim _sTATUS As Boolean
                    'For x = 34 To i
                    '    'check if the cell value matches the search string.
                    '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                    '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                    '        _sTATUS = True
                    '    End If

                    'Next
                    'If _sTATUS = False Then
                    '    _TotalFail = _TotalFail + 1
                    'End If
                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E13").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(13, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(13, 6) = _Last
                worksheet1.Cells(13, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '------------------------------------------------------------------------------------------
            'Aw Greige (Awaiting Greige)

            If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige' and r.PO_No='" & Trim(cboPO.Text) & "' "

                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'   and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and Awaiting_Grie='Awaiting Greige'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') and  r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and s.Customer in ('" & _Customer & "') "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and r.Department in ('" & _Department & "') "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'   and r.Merchant in ('" & _Merchant & "')"
                Else
                    SQL = "select r.PO_No,R.Prduct_Order, R.Del_Qty,r.Order_Qty,r.Delay_Qty,r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and  s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Awaiting_Grie='Awaiting Greige'"

                End If
            End If
            ' Dim _Last As Integer

            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            If isValidDataset(T01) Then
                i = i + 1
                worksheet1.Cells(i, 1) = "Awaiting Greige"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next
                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    'worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    worksheet1.Cells(i, 11) = "AW Greige"
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'Dim _sTATUS As Boolean
                    'For x = 34 To i
                    '    'check if the cell value matches the search string.
                    '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                    '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                    '        _sTATUS = True
                    '    End If

                    'Next
                    'If _sTATUS = False Then
                    '    _TotalFail = _TotalFail + 1
                    'End If

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E14").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(14, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(14, 6) = _Last
                worksheet1.Cells(14, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '-----------------------------------------------------------------------------------
            'To be Batch Card Send 

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and s.Customer in ('" & _Customer & "')"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "')"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')"
                End If

            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            If isValidDataset(T01) Then
                i = i + 1
                worksheet1.Cells(i, 1) = "To be Batch Card Send"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next
                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i

                ' Dim _Last As Integer
                _Last = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = "To be Batch Card Send "
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'Dim _sTATUS As Boolean
                    'For x = 34 To i
                    '    'check if the cell value matches the search string.
                    '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                    '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                    '        _sTATUS = True
                    '    End If

                    'Next
                    'If _sTATUS = False Then
                    '    _TotalFail = _TotalFail + 1
                    'End If

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E15").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(15, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(15, 6) = _Last
                worksheet1.Cells(15, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '------------------------------------------------------------------------------------------
            'Aw Recipe 01
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where   Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select * from View_1 where   Department in ('" & _Department & "')  and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where   Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select * from View_1 where  PO_No='" & Trim(cboPO.Text) & "' and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "')  and PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where   Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select * from View_1 where   Department in ('" & _Department & "')  and PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where   Merchant in ('" & _Merchant & "') and PO_No='" & Trim(cboPO.Text) & "' "
                Else
                    SQL = "select * from View_1 where  PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and Department in ('" & _Department & "')   and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "')  and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where   Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')  and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select * from View_1 where  Customer in ('" & _Customer & "') and  Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select * from View_1 where   Department in ('" & _Department & "')  and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from View_1 where   Merchant in ('" & _Merchant & "')  and Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                Else
                    SQL = "select * from View_1 where   Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    ' SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                    'SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') and Department in ('" & _Department & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"
                    ' SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') and   s.Customer in ('" & _Customer & "') "
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "')  and Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    '  SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Merchant in ('" & _Merchant & "') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') and   r.Department in ('" & _Department & "') "
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Department in ('" & _Department & "') and Merchant in ('" & _Merchant & "')"
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    '  SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and r.Customer in ('" & _Customer & "')"
                    ' SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') and  s.Customer in ('" & _Customer & "') "
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Customer in ('" & _Customer & "') "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Department in ('" & _Department & "') "
                    ' SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') and   r.Department in ('" & _Department & "')"
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and  Department in ('" & _Department & "') "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')   and  r.Merchant in ('" & _Merchant & "')"
                    'SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Merchant in ('" & _Merchant & "') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues')"
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Merchant in ('" & _Merchant & "')"
                Else
                    ' SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues')"
                    SQL = "select * from View_1 where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
                End If
            End If
            ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Dye','Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues')"

            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            If isValidDataset(T01) Then
                i = i + 1
                worksheet1.Cells(i, 1) = "Aw Recipe"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next
                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i

                _Last = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    ' R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues')

                    If Trim(cboPO.Text) <> "" Or Trim(cboSales_Order.Text) <> "" Then
                        worksheet1.Rows(i).Font.size = 8
                        worksheet1.Rows(i).Font.name = "Tahoma"

                        Dim _sTATUS As Boolean
                        For x = 34 To i
                            'check if the cell value matches the search string.
                            'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                            If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                                _sTATUS = True
                            End If

                        Next
                        If _sTATUS = False Then
                            _TotalFail = _TotalFail + 1
                        End If

                        worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                        worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                        worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                        worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                        worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                        '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                        ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                        worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                        range1 = worksheet1.Cells(i, 6)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                        '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                        worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                        range1 = worksheet1.Cells(i, 7)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                        worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                        range1 = worksheet1.Cells(i, 8)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                        '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                        range1 = worksheet1.Cells(i, 9)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                        worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                        worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 11) = "Aw Recipe" 'T01.Tables(0).Rows(Y)("Location")
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                        worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                        worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                        worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                        worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                        worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        'Dim _sTATUS As Boolean
                        'For x = 34 To i
                        '    'check if the cell value matches the search string.
                        '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                        '        _sTATUS = True
                        '    End If

                        'Next
                        'If _sTATUS = False Then
                        '    _TotalFail = _TotalFail + 1
                        'End If

                        _Last = _Last + 1
                        _Char = 65
                        For x = 1 To 17
                            ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            _Char = _Char + 1
                        Next
                        i = i + 1
                    Else
                        If T01.Tables(0).Rows(Y)("Del_Date") >= txtFromDate.Text Then
                            worksheet1.Rows(i).Font.size = 8
                            worksheet1.Rows(i).Font.name = "Tahoma"

                            Dim _sTATUS As Boolean
                            _sTATUS = False
                            For x = 34 To i
                                'check if the cell value matches the search string.
                                'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                                If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                                    _sTATUS = True
                                End If

                            Next
                            If _sTATUS = False Then
                                _TotalFail = _TotalFail + 1
                            End If

                            worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                            worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                            worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                            worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                            worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                            '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                            ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                            worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                            range1 = worksheet1.Cells(i, 6)
                            range1.NumberFormat = "0.00"
                            worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                            '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                            worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                            range1 = worksheet1.Cells(i, 7)
                            range1.NumberFormat = "0.00"
                            worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                            ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                            worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                            range1 = worksheet1.Cells(i, 8)
                            range1.NumberFormat = "0.00"
                            worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                            '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                            worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                            range1 = worksheet1.Cells(i, 9)
                            range1.NumberFormat = "0.00"
                            worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                            worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                            worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                            worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                            worksheet1.Cells(i, 11) = "Aw Recipe" 'T01.Tables(0).Rows(Y)("Location")
                            worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                            worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                            worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                            worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                            worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                            worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                            worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                            worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                            worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                            worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                            worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                            'Dim _sTATUS As Boolean
                            'For x = 34 To i
                            '    'check if the cell value matches the search string.
                            '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                            '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            '        _sTATUS = True
                            '    End If

                            'Next
                            'If _sTATUS = False Then
                            '    _TotalFail = _TotalFail + 1
                            'End If

                            _Last = _Last + 1
                            _Char = 65
                            For x = 1 To 17
                                ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                                _Char = _Char + 1
                            Next
                            i = i + 1
                        End If

                    End If
                    Y = Y + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E16").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(16, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(16, 6) = _Last
                worksheet1.Cells(16, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If

            ''Aw Recipe 02
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye','Finishing') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.NC_Comment in ('4.5 DRY AND HOLD','3.0 Shade Issues','3.3 Off Shade','3.4 Need to Over Dye')"
            'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            'If isValidDataset(T01) Then
            '    i = i + 1
            '    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Dye' and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')"
            '    T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            '    If isValidDataset(T03) Then

            '    Else
            '        worksheet1.Cells(i, 1) = "Aw Recipe"
            '        worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
            '        worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '        _Char = 65
            '        For x = 1 To 17
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            '            _Char = _Char + 1
            '        Next
            '    End If
            '    '-------------------------------------------------------------------------------------------------
            '    Y = 0
            '    i = i + 1
            '    _FirstChr = i
            '    For Each DTRow3 As DataRow In T01.Tables(0).Rows
            '        worksheet1.Rows(i).Font.size = 8
            '        worksheet1.Rows(i).Font.name = "Tahoma"

            '        worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
            '        worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '        worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
            '        worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '        worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
            '        worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
            '        worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
            '        '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
            '        ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
            '        worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '        worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

            '        range1 = worksheet1.Cells(i, 6)
            '        range1.NumberFormat = "0.00"
            '        worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
            '        '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

            '        worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
            '        range1 = worksheet1.Cells(i, 7)
            '        range1.NumberFormat = "0.00"
            '        worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
            '        ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

            '        worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
            '        range1 = worksheet1.Cells(i, 8)
            '        range1.NumberFormat = "0.00"
            '        worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
            '        '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

            '        worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
            '        range1 = worksheet1.Cells(i, 9)
            '        range1.NumberFormat = "0.00"
            '        worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            '        worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

            '        worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
            '        worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
            '        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '        worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
            '        worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '        worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
            '        worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '        worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
            '        worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

            '        worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
            '        worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
            '        worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
            '        worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
            '        _Char = 65
            '        For x = 1 To 17
            '            ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            '            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            '            _Char = _Char + 1
            '        Next
            '        Y = Y + 1
            '        i = i + 1
            '    Next
            '    worksheet1.Rows(i).Font.size = 8
            '    worksheet1.Rows(i).Font.name = "Tahoma"
            '    worksheet1.Rows(i).Font.bold = True
            '    worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
            '    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            '    range1 = worksheet1.Cells(i, 9)
            '    range1.NumberFormat = "0.00"
            '    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)
            '    _Char = 65
            '    For x = 1 To 17
            '        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

            '        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            '        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            '        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            '        _Char = _Char + 1
            '    Next

            'End If
            '-----------------------------------------------------------------------------------------------
            ' Aw prepare
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date  where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date  where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment='' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment='' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment='' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date  where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('AW Preparation') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.NC_Comment='' "
                End If

            End If
            '  SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw prepare') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw prepare"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    'worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    'worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    tblDye = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(tblDye) Then
                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " on " & Month(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(tblDye.Tables(0).Rows(0)("Dye_Pln_Date"))
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " not plan yet"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'Dim _sTATUS As Boolean
                    'For x = 34 To i
                    '    'check if the cell value matches the search string.
                    '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                    '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                    '        _sTATUS = True
                    '    End If

                    'Next
                    'If _sTATUS = False Then
                    '    _TotalFail = _TotalFail + 1
                    'End If

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E17").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(17, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(17, 6) = _Last
                worksheet1.Cells(17, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '-----------------------------------------------------------------------------------------------------------
            'Aw Presetting
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw prepare') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw prepare') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment='' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment='' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw prepare') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment='' "
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw prepare') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')"
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment='' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.NC_Comment=''"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment='' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment='' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.NC_Comment=''"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.NC_Comment=''"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment=''"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.NC_Comment=''"
                End If

            End If
            ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw Presetting"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    'worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    'worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    tblDye = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(tblDye) Then
                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " on " & Month(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(tblDye.Tables(0).Rows(0)("Dye_Pln_Date"))
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " not plan yet"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'Dim _sTATUS As Boolean
                    'For x = 34 To i
                    '    'check if the cell value matches the search string.
                    '    'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                    '    If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                    '        _sTATUS = True
                    '    End If

                    'Next
                    'If _sTATUS = False Then
                    '    _TotalFail = _TotalFail + 1
                    'End If

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E18").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(18, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(18, 6) = _Last
                worksheet1.Cells(18, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '---------------------------------------------------------------------------------------------
            'Aw Dyeing 

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                Else
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                Else
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                Else
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE') "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"

                Else
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.NC_Comment in ('22.RESCHEDULE - RE DYE','23.Reschedule – wash','24.RESCHEDULE - STRIPPED','25.RESCHEDULE - OVER DYE','26.RESCHEDULE - BOIL OFF','27.RESCHEDULE - SAMPLE')"
                End If
            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw Dyeing"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Prduct_Order='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "' and r.Location='Dye' and f.Stock_Code='' and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
                    T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    ' _Fail_Batch = 20
                    '_Fail_Batch = _Fail_Batch + 1
                    If isValidDataset(T03) Then
                    Else
                        _Fail_Batch = _Fail_Batch + 1

                        Dim _sTATUS As Boolean
                        _sTATUS = False
                        For x = 34 To i
                            'check if the cell value matches the search string.
                            'MsgBox(worksheet1.Cells(x, 2).text)
                            'MsgBox(T01.Tables(0).Rows(Y)("Line_Item"))
                            If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                                _sTATUS = True
                            End If

                        Next
                        If _sTATUS = False Then
                            _TotalFail = _TotalFail + 1
                        End If

                        worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                        worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                        worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                        worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                        worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                        '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                        ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                        worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                        range1 = worksheet1.Cells(i, 6)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                        '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                        worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                        range1 = worksheet1.Cells(i, 7)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                        worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                        range1 = worksheet1.Cells(i, 8)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                        '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                        range1 = worksheet1.Cells(i, 9)
                        range1.NumberFormat = "0.00"
                        worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                        worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                        worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        SQL = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                        tblDye = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        If isValidDataset(tblDye) Then
                            worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " on " & Month(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(tblDye.Tables(0).Rows(0)("Dye_Pln_Date"))
                            worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " not plane yet"
                            worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        End If
                        worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                        worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                        worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                        worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                        worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                        worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft


                      
                        _Last = _Last + 1
                        _Char = 65
                        For x = 1 To 17
                            ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            _Char = _Char + 1
                        Next
                        i = i + 1
                    End If
                    Y = Y + 1

                Next

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Aw Presetting') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing') and   s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')" ' and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')" ' and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                     SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')"' and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')   and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and s.Customer in ('" & _Customer & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and r.Department in ('" & _Department & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE') "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')   and r.Merchant in ('" & _Merchant & "') and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Finishing')  and r.NC_Comment in ('5.2 NEED TO WASH','5.1 NEED TO REDYE','5.3 NEED TO STRIPE','5.4 NEED TO OVER DYE')"
                End If

                'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('Dye') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)

                '_FirstChr = i
                Y = 0
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"
                    'worksheet1.Rows(i).Font.bold = True

                    _Fail_Batch = _Fail_Batch + 1


                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from FR_Update where Batch_No='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    tblDye = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(tblDye) Then
                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " on " & Month(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Microsoft.VisualBasic.Day(tblDye.Tables(0).Rows(0)("Dye_Pln_Date")) & "/" & Year(tblDye.Tables(0).Rows(0)("Dye_Pln_Date"))
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location") & " not plane yet"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If
                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                   

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    i = i + 1

                    Y = Y + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True

                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E19").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(19, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(19, 6) = _Last
                worksheet1.Cells(19, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '----------------------------------------------------------------------------------------
            '2062
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                End If


            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Department in ('" & _Department & "')"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
                End If

            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw for 2062 location  "
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If
                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from ZPP_DEL where Product_Order='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        worksheet1.Cells(i, 11) = Trim(dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")) & " days in 2062"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If
                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

              

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E20").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(20, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(20, 6) = _Last
                worksheet1.Cells(20, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '------------------------------------------------------------------------------------------------
            '2065
            If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2062') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "')"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "')  "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
                End If
            End If
            ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw for 2065 location  "
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from ZPP_DEL where Product_Order='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        worksheet1.Cells(i, 11) = Trim(dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")) & " days in 2065"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                  

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E21").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(21, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(21, 6) = _Last
                worksheet1.Cells(21, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '---------------------------------------------------------------------------
            'Pigmant
            If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'  "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%Pigment'  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "'   "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "'  "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'  "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing'  and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment like '%AW PAD%'  and  Location='Finishing'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and s.Customer in ('" & _Customer & "')   "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing'  and r.Department in ('" & _Department & "')  "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "')  "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%AW PAD%'  and  Location='Finishing' "
                End If

            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment like '%Pigment'  and  Location='Finishing' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw Pigment  "
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                   

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E22").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(22, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(22, 6) = _Last
                worksheet1.Cells(22, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '-----------------------------------------------------------------------------
            'AW Internal Shade Comment
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2,Recipy_Status  from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date  left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                Else
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'" '  and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                Else
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                Else
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales'  and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'" ' and Recipy_Status<>'Awaiting Recipe'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')" ' and Recipy_Status<>'Awaiting Recipe' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')" ' and Recipy_Status<>'Awaiting Recipe' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')" ' and Recipy_Status<>'Awaiting Recipe' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')" ' and Recipy_Status<>'Awaiting Recipe'"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Department in ('" & _Department & "')" ' and Recipy_Status<>'Awaiting Recipe'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')" ' and Recipy_Status<>'Awaiting Recipe' "
                Else
                    SQL = "select r.NC_Comment,Recipy_Status ,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date left join FR_Update on Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW SHADE COMMENTS')  and  Location='Finishing'" 'and Recipy_Status<>'Awaiting Recipe'"
                End If
            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('00 PILOT','2.1 AW 1ST BULK APP','2.6 AW CUS CARE COMMENT')  and  Location='Finishing' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "AW  Shade Comment"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("NC_Comment")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                   

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E23").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(23, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(23, 6) = _Last
                worksheet1.Cells(23, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '-----------------------------------------------------------------------------------
            'Aw  customer Shade comment  

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and  r.Department in ('" & _Department & "')"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing' and r.Merchant in ('" & _Merchant & "')  "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('AW 1st BULK APP','AW CUS APP','AW CUS CARE COMMENT','AW N/C APP')  and  Location='Finishing'"
                End If
            End If
            ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('2.2 AW N/C APP','2.3 AW CUS APP','2.7 SUBMIT AS ONGOING')  and  Location='Finishing' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw  customer Shade comment"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("NC_Comment")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                   


                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E24").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(24, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(24, 6) = _Last
                worksheet1.Cells(24, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '-------------------------------------------------------------------------
            'Held in  N/C due to  dyeing issues

            If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If

            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'"
                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Department in ('" & _Department & "') "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')  "
                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('DOWN GRADE','DYEING ISSUES','HOLD','HELD IN NEW ORDERS')  and  Location='Finishing' "
                End If
            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.0 HOLD','3.5 OTHERS','3.0 SHADE ISSUES','4.0 OTHER REASON','6.0 HELD IN NEW ORDERS','4.1 DYEING ISSUES')  and  Location='Finishing' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Held in  N/C due to  dyeing issues"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("NC_Comment")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                  

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E25").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(25, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(25, 6) = _Last
                worksheet1.Cells(25, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '--------------------------------------------------------------------------------------------------
            'Held in  N/C due to  Finishing issues
            If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    '  SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    '  SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    '  SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and  r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    '  SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "

                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and s.Customer in ('" & _Customer & "')  "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Department in ('" & _Department & "')  "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' and  r.Merchant in ('" & _Merchant & "') "

                Else
                    SQL = "select  r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' "
                End If
            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES')  and  Location='Finishing' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Held in  N/C due to  Finishing issues"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("NC_Comment")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    

                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E26").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(26, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(26, 6) = _Last
                worksheet1.Cells(26, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter


                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '----------------------------------------------------------------------------------
            'Aw Finishing 
            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'    and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'    and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT,'' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'    and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'    and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing' and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "

                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing' and r.PO_No='" & Trim(cboPO.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'    and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')   and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'    and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing' and r.Department in ('" & _Department & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "

                Else
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and  Location='Finishing' and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')  "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and s.Customer in ('" & _Customer & "') and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and r.Department in ('" & _Department & "') and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing'  and r.Merchant in ('" & _Merchant & "')  and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','') "

                Else
                    SQL = "select r.NC_Comment,r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' and NC_Comment in ('00 PILOT','1.0 AW PILOT','1.1 1ST BULK PILOT','1.2 ON GOING PILOT','')"
                End If
            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'   and  Location='Finishing' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            ' _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count

            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw Finishing"
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    If Trim(T01.Tables(0).Rows(Y)("Sales_Order")) = "1033698" Then
                        ' MsgBox("")
                    End If
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,f.Stock_Code,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date inner join FR_Update F on f.Batch_No=r.Prduct_Order where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location IN ('Finishing') and f.Recipy_Status in ('Awaiting Recipe','First Bulk Awaiting Recipe')  and r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') OR R.NC_Comment IN ('4.5 DRY AND HOLD','3.3 OFF SHADE','3.4 NEED TO OVER DYE','3.0 Shade Issues') AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(Y)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(Y)("Line_Item")) & "'"
                    'T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    'If isValidDataset(T03) Then

                    'Else
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('00 PILOT','2.1 AW 1ST BULK APP','2.6 AW CUS CARE COMMENT')  and  Location='Finishing' AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(Y)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(Y)("Line_Item")) & "'"
                    'T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    'If isValidDataset(T03) Then
                    'Else
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.0 HOLD','3.5 OTHERS','3.0 SHADE ISSUES','4.0 OTHER REASON','6.0 HELD IN NEW ORDERS','4.1 DYEING ISSUES')  and  Location='Finishing' AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(Y)("Sales_Order")) & "' "
                    'T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    'If isValidDataset(T03) Then
                    'Else
                    'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.NC_Comment in ('4.2 FINISHING ISSUES','5.2 NEED TO WASH','4.5 DRY AND HOLD')  and  Location='Finishing' AND R.Sales_Order='" & Trim(T01.Tables(0).Rows(Y)("Sales_Order")) & "' and r.Line_Item='" & Trim(T01.Tables(0).Rows(Y)("Line_Item")) & "'"
                    'T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    'If isValidDataset(T03) Then
                    'Else
                    _Fail_Batch = _Fail_Batch + 1
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from ZPP_DEL where Product_Order='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        worksheet1.Cells(i, 11) = Trim(dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")) & " days in Finishing"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        If Trim(dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")) > 2 Then
                            worksheet1.Range("k" & i, "k" & i).Font.Color = RGB(255, 0, 0)
                            worksheet1.Range("J" & i, "J" & i).Font.Color = RGB(255, 0, 0)
                        End If
                    End If
                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft



                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    i = i + 1
                    'End If
                    'End If
                    'End If
                    'End If
                    Y = Y + 1

                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E27").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(27, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(27, 6) = _Last
                worksheet1.Cells(27, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '---------------------------------------------------------------------------------------
            '2070
            If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')"
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn')  and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and s.Customer in ('" & _Customer & "')"

                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Department in ('" & _Department & "') "

                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and  r.Merchant in ('" & _Merchant & "')  "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
                End If
            End If
            'SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2070') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw for 2070 location  "
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If


                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 11) = T01.Tables(0).Rows(Y)("Location")
                    worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft


                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E28").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(28, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(28, 6) = _Last
                worksheet1.Cells(28, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter


                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If
            '------------------------------------------------------------------------
            'Aw Quality 

            If Trim(cboPO.Text) <> "" And Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam'  and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and r.PO_No='" & Trim(cboPO.Text) & "' and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If

            ElseIf Trim(cboPO.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "'  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam'  and r.Department in ('" & _Department & "') and r.PO_No='" & Trim(cboPO.Text) & "' "

                ElseIf Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "') and r.PO_No='" & Trim(cboPO.Text) & "' "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and r.PO_No='" & Trim(cboPO.Text) & "' "
                End If
            ElseIf Trim(cboSales_Order.Text) <> "" Then
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and s.Customer in ('" & _Customer & "') and  r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam'  and r.Department in ('" & _Department & "')  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"

                ElseIf Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "') and r.Sales_Order='" & Trim(cboSales_Order.Text) & "' "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales'  and r.Location='Exam'  and r.Sales_Order='" & Trim(cboSales_Order.Text) & "'"
                End If
            Else

                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location in ('2065') and  r.Awaiting_Grie not in ('Awaiting Greige','Aw Dyed Yarn','Aw yarn') and r.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') and r.Department in ('" & _Department & "') "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "')  and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and r.Department in ('" & _Department & "') and r.Merchant in ('" & _Merchant & "')  "

                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and s.Customer in ('" & _Customer & "') "

                ElseIf Trim(txtDepartment.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam'  and r.Department in ('" & _Department & "')"

                ElseIf Trim(txtMerchant.Text) <> "" Then

                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' and  r.Merchant in ('" & _Merchant & "')  "
                Else
                    SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item and s.Del_Date=r.Del_Date where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' "
                End If
            End If
            ' SQL = "select r.PO_No,r.Delay_Qty,r.Del_Qty,r.Order_Qty,r.Prduct_Order, r.Sales_Order,r.Line_Item,r.Metrrial,r.Met_Des,s.Customer,s.Department,s.Merchant,r.Del_Date,r.Location,r.PRD_Qty,r.Department_Comment,r.Comment2 from OTD_Records R inner join OTD_SMS S on r.Sales_Order=s.Sales_Order and s.Line_Item=r.Line_Item  where s.Status='Fales' and r.Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and r.Location='Exam' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Fail_Batch = _Fail_Batch + T01.Tables(0).Rows.Count
            _Last = 0
            If isValidDataset(T01) Then
                i = i + 1

                worksheet1.Cells(i, 1) = "Aw Quality "
                worksheet1.Range(worksheet1.Cells(i, 1), worksheet1.Cells(i, 3)).Merge()
                worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

                '-------------------------------------------------------------------------------------------------
                Y = 0
                i = i + 1
                _FirstChr = i
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(i).Font.size = 8
                    worksheet1.Rows(i).Font.name = "Tahoma"

                    Dim _sTATUS As Boolean
                    _sTATUS = False
                    For x = 34 To i
                        'check if the cell value matches the search string.
                        'MsgBox(worksheet1.Cells(_XLRow, 1).text)
                        If worksheet1.Cells(x, 1).value = T01.Tables(0).Rows(Y)("Sales_Order").ToString And worksheet1.Cells(x, 2).value = T01.Tables(0).Rows(Y)("Line_Item").ToString Then
                            _sTATUS = True
                        End If

                    Next
                    If _sTATUS = False Then
                        _TotalFail = _TotalFail + 1
                    End If

                    worksheet1.Cells(i, 1) = T01.Tables(0).Rows(Y)("Sales_Order")
                    worksheet1.Cells(i, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 2) = T01.Tables(0).Rows(Y)("Line_Item")
                    worksheet1.Cells(i, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 3) = T01.Tables(0).Rows(Y)("Metrrial")
                    worksheet1.Cells(i, 4) = T01.Tables(0).Rows(Y)("Met_Des")
                    worksheet1.Cells(i, 5) = T01.Tables(0).Rows(Y)("Del_Date")
                    '  worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(194, 214, 154)
                    ' worksheet1.Cells(i, 5).EntireColumn.NumberFormat = "dd-mm-yyyy"
                    worksheet1.Cells(i, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(i, 6) = T01.Tables(0).Rows(Y)("Order_Qty")

                    range1 = worksheet1.Cells(i, 6)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '   worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(252, 213, 180)

                    worksheet1.Cells(i, 7) = T01.Tables(0).Rows(Y)("Del_Qty")
                    range1 = worksheet1.Cells(i, 7)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    ' worksheet1.Range("E" & i, "E" & i).Interior.Color = RGB(204, 192, 218)

                    worksheet1.Cells(i, 8) = T01.Tables(0).Rows(Y)("Delay_Qty")
                    range1 = worksheet1.Cells(i, 8)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    '  worksheet1.Range("F" & i, "F" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 9) = T01.Tables(0).Rows(Y)("PRD_Qty")
                    range1 = worksheet1.Cells(i, 9)
                    range1.NumberFormat = "0.00"
                    worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Cells(i, 10) = T01.Tables(0).Rows(Y)("Prduct_Order")
                    worksheet1.Cells(i, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    SQL = "select * from ZPP_DEL where Product_Order='" & Trim(T01.Tables(0).Rows(Y)("Prduct_Order")) & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        worksheet1.Cells(i, 11) = Trim(dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")) & " days in Quality"
                        worksheet1.Cells(i, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        If Trim(dsUser.Tables(0).Rows(0)("No_Day_Same_Opp")) > 2 Then
                            worksheet1.Range("K" & i, "K" & i).Font.Color = RGB(255, 0, 0)
                            worksheet1.Range("J" & i, "J" & i).Font.Color = RGB(255, 0, 0)
                        End If

                    End If

                    worksheet1.Cells(i, 12) = T01.Tables(0).Rows(Y)("Department")
                    worksheet1.Cells(i, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 13) = T01.Tables(0).Rows(Y)("Merchant")
                    worksheet1.Cells(i, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(i, 15) = T01.Tables(0).Rows(Y)("Department_Comment")
                    worksheet1.Cells(i, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(i, 16) = T01.Tables(0).Rows(Y)("Comment2")
                    worksheet1.Cells(i, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(i, 17) = T01.Tables(0).Rows(Y)("PO_No")
                    worksheet1.Cells(i, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft



                    _Last = _Last + 1
                    _Char = 65
                    For x = 1 To 17
                        ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        _Char = _Char + 1
                    Next
                    Y = Y + 1
                    i = i + 1
                Next
                worksheet1.Rows(i).Font.size = 8
                worksheet1.Rows(i).Font.name = "Tahoma"
                worksheet1.Rows(i).Font.bold = True
                worksheet1.Range("i" & (i)).Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(i, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(i, 9)
                range1.NumberFormat = "0.00"
                worksheet1.Range("I" & i, "I" & i).Interior.Color = RGB(141, 180, 227)

                worksheet1.Range("E29").Formula = "=SUM(i" & _FirstChr & ":i" & (i - 1) & ")"
                worksheet1.Cells(29, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                'range1 = worksheet1.Cells(11, 5)
                'range1.NumberFormat = "0.00"
                worksheet1.Cells(29, 6) = _Last
                worksheet1.Cells(29, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _Char = 65
                For x = 1 To 17
                    ' worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Interior.Color = RGB(141, 180, 227)

                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    _Char = _Char + 1
                Next

            End If

            worksheet1.Cells(4, 5) = _TotalFail

            worksheet1.Range("E30").Formula = "=SUM(E11:E29)"
            worksheet1.Cells(30, 5).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet1.Range("F30").Formula = "=SUM(F11:F29)"
            worksheet1.Cells(30, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '            <<<===============================================================================>>>
            '     <<<=============================================================================================>>>
            '<<<======================================================================================================>>>
            '            <<<===============================================================================>>>
            '     <<<=============================================================================================>>>
            'SHEET 02
            If Trim(cboSales_Order.Text) <> "" Or Trim(cboPO.Text) <> "" Then

            Else
                workbooks.Application.Sheets.Add()
                Dim worksheet117 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
                worksheet117.Name = "Possible Fail item"
                worksheet117.Columns("A").ColumnWidth = 10
                worksheet117.Columns("B").ColumnWidth = 10
                worksheet117.Columns("C").ColumnWidth = 10
                worksheet117.Columns("D").ColumnWidth = 10
                worksheet117.Columns("E").ColumnWidth = 28
                worksheet117.Columns("F").ColumnWidth = 16
                worksheet117.Columns("G").ColumnWidth = 10
                worksheet117.Columns("H").ColumnWidth = 19

                worksheet117.Columns("J").ColumnWidth = 10
                worksheet117.Columns("K").ColumnWidth = 10
                worksheet117.Columns("L").ColumnWidth = 10
                worksheet117.Columns("M").ColumnWidth = 10
                worksheet117.Columns("N").ColumnWidth = 22
                worksheet117.Columns("O").ColumnWidth = 10
                worksheet117.Columns("P").ColumnWidth = 10
                'worksheet117.Columns("H").ColumnWidth = 18

                worksheet117.Cells(1, 1) = "Possible Failure Items"
                worksheet117.Rows(1).Font.size = 10
                worksheet117.Rows(1).Font.BOLD = True

                worksheet117.Range("A1:H1").VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet117.Range(worksheet117.Cells(1, 1), worksheet117.Cells(1, 8)).Merge()
                worksheet117.Range(worksheet117.Cells(1, 1), worksheet117.Cells(1, 8)).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet117.Cells(1, 10) = "Possible allocation"

                worksheet117.Range("J1:P1").VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet117.Range(worksheet117.Cells(1, 10), worksheet117.Cells(1, 16)).Merge()
                worksheet117.Range(worksheet117.Cells(1, 10), worksheet117.Cells(1, 16)).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet117.Rows(2).Font.size = 10
                worksheet117.Rows(2).Font.BOLD = True

                worksheet117.Cells(2, 1) = "30Class"
                worksheet117.Cells(2, 2) = "Sales Order"
                worksheet117.Cells(2, 3) = "Line Item"
                worksheet117.Cells(2, 4) = "Del Date"
                worksheet117.Cells(2, 5) = "Description"
                worksheet117.Cells(2, 6) = "Batch No"
                worksheet117.Cells(2, 7) = "Qty"
                worksheet117.Cells(2, 8) = "Reason"


                worksheet117.Cells(2, 10) = "Sales Order"
                worksheet117.Cells(2, 11) = "Line Item"
                worksheet117.Cells(2, 12) = "Batch No"
                worksheet117.Cells(2, 13) = "Qty"
                worksheet117.Cells(2, 14) = "Current Location"
                worksheet117.Cells(2, 15) = "Del Date"
                worksheet117.Cells(2, 16) = "NC"

                _Char = 65
                For i = 1 To 8
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Interior.Color = RGB(141, 180, 227)
                    ' MsgBox(ChrW(_Char) & "2:" & ChrW(_Char) & "2")
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").MergeCells = True
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter

                    _Char = _Char + 1
                Next
                _Char = _Char + 1
                For i = 1 To 7
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").Interior.Color = RGB(141, 180, 227)
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").MergeCells = True
                    worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
                    _Char = _Char + 1
                Next
                worksheet117.Rows(2).rowheight = 24.25
                worksheet117.Rows(2).Font.Bold = True
                worksheet117.Rows(2).Font.size = 8
                worksheet117.Rows(2).Font.name = "Tahoma"

                i = 0
                _cOUNT = 3

                If Trim(cboSales_Order.Text) <> "" And Trim(cboPO.Text) <> "" Then


                ElseIf Trim(cboPO.Text) <> "" Then

                ElseIf Trim(cboSales_Order.Text) <> "" Then

                Else

                    If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "') and Department in ('" & txtDepartment.Text & "') and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                    ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "') and Department in ('" & txtDepartment.Text & "')  order by Del_Date "
                    ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "')  and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                    ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Department in ('" & txtDepartment.Text & "') and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                    ElseIf Trim(txtCustomer.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "')  order by Del_Date "
                    ElseIf Trim(txtDepartment.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Department in ('" & txtDepartment.Text & "')  order by Del_Date "
                    ElseIf Trim(txtMerchant.Text) <> "" Then
                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                    Else

                        SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and Location in ('Finishing','Dye') and NC_Comment<>'' order by Del_Date "
                    End If

                End If
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet117.Rows(_cOUNT).Font.size = 8
                    worksheet117.Rows(_cOUNT).Font.name = "Tahoma"

                    worksheet117.Cells(_cOUNT, 1) = Trim(T01.Tables(0).Rows(i)("Metrrial"))
                    worksheet117.Cells(_cOUNT, 2) = Trim(T01.Tables(0).Rows(i)("Sales_Order"))
                    worksheet117.Cells(_cOUNT, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Cells(_cOUNT, 3) = Trim(T01.Tables(0).Rows(i)("Line_Item"))
                    worksheet117.Cells(_cOUNT, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Cells(_cOUNT, 4) = Trim(T01.Tables(0).Rows(i)("Del_Date"))
                    worksheet117.Cells(_cOUNT, 5) = Trim(T01.Tables(0).Rows(i)("Met_Des"))
                    worksheet117.Cells(_cOUNT, 6) = Trim(T01.Tables(0).Rows(i)("Prduct_Order"))
                    worksheet117.Cells(_cOUNT, 7) = Trim(T01.Tables(0).Rows(i)("PRD_Qty"))
                    worksheet117.Cells(_cOUNT, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Cells(_cOUNT, 8) = Trim(T01.Tables(0).Rows(i)("NC_Comment"))

                    range1 = worksheet1.Cells(_cOUNT, 7)
                    range1.NumberFormat = "0.00"

                    _Char = 65
                    For Y = 1 To 8
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                        worksheet117.Range("G" & _cOUNT & ":" & "G" & _cOUNT).Interior.Color = RGB(141, 180, 227)
                        'worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").MergeCells = True
                        'worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
                        _Char = _Char + 1
                    Next

                    _cOUNT = _cOUNT + 1
                    i = i + 1
                Next

                i = 0
                _cOUNT = 3

                'SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' order by Del_Date "
                If Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "') and Department in ('" & txtDepartment.Text & "') and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtDepartment.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "') and Department in ('" & txtDepartment.Text & "')  order by Del_Date "
                ElseIf Trim(txtCustomer.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "')  and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                ElseIf Trim(txtDepartment.Text) <> "" And Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Department in ('" & txtDepartment.Text & "') and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                ElseIf Trim(txtCustomer.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Customer in ('" & txtCustomer.Text & "')  order by Del_Date "
                ElseIf Trim(txtDepartment.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Department in ('" & txtDepartment.Text & "')  order by Del_Date "
                ElseIf Trim(txtMerchant.Text) <> "" Then
                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' and Merchant in ('" & txtMerchant.Text & "') order by Del_Date "
                Else

                    SQL = "select * from OTD_Records where Del_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and NC_Comment<>'' order by Del_Date "
                End If

                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                For Each DTRow3 As DataRow In T01.Tables(0).Rows
                    worksheet117.Rows(_cOUNT).Font.size = 8
                    worksheet117.Rows(_cOUNT).Font.name = "Tahoma"
                    Dim _PRD_Qty As Double
                    Dim _BatchNo As String

                    _PRD_Qty = 0

                    Y = 0
                    SQL = "select * from OTD_Records where  Location in ('Finishing','Exam') and Sales_Order<>'" & T01.Tables(0).Rows(i)("Sales_Order") & "' and Line_Item<>'" & T01.Tables(0).Rows(i)("Line_Item") & "' and Del_Date>'" & T01.Tables(0).Rows(i)("Del_Date") & "' and NC_Comment='' and Metrrial='" & Trim(T01.Tables(0).Rows(i)("Metrrial")) & "' "
                    T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(T03) Then
                        If T03.Tables(0).Rows(0)("PRD_Qty") >= T01.Tables(0).Rows(i)("PRD_Qty") Then
                            worksheet117.Cells(_cOUNT, 10) = Trim(T03.Tables(0).Rows(0)("Sales_Order"))
                            worksheet117.Cells(_cOUNT, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet117.Cells(_cOUNT, 11) = Trim(T03.Tables(0).Rows(0)("Line_Item"))
                            worksheet117.Cells(_cOUNT, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet117.Cells(_cOUNT, 15) = Trim(T03.Tables(0).Rows(0)("Del_Date"))
                            worksheet117.Cells(_cOUNT, 14) = Trim(T03.Tables(0).Rows(0)("Location"))
                            worksheet117.Cells(_cOUNT, 12) = Trim(T03.Tables(0).Rows(0)("Prduct_Order"))
                            worksheet117.Cells(_cOUNT, 13) = Trim(T03.Tables(0).Rows(0)("PRD_Qty"))
                            worksheet117.Cells(_cOUNT, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            ' worksheet117.Cells(_cOUNT, 8) = Trim(T01.Tables(0).Rows(i)("NC_Comment


                            range1 = worksheet1.Cells(_cOUNT, 13)
                            range1.NumberFormat = "0.00"
                        Else
                            worksheet117.Cells(_cOUNT, 10) = "No"
                        End If
                    Else
                        worksheet117.Cells(_cOUNT, 10) = "No"
                    End If

                    _Char = 74
                    For Y = 1 To 7
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet117.Range(ChrW(_Char) & _cOUNT & ":" & ChrW(_Char) & _cOUNT).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                        worksheet117.Range("M" & _cOUNT & ":" & "M" & _cOUNT).Interior.Color = RGB(141, 180, 227)
                        'worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").MergeCells = True
                        'worksheet117.Range(ChrW(_Char) & "2:" & ChrW(_Char) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
                        _Char = _Char + 1
                    Next

                    _cOUNT = _cOUNT + 1
                    i = i + 1
                Next

            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()
    End Sub

    Private Sub cboPO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPO.AfterCloseUp
        Call Search_SalesOrder()
    End Sub

    Private Sub cboPO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboPO.InitializeLayout

    End Sub

    Private Sub cboPO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPO.TextChanged
        cboSales_Order.Text = ""
    End Sub

    Function Search_SalesOrder()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M01Sales_Order as [Sales Order] from M01Sales_Order_SAP where M01PO='" & Trim(cboPO.Text) & "' group by M01Sales_Order"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboSales_Order
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
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
End Class