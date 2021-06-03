
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
Public Class frmNo_Orders

    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String
    Dim _Status As String

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
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
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
        Dim _Conqty As Double
        Dim _30Class As String

        Dim QualityNo As String

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
        Dim _Merchant1 As String
        Dim _Location1 As String
        Dim _Confact As Double

        Try


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\delsum.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 470 Then
                    ' MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then
                _PO_No = Trim(fields(3))
                _sales_Order = CInt(Trim(fields(0)))
                _LineItem = CInt(Trim(fields(5)))

                _Material = Trim(fields(7))
                _Material_Dis = Trim(fields(8))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                characterToRemove = """"
                _30Class = _Material

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                characterToRemove = "-"
                _30Class = (Replace(_30Class, characterToRemove, ""))

                _Customer = CInt(Trim(fields(1)))
                _CustomerName = Trim(fields(2))

                If Microsoft.VisualBasic.Right(_Material_Dis, 3) = "OCI" Then
                    _Location1 = "OCI"
                ElseIf Microsoft.VisualBasic.Right(_Material_Dis, 3) = "PTL" Then
                    _Location1 = "PTL"
                Else

                    _Location1 = "IN HOUSE"
                End If
                _PO_No = (Replace(_PO_No, characterToRemove, ""))

                '   _CustomerName = Microsoft.VisualBasic.Left(_Customer, 2)
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _CustomerName = (Replace(_CustomerName, characterToRemove, ""))
                _Department = Trim(fields(6))
                If Microsoft.VisualBasic.Left(_Department, 3) = "M&S" Then
                    _Department = Microsoft.VisualBasic.Left(_Department, 3)
                End If
                ' _Merchnat = Trim(fields(8))

                Dim TestString As String = _Material_Dis
                Dim TestArray() As String = Split(TestString)

                ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""} 
                Dim LastNonEmpty As Integer = -1
                For z As Integer = 0 To TestArray.Length - 1
                    If TestArray(z) <> "" Then
                        LastNonEmpty += 1
                        TestArray(LastNonEmpty) = TestArray(z)
                        QualityNo = TestArray(LastNonEmpty)
                        Exit For
                    End If
                Next

                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(4)), 6)
                _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(4)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                '_Del_Date = Trim(fields(9))
                _Order_Qty = Trim(fields(9))
                _Del_Qty = Trim(fields(12))
                _Confirm_Qty = Trim(fields(10))
                _FGStock = Trim(fields(13))
                _TollPLS = Trim(fields(18))
                _TollMIN = Trim(fields(17))
                _Merchant1 = Trim(fields(19))
                If Trim(fields(20)) <> "" Then
                    _Confact = Trim(fields(20))
                Else
                    _Confact = 0
                End If
                _depComm = Trim(fields(14))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _depComm = (Replace(_depComm, characterToRemove, ""))


                characterToRemove = ";"

                'MsgBox(Trim(fields(9)))
                _PO_No = (Replace(_PO_No, characterToRemove, ""))

                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _PO_No = (Replace(_PO_No, characterToRemove, ""))
                Dim _Week As Integer

                _Week = DatePart(DateInterval.WeekOfYear, _Del_Date)

                '_Confact = 0
                'nvcFieldList1 = "select * from M22Tec_Spec where M22Quality='" & QualityNo & "'"
                'dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(dsUser) Then
                '    _Confact = dsUser.Tables(0).Rows(0)("M22Con_Fact")
                'End If


                nvcFieldList1 = "select * from M01Sales_Order_SAP where M01Sales_Order='" & Trim(_sales_Order) & "' and M01Line_Item='" & Trim(_LineItem) & "' and M01Material_No='" & _Material & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "update M01Sales_Order_SAP set M01SO_Qty='" & _Order_Qty & "',M01Con_Qty='" & _Confirm_Qty & "',M01Delivary_Qty='" & _Del_Qty & "',M01SO_Date='" & _Del_Date & "' where M01Sales_Order='" & Trim(_sales_Order) & "' and M01Line_Item='" & Trim(_LineItem) & "' and M01Material_No='" & _Material & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M01Sales_Order_SAP(M01Sales_Order,M01PO,M01Customer_Code,M01Cuatomer_Name,M01SO_Date,M01Line_Item,M01Department,M01Material_No,M01Quality,M01SO_Qty,M01Con_Qty,M01Delivary_Qty,M01Cus_Tol_Min,M01Cus_Tol_Pls,M01Tobe_Deliverd,M01Reason_Rejection,M01Status,M01Merchant,M01Quality_No,M01Location,M01Con_Fact,M0130Class)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_PO_No) & "','" & Trim(_Customer) & "','" & Trim(_CustomerName) & "','" & _Del_Date & "','" & Trim(_LineItem) & "','" & Trim(_Department) & "','" & Trim(_Material) & "','" & Trim(_Material_Dis) & "','" & _Order_Qty & "','" & _Confirm_Qty & "','" & _Del_Qty & "','" & _TollMIN & "','" & _TollPLS & "','" & _FGStock & "','" & _depComm & "','A','" & _Merchant1 & "','" & QualityNo & "','" & _Location1 & "','" & _Confact & "','" & _30Class & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
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
                _Confirm_Qty = 0
                _Location1 = ""
                _Confact = 0
                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 16
            lblPro.Text = "Delsum.txt"
            lblPro.Refresh()
            pbCount.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_QulityRCode()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
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
        Dim _Conqty As Double
        Dim _30Class As String
        Dim QualityNo As String
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
        Dim _Merchant1 As String
        Dim _Location1 As String
        Dim _Confact As Double
        Dim _Rcode As String
        Dim _Shade1 As String
        Dim _Shade2 As String

        Try


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Quality_Rcode.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 470 Then
                    ' MsgBox("")
                End If

                

                _Material = Trim(fields(0))
                QualityNo = Trim(fields(1))
                _Material_Dis = Trim(fields(2))
                _Rcode = Trim(fields(3))
                _Shade1 = Trim(fields(4))
                _Shade2 = Trim(fields(5))

                characterToRemove = "-"

                'MsgBox(Trim(fields(9)))
                _Material = (Replace(_Material, characterToRemove, ""))

                characterToRemove = "'"
                _30Class = _Material

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))

                characterToRemove = "-"
                _30Class = (Replace(_30Class, characterToRemove, ""))

             

                nvcFieldList1 = "select * from M16Quality_RCode where M16Material='" & _30Class & "'  "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "update M16Quality_RCode set M16Quality='" & QualityNo & "',M16Description='" & _Material_Dis & "',M16R_Code='" & _Rcode & "',M16Product_Type='" & _Shade1 & "',M16Shade_Type='" & _Shade2 & "' where M16Material='" & _30Class & "'  "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M16Quality_RCode(M16Material,M16Quality,M16Description,M16R_Code,M16Product_Type,M16Shade_Type)" & _
                                                        " values('" & _Material & "', '" & QualityNo & "','" & _Material_Dis & "','" & _Rcode & "','" & _Shade1 & "','" & _Shade2 & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _Material = ""
                _Material_Dis = ""
                _Shade1 = ""
                _Shade2 = ""
                _Rcode = ""
                QualityNo = ""
                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 16
            lblPro.Text = "Quality RCode.txt"
            lblPro.Refresh()
            pbCount.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11 & "Quality_Rcode.txt")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_YarnPO()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
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
        Dim _Conqty As Double
        Dim _30Class As String
        Dim QualityNo As String
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
        Dim _Merchant1 As String
        Dim _Location1 As String
        Dim _10Class As String
        Dim _Item As String
        Dim _PO As String
        Dim _vender As String
        Dim _DelDate As Date
        Dim _St As String
        Dim _Description As String
        Dim _POQty As Double

        Dim _OpenQty As Double
        Dim _Price As Double
        Dim _NetValue As Double

        Try

            nvcFieldList1 = "delete from  M47Yarn_PO  "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\ZODS.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 470 Then
                    ' MsgBox("")
                End If


                _St = Trim(fields(1))
                _Del_Date = Microsoft.VisualBasic.Left(_St, 4) & "/" & Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(_St, 4), 2) & "/" & Microsoft.VisualBasic.Right(_St, 2)
                _vender = Trim(fields(4))
                _PO = Trim(fields(5))
                _Item = Trim(fields(7))
                _10Class = Trim(fields(9))
                _Description = Trim(fields(10))

                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Description = (Replace(_Description, characterToRemove, ""))

                _Description = Trim(fields(10))

                _POQty = Trim(fields(13))
                _POQty = Trim(fields(11))
                _Price = Trim(fields(24))
                _NetValue = Trim(fields(25))

              

                nvcFieldList1 = "Insert Into M47Yarn_PO(M47Ref_No,M47Del_Date,M47Vendor,M47PO_Order,M47Item,M47Material,M47Description,M47PO_Qty,M47Open_Qty,M47Net_Price,M47NetValue)" & _
                                                    " values('" & X11 & "', '" & _Del_Date & "','" & _vender & "','" & _PO & "','" & _Item & "','" & _10Class & "','" & _Description & "','" & _POQty & "','" & _OpenQty & "','" & _Price & "','" & _NetValue & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                _PO = ""
                _10Class = ""
                _Description = ""
                _NetValue = 0
                _PRD_Qty = 0
                _vender = ""
                _Item = ""

                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 16
            lblPro.Text = "Yarn PO.txt"
            lblPro.Refresh()
            pbCount.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11 & "Quality_Rcode.txt")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        cboMaterial.Text = ""
        cboQuality.Text = ""
        chk2.Checked = False
        chk3.Checked = False
        chk4.Checked = False
        chk5.Checked = False
        chk6.Checked = False
        chkUpload.Checked = False
        lblPro.Text = "Progress ...."
        pbCount.Value = 0
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()
    End Sub

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim _1stQualityROW As Integer
        Dim _PreShade As String

        Dim vcWhere As String
        Dim _FirstRow As Integer
        Dim _SHADE As String
        Dim _FromDate As Date
        Dim _ToDate As Date
        '  Dim M02 As DataSet
        Dim cargoWeights(5) As Double
        Dim _20sd(5) As String
        Dim _t As Integer

        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
            Dim range1 As Range

            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 15
            worksheet2.Columns("B").ColumnWidth = 10
            worksheet2.Columns("C").ColumnWidth = 10
            worksheet2.Columns("D").ColumnWidth = 20
            worksheet2.Columns("E").ColumnWidth = 15
            worksheet2.Columns("F").ColumnWidth = 22
            worksheet2.Columns("G").ColumnWidth = 12
            worksheet2.Columns("H").ColumnWidth = 10
            worksheet2.Columns("I").ColumnWidth = 10
            worksheet2.Columns("L").ColumnWidth = 10
            worksheet2.Columns("M").ColumnWidth = 15
            worksheet2.Columns("N").ColumnWidth = 10
            worksheet2.Columns("O").ColumnWidth = 30
            worksheet2.Columns("P").ColumnWidth = 13
            worksheet2.Columns("Q").ColumnWidth = 8
            worksheet2.Columns("R").ColumnWidth = 8
            worksheet2.Columns("S").ColumnWidth = 8
            worksheet2.Columns("T").ColumnWidth = 10
            worksheet2.Columns("U").ColumnWidth = 8
            worksheet2.Columns("V").ColumnWidth = 8
            worksheet2.Columns("W").ColumnWidth = 8
            worksheet2.Columns("X").ColumnWidth = 8
            worksheet2.Columns("Y").ColumnWidth = 8
            worksheet2.Columns("Z").ColumnWidth = 8
            worksheet2.Columns("AA").ColumnWidth = 10
            worksheet2.Columns("AB").ColumnWidth = 10
            worksheet2.Columns("AC").ColumnWidth = 10
            worksheet2.Columns("AD").ColumnWidth = 10


            worksheet2.Cells(1, 1) = "Greige Stock "
            worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Range("A1:Q1").Interior.Color = RGB(197, 217, 241)
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Rows(1).Font.size = 13
            worksheet2.Rows(1).rowheight = 35
            worksheet2.Rows(1).Font.name = "Times New Roman"
            worksheet2.Rows(1).Font.BOLD = True
            worksheet2.Range("A1:M1").MergeCells = True
            worksheet2.Range("A1:M1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 14) = "Colouring requirement "
            worksheet2.Cells(1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Range("N1:AA1").Interior.Color = RGB(197, 217, 241)
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("N1:AA1").MergeCells = True
            worksheet2.Range("N1:AA1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 28) = "SUMMARY"
            worksheet2.Cells(1, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Range("AB1:AI1").Interior.Color = RGB(197, 217, 241)
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("AB1:AI1").MergeCells = True
            worksheet2.Range("AB1:AI1").VerticalAlignment = XlVAlign.xlVAlignCenter

            Dim _Chr As Integer
            Dim I As Integer
            Dim X As Integer
            X = 1

            _Chr = 97
            For I = 1 To 35
                If I = 27 Then
                    _Chr = 97
                End If

                If I >= 27 Then
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                Else

                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                End If
                _Chr = _Chr + 1

            Next

            X = 2
            worksheet2.Rows(X).rowheight = 22
            worksheet2.Cells(X, 1) = "Quality No"
            worksheet2.Rows(X).Font.size = 9
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True
            worksheet2.Range("A2:A2").MergeCells = True
            worksheet2.Range("A2:A2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 2) = "20 Class"
            worksheet2.Range("B2:B2").MergeCells = True
            worksheet2.Range("B2:B2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 3) = "Shade"
            worksheet2.Range("C2:C2").MergeCells = True
            worksheet2.Range("C2:C2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 4) = "Unit "
            worksheet2.Range("D2:D2").MergeCells = True
            worksheet2.Range("D2:D2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 5) = "Fabric Type "
            worksheet2.Range("E2:E2").MergeCells = True
            worksheet2.Range("E2:E2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 6) = "Machine Type"
            worksheet2.Range("F2:F2").MergeCells = True
            worksheet2.Range("F2:F2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 7) = "Below One Month "
            worksheet2.Range("G2:G2").MergeCells = True
            worksheet2.Range("G2:G2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 8) = "02 Month"
            worksheet2.Range("H2:H2").MergeCells = True
            worksheet2.Range("H2:H2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 9) = "03 Month"
            worksheet2.Range("I2:I2").MergeCells = True
            worksheet2.Range("I2:I2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 10) = "04 Month"
            worksheet2.Range("J2:J2").MergeCells = True
            worksheet2.Range("J2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 11) = "05 Month"
            worksheet2.Range("K2:K2").MergeCells = True
            worksheet2.Range("K2:K2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 12) = "06 Month"
            worksheet2.Range("L2:L2").MergeCells = True
            worksheet2.Range("L2:L2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 13) = "Grand Total (Kg)"
            worksheet2.Range("M2:M2").MergeCells = True
            worksheet2.Range("M2:M2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 14) = "30 Class "
            worksheet2.Range("N2:N2").MergeCells = True
            worksheet2.Range("N2:N2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 15) = "Description"
            worksheet2.Range("O2:O2").MergeCells = True
            worksheet2.Range("O2:O2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 16) = "Grige Category"
            worksheet2.Range("p2:p2").MergeCells = True
            worksheet2.Range("p2:p2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 17) = "Width"
            worksheet2.Range("Q2:Q2").MergeCells = True
            worksheet2.Range("Q2:Q2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 18) = "GSM"
            worksheet2.Range("R2:R2").MergeCells = True
            worksheet2.Range("R2:R2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 19) = "C.Factor"
            worksheet2.Range("S2:S2").MergeCells = True
            worksheet2.Range("S2:S2").VerticalAlignment = XlVAlign.xlVAlignCenter



            worksheet2.Cells(X, 20) = "Backlog Qty"
            worksheet2.Range("T2:T2").MergeCells = True
            worksheet2.Range("T2:T2").VerticalAlignment = XlVAlign.xlVAlignCenter

            Dim _dATE As Date
            _dATE = Today
            worksheet2.Cells(X, 21) = "Week " & weekNumber(Today)
            worksheet2.Range("u2:u2").MergeCells = True
            worksheet2.Range("u2:u2").VerticalAlignment = XlVAlign.xlVAlignCenter

            _dATE = _dATE.AddDays(+7)

            worksheet2.Cells(X, 22) = "Week " & weekNumber(_dATE)
            worksheet2.Range("V2:V2").MergeCells = True
            worksheet2.Range("V2:V2").VerticalAlignment = XlVAlign.xlVAlignCenter
            _dATE = _dATE.AddDays(+7)
            worksheet2.Cells(X, 23) = "Week " & weekNumber(_dATE)
            worksheet2.Range("W2:W2").MergeCells = True
            worksheet2.Range("W2:W2").VerticalAlignment = XlVAlign.xlVAlignCenter

            _dATE = _dATE.AddDays(+7)
            worksheet2.Cells(X, 24) = "Week " & weekNumber(_dATE)
            worksheet2.Range("X2:X2").MergeCells = True
            worksheet2.Range("X2:X2").VerticalAlignment = XlVAlign.xlVAlignCenter

            _dATE = _dATE.AddDays(+7)

            worksheet2.Cells(X, 25) = "Week " & weekNumber(_dATE)
            worksheet2.Range("Y2:Y2").MergeCells = True
            worksheet2.Range("Y2:Y2").VerticalAlignment = XlVAlign.xlVAlignCenter

            _dATE = _dATE.AddDays(+7)
            worksheet2.Cells(X, 26) = "Week " & weekNumber(_dATE)
            worksheet2.Range("Z2:Z2").MergeCells = True
            worksheet2.Range("Z2:Z2").VerticalAlignment = XlVAlign.xlVAlignCenter

            _dATE = _dATE.AddDays(+7)
            worksheet2.Cells(X, 27) = "Week " & weekNumber(_dATE)
            worksheet2.Range("AA2:AA2").MergeCells = True
            worksheet2.Range("AA2:AA2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 28) = "Quality "
            worksheet2.Range("AB2:AB2").MergeCells = True
            worksheet2.Range("AB2:AB2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 29) = "Greige catogary "
            worksheet2.Range("AC2:AC2").MergeCells = True
            worksheet2.Range("AC2:AC2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 30) = "Greige Qty "
            worksheet2.Range("AD2:AD2").MergeCells = True
            worksheet2.Range("AD2:AD2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 31) = "Incomming Order "
            worksheet2.Range("AE2:AE2").MergeCells = True
            worksheet2.Range("AE2:AE2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 32) = "Actual Dyeing "
            worksheet2.Range("AF2:AF2").MergeCells = True
            worksheet2.Range("AF2:AF2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 33) = "Total dyeing"
            worksheet2.Range("AG2:AG2").MergeCells = True
            worksheet2.Range("AG2:AG2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 34) = "Balance (Kg)"
            worksheet2.Range("AH2:AH2").MergeCells = True
            worksheet2.Range("AH2:AH2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 35) = "Comments"
            worksheet2.Range("AI2:AI2").MergeCells = True
            worksheet2.Range("AI2:AI2").VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 35
                If I = 27 Then
                    _Chr = 97
                End If

                If I >= 27 Then
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                Else

                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                End If
                _Chr = _Chr + 1

            Next
            X = X + 1

            _Status = ""
            Dim _Last As Integer
            Dim _PreRetailer As String
            Dim z As Integer
            Dim Y As Integer
            Dim _ShadeDis As String
            Dim _SdCount As Integer
            Dim _QualityRow As Integer

            If chk2.Checked = True Then
                If _Status <> "" Then
                    _Status = _Status & ",'2'"
                Else
                    _Status = "'2'"
                End If

            End If

            If chk3.Checked = True Then
                If _Status <> "" Then
                    _Status = _Status & ",'3'"
                Else
                    _Status = "'3'"
                End If

            End If

            If chk4.Checked = True Then
                If _Status <> "" Then
                    _Status = _Status & ",'4'"
                Else
                    _Status = "'4'"
                End If

            End If

            If chk5.Checked = True Then
                If _Status <> "" Then
                    _Status = _Status & ",'5'"
                Else
                    _Status = "'5'"
                End If

            End If

            If chk6.Checked = True Then
                If _Status <> "" Then
                    _Status = _Status & ",'6'"
                Else
                    _Status = "'6'"
                End If

            End If

            If Trim(cboMaterial.Text) <> "" Then

                vcWhere = "M2120Class='" & Trim(cboMaterial.Text) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "LS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            ElseIf Trim(cboQuality.Text) <> "" Then
                vcWhere = "M21Material='" & Trim(cboQuality.Text) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "LS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            ElseIf Trim(cboBU.Text) <> "" Then
                vcWhere = "M14Name='" & Trim(cboBU.Text) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "LS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            ElseIf _Status <> "" Then

            Else
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "LST"))
            End If
            I = 0
            Dim _20Class As String
            _FirstRow = X
            _QualityRow = X
            _t = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _QualityRow = X
                _1stQualityROW = X
                worksheet2.Rows(X).Font.size = 10
                worksheet2.Rows(X).Font.name = "Times New Roman"
                worksheet2.Cells(X, 1) = T01.Tables(0).Rows(I)("M21Material")
                worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 28) = T01.Tables(0).Rows(I)("M21Material")
                worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 4) = T01.Tables(0).Rows(I)("M14Name")
                ' worksheet2.Cells(X, 2) = T01.Tables(0).Rows(I)("M2120Class")

                _20sd(0) = ""
                _20sd(1) = ""
                _20sd(2) = ""
                _20sd(3) = ""
                _20sd(4) = ""

                cargoWeights(0) = 0
                cargoWeights(1) = 0
                cargoWeights(2) = 0
                cargoWeights(3) = 0
                cargoWeights(4) = 0

                _20Class = ""
                Y = 0
                _Last = X
                _FirstRow = X
                _t = 0
                vcWhere = "M22Quality='" & T01.Tables(0).Rows(I)("M21Material") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(X, 5) = M01.Tables(0).Rows(0)("M22Fabric_Type")
                    worksheet2.Cells(X, 6) = M01.Tables(0).Rows(0)("M22Machine_Type")
                End If

                _SdCount = 0
                _SHADE = ""
                _SdCount = X
                vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "LSD"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    worksheet2.Rows(X).Font.size = 10
                    worksheet2.Rows(X).Font.name = "Times New Roman"
                    worksheet2.Cells(X, 2) = M01.Tables(0).Rows(Y)("M2120Class")
                    worksheet2.Cells(X, 3) = M01.Tables(0).Rows(Y)("M23Shade")

                    vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "' and M2120Class= '" & M01.Tables(0).Rows(Y)("M2120Class") & "'   and DATEDIFF(day, m21date, GETDATE()) <=30 "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "DAY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 7) = M02.Tables(0).Rows(0)("M21Qty")
                        worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet2.Cells(X, 7)
                        range1.NumberFormat = "0.00"
                    End If

                    vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "' and M2120Class= '" & M01.Tables(0).Rows(Y)("M2120Class") & "' and DATEDIFF(day, m21date, GETDATE()) >30 and DATEDIFF(day, m21date, GETDATE())<=60 "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "DAY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 8) = M02.Tables(0).Rows(0)("M21Qty")
                        worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet2.Cells(X, 8)
                        range1.NumberFormat = "0.00"
                    End If

                    vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "' and M2120Class= '" & M01.Tables(0).Rows(Y)("M2120Class") & "' and DATEDIFF(day, m21date, GETDATE()) >60 and DATEDIFF(day, m21date, GETDATE())<=90 "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "DAY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 9) = M02.Tables(0).Rows(0)("M21Qty")
                        worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet2.Cells(X, 9)
                        range1.NumberFormat = "0.00"
                    End If

                    vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "' and M2120Class= '" & M01.Tables(0).Rows(Y)("M2120Class") & "' and DATEDIFF(day, m21date, GETDATE()) >90 and DATEDIFF(day, m21date, GETDATE())<=120 "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "DAY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 10) = M02.Tables(0).Rows(0)("M21Qty")
                        worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet2.Cells(X, 10)
                        range1.NumberFormat = "0.00"
                    End If

                    vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "' and M2120Class= '" & M01.Tables(0).Rows(Y)("M2120Class") & "' and DATEDIFF(day, m21date, GETDATE()) >120 and DATEDIFF(day, m21date, GETDATE())<=150 "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "DAY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 11) = M02.Tables(0).Rows(0)("M21Qty")
                        worksheet2.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet2.Cells(X, 11)
                        range1.NumberFormat = "0.00"
                    End If

                    vcWhere = "M21Material='" & T01.Tables(0).Rows(I)("M21Material") & "' and M2120Class= '" & M01.Tables(0).Rows(Y)("M2120Class") & "' and DATEDIFF(day, m21date, GETDATE()) >150 and DATEDIFF(day, m21date, GETDATE())<=180 "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "DAY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 12) = M02.Tables(0).Rows(0)("M21Qty")
                        worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet2.Cells(X, 12)
                        range1.NumberFormat = "0.00"
                    End If
                    If _SHADE = Trim(M01.Tables(0).Rows(Y)("M23Shade")) Then

                        'worksheet2.Range("M" & (X)).Formula = "=SUM(G" & _SdCount & ":L" & X & ")"
                        ''worksheet2.Range("M" & _SdCount & ":M" & X).MergeCells = True
                        ' ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        ''worksheet2.Range("m" & X & ":m" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        ''worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
                        'range1 = worksheet2.Cells(X, 13)
                        'range1.NumberFormat = "0.00"
                    Else
                        If _SHADE <> "" Then
                            _20sd(_t) = _SHADE

                            worksheet2.Range("M" & (X - 1)).Formula = "=SUM(G" & _SdCount & ":L" & X - 1 & ")"
                            worksheet2.Range("M" & _SdCount & ":M" & X - 1).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("m" & X - 1 & ":m" & X - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(X - 1, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
                            range1 = worksheet2.Cells(X - 1, 13)
                            range1.NumberFormat = "0.00"
                            range1 = CType(worksheet2.Cells(_SdCount, 13), Microsoft.Office.Interop.Excel.Range)
                            cargoWeights(_t) = range1.Value
                            _t = _t + 1
                            _SdCount = X
                        Else
                            _SHADE = Trim(M01.Tables(0).Rows(Y)("M23Shade"))
                        End If
                    End If

                    _SHADE = M01.Tables(0).Rows(Y)("M23Shade")
                    _PreRetailer = Trim(T01.Tables(0).Rows(I)("M14Name"))

                    _Chr = 97
                    For z = 1 To 35
                        If z = 27 Then
                            _Chr = 97
                        End If

                        If z >= 27 Then
                            worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            If _PreRetailer = Trim(T01.Tables(0).Rows(I)("M14Name")) Then
                                If Trim(T01.Tables(0).Rows(I)("M14Name")) = "EMG- BRANDS" Then
                                    worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(255, 255, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "INTIMISIMI/ TEZENEIS" Then
                                    worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(0, 176, 240)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "M&S" Then
                                    worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(255, 192, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "VSS" Then
                                    worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(169, 208, 142)
                                End If
                            Else

                            End If
                        Else

                            worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                            If _PreRetailer = Trim(T01.Tables(0).Rows(I)("M14Name")) Then
                                If Trim(T01.Tables(0).Rows(I)("M14Name")) = "EMG- BRANDS" Then
                                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(255, 255, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "INTIMISIMI/ TEZENEIS" Then
                                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(0, 176, 240)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "M&S" Then
                                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(255, 192, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "VSS" Then
                                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(169, 208, 142)
                                End If
                            Else

                            End If
                        End If
                        _Chr = _Chr + 1

                    Next
                    X = X + 1
                    Y = Y + 1
                Next


                worksheet2.Range("M" & (X - 1)).Formula = "=SUM(G" & _SdCount & ":L" & X - 1 & ")"
                worksheet2.Range("M" & _SdCount & ":M" & X - 1).MergeCells = True
                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                worksheet2.Range("m" & X - 1 & ":m" & X - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Cells(X - 1, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet2.Cells(X - 1, 13)
                range1.NumberFormat = "0.00"

                _20sd(_t) = _SHADE
                If _SHADE <> "" Then
                    'If _ShadeDis = "Commen" Then
                    '    _20sd(_t) = "-"
                    'End If
                Else
                    _20sd(_t) = "-"
                End If
                range1 = CType(worksheet2.Cells(_SdCount, 13), Microsoft.Office.Interop.Excel.Range)
                cargoWeights(_t) = range1.Value
                _t = _t + 1

                'worksheet2.Range("A" & _Last & ":A" & X - 1).MergeCells = True
                ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                'worksheet2.Range("A" & _Last & ":A" & X - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                'worksheet2.Cells(_Last, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                'range1 = worksheet2.Cells(_Last, 1)
                '=========================================================================================
                _ShadeDis = ""
                Dim _SHDCOUNT As Integer
                Dim _quality As String
                _quality = ""
                _quality = T01.Tables(0).Rows(I)("M21Material")

                vcWhere = "M26Quality20='" & T01.Tables(0).Rows(I)("M21Material") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "ALT"), New SqlParameter("@vcWhereClause1", vcWhere))

                Y = 0
                For Each DTRow4 As DataRow In M03.Tables(0).Rows
                    _quality = _quality & "','" & M03.Tables(0).Rows(Y)("M26Quality30") & ""
                    Y = Y + 1
                Next
                _SHDCOUNT = _FirstRow
                _PreShade = ""
                '30 class
                vcWhere = "M07Quality='" & T01.Tables(0).Rows(I)("M21Material") & "'"
                vcWhere = "M07Quality IN ('" & _quality & "')"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "30C"), New SqlParameter("@vcWhereClause1", vcWhere))
                Y = 0
                For Each DTRow4 As DataRow In M01.Tables(0).Rows

                    worksheet2.Rows(_FirstRow).Font.size = 10
                    worksheet2.Rows(_FirstRow).Font.name = "Times New Roman"

                    worksheet2.Cells(_FirstRow, 14) = M01.Tables(0).Rows(Y)("M07Material")
                    worksheet2.Cells(_FirstRow, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Cells(_FirstRow, 15) = M01.Tables(0).Rows(Y)("M07Met_Dis")
                    worksheet2.Cells(_FirstRow, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    _FromDate = Today
                    Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                    Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromDate)
                    Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

                    If dayName = "Sunday" Then
                        _FromDate = CDate(_FromDate).AddDays(-6)
                    ElseIf dayName = "Tuesday" Then
                        _FromDate = CDate(_FromDate).AddDays(-1)
                    ElseIf dayName = "Wednesday" Then
                        _FromDate = CDate(_FromDate).AddDays(-2)
                    ElseIf dayName = "Thursday" Then
                        _FromDate = CDate(_FromDate).AddDays(-3)
                    ElseIf dayName = "Friday" Then
                        _FromDate = CDate(_FromDate).AddDays(-4)
                    ElseIf dayName = "Saturday" Then
                        _FromDate = CDate(_FromDate).AddDays(-5)
                    End If

                    _ToDate = _FromDate.AddDays(+6)
                    Dim _redQty As Double
                    _redQty = 0
                    Dim _Material As String
                    Dim _Prd_Qty As Double

                    _Material = M01.Tables(0).Rows(Y)("M07Material")
                    _Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, 5)

                    vcWhere = "Metrrial='" & _Material & "' and del_date<'" & _ToDate & "' " 'and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                        '  _redQty = _redQty + M03.Tables(0).Rows(0)("Del_Qty")

                    End If
                    _Prd_Qty = 0
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    ' _redQty = _redQty + M03.Tables(0).Rows(0)("PRD_Qty")
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            '  _redQty = _redQty + M03.Tables(0).Rows(0)("PRD_Qty")
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date<'" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 20) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 20)
                        range1.NumberFormat = "0.00"


                    End If

                 

                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0
                    _Prd_Qty = 0
                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between '" & _FromDate & "'  and '" & _ToDate & "'"' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                    End If
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 21) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 21)
                        range1.NumberFormat = "0.00"

                    End If


                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0

                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between '" & _FromDate & "' and '" & _ToDate & "'" ' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                        '_redQty = _redQty + M03.Tables(0).Rows(0)("Del_Qty")
                    End If


                    _Prd_Qty = 0

                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    ' _redQty = _redQty + M03.Tables(0).Rows(0)("PRD_Qty")
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            ' _redQty = _redQty + M03.Tables(0).Rows(0)("PRD_Qty")
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next


                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 22) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 22)
                        range1.NumberFormat = "0.00"

                    End If
                 

                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0
                    _Prd_Qty = 0
                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between'" & _FromDate & "' and '" & _ToDate & "'" ' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                    End If
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 23) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 23)
                        range1.NumberFormat = "0.00"

                    End If

                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0
                    _Prd_Qty = 0
                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between '" & _FromDate & "' and '" & _ToDate & "'" ' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                    End If
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 24) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 24)
                        range1.NumberFormat = "0.00"

                    End If


                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0
                    _Prd_Qty = 0
                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between '" & _FromDate & "' and '" & _ToDate & "'" ' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                    End If
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 25) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 25)
                        range1.NumberFormat = "0.00"

                    End If

                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0
                    _Prd_Qty = 0
                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between '" & _FromDate & "' and '" & _ToDate & "'" ' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                    End If
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 26) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 26)
                        range1.NumberFormat = "0.00"

                    End If

                    _FromDate = _ToDate.AddDays(+1)
                    _ToDate = _FromDate.AddDays(+6)

                    _redQty = 0
                    _Prd_Qty = 0
                    vcWhere = "Metrrial='" & M01.Tables(0).Rows(Y)("M07Material") & "' and del_date between '" & _FromDate & "' and '" & _ToDate & "'" ' and location in ('2065','2062','2055','2070','AW PREPARATION','AW PRESETTING','EXAM','FINISHING','dye')"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    z = 0
                    If isValidDataset(M03) Then
                        _redQty = M03.Tables(0).Rows(0)("FG_Stock")
                    End If
                    For Each DTRow5 As DataRow In M03.Tables(0).Rows
                        If Trim(M03.Tables(0).Rows(z)("Location")) = "Dye" Then
                            vcWhere = "Batch_No='" & M03.Tables(0).Rows(z)("Prduct_Order") & "'"
                            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(dsUser) Then
                                If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                    _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                                End If
                            End If
                        Else
                            _Prd_Qty = _Prd_Qty + M03.Tables(0).Rows(0)("PRD_Qty")
                        End If
                        z = z + 1
                    Next

                    vcWhere = "M07Material='" & M01.Tables(0).Rows(Y)("M07Material") & "' and m07date BETWEEN '" & _FromDate & "' AND '" & _ToDate & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(_FirstRow, 27) = Trim(M02.Tables(0).Rows(0)("Qty")) - (_redQty + _Prd_Qty)
                        range1 = worksheet2.Cells(_FirstRow, 27)
                        range1.NumberFormat = "0.00"

                    End If

                    If M01.Tables(0).Rows(Y)("M14Grige") <> "" Then
                        If _PreShade = Trim(M01.Tables(0).Rows(Y)("M14Grige")) Then
                            If _SHDCOUNT = _FirstRow Then
                                _ShadeDis = Trim(M01.Tables(0).Rows(Y)("M14Grige"))
                                worksheet2.Cells(_FirstRow, 16) = Trim(M01.Tables(0).Rows(Y)("M14Grige"))
                                worksheet2.Cells(_FirstRow, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                                worksheet2.Cells(_FirstRow, 29) = Trim(M01.Tables(0).Rows(Y)("M14Grige"))
                                worksheet2.Cells(_FirstRow, 29).HorizontalAlignment = XlHAlign.xlHAlignLeft
                            End If
                        Else
                            _ShadeDis = Trim(M01.Tables(0).Rows(Y)("M14Grige"))
                            worksheet2.Cells(_FirstRow, 16) = Trim(M01.Tables(0).Rows(Y)("M14Grige"))
                            worksheet2.Cells(_FirstRow, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                            worksheet2.Cells(_FirstRow, 29) = Trim(M01.Tables(0).Rows(Y)("M14Grige"))
                            worksheet2.Cells(_FirstRow, 29).HorizontalAlignment = XlHAlign.xlHAlignLeft

                            If _SHDCOUNT = _FirstRow Then
                                'worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow).MergeCells = True
                                ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                'worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                'worksheet2.Cells(_FirstRow, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                'range1 = worksheet2.Cells(_FirstRow, 16)

                            Else
                                worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 16)

                              

                                worksheet2.Range("AF" & _FirstRow - 1).Formula = "=SUM(T" & _SHDCOUNT & ":T" & _FirstRow - 1 & ")+SUM(U" & _SHDCOUNT & ":U" & _FirstRow - 1 & ")+SUM(V" & _SHDCOUNT & ":V" & _FirstRow - 1 & ")+SUM(W" & _SHDCOUNT & ":W" & _FirstRow - 1 & ")+SUM(X" & _SHDCOUNT & ":X" & _FirstRow - 1 & ")+SUM(Y" & _SHDCOUNT & ":Y" & _FirstRow - 1 & ")+SUM(Z" & _SHDCOUNT & ":Z" & _FirstRow - 1 & ")"
                                worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 32)
                                range1.NumberFormat = "0.00"


                                worksheet2.Range("AG" & _FirstRow - 1).Formula = "=SUM(AF" & _SHDCOUNT & ":AE" & _SHDCOUNT & ")"
                                worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 33)
                                range1.NumberFormat = "0.00"

                                Dim G As Integer
                                For G = 0 To 5

                                    If _PreShade = Microsoft.VisualBasic.Left(_20sd(G), 1) Then
                                        If _SHDCOUNT = _FirstRow Then
                                            'worksheet2.Cells(_FirstRow, 30) = cargoWeights(_F)
                                            'worksheet2.Cells(_FirstRow, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                            'worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            'worksheet2.Cells(_FirstRow, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            'range1 = worksheet2.Cells(_FirstRow, 30)
                                            'range1.NumberFormat = "0.00"
                                        Else
                                            worksheet2.Range("Ad" & _SHDCOUNT & ":Ad" & _FirstRow - 1).MergeCells = True
                                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                            worksheet2.Range("Ad" & _SHDCOUNT & ":Ad" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                            range1.NumberFormat = "0.00"

                                            worksheet2.Cells(_SHDCOUNT, 30) = cargoWeights(G)
                                            worksheet2.Cells(_SHDCOUNT, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                            range1.NumberFormat = "0.00"

                                            '_SHDCOUNT = _FirstRow
                                        End If

                                        Exit For
                                    ElseIf _PreShade = "Common" Then
                                        If _20sd(G) = "-" Then
                                            worksheet2.Cells(_FirstRow - 1, 30) = cargoWeights(G)
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                            range1.NumberFormat = "0.00"
                                            Exit For
                                        End If
                                    End If
                                Next
                                worksheet2.Range("AG" & _FirstRow - 1).Formula = "=SUM(AF" & _SHDCOUNT & ":AE" & _SHDCOUNT & ")"
                                worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 33)
                                range1.NumberFormat = "0.00"
                                If _20sd(1) = "Dark" Or _20sd(1) = "Light" Or _20sd(1) = "Marl" Or _20sd(1) = "Dyed yarn" Then
                                    worksheet2.Cells(_FirstRow - 1, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                    range1.NumberFormat = "0.00"

                                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                Else
                                    If _20sd(0) = "Dark" Or _20sd(0) = "Light" Or _20sd(0) = "Marl" Or _20sd(0) = "Dyed yarn" Then
                                        worksheet2.Cells(_FirstRow - 1, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                                        worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                        range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                        range1.NumberFormat = "0.00"

                                        worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).MergeCells = True
                                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                        worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                        range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                    End If
                                End If
                                If _20sd(1) = "Dark" Or _20sd(1) = "Light" Or _20sd(1) = "Marl" Or _20sd(1) = "Dyed yarn" Then
                                    range1 = CType(worksheet2.Cells(_SHDCOUNT, 34), Microsoft.Office.Interop.Excel.Range)
                                    If range1.Value > 0 Then
                                        worksheet2.Cells(_FirstRow - 1, 35) = "Orders Available"
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    Else
                                        worksheet2.Cells(_FirstRow - 1, 35) = "No Orders"
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    End If
                                Else
                                    'new modification on 2015.7.1
                                    range1 = CType(worksheet2.Cells(_SHDCOUNT, 34), Microsoft.Office.Interop.Excel.Range)
                                    If range1.Value > 0 Then
                                        worksheet2.Cells(_FirstRow - 1, 35) = "Orders Available"
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    Else
                                        worksheet2.Cells(_FirstRow - 1, 35) = "No Orders"
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                        worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    End If
                                End If
                                End If

                                If _SHDCOUNT = _FirstRow Then
                                    'worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow).MergeCells = True
                                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    'worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 29)

                                    'worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow).MergeCells = True
                                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    'worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 31)

                                    'worksheet2.Range("AF" & _FirstRow).Formula = "=SUM(T" & _QualityRow & ":T" & _FirstRow - 1 & ")+SUM(U" & _QualityRow & ":U" & _FirstRow - 1 & ")+SUM(V" & _QualityRow & ":V" & _FirstRow - 1 & ")+SUM(W" & _QualityRow & ":W" & _FirstRow - 1 & ")+SUM(X" & _QualityRow & ":X" & _FirstRow - 1 & ")+SUM(Y" & _QualityRow & ":Y" & _FirstRow - 1 & ")+SUM(Z" & _QualityRow & ":Z" & _FirstRow - 1 & ")"
                                    'worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow).MergeCells = True
                                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    'worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 32)
                                    'range1.NumberFormat = "0.00"


                                    'worksheet2.Range("AG" & _FirstRow).Formula = "=SUM(AF" & _SHDCOUNT & ":AE" & _SHDCOUNT & ")"
                                    'worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow).MergeCells = True
                                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    'worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 33)
                                    'range1.NumberFormat = "0.00"
                                Else
                                    worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 29)

                                    worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 31)

                                    worksheet2.Range("AF" & _FirstRow - 1).Formula = "=SUM(T" & _SHDCOUNT & ":T" & _FirstRow - 1 & ")+SUM(U" & _SHDCOUNT & ":U" & _FirstRow - 1 & ")+SUM(V" & _SHDCOUNT & ":V" & _FirstRow - 1 & ")+SUM(W" & _SHDCOUNT & ":W" & _FirstRow - 1 & ")+SUM(X" & _SHDCOUNT & ":X" & _FirstRow - 1 & ")+SUM(Y" & _SHDCOUNT & ":Y" & _FirstRow - 1 & ")+SUM(Z" & _SHDCOUNT & ":Z" & _FirstRow - 1 & ")"
                                    worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 32)
                                    range1.NumberFormat = "0.00"


                                    worksheet2.Range("AG" & _FirstRow - 1).Formula = "=SUM(AF" & _SHDCOUNT & ":AE" & _SHDCOUNT & ")"
                                    worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 33)
                                    range1.NumberFormat = "0.00"
                                End If
                                Dim _F As Integer
                                For _F = 0 To 5

                                    If _PreShade = Microsoft.VisualBasic.Left(_20sd(_F), 1) Then
                                        If _SHDCOUNT = _FirstRow Then
                                            'worksheet2.Cells(_FirstRow, 30) = cargoWeights(_F)
                                            'worksheet2.Cells(_FirstRow, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                            'worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            'worksheet2.Cells(_FirstRow, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            'range1 = worksheet2.Cells(_FirstRow, 30)
                                            'range1.NumberFormat = "0.00"
                                        Else
                                            worksheet2.Range("Ad" & _SHDCOUNT & ":Ad" & _FirstRow - 1).MergeCells = True
                                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                            worksheet2.Range("Ad" & _SHDCOUNT & ":Ad" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                            range1.NumberFormat = "0.00"

                                            worksheet2.Cells(_FirstRow - 1, 30) = cargoWeights(_F)
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                            range1.NumberFormat = "0.00"

                                            _SHDCOUNT = _FirstRow
                                        End If

                                        Exit For
                                    ElseIf _ShadeDis = "Commen" Then
                                        If _20sd(_F) = "-" Then
                                            worksheet2.Cells(_FirstRow - 1, 30) = cargoWeights(_F)
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                            range1.NumberFormat = "0.00"
                                            Exit For
                                        End If
                                    End If
                                Next

                                If _SHDCOUNT = _FirstRow Then
                                    'worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow).MergeCells = True
                                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    'worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 30)

                                    'worksheet2.Cells(_FirstRow, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                                    'worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    'worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 34)
                                    'range1.NumberFormat = "0.00"

                                    'worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow).MergeCells = True
                                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    'worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    'worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    'range1 = worksheet2.Cells(_FirstRow, 34)

                                Else
                                    worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 30)

                                    worksheet2.Cells(_FirstRow - 1, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                    range1.NumberFormat = "0.00"

                                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                End If
                                _SHDCOUNT = _FirstRow
                        End If
                    Else
                            If Trim(M01.Tables(0).Rows(Y)("M14Grige")) <> "" Then
                            Else
                            If _ShadeDis = "Common" Then
                            Else
                                _ShadeDis = "Common"
                                worksheet2.Cells(_FirstRow, 16) = "Common"
                                worksheet2.Cells(_FirstRow, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

                                worksheet2.Cells(_FirstRow, 29) = "Common"
                                worksheet2.Cells(_FirstRow, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter

                                'worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow - 1).MergeCells = True
                                ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                'worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                'worksheet2.Cells(_FirstRow, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                'range1 = worksheet2.Cells(_FirstRow, 16)
                                _SHDCOUNT = _FirstRow
                            End If
                            End If
                    End If

                    _Chr = 97
                    For z = 1 To 35
                        If z = 27 Then
                            _Chr = 97
                        End If

                        If z >= 27 Then
                            worksheet2.Range("A" & Chr(_Chr) & _FirstRow, "A" & Chr(_Chr) & _FirstRow).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range("A" & Chr(_Chr) & _FirstRow, "A" & Chr(_Chr) & _FirstRow).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range("A" & Chr(_Chr) & _FirstRow, "A" & Chr(_Chr) & _FirstRow).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            If _PreRetailer = Trim(T01.Tables(0).Rows(I)("M14Name")) Then
                                If Trim(T01.Tables(0).Rows(I)("M14Name")) = "EMG- BRANDS" Then
                                    worksheet2.Range("A" & Chr(_Chr) & _FirstRow & ":A" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(255, 255, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "INTIMISIMI/ TEZENEIS" Then
                                    worksheet2.Range("A" & Chr(_Chr) & _FirstRow & ":A" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(0, 176, 240)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "M&S" Then
                                    worksheet2.Range("A" & Chr(_Chr) & _FirstRow & ":A" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(255, 192, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "VSS" Then
                                    worksheet2.Range("A" & Chr(_Chr) & _FirstRow & ":A" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(169, 208, 142)
                                End If
                            Else

                            End If
                        Else

                            worksheet2.Range(Chr(_Chr) & _FirstRow, Chr(_Chr) & _FirstRow).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range(Chr(_Chr) & _FirstRow, Chr(_Chr) & _FirstRow).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range(Chr(_Chr) & _FirstRow, Chr(_Chr) & _FirstRow).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                            If _PreRetailer = Trim(T01.Tables(0).Rows(I)("M14Name")) Then
                                If Trim(T01.Tables(0).Rows(I)("M14Name")) = "EMG- BRANDS" Then
                                    worksheet2.Range(Chr(_Chr) & _FirstRow & ":" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(255, 255, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "INTIMISIMI/ TEZENEIS" Then
                                    worksheet2.Range(Chr(_Chr) & _FirstRow & ":" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(0, 176, 240)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "M&S" Then
                                    worksheet2.Range(Chr(_Chr) & _FirstRow & ":" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(255, 192, 0)
                                ElseIf Trim(T01.Tables(0).Rows(I)("M14Name")) = "VSS" Then
                                    worksheet2.Range(Chr(_Chr) & _FirstRow & ":" & Chr(_Chr) & _FirstRow).Interior.Color = RGB(169, 208, 142)
                                End If
                            Else

                            End If
                        End If
                        _Chr = _Chr + 1

                    Next
                    _PreShade = _ShadeDis
                    _FirstRow = _FirstRow + 1
                    Y = Y + 1
                Next
                '========================================================================================
                ''alternative 30class
            
                If isValidDataset(M01) Then
                    If _SHDCOUNT = _FirstRow Then
                    Else
                        worksheet2.Range("Q" & _QualityRow & ":Q" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("Q" & _QualityRow & ":Q" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 17)

                        worksheet2.Range("R" & _QualityRow & ":R" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("R" & _QualityRow & ":R" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 18)
                    End If
                    'worksheet2.Range("T" & _QualityRow & ":T" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("T" & _QualityRow & ":T" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 20)

                    'worksheet2.Range("U" & _QualityRow & ":U" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("U" & _QualityRow & ":U" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 21)

                    'worksheet2.Range("V" & _QualityRow & ":V" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("V" & _QualityRow & ":V" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 22)

                    'worksheet2.Range("W" & _QualityRow & ":W" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("W" & _QualityRow & ":W" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 23)

                    'worksheet2.Range("X" & _QualityRow & ":X" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("X" & _QualityRow & ":X" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 24)

                    'worksheet2.Range("Y" & _QualityRow & ":Y" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("Y" & _QualityRow & ":Y" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 25)

                    'worksheet2.Range("Z" & _QualityRow & ":Z" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("Z" & _QualityRow & ":Z" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 26)

                    'worksheet2.Range("AA" & _QualityRow & ":AA" & _FirstRow - 1).MergeCells = True
                    ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    'worksheet2.Range("AA" & _QualityRow & ":AA" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet2.Cells(_FirstRow - 1, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    'range1 = worksheet2.Cells(_FirstRow - 1, 27)
                    If _SHDCOUNT = _FirstRow Then
                        worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 31)

                        worksheet2.Range("AF" & _FirstRow).Formula = "=SUM(T" & _QualityRow & ":T" & _FirstRow - 1 & ")+SUM(U" & _QualityRow & ":U" & _FirstRow - 1 & ")+SUM(V" & _QualityRow & ":V" & _FirstRow - 1 & ")+SUM(W" & _QualityRow & ":W" & _FirstRow - 1 & ")+SUM(X" & _QualityRow & ":X" & _FirstRow - 1 & ")+SUM(Y" & _QualityRow & ":Y" & _FirstRow - 1 & ")+SUM(Z" & _QualityRow & ":Z" & _FirstRow - 1 & ")+SUM(AA" & _QualityRow & ":AA" & _FirstRow - 1 & ")"
                        worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 32)
                        range1.NumberFormat = "0.00"


                        worksheet2.Range("AG" & _FirstRow - 1).Formula = "=SUM(AF" & _SHDCOUNT & ":AE" & _SHDCOUNT & ")"
                        worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 33)
                        range1.NumberFormat = "0.00"

                        worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 16)


                        worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 29)

                    Else
                        worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AE" & _SHDCOUNT & ":AE" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 31)

                        worksheet2.Range("AF" & _FirstRow - 1).Formula = "=SUM(T" & _SHDCOUNT & ":T" & _FirstRow - 1 & ")+SUM(U" & _SHDCOUNT & ":U" & _FirstRow - 1 & ")+SUM(V" & _SHDCOUNT & ":V" & _FirstRow - 1 & ")+SUM(W" & _SHDCOUNT & ":W" & _FirstRow - 1 & ")+SUM(X" & _SHDCOUNT & ":X" & _FirstRow - 1 & ")+SUM(Y" & _SHDCOUNT & ":Y" & _FirstRow - 1 & ")+SUM(Z" & _SHDCOUNT & ":Z" & _FirstRow - 1 & ")+SUM(AA" & _SHDCOUNT & ":AA" & _FirstRow - 1 & ")"
                        worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AF" & _SHDCOUNT & ":AF" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 32)
                        range1.NumberFormat = "0.00"


                        worksheet2.Range("AG" & _FirstRow - 1).Formula = "=SUM(AF" & _SHDCOUNT & ":AE" & _SHDCOUNT & ")"
                        worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AG" & _SHDCOUNT & ":AG" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 33)
                        range1.NumberFormat = "0.00"

                        'worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("P" & _SHDCOUNT & ":P" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 16)


                        worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow - 1).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AC" & _SHDCOUNT & ":AC" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow - 1, 29)
                    End If
                    Dim _R As Integer
                    For _R = 0 To 5

                        If _ShadeDis = Microsoft.VisualBasic.Left(_20sd(_R), 1) Then
                            worksheet2.Cells(_FirstRow - 1, 30) = cargoWeights(_R)
                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(_FirstRow - 1, 30)
                            range1.NumberFormat = "0.00"
                            Exit For
                        ElseIf _ShadeDis = "Common" Then
                            If _20sd(_R) = "-" Then
                                worksheet2.Cells(_FirstRow - 1, 30) = cargoWeights(_R)
                                worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                                worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                range1.NumberFormat = "0.00"
                                Exit For
                            End If
                        Else
                            'worksheet2.Cells(_FirstRow - 1, 30) = cargoWeights(_R)
                            'worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                            'worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            'worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            'range1 = worksheet2.Cells(_FirstRow - 1, 30)
                            'range1.NumberFormat = "0.00"
                            'Exit For
                        End If
                    Next

                    If _SHDCOUNT = _FirstRow Then

                        worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 30)

                        worksheet2.Cells(_FirstRow, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                        worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                        worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 34)
                        range1.NumberFormat = "0.00"

                        worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(_FirstRow, 34)
                        range1 = CType(worksheet2.Cells(_FirstRow, 34), Microsoft.Office.Interop.Excel.Range)
                        If range1.Value > 0 Then
                            worksheet2.Cells(_FirstRow, 35) = "Orders Available"
                            worksheet2.Cells(_FirstRow, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            worksheet2.Cells(_FirstRow, 35) = "No Orders"
                            worksheet2.Cells(_FirstRow, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        End If
                    Else
                        If _20sd(1) = "Dark" Or _20sd(1) = "Light" Or _20sd(1) = "Marl" Or _20sd(1) = "Dyed yarn" Then
                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(_FirstRow - 1, 30)

                            worksheet2.Cells(_FirstRow - 1, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                            worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                            worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(_FirstRow - 1, 34)
                            range1.NumberFormat = "0.00"

                            worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(_FirstRow - 1, 34)

                            range1 = worksheet2.Cells(_FirstRow - 1, 34)
                            range1 = CType(worksheet2.Cells(_SHDCOUNT, 34), Microsoft.Office.Interop.Excel.Range)
                            If range1.Value > 0 Then
                                worksheet2.Cells(_FirstRow - 1, 35) = "Orders Available"
                                worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            Else
                                worksheet2.Cells(_FirstRow - 1, 35) = "No Orders"
                                worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            End If
                        Else
                            'new modification on 2015.6.30
                            If _20sd(0) = "-" Then
                                worksheet2.Range("AD" & _QualityRow & ":AD" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AD" & _QualityRow & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 30)
                                If _SHDCOUNT = _QualityRow Then
                                    worksheet2.Cells(_QualityRow, 34) = "=(AG" & _QualityRow & ")" & "-AD" & _QualityRow
                                Else
                                    worksheet2.Cells(_QualityRow, 34) = "=(AG" & _QualityRow & "+AG" & _SHDCOUNT & ")" & "-AD" & _QualityRow
                                End If
                                ' worksheet2.Cells(_FirstRow - 1, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                                worksheet2.Cells(_QualityRow, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                                worksheet2.Range("AH" & _QualityRow & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_QualityRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                range1.NumberFormat = "0.00"

                                worksheet2.Range("AH" & _QualityRow & ":AH" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AH" & _QualityRow & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_QualityRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_QualityRow, 34)

                                range1 = worksheet2.Cells(_QualityRow, 34)
                                range1 = CType(worksheet2.Cells(_QualityRow, 34), Microsoft.Office.Interop.Excel.Range)
                                If range1.Value > 0 Then
                                    worksheet2.Cells(_QualityRow, 35) = "Orders Available"
                                    worksheet2.Cells(_QualityRow, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    worksheet2.Range("AI" & _QualityRow & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_QualityRow, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    worksheet2.Range("AI" & _QualityRow & ":AI" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AI" & _QualityRow & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_QualityRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                Else
                                    worksheet2.Cells(_QualityRow, 35) = "No Orders"
                                    worksheet2.Cells(_QualityRow, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    worksheet2.Range("AI" & _QualityRow & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_QualityRow, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    worksheet2.Range("AI" & _QualityRow & ":AI" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AI" & _QualityRow & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_QualityRow, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                End If
                            Else
                                worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 30)

                                worksheet2.Cells(_FirstRow - 1, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                                worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                                worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                range1.NumberFormat = "0.00"

                                worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).MergeCells = True
                                '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                range1 = worksheet2.Cells(_FirstRow - 1, 34)
                                range1 = CType(worksheet2.Cells(_FirstRow - 1, 34), Microsoft.Office.Interop.Excel.Range)
                                If range1.Value > 0 Then
                                    worksheet2.Cells(_FirstRow - 1, 35) = "Orders Available"
                                    worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                Else
                                    worksheet2.Cells(_FirstRow - 1, 35) = "No Orders"
                                    worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                                    worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                    worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                                    worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                                    worksheet2.Cells(_FirstRow - 1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                End If
                            End If
                        End If

                    End If
                End If

                If isValidDataset(M01) Then
                Else
                    worksheet2.Cells(_Last, 30) = cargoWeights(0)
                    worksheet2.Cells(_Last, 30).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet2.Range("AD" & _SHDCOUNT & ":AD" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_Last, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_Last, 30)
                    range1.NumberFormat = "0.00"

                    worksheet2.Cells(_Last, 34) = "=AG" & _SHDCOUNT & "-AD" & _SHDCOUNT
                    worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_Last, 34)
                    range1.NumberFormat = "0.00"

                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _Last).MergeCells = True
                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    worksheet2.Range("AH" & _SHDCOUNT & ":AH" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_Last, 34)
                    range1 = CType(worksheet2.Cells(_Last, 34), Microsoft.Office.Interop.Excel.Range)
                    If range1.Value > 0 Then
                        worksheet2.Cells(_Last, 35) = "Orders Available"
                        worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_FirstRow - 1, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _Last).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet2.Cells(_Last, 35) = "No Orders"
                        worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _Last).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _Last).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If
                End If

                'modify by suranga 2015.7.9 request by Sameera
                If _PreShade = "L" Or _PreShade = "D" Then
                    If _20sd(0) = "-" Then
                        worksheet2.Cells(_Last, 30) = cargoWeights(0)
                        worksheet2.Cells(_Last, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        range1 = worksheet2.Cells(_Last, 34)
                        range1 = CType(worksheet2.Cells(_Last, 34), Microsoft.Office.Interop.Excel.Range)
                        If range1.Value > 0 Then
                            worksheet2.Cells(_Last, 35) = "Orders Available"
                            worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            worksheet2.Cells(_Last, 35) = "No Orders"
                            worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignRight
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_Last, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).MergeCells = True
                            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                            worksheet2.Range("AI" & _SHDCOUNT & ":AI" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(_Last, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        End If
                    End If
                End If
                If X > _FirstRow Then
                    worksheet2.Range("A" & _Last & ":A" & X - 1).MergeCells = True
                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    worksheet2.Range("A" & _Last & ":A" & X - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_Last, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_Last, 1)

                    worksheet2.Range("AB" & _Last & ":AB" & X - 1).MergeCells = True
                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    worksheet2.Range("AB" & _Last & ":AB" & X - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_Last, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_Last, 1)

                    _Last = X
                Else
                    worksheet2.Range("A" & _Last & ":A" & _FirstRow - 1).MergeCells = True
                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    worksheet2.Range("A" & _Last & ":A" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_Last, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_Last, 1)

                    worksheet2.Range("AB" & _Last & ":AB" & _FirstRow - 1).MergeCells = True
                    '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                    worksheet2.Range("AB" & _Last & ":AB" & _FirstRow - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Cells(_FirstRow, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(_FirstRow, 28)

                    _Last = _FirstRow
                    X = _FirstRow
                End If
                '  X = X + 1
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

    Public Function weekNumber(ByVal d As Date) As Integer
        weekNumber = DatePart(DateInterval.WeekOfYear, d, FirstDayOfWeek.Monday, FirstWeekOfYear.System)

    End Function

    Function Upload_Grige()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _PO_No As String
        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _Shadule As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Customer As String
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _20Cls As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double
        Dim _RollNo As String


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
        Dim _BatchNo As String

        Try
            'nvcFieldList1 = "delete from M21Stock_Grige"
            DBEngin.ExecuteScalar(connection, transaction, "up_GetSetStock_Grige", New SqlParameter("@cQryType", "DEL"))

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Stock_Greige.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 4429 Then
                    ' MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then
                If Trim(fields(12)) = "2040" Then
                    If (Trim(fields(0))) <> "" Then
                        _sales_Order = (Trim(fields(0)))
                    Else
                        _sales_Order = "0"
                    End If

                    'If (Trim(fields(1))) <> "" Then
                    _LineItem = (Trim(fields(1)))
                Else
                    _LineItem = 0
                End If
                _20Cls = Trim(fields(2))
                _Material_Dis = Trim(fields(3))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))


                _Merchant = Trim(fields(5))
                _Department = Trim(fields(4))
                _RollNo = Trim(fields(6))

                Dim TestString As String = _Material_Dis
                Dim TestArray() As String = Split(TestString)

                ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""} 
                Dim LastNonEmpty As Integer = -1
                For z As Integer = 0 To TestArray.Length - 1
                    If TestArray(z) <> "" Then
                        LastNonEmpty += 1
                        TestArray(LastNonEmpty) = TestArray(z)
                        _Material = TestArray(LastNonEmpty)
                        'If Microsoft.VisualBasic.Left(_Material, 1) = "Q" Then
                        '    _Material = Microsoft.VisualBasic.Right(_Material, Len(_Material) - 1)
                        'End If
                        Exit For
                    End If
                Next

                ' _Merchnat = Trim(fields(8))
                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(9)), 6)
                _Del_Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(9)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                '_Del_Date = Trim(fields(9))
                _Order_Qty = Trim(fields(10))
                _BatchNo = Trim(fields(7))
                'Dim _Week As Integer

                '_Week = DatePart(DateInterval.WeekOfYear, _Del_Date)
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Department = (Replace(_Department, characterToRemove, ""))

                nvcFieldList1 = "select * from M21Stock_Grige where M21RollNo='" & _RollNo & "' and M21Qty=" & Trim(_Order_Qty) & "  "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into M21Stock_Grige(M21Sales_Order,M21Line_Item,M2120Class,M21Dis,M21Merch,M21RollNo,M21Date,M21Qty,M21Location,M21Material,M21Batch_No)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_20Cls) & "','" & Trim(_Material_Dis) & "','" & _Merchant & "','" & Trim(_RollNo) & "','" & Trim(_Del_Date) & "','" & Trim(_Order_Qty) & "','" & Trim(_Department) & "','" & _Material & "','" & _BatchNo & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
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
                _Confirm_Qty = 0

                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 32
            lblPro.Text = lblPro.Text & ",Stock_Greige.txt"
            pbCount.Refresh()
            lblPro.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_Rcode()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _Quality As String
        Dim _Order As String
        Dim _Type As String
        Dim _Shade As String
        Dim _Shade_Cat As String
        Dim _Status As String
        Dim _Material As String
        Dim _YarnCnt As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Res_Date As String


        Dim _Buyer As String
        Dim _Customer As String
        Dim _Graige As String

        Dim _Criticle As String
       
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
            'nvcFieldList1 = "delete from M21Stock_Grige"
            '  DBEngin.ExecuteScalar(connection, transaction, "up_GetSetStock_Grige", New SqlParameter("@cQryType", "DEL"))

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\R_code.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 36994 Then
                    ' MsgBox("")
                End If


                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then

                _Order = (Trim(fields(0)))
                _Type = (Trim(fields(1)))
                _Shade = (Trim(fields(2)))
                _Graige = (Trim(fields(3)))
                _Status = (Trim(fields(4)))
                _Material = (Trim(fields(5)))
                _YarnCnt = (Trim(fields(6)))
                _Res_Date = (Trim(fields(7)))
                _Buyer = (Trim(fields(8)))
                _Customer = (Trim(fields(9)))
                characterToRemove = "'"
                _Customer = (Replace(_Customer, characterToRemove, ""))
                characterToRemove = "'"
                _Status = (Replace(_Status, characterToRemove, ""))
                characterToRemove = "'"
                _Shade = (Replace(_Shade, characterToRemove, ""))
                _Buyer = (Replace(_Buyer, characterToRemove, ""))
                _YarnCnt = (Replace(_YarnCnt, characterToRemove, ""))
                _Shade_Cat = (Trim(fields(10)))
                _Quality = (Trim(fields(11)))
                _Criticle = (Trim(fields(12)))
               
                characterToRemove = "#"
                _Shade = (Replace(_Shade, characterToRemove, ""))

                nvcFieldList1 = "select * from M14R_Code where M14Order='" & Trim(_Order) & "'  "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "UPDATE M14R_Code SET M14Type='" & _Type & "',M14Shade='" & _Shade & "',M14Grige='" & _Graige & "',M14Status='" & _Status & "',M14Material='" & _Material & "',M14Shade_Cat='" & _Shade_Cat & "' WHERE M14Order='" & Trim(_Order) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M14R_Code(M14Order,M14Type,M14Shade,M14Grige,M14Status,M14Material,M14Yarn_Cnt,M14Rec_Date,M14Buyer,M14Customer,M14Shade_Cat,M14Quality,M14Criticle)" & _
                                                        " values('" & Trim(_Order) & "', '" & Trim(_Type) & "','" & Trim(_Shade) & "','" & Trim(_Graige) & "','" & _Status & "','" & Trim(_Material) & "','" & _YarnCnt & "','" & _Res_Date & "','" & _Buyer & "','" & _Customer & "','" & _Shade_Cat & "','" & _Quality & "','" & _Criticle & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _Quality = ""
                _Order = ""
                _Shade = ""
                _Shade_Cat = ""
                '_LineItem = ""
                _Status = ""
                _Buyer = ""
                _Criticle = ""
                _Customer = ""
                _Graige = ""
                _Merchant = ""


                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 80
            lblPro.Text = lblPro.Text & ",R_code.txt"
            pbCount.Refresh()
            lblPro.Refresh()
            '  MsgBox("Files upload suceesfully...", MsgBoxStyle.Information, "Technova .......")
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_BomWst()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _Material As String
        Dim _Order As String
        Dim _Type As String
        Dim _Shade As String
        Dim _BQ As Double
        Dim _WQ As Double
        Dim _Diff As Double
        Dim _YarnCnt As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Res_Date As String


        Dim _Buyer As String
        Dim _Customer As String
        Dim _Graige As String

        Dim _Criticle As String

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
        Dim _cF As Double
        Dim _Material_Type As String
        Dim _Compo As Integer

        Try
            'nvcFieldList1 = "delete from M21Stock_Grige"
            '  DBEngin.ExecuteScalar(connection, transaction, "up_GetSetStock_Grige", New SqlParameter("@cQryType", "DEL"))

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\BOM_WASTAGE.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 36994 Then
                    'MsgBox("")
                End If


                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then

                _Material = (Trim(fields(0)))
                _BQ = (Trim(fields(1)))
                _WQ = (Trim(fields(2)))
                _cF = (Trim(fields(4)))
                _Material_Type = (Trim(fields(5)))
                _Compo = (Trim(fields(6)))

                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))
               
                _Diff = (_WQ - _BQ) / _WQ

                nvcFieldList1 = "select * from M24Zbomwaste where M24Material='" & Trim(_Material) & "'  "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into M24Zbomwaste(M24Material,M24Btc_Qty,M24Wast_Qty,M24WST,M24Con,M24MT_Type,M24Componant)" & _
                                                        " values('" & Trim(_Material) & "', '" & Trim(_BQ) & "','" & Trim(_WQ) & "','" & Trim(_Diff) & "','" & _cF & "','" & _Material_Type & "','" & _Compo & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _Material = ""
                _BQ = 0
                _WQ = 0
                _Diff = 0

                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 100
            lblPro.Text = lblPro.Text & ",zbomwaste.txt"
            pbCount.Refresh()
            lblPro.Refresh()
            MsgBox("Files upload suceesfully...", MsgBoxStyle.Information, "Technova .......")
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_TecSpec()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _Quality As String
        Dim _MachineType As String
        Dim _Yarn As String
        Dim _Fabric_Weight As Double
        Dim _Userble_Width As Double
        Dim _Strich As Double
        Dim _Customer As String
        Dim _CustomerName As String
        Dim _Department As String
        Dim _Merchnat As String
        Dim _Del_Date As Date
        Dim _Order_Qty As Double
        Dim _Del_Qty As Double
        Dim _Confirm_Qty As Double
        Dim _FGStock As Double
        Dim _Balance As Double
        Dim _Location As String
        Dim _PRD_Qty As String
        Dim _Grg_Qty As Double
        Dim _NCComment As String
        Dim _Awaiting As String
        Dim _depComm As String
        Dim _Comm2 As String
        Dim _20Cls As String
        Dim _OTDStatus As String
        Dim _PRD_OrderQty As Double
        Dim _RollNo As String
        Dim _Fabric_type As String
        Dim _YarnConsu As Double
        Dim _Needls As Integer
        Dim _Feedes As Integer
        Dim _YNCount As Integer
        Dim _PRODUCT As String
        Dim _MClass As String
        Dim _KgHR As Double

        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim _Confact As Double

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
            'nvcFieldList1 = "delete from M21Stock_Grige"
            '  DBEngin.ExecuteScalar(connection, transaction, "up_GetSetStock_Grige", New SqlParameter("@cQryType", "DEL"))

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\TechSpechData.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                'If X11 = 20882 Then
                '    MsgBox("")
                'End If


                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then

                _Quality = (Trim(fields(0)))
                If Microsoft.VisualBasic.Left(_Quality, 1) = "Q" Then

                    _Quality = Microsoft.VisualBasic.Right(_Quality, Microsoft.VisualBasic.Len(_Quality) - 1)
                End If
                _MachineType = (Trim(fields(2)))
                _Fabric_type = (Trim(fields(1)))

                _Yarn = Trim(fields(5))
                _Fabric_Weight = Trim(fields(9))
                If IsNumeric(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Yarn, 5), 1)) Then
                    _YNCount = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Yarn, 5), 3)

                Else
                    If IsNumeric(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Yarn, 4), 1)) Then

                        _YNCount = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Yarn, 4), 2)
                    Else
                        If IsNumeric(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Yarn, 3), 1)) Then
                            _YNCount = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Yarn, 3), 2)
                        End If
                    End If
                End If
                _Userble_Width = Trim(fields(10))
                If Trim(fields(11)) <> "" Then

                    _Strich = Trim(fields(11))
                End If
                _Confact = Trim(fields(3))
                If Trim(fields(6)) <> "" Then
                    _YarnConsu = Trim(fields(6))
                Else
                    _YarnConsu = 0
                End If
                If Trim(fields(13)) <> "" Then
                    _Needls = Trim(fields(13))
                Else
                    _Needls = 0
                End If
                If Trim(fields(14)) <> "" Then
                    _Feedes = Trim(fields(14))
                Else
                    _Feedes = 0
                End If
                _PRODUCT = Trim(fields(7))
                _MClass = Trim(fields(16))
                _KgHR = Trim(fields(4))

                nvcFieldList1 = "select * from M22Tec_Spec where M22Quality='" & Trim(_Quality) & "' and M22Machine_Type='" & Trim(_MachineType) & "' and M22Fabric_Weight='" & _Fabric_Weight & "' and M22Userble_Width='" & _Userble_Width & "' and M22Strich_Lenth='" & _Strich & "' and M22Yarn_Cons='" & _YarnConsu & "' and M22Yarn='" & _Yarn & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into M22Tec_Spec(M22Quality,M22Machine_Type,M22Yarn,M22Fabric_Weight,M22Userble_Width,M22Strich_Lenth,M22Fabric_Type,M22Con_Fact,M22Yarn_Cons,M22Needles,M22Feeds,m22Yarn_Cnt,M22Product_Type,M22M_Class,M22Kg_Hr)" & _
                                                        " values('" & Trim(_Quality) & "', '" & Trim(_MachineType) & "','" & Trim(_Yarn) & "','" & Trim(_Fabric_Weight) & "','" & _Userble_Width & "','" & Trim(_Strich) & "','" & _Fabric_type & "','" & _Confact & "','" & _YarnConsu & "'," & _Needls & "," & _Feedes & "," & _YNCount & ",'" & _PRODUCT & "','" & _MClass & "','" & _KgHR & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _Quality = ""
                _Yarn = ""
                _MachineType = ""
                _Fabric_type = ""
                '_LineItem = ""
                _Fabric_Weight = 0
                _Userble_Width = 0
                _Strich = 0
                _YarnConsu = 0
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 48
            lblPro.Text = lblPro.Text & ",TechSpechData.txt"
            pbCount.Refresh()
            lblPro.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Function Upload_zgreigeshade()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim _Material As String
        Dim _Material_Dis As String
        Dim _Len As String
        Dim _Shade As String
        

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
            'nvcFieldList1 = "delete from M21Stock_Grige"

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\MATERIAL_SHADES.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim _Lenth As Integer

                Dim fields() As String = line.Split(vbTab)
                'If X11 = 50 Then
                '    MsgBox("")
                'End If


                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then

                _Material = CInt(Trim(fields(0)))
                _Material_Dis = "-" '(Trim(fields(1)))
                _Len = Microsoft.VisualBasic.Len(_Material)
                _Material = Microsoft.VisualBasic.Left(_Material, 2) & "-" & Microsoft.VisualBasic.Right(_Material, _Len - 2)
                _Len = "-" ' Trim(fields(2))
                _Shade = Trim(fields(1))

                If _Shade = "MARL" Then
                ElseIf _Shade = "SPCL" Then
                    _Shade = "SP"
                Else
                    _Shade = Microsoft.VisualBasic.Left(_Shade, 1)
                End If
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Material_Dis = (Replace(_Material_Dis, characterToRemove, ""))



                If Microsoft.VisualBasic.Left(_Material, 2) = "20" Then
                    nvcFieldList1 = "select * from M23zgreigeshade where M23Material='" & Trim(_Material) & "'  "
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then

                        nvcFieldList1 = "update M23zgreigeshade set M23Shade='" & Trim(_Shade) & "' where M23Material='" & Trim(_Material) & "'  "
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into M23zgreigeshade(M23Material,M23Discription,M23Language,M23Shade)" & _
                                                            " values('" & Trim(_Material) & "', '" & Trim(_Material_Dis) & "','" & Trim(_Len) & "','" & Trim(_Shade) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                Else
                    nvcFieldList1 = "select * from M34Yarn_Shade where M34Class10='" & Trim(_Material) & "'  "
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then

                        nvcFieldList1 = "update M34Yarn_Shade  set M34Shade='" & Trim(_Shade) & "' where M34Class10='" & Trim(_Material) & "'  "
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into M34Yarn_Shade (M34Class10,M34Description,M34Shade)" & _
                                                            " values('" & Trim(_Material) & "', '" & Trim(_Material_Dis) & "','" & Trim(_Shade) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                End If
                ' pbCount.Value = pbCount.Value + 1


                '  lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))
                _Material = ""
                _Material_Dis = ""
                _Len = ""
                '_LineItem = ""
                _Shade = ""


                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            pbCount.Value = 64
            lblPro.Text = lblPro.Text & ",zgreigeshade.txt"
            pbCount.Refresh()
            lblPro.Refresh()



        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub chkUpload_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpload.CheckedChanged
        If chkUpload.Checked = True Then
            Call Upload_QulityRCode()
            Call Upload_File()
            Call Upload_Grige()
            Call Upload_TecSpec()
            Call Upload_zgreigeshade()
            Call Upload_Rcode()
            Call Upload_YarnPO()
            Call Upload_BomWst()

        End If
    End Sub

    Private Sub chk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk2.CheckedChanged

    End Sub

    Private Sub frmNo_Orders_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Quality()
        Call Load_Material()
        Call Load_BU()
        Call Load_Merchant()

    End Sub

    Function Load_Quality()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M21Material as [Quality] from M21Stock_Grige group by M21Material"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboQuality
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

    Function Load_Material()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M2120Class as [20Class] from M21Stock_Grige order by M2120Class"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboMaterial
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


    Function Load_Merchant()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M13Merchant as [Merchant] from M13Biz_Unit"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboMerchant
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

    Function Load_BU()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M14Name as [BU] from M14Retailer"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboBU
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

    Private Sub cboQuality_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboQuality.InitializeLayout

    End Sub
End Class