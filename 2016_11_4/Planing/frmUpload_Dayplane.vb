Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Imports Microsoft.VisualBasic.FileIO
Imports System.IO

Public Class frmUpload_Dayplane
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableWIP
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(7).Width = 80

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub frmUpload_Dayplane_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Dim X11 As Integer

        Try
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\ZPL_ORDER.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                If Microsoft.VisualBasic.Left((Trim(fields(0))), 5) = "00000" Then
                    newRow("Batch No") = CInt(Trim(fields(0)))   '0
                Else
                    newRow("Batch No") = (Trim(fields(0)))
                End If
                newRow("Customer") = (Trim(fields(1))) '1
                newRow("Material") = Trim(fields(2)) '3
                newRow("Description") = Trim(fields(3)) '4
                newRow("Delivary Date") = Trim(fields(4)) '5
                newRow("Last Confirmation Date") = Trim(fields(5)) '6
                newRow("Qty (Kg)") = Trim(fields(6)) '7
                newRow("Next Oparation") = Trim(fields(7)) '8
                newRow("Planing Comment") = Trim(fields(8)) '9
                newRow("Order Type") = Trim(fields(9)) '10
                newRow("Sales Order") = Trim(fields(10)) '11
                newRow("Line Item") = Trim(fields(11)) '11
                newRow("Qty (Mtr)") = Trim(fields(12)) '12
                newRow("Merchant") = Trim(fields(13)) '13

                c_dataCustomer1.Rows.Add(newRow)

                X11 = X11 + 1
                ' pbCount.Value = pbCount.Value + 1
                lblDis.Text = Trim(fields(0)) & "-" & Trim(fields(4))
                cmdEdit.Enabled = True
            Next
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox("Error Record in txt File-Line -" & X11, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Load_Gride()
        Call Upload_File()

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
       

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
        Dim p_Date As Date
        Dim _Convenshion As Double
        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim T03 As DataSet

        Dim _BatchNo As String
        Dim _Customer As String
        Dim _Material As String
        Dim _Dis As String
        Dim _DDate As Date
        Dim _LCDate As Date
        Dim _QtyKG As Double
        Dim _NextOP As String
        Dim _PLCom As String
        Dim _OrderType As String
        Dim _SalesOrder As String
        Dim _LineItem As String
        Dim _QtyMtr As Double
        Dim _Merchant As String


        Dim t_Date As Date
        Dim _WeekNo As Integer
        Dim X11 As Integer
        Dim Y As Integer
        Dim _Status As Boolean
        Dim nvNo As Integer


        Try
            Me.Cursor = Cursors.WaitCursor
            nvNo = 0
            'Set M18WIP RefNo
            nvcFieldList1 = "select * from P02Parameter where P02Date='" & Today & "' and P02Code='DL'"
            T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(T02) Then
                nvNo = T02.Tables(0).Rows(0)("P02No")
                nvcFieldList1 = "update P02Parameter set P02No=P02No +" & 1 & " where P02Date='" & Today & "' and P02Code='DL'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvNo = "1"
                nvcFieldList1 = "Insert Into P02Parameter(P02No,P02Date,P02Code)" & _
                                                        " values('" & nvNo & "', '" & Today & "','DL')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            nvcFieldList1 = "update M18WIP set M18Status='N'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\ZPL_ORDER.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)



                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                If Microsoft.VisualBasic.Left((Trim(fields(0))), 5) = "00000" Then
                    _BatchNo = CInt(Trim(fields(0)))   '0
                Else
                    _BatchNo = (Trim(fields(0)))
                End If
                _Customer = (Trim(fields(1))) '1
                _Material = Trim(fields(2)) '3
                _Dis = Trim(fields(3)) '4
                _DDate = Microsoft.VisualBasic.Left(Trim(fields(4)), 4) & "/" & Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(4)), 6), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(4)), 2)
                If Microsoft.VisualBasic.Left(Trim(fields(5)), 2) = "00" Then
                    _LCDate = "1900/1/1"
                Else
                    _LCDate = Microsoft.VisualBasic.Left(Trim(fields(5)), 4) & "/" & Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(5)), 6), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(5)), 2)
                End If
                _QtyKG = Trim(fields(6)) '7
                _NextOP = Trim(fields(7)) '8
                'If Microsoft.VisualBasic.Left(Trim(fields(8)), 6) = "?Don't" Then
                '    _PLCom = "?Dont " & Microsoft.VisualBasic.Right(Trim(fields(8)), Microsoft.VisualBasic.Len(Trim(fields(8))) - 6)
                'Else
                'For Y = 1 To Microsoft.VisualBasic.Len(Trim(fields(8)))
                '    'MsgBox(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(8)), Y), 1))
                '    If Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(8)), Y), 1) = "'" Then
                '        Exit For
                '    Else
                '        _PLCom = Microsoft.VisualBasic.Left(Trim(fields(8)), Y)
                '    End If
                'Next
                Dim stringToCleanUp As String
                Dim characterToRemove As String

                stringToCleanUp = Trim(fields(8))
                characterToRemove = "'"
                _PLCom = Replace(stringToCleanUp, characterToRemove, "")
                ' _PLCom = Microsoft.VisualBasic.Left(Trim(fields(8)), Y - 1) & Microsoft.VisualBasic.Right(Trim(fields(8)), Microsoft.VisualBasic.Len(Trim(fields(8))) - (Y - 1)) '9
                ' End If
                _OrderType = Trim(fields(9)) '10
                _SalesOrder = Trim(fields(10)) '11
                _LineItem = Trim(fields(11)) '11
                _QtyMtr = Trim(fields(12)) '12
                _Merchant = Trim(fields(13)) '13

               
                nvcFieldList1 = "select * from M16Meterial where Dis='" & CInt(Trim(_Material)) & "'"
                T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(T01) Then

                Else

                    Y = 0
                    _Status = False
                    nvcFieldList1 = "select * from M17Planing_Comment "
                    T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    For Each DTRow4 As DataRow In T02.Tables(0).Rows
                        ' MsgBox(UCase(Microsoft.VisualBasic.Left(_PLCom, T02.Tables(0).Rows(Y)("M17Lenth"))))
                        If UCase(Microsoft.VisualBasic.Left(_PLCom, T02.Tables(0).Rows(Y)("M17Lenth"))) = Trim(T02.Tables(0).Rows(Y)("M17Dis")) Then
                            _Status = True
                            Exit For
                        Else

                        End If
                        Y = Y + 1
                    Next

                    If _Status = False Then

                        nvcFieldList1 = "select * from M18WIP where M18Batch='" & _BatchNo & "' and M18Status='N' and M18No=" & nvNo & ""
                        T03 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(T03) Then
                            '_WeekNo = DatePart(DateInterval.WeekOfYear, _DDate)
                            'nvcFieldList1 = "Insert Into M18WIP(M18Batch,M18Customer,M18Material,M18Dis,M18DDate,M18LCDate,M18QtyKG,M18NextOparation,M18PComment,M18OrderType,M18SalesOrder,M18LineItem,M18Qty,M18Merchant,M18Week,M18Year,M18Status)" & _
                            '                            " values('" & _BatchNo & "', '" & _Customer & "'," & _Material & ",'" & _Dis & "','" & _DDate & "','" & _LCDate & "','" & _QtyKG & "','" & _NextOP & "','" & _PLCom & "','" & _OrderType & "','" & _SalesOrder & "','" & _LineItem & "','" & _QtyMtr & "','" & _Merchant & "'," & _WeekNo & "," & Year(_DDate) & ",'Y' )"
                            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        Else
                            _WeekNo = DatePart(DateInterval.WeekOfYear, _DDate)
                            ' _LCDate = "1981/1/1"
                            nvcFieldList1 = "Insert Into M18WIP(M18Batch,M18Customer,M18Material,M18Dis,M18DDate,M18LCDate,M18QtyKG,M18NextOparation,M18PComment,M18OrderType,M18SalesOrder,M18LineItem,M18Qty,M18Merchant,M18Week,M18Year,M18Status,M18No)" & _
                                                        " values('" & _BatchNo & "', '" & _Customer & "'," & _Material & ",'" & _Dis & "','" & Microsoft.VisualBasic.Format(_DDate, "MM/dd/yyyy") & "','" & Microsoft.VisualBasic.Format(_LCDate, "MM/dd/yyyy") & "','" & _QtyKG & "','" & _NextOP & "','" & _PLCom & "','" & _OrderType & "','" & _SalesOrder & "','" & _LineItem & "','" & _QtyMtr & "','" & _Merchant & "'," & _WeekNo & "," & Year(_DDate) & ",'Y'," & nvNo & " )"

                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        End If

                        nvcFieldList1 = "select * from M19Segrigrade where M19Dis='" & _NextOP & "'"
                        T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(T01) Then
                        Else
                            strFileName = ConfigurationManager.AppSettings("UploadPath") + "\segrigrade" & Year(Today) & Month(Today) & Microsoft.VisualBasic.Day(Today) & ".txt"
                            FileOpen(1, strFileName, OpenMode.Append)

                            FileClose(1)

                            Dim TargetFile As StreamWriter
                            TargetFile = New StreamWriter(strFileName, True)
                            TargetFile.Write(_NextOP & vbTab)
                            TargetFile.WriteLine()
                            TargetFile.Close()
                        End If
                    End If
                End If



                _BatchNo = ""
                _Customer = ""
                _Material = ""
                _Dis = ""

                _QtyKG = 0
                _NextOP = ""
                _PLCom = ""
                _OrderType = ""
                _SalesOrder = ""
                _LineItem = ""
                _QtyMtr = 0
                _Merchant = ""

                X11 = X11 + 1
                ' pbCount.Value = pbCount.Value + 1
                lblDis.Text = Trim(fields(0)) & "-" & Trim(fields(0))
                cmdEdit.Enabled = True
            Next
            '  MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            Call Upload_BLOCKSTOCK()
            Call Upload_NC()


        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox("Error Record in txt File-Line -" & X11, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Sub

    Function Upload_NC()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Dim PONo As String
        Dim _Sold As String
        Dim _Sales_Doc As String
        Dim _Dis As String
        Dim _Item As String
        Dim _Meterial As String
        Dim _Create As Date
        Dim _Dilivary As String
        Dim _Qty As Double
        Dim _CreateBy As String
        Dim _Su As String
        Dim _Reject As String
        Dim _Db As String
        Dim _DS As String

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
        Dim p_Date As Date
        Dim _Convenshion As Double
        Dim T01 As DataSet

        Dim t_Date As Date
        Dim _WeekNo As Integer
        Dim X11 As Integer

        Try
            Me.Cursor = Cursors.WaitCursor


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\NC.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                ' Use the currentField variable to loop
                ' through fields in the currentRow.



                nvcFieldList1 = "select * from M18WIP where M18Batch='" & Trim(fields(0)) & "' "
                dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(dsUser) Then
                    nvcFieldList1 = "update M18WIP set M18NonCon='" & Trim(fields(1)) & "',M18NonConDis='" & Trim(fields(2)) & "' where M18Batch='" & Trim(fields(0)) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else

                End If



                X11 = X11 + 1
            Next

            MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Me.Cursor = Cursors.Arrow
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox("Error Record in txt File-Line -" & X11, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function


    Function Upload_BLOCKSTOCK()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Dim PONo As String
        Dim _Sold As String
        Dim _Sales_Doc As String
        Dim _Dis As String
        Dim _Item As String
        Dim _Meterial As String
        Dim _Create As Date
        Dim _Dilivary As String
        Dim _Qty As Double
        Dim _CreateBy As String
        Dim _Su As String
        Dim _Reject As String
        Dim _Db As String
        Dim _DS As String

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
        Dim p_Date As Date
        Dim _Convenshion As Double
        Dim T01 As DataSet

        Dim t_Date As Date
        Dim _WeekNo As Integer
        Dim X11 As Integer
        Dim T02 As DataSet
        Dim nvNo As String
        Dim _LCDate As Date

        Try
            Me.Cursor = Cursors.WaitCursor
            nvcFieldList1 = "select * from P02Parameter where P02Date='" & Today & "' and P02Code='BS'"
            T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(T02) Then
                nvNo = T02.Tables(0).Rows(0)("P02No")
                nvcFieldList1 = "update P02Parameter set P02No=P02No +" & 1 & " where P02Date='" & Today & "' and P02Code='BS'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvNo = "1"
                nvcFieldList1 = "Insert Into P02Parameter(P02No,P02Date,P02Code)" & _
                                                        " values('" & nvNo & "', '" & Today & "','BS')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If


            nvcFieldList1 = "update BLOCK_STOCK set Status='N'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\BLOCK_STOCK.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                Dim _DelivaryDate As Date

                ' Use the currentField variable to loop
                ' through fields in the currentRow.
                '  _DelivaryDate = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(Trim(fields(8)), 4), 2) & "\" & Microsoft.VisualBasic.Right(Trim(fields(8)), 2) & "\" & Microsoft.VisualBasic.Left(Trim(fields(8)), 4)

                If Microsoft.VisualBasic.Left(Trim(fields(8)), 2) = "00" Then
                    _DelivaryDate = "1900/1/1"
                Else
                    _DelivaryDate = Microsoft.VisualBasic.Left(Trim(fields(8)), 4) & "/" & Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(Trim(fields(8)), 4), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(8)), 2)
                End If

                If Microsoft.VisualBasic.Left(Trim(fields(9)), 2) = "00" Then
                    _LCDate = "1900/1/1"
                Else
                    _LCDate = Microsoft.VisualBasic.Left(Trim(fields(9)), 4) & "/" & Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(9)), 6), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(9)), 2)
                End If
                'If Trim(fields(7)).IndexOf("", 0, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                '    MessageBox.Show(Trim(fields(7)).Replace("""", ""))
                'End If

                ' MsgBox((Trim(fields(2))))
                If _DelivaryDate = "1900/1/1" Then
                Else

                    nvcFieldList1 = "select * from BLOCK_STOCK where Batch='" & (Trim(fields(3))) & "' AND Stock_Loc='" & Trim(fields(1)) & "' AND Sales_order='" & CInt(Trim(fields(2))) & "' AND LineItem='" & CInt(Trim(fields(4))) & "' AND Dilivary_Date='" & Microsoft.VisualBasic.Format(_DelivaryDate, "MM/dd/yyyy") & "' AND STATUS='Y'"
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(dsUser) Then
                        nvcFieldList1 = "update BLOCK_STOCK set Qty_Kg=Qty_Kg +" & Trim(fields(10)) & ",Qty_Mtr=Qty_Mtr +" & Trim(fields(11)) & " where Batch='" & (Trim(fields(3))) & "' AND Stock_Loc='" & Trim(fields(1)) & "' AND Sales_order='" & CInt(Trim(fields(2))) & "' AND LineItem='" & CInt(Trim(fields(4))) & "' AND Dilivary_Date='" & Microsoft.VisualBasic.Format(_DelivaryDate, "MM/dd/yyyy") & "' AND STATUS='Y'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into BLOCK_STOCK(Batch,Stock_Loc,Sales_order,LineItem,Customer,Material,Mat_Dis,Dilivary_Date,GRN_date,Qty_Kg,Qty_Mtr,Status,BNo)" & _
                                                           " values('" & (Trim(fields(3))) & "', '" & Trim(fields(1)) & "','" & CInt(Trim(fields(2))) & "','" & CInt(Trim(fields(4))) & "','" & Trim(fields(5)) & "','" & CInt(Trim(fields(6))) & "','" & Trim(fields(7)).Replace("""", "") & "','" & Microsoft.VisualBasic.Format(_DelivaryDate, "MM/dd/yyyy") & "','" & Microsoft.VisualBasic.Format(_LCDate, "MM/dd/yyyy") & "','" & Trim(fields(10)) & "','" & Trim(fields(11)) & "','Y'," & nvNo & ")"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If

                End If

                X11 = X11 + 1
            Next

            '   MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Me.Cursor = Cursors.Arrow
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox("Error Record in txt File-Line -" & X11, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Public Function convertQuotes(ByVal str As String) As String
        convertQuotes = str.Replace("'", "''")
    End Function
End Class