Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Public Class frmUploadPlaning
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTablePlaning
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

    Private Sub frmUploadPlaning_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Upload_File()

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
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

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\PlaningSAP.CSV"
            Dim theTextFieldParser As FileIO.TextFieldParser

            theTextFieldParser = My.Computer.FileSystem.OpenTextFieldParser(strFileName)
            theTextFieldParser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            theTextFieldParser.Delimiters = New String() {","}
            Dim currentRow() As String
            While Not theTextFieldParser.EndOfData
                currentRow = theTextFieldParser.ReadFields()

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                ' Use the currentField variable to loop
                ' through fields in the currentRow.
                i = 0
                If X11 = 2051 Then
                    '   MsgBox("")
                End If

                For Each currentField In currentRow




                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    If i = 0 Then
                        PONo = currentField
                    ElseIf i = 1 Then
                        _Sold = currentField
                    ElseIf i = 2 Then
                        _Sales_Doc = currentField
                    ElseIf i = 3 Then
                        _Item = currentField
                    ElseIf i = 4 Then
                        _Meterial = currentField
                    ElseIf i = 5 Then
                        _Dis = currentField
                    ElseIf i = 6 Then
                        _Create = currentField
                    ElseIf i = 7 Then
                        _CreateBy = currentField
                    ElseIf i = 8 Then
                        _Dilivary = currentField
                    ElseIf i = 9 Then
                        _Qty = currentField
                    ElseIf i = 10 Then
                        _Su = currentField
                    ElseIf i = 11 Then
                        _Reject = currentField
                    ElseIf i = 12 Then
                        _Db = currentField
                    ElseIf i = 13 Then
                        _DS = currentField
                    End If
                 

                    i = i + 1
                Next

            

                If Microsoft.VisualBasic.Left(_CreateBy, 3) = "JKH" Then
                Else
                    If Trim(_Reject) = "7" And Trim(_Db) = "A" Then
                    Else
                        If Microsoft.VisualBasic.Left(_Sales_Doc, 1) = "1" Then

                            If Microsoft.VisualBasic.Left(_Dis, 1) = "Y" Then
                                nvcFieldList1 = "select * from M15Convension where M15Code='" & Microsoft.VisualBasic.Left(_Dis, 8) & "'"
                                T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                                If isValidDataset(T01) Then
                                    _Convenshion = T01.Tables(0).Rows(0)("M15Amount")
                                End If
                            Else
                                If IsNumeric(Microsoft.VisualBasic.Left(_Dis, 1)) Then
                                    If Microsoft.VisualBasic.Right(_Dis, 6) = "L" Or Microsoft.VisualBasic.Right(_Dis, 6) = "D" Then
                                        nvcFieldList1 = "select * from M15Convension where M15Code='" & Microsoft.VisualBasic.Left(_Dis, 8) & "'"
                                        T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                                        If isValidDataset(T01) Then
                                            _Convenshion = T01.Tables(0).Rows(0)("M15Amount")
                                        End If
                                    Else
                                        nvcFieldList1 = "select * from M15Convension where M15Code='" & Microsoft.VisualBasic.Left(_Dis, 5) & "'"
                                        T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                                        If isValidDataset(T01) Then
                                            _Convenshion = T01.Tables(0).Rows(0)("M15Amount")
                                        End If
                                    End If
                                End If
                            End If
                            nvcFieldList1 = "select * from M14Planing where M14Ref='" & _Sales_Doc & "' and M14Item='" & _Item & "' and M14Delevary='" & _Dilivary & "' "
                            dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                            If isValidDataset(dsUser) Then
                            Else
                                _WeekNo = DatePart(DateInterval.WeekOfYear, _Create)
                                _Dis = Microsoft.VisualBasic.Left(_Dis, 20) & "-" & Microsoft.VisualBasic.Right(_Dis, 5)
                                nvcFieldList1 = "Insert Into M14Planing(M14PO,M14Sold,M14Ref,M14Item,M14Material,M14Dis,M14CreateOn,M14Delevary,M14CreateBy,M14Qty,M14Su,M14Rej,M14DB,M14DS,M14Week,M14Year,M14QtyKg)" & _
                                                     " values('" & PONo & "', '" & _Sold & "'," & _Sales_Doc & ",'" & _Item & "','" & _Meterial & "','" & _Dis & "','" & _Create & "','" & _Dilivary & "','" & _CreateBy & "','" & _Qty & "','" & _Su & "','" & _Reject & "','" & _Db & "','" & _DS & "','" & _WeekNo & "','" & Year(_Create) & "','" & _Qty / _Convenshion & "')"
                                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            End If
                        End If
                        End If
                    End If


                PONo = ""
                _Sold = ""
                _Sales_Doc = ""
                _Dis = ""
                _Item = ""
                _Meterial = ""
                '  _Create = ""
                _Dilivary = ""
                _Qty = 0
                _CreateBy = ""
                _Su = ""
                _Reject = ""
                _Db = ""
                _DS = ""
                X11 = X11 + 1
            End While

            MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            Me.Cursor = Cursors.Arrow
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            '  MsgBox(X11)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim PONo As String
        Dim _Sold As String
        Dim _Sales_Doc As String
        Dim _Dis As String
        Dim _Item As String
        Dim _Meterial As String
        Dim _Create As String
        Dim _Dilivary As String
        Dim _Qty As Double
        Dim _CreateBy As String
        Dim _Su As String
        Dim _Reject As String
        Dim _Db As String
        Dim _DS As String

        Try
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\PlaningSAP.CSV"
            Dim theTextFieldParser As FileIO.TextFieldParser

            theTextFieldParser = My.Computer.FileSystem.OpenTextFieldParser(strFileName)
            theTextFieldParser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            theTextFieldParser.Delimiters = New String() {","}
            Dim currentRow() As String
            While Not theTextFieldParser.EndOfData
                currentRow = theTextFieldParser.ReadFields()

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                ' Use the currentField variable to loop
                ' through fields in the currentRow.
                i = 0

             
                For Each currentField In currentRow
                  



                    'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    If i = 0 Then
                        PONo = currentField
                    ElseIf i = 1 Then
                        _Sold = currentField
                    ElseIf i = 2 Then
                        _Sales_Doc = currentField
                    ElseIf i = 3 Then
                       _item = currentField
                    ElseIf i = 4 Then
                        _Meterial = currentField
                    ElseIf i = 5 Then
                        _Dis = currentField
                    ElseIf i = 6 Then
                        _Create = currentField
                    ElseIf i = 7 Then
                        _CreateBy = currentField
                    ElseIf i = 8 Then
                        _Dilivary = currentField
                    ElseIf i = 9 Then
                        _Qty = currentField
                    ElseIf i = 10 Then
                        _Su = currentField
                    ElseIf i = 11 Then
                        _Reject = currentField
                    ElseIf i = 12 Then
                        _Db = currentField
                    ElseIf i = 13 Then
                        _DS = currentField
                    End If




                    ' pbCount.Value = pbCount.Value + 1
                    '  lblDis.Text = Trim(fields(0)) & "-" & Trim(fields(4))
                    cmdEdit.Enabled = True
                    i = i + 1
                Next
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("P/O No") = PONo
                newRow("Sold-to pt") = _Sold
                newRow("Sales Doc") = _Sales_Doc
                newRow("Item") = _Item
                newRow("Material") = _Meterial
                newRow("Description") = _Dis
                newRow("Created On") = _Create
                newRow("Created by") = _CreateBy
                newRow("Delivary Date") = _Dilivary
                newRow("Qty") = _Qty
                newRow("Su") = _Su
                newRow("Reject") = _Reject
                newRow("DB") = _Db
                newRow("DS") = _DS



                c_dataCustomer1.Rows.Add(newRow)

            End While
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            '  MsgBox(X11)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

End Class