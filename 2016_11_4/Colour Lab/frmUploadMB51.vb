Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Public Class frmUploadMB51
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableMB51
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


    Private Sub frmUploadMB51_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Try
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\MB51.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                MTNo = Trim(fields(0))
                _Plant = Trim(fields(1))
                _StockLoc = Trim(fields(2))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Material No") = CInt(Trim(fields(0)))
                newRow("Plant") = Trim(fields(1))
                newRow("Stock Loc") = Trim(fields(2))
                newRow("Movement Type") = Trim(fields(3))
                newRow("No of Mat Doc") = Trim(fields(4))
                newRow("Posting Date") = Trim(fields(5))
                newRow("Quantity") = Trim(fields(6))
                newRow("Unit") = Trim(fields(7))
                newRow("Batch No") = Trim(fields(8))
                newRow("PO No") = Trim(fields(9))
                newRow("DOADE") = Trim(fields(10))


                c_dataCustomer1.Rows.Add(newRow)

                ' pbCount.Value = pbCount.Value + 1
                lblDis.Text = Trim(fields(0)) & "-" & Trim(fields(4))
                cmdEdit.Enabled = True
            Next
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

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Upload_File()

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String

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
        Dim t_Date As Date

        Try
            Me.Cursor = Cursors.WaitCursor

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\MB51.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                MTNo = Trim(fields(0))
                _Plant = Trim(fields(1))
                _StockLoc = Trim(fields(2))
                nvcFieldList1 = "select * from MB51 where MtrNo='" & CInt(Trim(fields(0))) & "' and NoofDoc='" & Trim(fields(4)) & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                Else

                    p_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(5)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(5)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(5)), 4)
                    t_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(10)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(10)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(10)), 4)


                    nvcFieldList1 = "Insert Into MB51(MtrNo,Plant,Stock_Loc,MovementType,NoofDoc,PDOD,Qty,Unit,BatchNo,PONo,DOA)" & _
                                                             " values('" & CInt(Trim(fields(0))) & "', " & Trim(fields(1)) & "," & Trim(fields(2)) & ",'" & Trim(fields(3)) & "','" & Trim(fields(4)) & "','" & p_Date & "','" & Trim(fields(6)) & "','" & Trim(fields(7)) & "','" & Trim(fields(8)) & "','" & Trim(fields(9)) & "','" & t_Date & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If


            Next
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

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

    End Sub
End Class