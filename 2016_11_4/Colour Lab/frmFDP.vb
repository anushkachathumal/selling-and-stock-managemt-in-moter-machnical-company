Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Public Class frmFDP
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableFDP
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 270
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 90
       
        End With
    End Function

    Private Sub frmFDP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Try
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\FDP.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                MTNo = Trim(fields(0))
                _Plant = Trim(fields(1))
                _StockLoc = Trim(fields(2))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("SAP Code") = CInt(Trim(fields(0)))   '0
                newRow("Description") = (Trim(fields(1))) '1
                newRow("Week") = Trim(fields(2)) '3
                newRow("Qty") = Trim(fields(3)) '4
               
                c_dataCustomer1.Rows.Add(newRow)

                ' pbCount.Value = pbCount.Value + 1
                lblDis.Text = Trim(fields(0)) & "-" & Trim(fields(1))
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

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Call Load_Gride()
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
        Dim p_Date As String
        Dim t_Date As String
        Dim ETD As String
        Dim ETA As String
        Dim X11 As Integer

        Dim Category As String
        Dim SAPCode As String
        Dim Dis As String
        Dim SS As String
        Dim EndStock As Double
        Dim Avg As Double
        Dim M6 As Double
        Dim N3 As Double
        Dim CQ_MRS As Double
        Dim Last_MntQty As Double
        Dim Last_30Day As Double
        Dim SD As Double
        Dim L_Week2nd As Double
        Dim L_week2 As Double
        Dim L_30Day2 As Double
        Dim SD2 As Double
        Dim L_Week2nd2 As Double
        Dim L_Week3 As Double
        Dim L_14Day As Double
        Dim RL As Double
        Dim RQ As Double
        Dim LT_day As Double
        Dim Pending_PO As String
        Dim PO As String
        Dim ItemNo As String
        Dim L14Day_Con As Double
        Dim Qty As Double
        Dim _Week As Integer
        Dim _Year As Integer
        Dim _Year1 As String


        Try
            Me.Cursor = Cursors.WaitCursor
            nvcFieldList1 = "delete from  M12FDP"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\FDP.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                SAPCode = CInt(Trim(fields(0))) '1
                Dis = Trim(fields(1))
                _Year1 = Microsoft.VisualBasic.Left(Trim(fields(2)), 9) '3
                _Year = Microsoft.VisualBasic.Right(_Year1, 4)
                _Week = Microsoft.VisualBasic.Right(Trim(fields(2)), 2)
                Qty = Trim(fields(3)) '5
               
               
                nvcFieldList1 = "select * from M12FDP where M12SAPCode='" & CInt(Trim(fields(0))) & "' and M12Week='" & _Week & "' and M12Year='" & _Year & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                Else
                   
                  

                    nvcFieldList1 = "Insert Into M12FDP(M12SAPCode,M12Dis,M12Week,M12Year,M12Qty)" & _
                                                             " values('" & SAPCode & "', '" & Dis & "','" & _Week & "','" & _Year & "','" & Qty & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                Qty = 0
                _Year = 0
                _Week = 0
                Dis = ""
                SAPCode = ""



                X11 = X11 + 1
            Next
            MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            Me.Cursor = Cursors.Arrow
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            ' MsgBox(X11)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Sub
End Class