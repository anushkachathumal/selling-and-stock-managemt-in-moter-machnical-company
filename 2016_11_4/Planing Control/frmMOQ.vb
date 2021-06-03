
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader

Public Class frmMOQ
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmMOQ_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableQuality
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
     

        Dim nvcFieldList1 As String

      
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
        Dim _quality As String
        Dim _Qty As Double

        Try


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\MOQ.txt"
            pbCount.Maximum = System.IO.File.ReadAllLines(strFileName).Length
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 10 Then
                    '   MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then
                _quality = Trim(fields(0))
                _Qty = CInt(Trim(fields(1)))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Quality No") = _quality
                newRow("Quantity") = _Qty
                c_dataCustomer1.Rows.Add(newRow)

                pbCount.Value = pbCount.Value + 1
                lblPro.Text = "Delsum.txt"
                lblPro.Refresh()
                pbCount.Refresh()
                I = I + 1
                'cmdEdit.Enabled = True
            Next
           

        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            'DBEngin.CloseConnection(connection)
            'connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            'DBEngin.CloseConnection(connection)
            'connection.ConnectionString = ""
            'connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub UltraButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton7.Click
        pbCount.Value = 0
        lblPro.Text = ""
        Call Upload_File()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            Dim i As Integer

            For Each uRow As UltraGridRow In UltraGrid1.Rows
                nvcFieldList1 = "select * from M31MOQ where M31Quality='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "'"
                dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(dsUser) Then

                    nvcFieldList1 = "update M31MOQ set M31Qty='" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "' where M31Quality='" & UltraGrid1.Rows(i).Cells(0).Text & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M31MOQ(M31Quality,M31Qty)" & _
                                                                      " values('" & UltraGrid1.Rows(i).Cells(0).Text & "', '" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If
                i = i + 1
            Next
            MsgBox("Data update successfully", MsgBoxStyle.Information, "Information ......")
            transaction.Commit()

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try

    End Sub
End Class