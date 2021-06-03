Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Public Class frmUploadMRS
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableMRS
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

    Private Sub frmUploadMRS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Call Load_Gride()
       

    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Dim X11 As Integer
        Try

            X11 = 0
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\MRS.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                MTNo = Trim(fields(0))
                _Plant = Trim(fields(1))
                _StockLoc = Trim(fields(2))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("SAP-Code") = CInt(Trim(fields(0)))
                newRow("Description") = Trim(fields(1))
                newRow("Prc-$/Kg") = Trim(fields(2))
                newRow("PkSiz") = Trim(fields(3))
                newRow("MR") = Trim(fields(4))
                newRow("Month") = Trim(fields(5))
                newRow("Value") = Trim(fields(6))
                newRow("12 MAvg") = Trim(fields(7))
                newRow("Hgt 6 MAvg") = Trim(fields(8))
                newRow("Hgt 3 MAvg") = Trim(fields(9))
                newRow("CQ-MRS") = Trim(fields(10))
                newRow("Purchase") = Trim(fields(11))
                newRow("NLTW") = Trim(fields(12))
                newRow("NLTD") = Trim(fields(13))
                newRow("SS") = Trim(fields(14))
                newRow("$SS") = Trim(fields(15))
                newRow("RL") = Trim(fields(16))
                newRow("$RL") = Trim(fields(17))
                newRow("RQ") = Trim(fields(18))
                newRow("$RQ") = Trim(fields(19))
                newRow("ML") = Trim(fields(20))
                newRow("End Stock") = Trim(fields(21))
                newRow("stk-Holding") = Trim(fields(22))
                newRow("Category") = Trim(fields(23))


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
            MsgBox(X11)
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

        Dim SAPCode As Integer
        Dim _Prc As String
        Dim PkSiz As String
        Dim MR As String
        Dim Value As String
        Dim n_MAvg As String
        Dim HgtMAvg As String
        Dim Hgt3MAvg As String
        Dim CQ_MRS As String
        Dim Purchase As String
        Dim NLTW As String
        Dim NLTD As String
        Dim SS As String
        Dim nSS As String
        Dim RL As String
        Dim nRL As String
        Dim RQ As String
        Dim nRQ As String
        Dim ML As String
        Dim EnStock As String
        Dim stk_Holding As String
        Dim Category As String
        Dim X11 As Integer
        Dim Discripsion As String
        Dim _SD As String
        Dim n_SD As String
        Dim _N3 As String
        Dim _RQLT As String
        Dim _RLLT As String
        Dim n_SS As String
        Dim n_RL As String
        Dim n_RQ As String
        Dim WH_Stock As String
        Dim TOD_CQ As String
        Dim TOD_N3 As String
        Dim Tot_Req As String
        Dim Old_PR As String
        Dim PO_Qty As String
        Dim NewPR As String
        Dim _DisFilePath As String

        X11 = 0
        Try
            Me.Cursor = Cursors.WaitCursor

            ' nvcFieldList1 = "Update M11MRS set M11Status='N'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\MRS.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 3692 Then
                    ' MsgBox("")
                End If
                SAPCode = Trim(fields(0))
                Discripsion = Trim(fields(1))
                _Prc = Trim(fields(2))
                PkSiz = Trim(fields(3))
                If Trim(PkSiz) <> "" Then
                Else
                    PkSiz = "0"
                End If

                MR = Trim(fields(4))
                '  newRow("Month") = Trim(fields(5))
                Value = Trim(fields(6))
                _SD = Trim(fields(7))
                n_MAvg = Trim(fields(8))
                n_SD = Trim(fields(9))
                HgtMAvg = Trim(fields(10))
                Hgt3MAvg = Trim(fields(11))
                _N3 = Trim(fields(12))
                CQ_MRS = Trim(fields(13))
                Purchase = Trim(fields(14))
                _RQLT = Trim(fields(15))
                _RLLT = Trim(fields(16))
                SS = Trim(fields(17))
                n_SS = Trim(fields(18))
                RL = Trim(fields(19))
                n_RL = Trim(fields(20))
                RQ = Trim(fields(21))
                n_RQ = Trim(fields(22))
                ML = Trim(fields(23))
                WH_Stock = Trim(fields(24))
                EnStock = Trim(fields(25))
                stk_Holding = Trim(fields(26))
                Category = Trim(fields(27))
                TOD_CQ = Trim(fields(28))
                TOD_N3 = Trim(fields(29))
                Tot_Req = Trim(fields(30))
                Old_PR = Trim(fields(31))
                PO_Qty = Trim(fields(32))
                NewPR = Trim(fields(33))

                ' MsgBox(Trim(fields(5)))
                nvcFieldList1 = "select * from M11MRS where M11SAPCode='" & CInt(Trim(fields(0))) & "' and M11Year=" & Microsoft.VisualBasic.Left(Trim(fields(5)), 4) & " and M11Month='" & Month(Trim(fields(5))) & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    nvcFieldList1 = "update M11MRS set M11MR='" & MR & "',M11Value='" & Value & "',M11SD='" & _SD & "',M11MAvg='" & n_MAvg & "',M11NSD='" & n_SD & "',M11HGT6='" & HgtMAvg & "',M11HGT3='" & Hgt3MAvg & "',M11N3='" & _N3 & "',M11CQ='" & CQ_MRS & "',M11Purchase='" & Purchase & "',M11RQLT='" & _RQLT & "' where M11SAPCode='" & CInt(Trim(fields(0))) & "' and M11Year=" & Microsoft.VisualBasic.Left(Trim(fields(5)), 4) & " and M11Month='" & Month(Trim(fields(5))) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "update M11MRS set M11StkH='" & stk_Holding & "',M11EndStock='" & EnStock & "',M11WHStock='" & WH_Stock & "',M11ML='" & ML & "',M11$RQ='" & n_RQ & "',M11RQ='" & RQ & "',M11$RL='" & n_RL & "',M11RL='" & RL & "',M11$ss='" & n_SS & "',M11SS='" & SS & "',M11RLLT='" & _RLLT & "',M11Category='" & Category & "',M11TOD_CQ='" & TOD_CQ & "',M11TOD_N3='" & TOD_N3 & "',M11Tot_Req='" & Tot_Req & "',M11OLD_PR='" & Old_PR & "',M11PO_QTY='" & PO_Qty & "',M11NewPR='" & NewPR & "' where M11SAPCode='" & CInt(Trim(fields(0))) & "' and M11Year=" & Microsoft.VisualBasic.Left(Trim(fields(5)), 4) & " and M11Month='" & Month(Trim(fields(5))) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                Else

                    p_Date = Month(Trim(fields(5))) & "/" & Microsoft.VisualBasic.Left(Trim(fields(5)), 4)
                    ' t_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(10)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(10)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(10)), 4)

                    nvcFieldList1 = "Insert Into M11MRS(M11SAPCode,M11Dis,M11PRC,M11PckSize,M11MR,M11Value,M11SD,M11MAvg,M11NSD,M11HGT6,M11HGT3,M11N3,M11CQ,M11Purchase,M11RQLT,M11RLLT,M11SS,M11$ss,M11RL,M11$RL,M11RQ,M11$RQ,M11ML,M11WHStock,M11EndStock,M11StkH,M11Category,M11TOD_CQ,M11TOD_N3,M11Tot_Req,M11OLD_PR,M11PO_QTY,M11NewPR,M11Month,M11Year,M11Date,M11Status)" & _
                                                             " values('" & SAPCode & "','" & Discripsion & "','" & _Prc & "','" & PkSiz & "','" & MR & "','" & Value & "','" & _SD & "','" & n_MAvg & "','" & n_SD & "','" & HgtMAvg & "','" & Hgt3MAvg & "','" & _N3 & "','" & CQ_MRS & "','" & Purchase & "','" & _RQLT & "','" & _RLLT & "','" & SS & "','" & n_SS & "','" & RL & "','" & n_RL & "','" & RQ & "','" & n_RQ & "','" & ML & "','" & WH_Stock & "','" & EnStock & "','" & stk_Holding & "','" & Category & "','" & TOD_CQ & "','" & TOD_N3 & "','" & Tot_Req & "','" & Old_PR & "','" & PO_Qty & "','" & NewPR & "','" & Month(Trim(fields(5))) & "','" & Year(Trim(fields(5))) & "','" & p_Date & "','Y')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                SAPCode = 0
                _Prc = ""
                PkSiz = ""
                MR = ""
                '  newRow("Month") = Trim(fields(5))
                Value = ""
                n_MAvg = ""
                HgtMAvg = ""
                Hgt3MAvg = ""
                CQ_MRS = ""
                Purchase = ""
                NLTW = ""
                NLTD = ""
                SS = ""
                nSS = ""
                RL = ""
                nRL = ""
                RQ = ""
                nRQ = ""
                ML = ""
                EnStock = ""
                stk_Holding = ""
                Category = ""
                X11 = X11 + 1
            Next
            MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Me.Cursor = Cursors.Arrow

            Dim _Renamepath As String

            _DisFilePath = ConfigurationManager.AppSettings("FilePath") + "\MRS_OLD\MRS.txt"
            _Renamepath = ConfigurationManager.AppSettings("FilePath") + "\MRS_OLD\" & Year(Today) & Month(Today) & Microsoft.VisualBasic.Day(Today) & ".txt"
            ' FileCopy(strFileName, _DisFilePath)
            FileCopy(strFileName, _Renamepath)
            Kill(strFileName)



        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Sub
End Class