Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Public Class frmAlarm
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableAlarm
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

    Private Sub frmAlarm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Try
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\DCA_ALARM.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                MTNo = Trim(fields(0))
                _Plant = Trim(fields(1))
                _StockLoc = Trim(fields(2))

                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Category") = (Trim(fields(0)))   '0
                newRow("SAP Code") = CInt(Trim(fields(1))) '1
                newRow("SS") = Trim(fields(3)) '3
                newRow("End Stock") = Trim(fields(4)) '4
                newRow("Avg 12") = Trim(fields(5)) '5
                newRow("M6") = Trim(fields(6)) '6
                newRow("N3") = Trim(fields(7)) '7
                newRow("CQ-MRS") = Trim(fields(8)) '8
                newRow("Last Month Qty") = Trim(fields(9)) '9
                newRow("Last 30Day Qty") = Trim(fields(10)) '10
                newRow("SD") = Trim(fields(11)) '11
                newRow("2nd Last Week") = Trim(fields(12)) '12
                newRow("Last Week") = Trim(fields(13)) '13
                newRow("Com Last 30Day") = Trim(fields(14)) '14
                newRow("Com SD") = Trim(fields(15)) '15
                newRow("Com 2ndLast Week") = Trim(fields(16)) '16
                newRow("Com Last Week") = Trim(fields(17)) '17
                newRow("Com Last 14Days") = Trim(fields(18)) '18
                newRow("RL") = Trim(fields(19)) '19
                newRow("RQ") = Trim(fields(20)) '20
                newRow("LT Days") = Trim(fields(21)) '21
                newRow("Outstanding P/O") = Trim(fields(22)) '22
                newRow("#PO") = Trim(fields(23)) '23
                newRow("PID") = Trim(fields(24))
                newRow("PDD") = Trim(fields(25))
                newRow("Reg No") = Trim(fields(26))
                newRow("ETA") = Trim(fields(27))
                newRow("ETD") = Trim(fields(28))
                newRow("L14D - Con") = Trim(fields(29))

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

        Try
            Me.Cursor = Cursors.WaitCursor

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\DCA_ALARM.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                MTNo = Trim(fields(0))
                _Plant = Trim(fields(1))
                _StockLoc = Trim(fields(2))


                Category = (Trim(fields(0)))   '0
                SAPCode = CInt(Trim(fields(1))) '1
                Dis = Trim(fields(2))
                SS = Trim(fields(3)) '3
                EndStock = Trim(fields(4)) '4
                Avg = Trim(fields(5)) '5
                M6 = Trim(fields(6)) '6
                N3 = Trim(fields(7)) '7
                CQ_MRS = Trim(fields(8)) '8
                Last_MntQty = Trim(fields(9)) '9
                Last_30Day = Trim(fields(10)) '10
                SD = Trim(fields(11)) '11
                L_Week2nd = Trim(fields(12)) '12
                L_week2 = Trim(fields(13)) '13
                L_30Day2 = Trim(fields(14)) '14
                SD2 = Trim(fields(15)) '15
                L_Week2nd2 = Trim(fields(16)) '16
                L_Week3 = Trim(fields(17)) '17

                'MsgBox(Trim(fields(18)))
                L_14Day = Trim(fields(18)) '18
                RL = Trim(fields(19)) '19
                RQ = Trim(fields(20)) '20
                '  MsgBox(Trim(fields(21)))
                If Trim(fields(21)) <> "" Then
                    LT_day = Trim(fields(21)) '21
                End If
                Pending_PO = Trim(fields(22)) '22
                PO = Trim(fields(23)) '23
                ' newRow("PID") = Trim(fields(24))
                'newRow("PDD") = Trim(fields(25))
                ItemNo = Trim(fields(26))
                ' newRow("ETA") = Trim(fields(27))
                'newRow("ETD") = Trim(fields(28))
                L14Day_Con = Trim(fields(29))



                If Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(25)), 5), 2) = "00" Then
                    nvcFieldList1 = "select * from Alarm where SAPCode='" & CInt(Trim(fields(1))) & "' and Pending_PO='" & Trim(fields(22)) & "'"
                Else
                    t_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(25)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(25)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(25)), 4)
                    nvcFieldList1 = "select * from Alarm where SAPCode='" & CInt(Trim(fields(1))) & "' and Pending_PO='" & Trim(fields(22)) & "' and PODelDate='" & t_Date & "'"
                End If


                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                Else
                    p_Date = ""
                    t_Date = ""
                    If Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(24)), 5), 2) = "00" Then
                    Else
                        p_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(24)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(24)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(24)), 4)
                    End If

                    If Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(25)), 5), 2) = "00" Then
                    Else
                        t_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(25)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(25)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(25)), 4)
                    End If
                    'ETA
                    If Trim(fields(28)) <> "" Then
                        ETA = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(28)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(28)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(28)), 4)
                    Else
                        '  ETA = "00/00/0000"
                    End If
                    '-----------------------------------------------------
                    'ETD
                    If Trim(fields(27)) <> "" Then
                        ETD = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(27)), 5), 2) & "/" & Microsoft.VisualBasic.Left(Trim(fields(27)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(27)), 4)
                    Else
                        ' ETD = "00/00/0000"
                    End If

                    nvcFieldList1 = "Insert Into Alarm(Category,SAPCode,SS,EndStock,Avg,M6,N3,CQ_MRS,Last_MntQty,Last_30Day,SD,L_Week2nd,L_week2,L_30Day2,SD2,L_Week2nd2,L_Week3,L_14Day,RL,RQ,LT_day,Pending_PO,PO,PODate,PODelDate,ItemNo,ETD,ETA,L14Day_Con,Dis,RunDate)" & _
                                                             " values('" & Category & "', " & CInt(SAPCode) & ",'" & SS & "','" & EndStock & "','" & Avg & "','" & M6 & "','" & N3 & "','" & CQ_MRS & "','" & Last_MntQty & "','" & Last_30Day & "','" & SD & "','" & L_Week2nd & "','" & L_week2 & "','" & L_30Day2 & "','" & SD2 & "','" & L_Week2nd2 & "','" & L_Week3 & "','" & L_14Day & "','" & RL & "','" & RQ & "','" & LT_day & "','" & Pending_PO & "','" & PO & "','" & p_Date & "','" & t_Date & "','" & ItemNo & "','" & ETD & "','" & ETA & "','" & L14Day_Con & "','" & Dis & "','" & Today & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                p_Date = ""
                t_Date = ""
                ETA = ""
                ETD = ""

                Category = ""   '0
                SAPCode = "" '1
                Dis = ""
                SS = "" '3
                EndStock = 0 '4
                Avg = 0 '5
                M6 = 0 '6
                N3 = 0 '7
                CQ_MRS = 0 '8
                Last_MntQty = 0 '9
                Last_30Day = 0 '10
                SD = 0 '11
                L_Week2nd = 0 '12
                L_week2 = 0 '13
                L_30Day2 = 0 '14
                SD2 = 0 '15
                L_Week2nd2 = 0 '16
                L_Week3 = 0 '17
                L_14Day = 0 '18
                RL = 0 '19
                RQ = 0 '20
                LT_day = 0 '21
                Pending_PO = "" '22
                PO = "" '23
                ' newRow("PID") = Trim(fields(24))
                'newRow("PDD") = Trim(fields(25))
                ItemNo = ""
                ' newRow("ETA") = Trim(fields(27))
                'newRow("ETD") = Trim(fields(28))
                L14Day_Con = 0



                X11 = X11 + 1
            Next
            MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            Me.Cursor = Cursors.Arrow
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