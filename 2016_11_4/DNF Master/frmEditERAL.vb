Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmEditERAL
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableEral1
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Private Sub frmEditERAL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        txtDate.Text = Today
        txtTo.Text = Today
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Call Load_Gride()
        cmdEdit.Enabled = False
        txtDate.Text = Today
        txtTo.Text = Today
        chk1.Checked = False
        chkNowork.Checked = False
        chkOthers.Checked = False
        chkRe.Checked = False
        chkSt.Checked = False
        chkWash.Checked = False

    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim Sql As String
        Dim M01 As DataSet
        Dim I As Integer
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim _No As String
        Dim _Re As String
        Dim _Wt As String
        Dim _Ot As String
        Dim _Sp As String

        _No = ""
        _Re = ""
        _Ot = ""
        _Sp = ""
        _Wt = ""
        Try
            Dim From_Date As Date  ' Declaring From Date
            Dim To_Date As Date    ' Declaring To Date  

            From_Date = txtDate.Value & " " & "7:30AM"
            To_Date = txtTo.Value & " " & "7:30AM"

            Me.Cursor = Cursors.WaitCursor
            Dim strSt As String
            strSt = ""
            If chk1.Checked = True Then
                If chkNowork.Checked = True Then
                    _No = "N"
                ElseIf chkRe.Checked = True Then
                    _Re = "R"
                ElseIf chkWash.Checked = True Then
                    _Wt = "W"
                ElseIf chkSt.Checked = True Then
                    _Sp = "S"
                ElseIf chkOthers.Checked = True Then
                    _Ot = "O"
                End If

                If _No <> "" Or _Re <> "" Or _Sp <> "" Or _Wt <> "" Or _Ot <> "" Then
                    strSt = ""
                    If _No <> "" Then
                        strSt = _No
                    ElseIf _Re <> "" Then
                        If strSt <> "" Then
                            strSt = strSt & "," & _Re
                        Else
                            strSt = _Re
                        End If
                    ElseIf _Wt <> "" Then
                        If strSt <> "" Then

                            strSt = strSt & "," & _Wt
                        Else
                            strSt = _Wt
                        End If
                    ElseIf _Sp <> "" Then
                        If strSt <> "" Then

                            strSt = strSt & "," & _Sp
                        Else
                            strSt = _Sp
                        End If
                    ElseIf _Ot <> "" Then
                        If strSt <> "" Then

                            strSt = strSt & "," & _Ot
                        Else
                            strSt = _Ot
                        End If
                    End If
                End If
            End If
            If Trim(strSt) <> "" Then
                Sql = "select * from M04Lot where M04Etime between '" & From_Date & "' and '" & To_Date & "' and M04ProgrameType in ('" & strSt & "')"
            Else
                Sql = "select * from M04Lot where M04Etime between '" & From_Date & "' and '" & To_Date & "'"
            End If

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            pbCount.Minimum = 0
            lblDis.Text = ""
            pbCount.Value = pbCount.Minimum
            pbCount.Maximum = M01.Tables(0).Rows.Count
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                With M01.Tables(0)
                    newRow("Ref.Doc") = .Rows(I)("M04Ref")
                    newRow("Lot No") = .Rows(I)("M04Lotno")
                    newRow("Machine No") = .Rows(I)("M04Machine_No")
                    newRow("Programe No") = .Rows(I)("M04Program")
                    newRow("Programe Type") = .Rows(I)("M04ProgrameType")
                    newRow("Lot Type") = .Rows(I)("M04Type")
                    newRow("Standed Time") = .Rows(I)("M04STD")
                    newRow("Start Date") = .Rows(I)("M04DateIn")
                    newRow("Start Time") = .Rows(I)("M04TimeIn")
                    newRow("End Date") = .Rows(I)("M04Date_Out")
                    newRow("End Time") = .Rows(I)("M04Time_Out")
                    newRow("Total Hour") = .Rows(I)("M04Taken")
                    newRow("Quality") = .Rows(I)("M04Quality")
                    '  newRow("Quality Group") = .Rows(I)("M04Quality")
                    newRow("Shade Code") = .Rows(I)("M04Shade_Code")
                    newRow("Shade") = .Rows(I)("M04Shade")
                    newRow("Shade Type") = .Rows(I)("M04Shade_Type")
                    newRow("Weight") = .Rows(I)("M04Batchwt")

                    c_dataCustomer1.Rows.Add(newRow)
                End With

                ' If Not x = M01.Tables(0).Rows.Count - 1 Then
                ' rsT24RecHeader.MoveNext()
                pbCount.Value = pbCount.Value + 1
                lblDis.Text = M01.Tables(0).Rows(I)("M04Lotno") & "-" & M01.Tables(0).Rows(I)("M04ref")
                Me.Refresh()

                I = I + 1
            Next
            Me.Cursor = Cursors.Arrow
            cmdEdit.Enabled = True

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim i As Integer

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            pbCount.Minimum = 0
            lblDis.Text = ""
            pbCount.Value = pbCount.Minimum
            pbCount.Maximum = UltraGrid1.Rows.Count

            For Each uRow As UltraGridRow In UltraGrid1.Rows
                nvcFieldList1 = "update M04Lot set M04Lotno='" & UltraGrid1.Rows(i).Cells(1).Value & "',M04ProgrameType='" & UltraGrid1.Rows(i).Cells(4).Value & "',M04STD=" & Val(UltraGrid1.Rows(i).Cells(6).Value) & ",M04Quality='" & UltraGrid1.Rows(i).Cells(12).Value & "',M04Shade_Code='" & UltraGrid1.Rows(i).Cells(14).Value & "',M04Shade='" & UltraGrid1.Rows(i).Cells(15).Value & "',M04Shade_Type='" & UltraGrid1.Rows(i).Cells(16).Value & "',M04Batchwt=" & UltraGrid1.Rows(i).Cells(17).Value & " where M04Ref=" & UltraGrid1.Rows(i).Cells(0).Value & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                pbCount.Value = pbCount.Value + 1
                lblDis.Text = UltraGrid1.Rows(i).Cells(0).Value
                Me.Refresh()

                i = i + 1
            Next
            MsgBox("Record Updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

End Class