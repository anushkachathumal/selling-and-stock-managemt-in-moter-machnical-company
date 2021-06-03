Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmRoot_Cnt
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _Comcode As String
    Dim _Disname As String


    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_District()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M15Name as [##] from M15District order by M15Province"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDistrict
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 237
                '  .Rows.Band.Columns(1).Width = 160


            End With

            With cboDistrict_1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 258
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmRoot_Cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _PrintStatus = "A"
        Call Load_Entry()
        Call Load_District()
        Call Load_Grid()
    End Sub

    Function Load_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='RT' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtCode.Text = "RT/00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtCode.Text = "RT/0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtCode.Text = "RT/" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Clear_Text()
        Call Load_Entry()
        Me.txtDescription.Text = ""
        Me.cboDistrict.Text = ""
        cboDistrict.ToggleDropdown()
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Call Clear_Text()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_Text()
        Panel7.Visible = True
        cboDistrict_1.Text = ""
    End Sub

    Private Sub cboDistrict_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDistrict.AfterCloseUp
        Call Load_Grid_District(Trim(cboDistrict.Text))
    End Sub

    Private Sub cboDistrict_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDistrict.KeyUp
        If e.KeyCode = 13 Then
            txtDescription.Focus()
        End If
    End Sub

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtDescription.Text) <> "" Then
                cmdAdd.Focus()
            End If
        End If
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M02Root_Code  ) ##,M02Root_Code as [Root Code],M02Dis_Name as [District],M02Root_Name  as [Root Name]  from M02New_Root where M02Status='A' order by M02id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 30
            UltraGrid1.Rows.Band.Columns(1).Width = 80
            UltraGrid1.Rows.Band.Columns(2).Width = 120
            UltraGrid1.Rows.Band.Columns(3).Width = 230
            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
           
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_District(ByVal str_Disname As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M02Root_Code  ) ##,M02Root_Code as [Root Code],M02Dis_Name as [District],M02Root_Name  as [Root Name]  from M02New_Root where M02Status='A' and M02Dis_Name='" & str_Disname & "' order by M02id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 30
            UltraGrid1.Rows.Band.Columns(1).Width = 80
            UltraGrid1.Rows.Band.Columns(2).Width = 120
            UltraGrid1.Rows.Band.Columns(3).Width = 230
            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_Root()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M02Root_Code  ) ##,M02Root_Code as [Root Code],M02Dis_Name as [District],M02Root_Name  as [Root Name]  from M02New_Root where M02Status='A' and M02Root_Name like '%" & Trim(txtDescription.Text) & "%' order by M02id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 30
            UltraGrid1.Rows.Band.Columns(1).Width = 80
            UltraGrid1.Rows.Band.Columns(2).Width = 120
            UltraGrid1.Rows.Band.Columns(3).Width = 230
            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_District() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M15District where  M15Name='" & Trim(cboDistrict.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_District = True
            End If


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Search_District() = True Then
        Else
            MsgBox("Please select the correct District", MsgBoxStyle.Information, "Information ........")
            Exit Sub
        End If

        If Trim(txtDescription.Text) <> "" Then
        Else
            MsgBox("Please enter the root name", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        Call Save_Data()
    End Sub

    Function Save_Data()
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
        Try
            
            nvcFieldList1 = "SELECT * FROM M02New_Root WHERE M02Root_Code='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M02New_Root set M02Root_Name='" & Trim(txtDescription.Text) & "',M02Dis_Name='" & Trim(cboDistrict.Text) & "' where M02Root_Code='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_ROOTS','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='RT' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M02New_Root(M02Root_Code,M02Root_Name,M02Status,M02Dis_Name)" & _
                                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(txtDescription.Text) & "', 'A','" & Trim(cboDistrict.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                " values('NEW_ROOTS','SAVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Root create successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            txtDescription.Text = ""
            txtDescription.Focus()
            Call Load_Grid()
            Call Load_Entry()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    
    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M02New_Root where  M02Root_Code='" & Trim(txtCode.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboDistrict.Text = Trim(M01.Tables(0).Rows(0)("M02Dis_Name"))
                txtDescription.Text = Trim(M01.Tables(0).Rows(0)("M02Root_Name"))
            End If


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Private Sub txtDescription_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.ValueChanged
        Call Load_Grid_Root()
    End Sub

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid1.ActiveRow.Index
        txtCode.Text = Trim(UltraGrid1.Rows(_Row).Cells(1).Text)
        Call Search_Records()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
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
        Dim A As String

        Try
            A = MsgBox("Are you sure you want to Delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel ....")
            If A = vbYes Then
                nvcFieldList1 = "SELECT * FROM M02New_Root WHERE M02Root_Code='" & Trim(txtCode.Text) & "' "
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    nvcFieldList1 = "update M02New_Root set M02Status='I' where M02Root_Code='" & Trim(txtCode.Text) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                   " values('NEW_ROOTS','DELETE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Root create successfully", MsgBoxStyle.Information, "Information ..........")
                Else
                    
                    MsgBox("Record not available  for delete", MsgBoxStyle.Information, "Information .....")
                End If
            End If

            transaction.Commit()
            txtDescription.Text = ""
            txtDescription.Focus()
            Call Load_Grid()
            Call Load_Entry()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub EditDeleteRootToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditDeleteRootToolStripMenuItem.Click
        Panel7.Visible = True
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        _PrintStatus = "B"
        _Disname = Trim(cboDistrict_1.Text)
        Call Load_Grid_District(Trim(cboDistrict_1.Text))
        Panel7.Visible = False
    End Sub
End Class