Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.Drawing.Image
Imports System.IO
Public Class frmElectrician_cnt
    Dim _PrintStatus As String
    Dim _From As String
    Dim _TO As String
    Dim _Sales_ref As String
    Dim _Root As String

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        OPR0.Visible = False
    End Sub
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
        cboDistrict.ToggleDropdown()
    End Sub
    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim A As String

        A = MsgBox("Are you sure you want to deactive Electrician", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Deactive Electrician.........")
        If A = vbYes Then
            Call Deactive_Data()
        End If
    End Sub
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Trim(cboDistrict.Text) <> "" Then
        Else
            MsgBox("Please select the District", MsgBoxStyle.Information, "Information ........")
            cboDistrict.ToggleDropdown()
            Exit Sub
        End If

        If Trim(txtName.Text) <> "" Then
        Else
            MsgBox("Please enter the Electrician Name", MsgBoxStyle.Information, "Information ........")
            Exit Sub
        End If

        If Trim(txtAddress.Text) <> "" Then
        Else
            txtAddress.Text = "-"
        End If

        If txtNIC.Text <> "" Then
        Else
            MsgBox("Please enter the NIC No", MsgBoxStyle.Information, "Information .......")
            Exit Sub
        End If

        If Trim(txtTP.Text) <> "" Then
        Else
            MsgBox("Please enter the Contact No", MsgBoxStyle.Information, "Information .......")
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

            nvcFieldList1 = "SELECT * FROM M08Electritions WHERE M08Ref_No='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M08Electritions set M08District='" & Trim(cboDistrict.Text) & "',M08Elc_Name='" & Trim(txtName.Text) & "',M08Address='" & Trim(txtAddress.Text) & "',M08Contact_No='" & Trim(txtTP.Text) & "',M08NIC='" & Trim(txtNIC.Text) & "',M08Status='A' where M08Ref_No='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_ELECTRICIAN','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='EL' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M08Electritions(M08Ref_No,M08District,M08Elc_Name,M08Address,M08Contact_No,M08NIC,M08Status)" & _
                                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(cboDistrict.Text) & "', '" & Trim(txtName.Text) & "','" & Trim(txtAddress.Text) & "','" & Trim(txtTP.Text) & "','" & Trim(txtNIC.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                " values('NEW_ELECTRICIAN','SAVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()
            

            Call Load_Grid()
            Call Clear_Text()
            Call Load_Entry()
            cboDistrict.ToggleDropdown()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function


    Function Deactive_Data()
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

            nvcFieldList1 = "SELECT * FROM M08Electritions WHERE M08Ref_No='" & Trim(txtCode.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M08Electritions set M08Status='I' where M08Ref_No='" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_ELECTRICIAN','DELETE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
           
            End If
            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()


            Call Load_Grid()
            Call Clear_Text()
            Call Load_Entry()
            cboDistrict.ToggleDropdown()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function


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
                .Rows.Band.Columns(0).Width = 285
                '  .Rows.Band.Columns(1).Width = 160


            End With

            'With cboDistrict_1
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 258
            '    '  .Rows.Band.Columns(1).Width = 160


            'End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Clear_Text()
        Me.txtAddress.Text = ""
        Me.txtNIC.Text = ""
        Me.txtTP.Text = ""
        Me.txtName.Text = ""
        Me.cboDistrict.Text = ""
        Call Load_Entry()
    End Function

    Private Sub frmElectrician_cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_District()
        Call Load_Entry()
        Call Load_Grid()
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Function Load_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='EL' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtCode.Text = "EL/00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtCode.Text = "EL/0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtCode.Text = "EL/" & M01.Tables(0).Rows(0)("P01No")
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

    Private Sub cboDistrict_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDistrict.KeyUp
        If e.KeyCode = 13 Then
            txtName.Focus()
        End If
    End Sub

    Private Sub txtName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtName.Text) <> "" Then
                txtAddress.Focus()
            End If
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = Keys.F1 Then
            txtNIC.Focus()
        End If
    End Sub

    Private Sub txtNIC_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNIC.KeyUp
        If e.KeyCode = 13 Then
            txtTP.Focus()
        End If
    End Sub

    Private Sub txtTP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTP.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_Text()
        Call Load_Grid()
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        OPR0.Visible = True
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M08Ref_No) ##,M08Ref_No as [Ref.Doc.No],M08District as [District Name],M08Elc_Name as [Electrician Name],M08Contact_No as [Contact No]  from M08Electritions where M08Status='A' order by M08Id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 30
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 110
            UltraGrid1.Rows.Band.Columns(3).Width = 320
            UltraGrid1.Rows.Band.Columns(4).Width = 120
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Load_Grid_1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M08Ref_No) ##,M08Ref_No as [Ref.Doc.No],M08District as [District Name],M08Elc_Name as [Electrician Name],M08Contact_No as [Contact No]  from M08Electritions where M08Status='I' order by M08Id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 30
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 110
            UltraGrid1.Rows.Band.Columns(3).Width = 320
            UltraGrid1.Rows.Band.Columns(4).Width = 120
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer
        OPR0.Visible = True
        _Row = UltraGrid1.ActiveRow.Index
        txtCode.Text = Trim(UltraGrid1.Rows(_Row).Cells(1).Text)
        Call Search_Records()

    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select  *  from M08Electritions where M08Ref_No='" & Trim(txtCode.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboDistrict.Text = Trim(M01.Tables(0).Rows(0)("M08District"))
                txtName.Text = Trim(M01.Tables(0).Rows(0)("M08Elc_Name"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M08Address"))
                txtNIC.Text = Trim(M01.Tables(0).Rows(0)("M08NIC"))
                txtTP.Text = Trim(M01.Tables(0).Rows(0)("M08Contact_No"))

            End If


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub DeactivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactivateToolStripMenuItem.Click
        Call Load_Grid_1()
    End Sub
End Class