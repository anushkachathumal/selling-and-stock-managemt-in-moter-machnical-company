Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmNew_Roots
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _Comcode As String
    Dim _Edit_Status As Boolean
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M02Root_Code as [##],M02Name as [Root Name] from M02Root where M02Status='A' order by M02Root_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmNew_Roots_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'lblDisplay.Text = strCompany_Profile
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCode.ReadOnly = True

        txtCode1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCode1.ReadOnly = True

        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Grid()
        Call Load_Entry()
    End Sub

    Function Load_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='RT' and P01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtCode.Text = "RT/00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtCode.Text = "RT/0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtCode.Text = "RT/" & M01.Tables(0).Rows(0)("P01LastNo")
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

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Call Load_Entry()
        OPR0.Visible = True
        txtDescription.Focus()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        OPR0.Visible = False
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtDescription.Text = ""
    End Sub

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = 13 Then
            If txtDescription.Text <> "" Then
                cmdAdd.Focus()
            End If
        End If
    End Sub


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
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
            If txtDescription.Text <> "" Then
            Else
                MsgBox("Please enter the Root Name", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                Exit Sub
            End If

            nvcFieldList1 = "SELECT * FROM M02Root WHERE M02Root_Code='" & Trim(txtCode.Text) & "' and m02Come_Code='" & _Comcode & "'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                MsgBox("This Root code alrady exist", MsgBoxStyle.Information, "Information ........")
                connection.Close()
            Else
                nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='RT'  and P01Com_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M02Root(M02Root_Code,M02Name,M02Status,m02Come_Code)" & _
                                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(txtDescription.Text) & "', 'A','" & _Comcode & "')"
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
    End Sub

    Private Sub EditDeleteRootToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditDeleteRootToolStripMenuItem.Click
        _Edit_Status = True
    End Sub

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _row As Integer

        _row = UltraGrid1.ActiveRow.Index
        If _Edit_Status = True Then
            OPR0.Visible = False
            OPR1.Visible = True
            Call Search_Records(Trim(UltraGrid1.Rows(_row).Cells(0).Text))
        End If
    End Sub

    Function Search_Records(ByVal strcode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M02Root_Code ,M02Name  from M02Root where M02Status='A' and M02Root_Code='" & strcode & "' and m02Come_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtCode1.Text = Trim(dsUser.Tables(0).Rows(0)("M02Root_Code"))
                txtDescription1.Text = Trim(dsUser.Tables(0).Rows(0)("M02Name"))
            End If
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
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
            If txtDescription1.Text <> "" Then
            Else
                MsgBox("Please enter the Root Name", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                Exit Sub
            End If
            nvcFieldList1 = "SELECT * FROM M02Root WHERE M02Root_Code='" & Trim(txtCode1.Text) & "' and m02Come_Code='" & _Comcode & "'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then

                nvcFieldList1 = "UPDATE M02Root SET M02Name='" & Trim(txtDescription1.Text) & "' WHERE M02Root_Code='" & Trim(txtCode1.Text) & "' and m02Come_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                'nvcFieldList1 = "Insert Into M02Root(M02Root_Code,M02Name,M02Status)" & _
                '                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(txtDescription.Text) & "', 'A')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                " values('NEW_ROOTS','EDIT', '" & Now & "','" & strDisname & "','" & txtCode1.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Root Change successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            txtDescription1.Text = ""
            txtCode1.Text = ""
            OPR1.Visible = False
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
            A = MsgBox("Are you sure you want to cancel this root", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel ..........")
            If A = vbYes Then

                nvcFieldList1 = "SELECT * FROM M02Root WHERE M02Root_Code='" & Trim(txtCode1.Text) & "' and m02Come_Code='" & _Comcode & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then

                    nvcFieldList1 = "UPDATE M02Root SET M02Name='" & Trim(txtDescription1.Text) & "' WHERE M02Root_Code='" & Trim(txtCode1.Text) & "' and m02Come_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'nvcFieldList1 = "Insert Into M02Root(M02Root_Code,M02Name,M02Status)" & _
                    '                                                  " values('" & Trim(txtCode.Text) & "','" & Trim(txtDescription.Text) & "', 'A')"
                    'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                    " values('NEW_ROOTS','EDIT', '" & Now & "','" & strDisname & "','" & txtCode1.Text & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If
                MsgBox("Root Change successfully", MsgBoxStyle.Information, "Information ..........")
                transaction.Commit()
                txtDescription1.Text = ""
                txtCode1.Text = ""
                OPR1.Visible = False
                Call Load_Grid()
                Call Load_Entry()

            End If
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        _Edit_Status = False
        OPR0.Visible = False
        OPR1.Visible = False
        txtDescription.Text = ""
        txtDescription1.Text = ""
        Call Load_Grid()
    End Sub


    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        OPR1.Visible = False
        Me.txtDescription1.Text = ""
        Me.txtCode1.Text = ""
    End Sub
End Class