Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Imports Infragistics.Win
Imports Infragistics.Win.Layout
Imports Infragistics.Win.UltraWinTree
Public Class frmLeger_Acc
    Dim _Acc_Type As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _Main_Acc As String
    Dim _SUB_Acc As String
    Dim _Acc_Doc As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OPR0.Enabled = True
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M28Name as [##] from M28Main_Acc "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMain
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 320
                '  .Rows.Band.Columns(1).Width = 160


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmLeger_Acc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCODE")
        Call Load_Acc_Type()
        txtYear.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtOB_Chq.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCreate_Date.Text = Today
        ' Call Load_Status()
        txtYear.Text = Year(Today)
        Call Load_Combo()
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride1()
    End Sub

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet

        Try
            Sql = "select M30Dis as [Acc Of],M01Acc_Code as [Account Code],M01Acc_Name as [Description] from View_Acc where M01Com_Code='" & _Comcode & "' and m01ACC='A' and M28Name='" & cboMain.Text & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 80
            UltraGrid1.Rows.Band.Columns(2).Width = 320

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet

        Try
            Sql = "select M30Dis as [Acc Of],M01Acc_Code as [Account Code],M01Acc_Name as [Description] from View_Acc where M01Com_Code='" & _Comcode & "' and m01ACC='A' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 80
            UltraGrid1.Rows.Band.Columns(2).Width = 320

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Acc_Type()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M30Dis as [##] from M30Acc_Statment"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboType
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 175
            End With
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Main_Acccode() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select * from M28Main_Acc where M28Name='" & cboMain.Text & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then

                Search_Main_Acccode = True
                _Main_Acc = Trim(T01.Tables(0).Rows(0)("M28Code"))

                SQL = "select M29Dis as [##] from M29Acc_Type where M29Main_Acc='" & _Main_Acc & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                With cboSub
                    .DataSource = T01
                    .Rows.Band.Columns(0).Width = 175
                End With
                'Call Load_Acc_Code(_Main_Acc)
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Search_SUB_Acccode() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select * from M29Acc_Type where M29Main_Acc='" & _Main_Acc & "' AND M29Dis='" & cboSub.Text & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then

                Search_SUB_Acccode = True
                _SUB_Acc = Trim(T01.Tables(0).Rows(0)("M29Code"))

              
                'Call Load_Acc_Code(_Main_Acc)
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    'Function Load_Status()
    '    Dim SQL As String
    '    Dim con = New SqlConnection()
    '    con = DBEngin.GetConnection()
    '    Dim T01 As DataSet
    '    Try
    '        SQL = "select M06Name as [##] from M06Status"
    '        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '        With cboStatus
    '            .DataSource = T01
    '            .Rows.Band.Columns(0).Width = 65
    '        End With
    '    Catch returnMessage As Exception
    '        If returnMessage.Message <> Nothing Then
    '            MessageBox.Show(returnMessage.Message)
    '        End If
    '    End Try
    'End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtAmount.Text) Then
                Value = txtAmount.Text
                txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtOB_Chq.Focus()
        End If
    End Sub

    Private Sub txtAmount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.ValueChanged

    End Sub

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
            cboType.ToggleDropdown()

        End If
    End Sub

    Private Sub txtCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.LostFocus
        Call Search_Records()
    End Sub

    Private Sub txtCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.ValueChanged

    End Sub

    Private Sub cboType_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboType.InitializeLayout

    End Sub

    Private Sub cboType_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboType.KeyUp
        If e.KeyCode = 13 Then
            ' Call Search_Records()
            txtYear.Focus()

        ElseIf e.KeyCode = Keys.Tab Then
            txtYear.Focus()
        End If
    End Sub

    Private Sub cboStatus_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub cboStatus_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtDescription.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDescription.Focus()
        End If
    End Sub

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            cmdAdd.Focus()
        End If
    End Sub

   
    Private Sub txtAdd1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    

   

    Private Sub txtYear_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear.KeyUp
        If e.KeyCode = Keys.Enter Then '
            txtDescription.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDescription.Focus()
        End If
    End Sub

    Private Sub txtYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtYear.ValueChanged

    End Sub

    Private Sub cmdExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub txtDescription_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.ValueChanged

    End Sub

    Private Sub txtContactNo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtCreate_Date.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCreate_Date.Focus()
        End If
    End Sub

    Private Sub txtContactNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtCreate_Date_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCreate_Date.BeforeDropDown

    End Sub

    Private Sub txtCreate_Date_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCreate_Date.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            cmdAdd.Focus()
        End If
    End Sub

    Function Search_ACCType() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select * from M30Acc_Statment where M30Dis='" & Trim(cboType.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Acc_Type = Trim(M01.Tables(0).Rows(0)("M30Code"))
                Search_ACCType = True
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
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
        Dim i As Integer
        Dim result1 As String
        Try
            If cboMain.Text <> "" Then
            Else
                result1 = MessageBox.Show("Please Select the Main Account", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboMain.ToggleDropdown()
                    Exit Sub
                End If
            End If


            If Search_ACCType() = True Then
            Else
                result1 = MessageBox.Show("Please Select the Account Doc", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboType.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please Select the Account type", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboSub.ToggleDropdown()
                    Exit Sub
                End If
            End If

            ' Call Load_Acc_Code(_Main_Acc)

            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Account Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            'If Trim(cboStatus.Text) <> "" Then
            'Else
            '    result1 = MessageBox.Show("Please enter the Status", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    If result1 = Windows.Forms.DialogResult.OK Then
            '        cboStatus.ToggleDropdown()
            '        Exit Sub
            '    End If
            'End If

            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Account Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If

            If txtAmount.Text <> "" Then
                If IsNumeric(txtAmount.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the correct Amount", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtAmount.Focus()
                        Exit Sub
                    End If
                End If
            End If


            '----------------------------------------------------------------------
            'If txtAdd1.Text <> "" Then
            'Else
            '    txtAdd1.Text = " "
            'End If

            'If txtAdd2.Text <> "" Then
            'Else
            '    txtAdd2.Text = " "
            'End If

            'If txtContactNo.Text <> "" Then
            'Else
            '    txtContactNo.Text = " "
            'End If

            'nvcFieldList1 = "select * from P02PARAMETER  where P02CODE='" & _Main_Acc & "' and P02DIS='" & _Comcode & "'"
            'MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            'If isValidDataset(MB51) Then
            '    If MB51.Tables(0).Rows(0)("P02SUB") = 0 Then
            '        nvcFieldList1 = "update P02PARAMETER set P02NO=P02NO + " & 1 & " where P02CODE='" & _Main_Acc & "' and P02DIS='" & _Comcode & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '    ElseIf MB51.Tables(0).Rows(0)("P02SUB") = 100 Then
            '        'nvcFieldList1 = "update P02PARAMETER set P02SUB='1', where P02CODE='" & _Main_Acc & "' and P02DIS='" & _Comcode & "'"
            '        'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '    End If
            'End If
            nvcFieldList1 = "Insert Into M01Account_Master(M01Acc_Type,M01Main_Acc,M01Acc_Statment,M01Acc_Code,M01Acc_Name,M01Address,M01Address2,M01TP,M01Acc_Limit,M01DOC,M01User,M01Status,M01year,M01Com_Code,M01Acc_Of,M01OB_Chq,m01ACC)" & _
                                                            " values('" & _Loccode & "','" & _Main_Acc & "','" & _Acc_Type & "', '" & (Trim(txtCode.Text)) & "','" & txtDescription.Text & "','-','-','-','" & txtAmount.Text & "','" & txtCreate_Date.Text & "','" & strDisname & "','A','" & txtYear.Text & "','" & _Comcode & "','" & _Comcode & "','" & txtOB_Chq.Text & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            nvcFieldList1 = "UPDATE P02PARAMETER SET P02NO=P02NO +" & 1 & " WHERE P02CODE='" & _Main_Acc & "' AND P02SUB='" & _Loccode & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            connection.Close()
            txtCode.Text = ""
            txtDescription.Text = ""
            ' txtAdd1.Text = ""
            'txtAdd2.Text = ""
            cboType.Text = ""
            'txtContactNo.Text = ""
            txtYear.Text = Year(Today)
            cboType.ToggleDropdown()
            Call Load_Gride()
            Call Load_Acc_Code(_Main_Acc)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Call Search_ACCType()

            Sql = "select * from View_Acc  where M01Acc_Code='" & Trim(txtCode.Text) & "'  and M01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                '  txtYear.Text = Trim(M01.Tables(0).Rows(0)("M01year"))
                txtDescription.Text = Trim(M01.Tables(0).Rows(0)("M01Acc_Name"))
                'txtAdd1.Text = Trim(M01.Tables(0).Rows(0)("M01Address"))
                'txtAdd2.Text = Trim(M01.Tables(0).Rows(0)("M01Address2"))
                'txtContactNo.Text = Trim(M01.Tables(0).Rows(0)("M01TP"))
                txtCreate_Date.Text = Trim(M01.Tables(0).Rows(0)("M01DOC"))
                'Value = Trim(M01.Tables(0).Rows(0)("M01Acc_Limit"))
                'txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                ' cboStatus.Text = Trim(M01.Tables(0).Rows(0)("M01Status"))
                cboMain.Text = Trim(M01.Tables(0).Rows(0)("M28Name"))
                Call Search_Main_Acccode()
                cboSub.Text = Trim(M01.Tables(0).Rows(0)("M29Dis"))
                cboType.Text = Trim(M01.Tables(0).Rows(0)("M30Dis"))

                'cboLocation.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                cmdAdd.Enabled = False
                cmdDelete.Enabled = True
                cmdEdit.Enabled = True
            Else
                cmdAdd.Enabled = True
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        ' OPR2.Enabled = False
        'OPR1.Enabled = False
        OPR0.Enabled = True
        ' OPR3.Enabled = False
        cmdAdd.Enabled = True
        txtCode.Focus()
        ' cmdSave.Enabled = False
        cmdDelete.Enabled = False
        'Call Load_Gride()
        txtYear.Text = Year(Today)
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
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
        Dim i As Integer
        Dim result1 As String
        Try

            If Search_ACCType() = True Then
            Else
                result1 = MessageBox.Show("Please Select the Account Type", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboType.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Account Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            'If Trim(cboStatus.Text) <> "" Then
            'Else
            '    result1 = MessageBox.Show("Please enter the Status", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    If result1 = Windows.Forms.DialogResult.OK Then
            '        cboStatus.ToggleDropdown()
            '        Exit Sub
            '    End If
            'End If

            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Account Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If

            If txtAmount.Text <> "" Then
                If IsNumeric(txtAmount.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the correct Amount", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtAmount.Focus()
                        Exit Sub
                    End If
                End If
            End If
            '----------------------------------------------------------------------
            'If txtAdd1.Text <> "" Then
            'Else
            '    txtAdd1.Text = " "
            'End If

            'If txtAdd2.Text <> "" Then
            'Else
            '    txtAdd2.Text = " "
            'End If

            'If txtContactNo.Text <> "" Then
            'Else
            '    txtContactNo.Text = " "
            'End If

            nvcFieldList1 = "update M01Account_Master set M01Acc_Name='" & Trim(txtDescription.Text) & "',M01Address='-',M01Address2='-',M01TP='-',M01Acc_Limit='" & txtAmount.Text & "',M01DOC='" & txtCreate_Date.Text & "',M01Status='-',M01year='" & txtYear.Text & "',M01OB_Chq='" & txtOB_Chq.Text & "' where  M01Acc_Code='" & Trim(txtCode.Text) & "' and M01Com_Code='" & _Comcode & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR0.Enabled = True
            ' OPR3.Enabled = False
            cmdAdd.Enabled = True
            txtCode.Focus()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride1()
            'Call Load_Gride()
            txtYear.Text = Year(Today)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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
        Dim i As Integer
        Dim result1 As String
        Try

            result1 = MsgBox("Are you sure you want to delete this Account", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete ......")
            If result1 = vbYes Then
                Call Search_ACCType()

                'nvcFieldList1 = "select * from P02PARAMETER  where P02CODE='" & _Main_Acc & "' and P02DIS='" & _Comcode & "'"
                'MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(MB51) Then
                '    If MB51.Tables(0).Rows(0)("P02SUB") = 0 Then
                '        nvcFieldList1 = "update P02PARAMETER set P02NO=P02NO - " & 1 & " where P02CODE='" & _Main_Acc & "' and P02DIS='" & _Comcode & "'"
                '        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '    ElseIf MB51.Tables(0).Rows(0)("P02SUB") = 100 Then
                '        'nvcFieldList1 = "update P02PARAMETER set P02SUB='1', where P02CODE='" & _Main_Acc & "' and P02DIS='" & _Comcode & "'"
                '        'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '    End If
                'End If

                nvcFieldList1 = "delete from M01Account_Master where M01Com_Code='" & _Comcode & "'  and M01Acc_Code='" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
                transaction.Commit()
            End If
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR0.Enabled = True
            ' OPR3.Enabled = False
            cmdAdd.Enabled = True
            txtCode.Focus()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()
            txtYear.Text = Year(Today)
            '   Call Load_Acc_Code()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub UltraGroupBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox5.Click

    End Sub

    Private Sub cboLocation_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    'Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = 13 Then
    '        txtContactNo.Focus()
    '    ElseIf e.KeyCode = Keys.Tab Then
    '        txtContactNo.Focus()
    '    End If
    ' End Sub

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet

        Dim i As Integer
        Try
            Sql = "select * from M29Acc_Type where M29Main_Acc='" & _Main_Acc & "' and M29Dis='" & Trim(cboSub.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Loccode = Trim(M01.Tables(0).Rows(0)("M29Code"))
                Search_Location = True
            End If



            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtOB_Chq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOB_Chq.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtOB_Chq.Text) Then
                Value = txtOB_Chq.Text
                txtOB_Chq.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtOB_Chq.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtOB_Chq.Focus()
        End If
    End Sub

    Function Load_Acc_Code(ByVal ACCCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P02PARAMETER where  P02CODE='" & ACCCode & "' AND P02SUB='" & _SUB_Acc & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If (Trim(M01.Tables(0).Rows(0)("P02no"))) < 10 Then
                    txtCode.Text = Trim(_Comcode) & "-" & Trim(M01.Tables(0).Rows(0)("P02CODE")) & "/" & Trim(M01.Tables(0).Rows(0)("P02DIS")) & "/00" & M01.Tables(0).Rows(0)("P02no")
                ElseIf (Trim(M01.Tables(0).Rows(0)("P02no"))) >= 10 And (Trim(M01.Tables(0).Rows(0)("P02no"))) < 100 Then
                    txtCode.Text = Trim(_Comcode) & "-" & Trim(M01.Tables(0).Rows(0)("P02CODE")) & "/" & Trim(M01.Tables(0).Rows(0)("P02DIS")) & "/0" & M01.Tables(0).Rows(0)("P02no")
                Else
                    txtCode.Text = Trim(_Comcode) & "-" & Trim(M01.Tables(0).Rows(0)("P02CODE")) & "/" & Trim(M01.Tables(0).Rows(0)("P02DIS")) & "/" & M01.Tables(0).Rows(0)("P02no")
                End If
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtOB_Chq_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOB_Chq.ValueChanged

    End Sub

    Private Sub cboMain_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMain.AfterCloseUp
        Call Search_Main_Acccode()
        Call Load_Gride()
    End Sub

    Private Sub cboMain_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboMain.InitializeLayout

    End Sub

    Private Sub cboMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMain.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Main_Acccode()
            cboSub.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_Main_Acccode()
            cboSub.ToggleDropdown()
        End If
    End Sub

    Private Sub cboSub_AfterDropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSub.AfterDropDown
        Call Search_Main_Acccode()
        Call Search_SUB_Acccode()
        Call Load_Acc_Code(_Main_Acc)
    End Sub


    Private Sub cboSub_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSub.KeyUp
        If e.KeyCode = 13 Then
            ' Call Search_Main_Acccode()
            Call Search_SUB_Acccode()
            cboType.ToggleDropdown()
            Call Load_Acc_Code(_Main_Acc)
        ElseIf e.KeyCode = Keys.Tab Then
            cboType.ToggleDropdown()
        End If
    End Sub


    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If UltraGroupBox1.Visible = True Then
            UltraGroupBox1.Visible = False
        Else
            UltraGroupBox1.Visible = True
            Dim dataSet As DataSet = Me.GetData()

            '	Set 'AutoGenerateColumnSets' to true so that UltraTreeCOlumnSets are automatically generated
            Me.UltraTree1.ColumnSettings.AutoGenerateColumnSets = True

            '	Set 'SynchronizeCurrencyManager' to false to optimize performance
            Me.UltraTree1.SynchronizeCurrencyManager = False

            '	Set 'ViewStyle' to OutlookExpress for a relational display
            Me.UltraTree1.ViewStyle = UltraWinTree.ViewStyle.OutlookExpress

            '	Set 'AutoFitColumns' to true to automatically fit all columns
            Me.UltraTree1.ColumnSettings.AutoFitColumns = AutoFitColumns.ResizeAllColumns

            '	Set DataSource/DataMember to bind the UltraTree
            Me.UltraTree1.DataSource = dataSet
            Me.UltraTree1.DataMember = "TVSeries"

        End If


    End Sub

    Private Function GetData() As DataSet
        Dim dataSet As DataSet = New DataSet()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim I As Integer
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim X As Integer
        Dim Z As Integer
        Try
            Dim dataTableSeries As DataTable = New DataTable("TVSeries")
            Dim dataTableSpinoffs As DataTable = New DataTable("Spinoffs")
            Dim dataTableSpinoffs1 As DataTable = New DataTable("Spinoffs1")

            dataTableSeries.Columns.Add("RecID", GetType(Integer))
            dataTableSeries.Columns.Add("Account Name", GetType(String))

            dataTableSpinoffs.Columns.Add("RecordID", GetType(Integer))
            dataTableSpinoffs.Columns.Add("ParentRecordID", GetType(Integer))
            dataTableSpinoffs.Columns.Add("Account Name", GetType(String))

            dataTableSpinoffs1.Columns.Add("RecordID", GetType(Integer))
            dataTableSpinoffs1.Columns.Add("ParentRecordID", GetType(Integer))
            dataTableSpinoffs1.Columns.Add("Account Name", GetType(String))

            I = 0
            Sql = "select * from M28Main_Acc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                dataTableSeries.Rows.Add(New Object() {I, Trim(M01.Tables(0).Rows(I)("M28Name"))})
                Sql = "select * from M01Account_Master where M01Main_Acc='" & Trim(M01.Tables(0).Rows(I)("M28Code")) & "'  and M01Com_Code='" & _Comcode & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                X = 0
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    dataTableSpinoffs.Rows.Add(New Object() {X, I, M02.Tables(0).Rows(X)("M01Acc_Code") & "-" & Trim(M02.Tables(0).Rows(X)("M01Acc_Name"))})
                    Z = 0
                    Sql = "select * from M01Account_Master where M01Main_Acc='" & Trim(M01.Tables(0).Rows(I)("M28Code")) & "' and len(M01Acc_Code)>4 and left(M01Acc_Code,4)='" & M02.Tables(0).Rows(X)("M01Acc_Code") & "' and M01Com_Code='" & _Comcode & "' order by convert(int,right(M01Acc_Code,2))"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    For Each DTRow4 As DataRow In M03.Tables(0).Rows
                        dataTableSpinoffs.Rows.Add(New Object() {M03.Tables(0).Rows(Z)("M01Acc_Code"), M02.Tables(0).Rows(X)("M01Acc_Code"), M03.Tables(0).Rows(Z)("M01Acc_Code") & "-" & Trim(M03.Tables(0).Rows(Z)("M01Acc_Name"))})
                        Z = Z + 1
                    Next
                    X = X + 1
                Next


                Sql = "select * from M01Account_Master where M01Main_Acc='" & Trim(M01.Tables(0).Rows(I)("M28Code")) & "' and len(M01Acc_Code)=5 and M01Com_Code='" & _Comcode & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                X = 0
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    dataTableSpinoffs.Rows.Add(New Object() {M02.Tables(0).Rows(X)("M01Acc_Code"), I, M02.Tables(0).Rows(X)("M01Acc_Code") & "-" & Trim(M02.Tables(0).Rows(X)("M01Acc_Name"))})
                    '    Z = 0
                    '   Sql = "select * from M01Account_Master where M01Main_Acc='" & Trim(M01.Tables(0).Rows(I)("M28Code")) & "' and len(M01Acc_Code)>4 and left(M01Acc_Code,4)='" & M02.Tables(0).Rows(X)("M01Acc_Code") & "' and M01Com_Code='" & _Comcode & "' order by M01Acc_Code"
                    '  M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    ' For Each DTRow4 As DataRow In M03.Tables(0).Rows
                    'dataTableSpinoffs.Rows.Add(New Object() {M03.Tables(0).Rows(Z)("M01Acc_Code"), M02.Tables(0).Rows(X)("M01Acc_Code"), M03.Tables(0).Rows(Z)("M01Acc_Code") & "-" & Trim(M03.Tables(0).Rows(Z)("M01Acc_Name"))})
                    'Z = Z + 1
                    'Next
                    X = X + 1
                Next

                I = I + 1
            Next

            dataSet.Tables.Add(dataTableSeries)
            dataSet.Tables.Add(dataTableSpinoffs)

            dataSet.Relations.Add("SpinoffToSeries", dataTableSeries.Columns("RecID"), dataTableSpinoffs.Columns("ParentRecordID"), False)
            dataSet.Relations.Add("SpinoffToSpinoff", dataTableSpinoffs.Columns("RecordID"), dataTableSpinoffs.Columns("ParentRecordID"), False)
            ' dataSet.Relations.Add("SpinoffToSpinoff1", dataTableSpinoffs.Columns("RecordID"), dataTableSpinoffs.Columns("ParentRecordID"), False)

            '' dataSet.Relations.Add("SpinoffToSeries", dataTableSeries.Columns("RecID"), dataTableSpinoffs.Columns("ParentRecordID"), False)
            'dataSet.Relations.Add("SpinoffToSpinoff", dataTableSpinoffs.Columns("RecordID"), dataTableSpinoffs.Columns("ParentRecordID"), False)
            'dataSet.Relations.Add("SpinoffToSpinoff1", dataTableSpinoffs.Columns("RecordID"), dataTableSpinoffs.Columns("ParentRecordID"), False)

            Return dataSet

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub ultraTree1_InitializeDataNode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.InitializeDataNodeEventArgs) Handles UltraTree1.InitializeDataNode

        '	If the node is being initialized for the first time,
        '	set some appearance properties for root nodes
        If e.Reinitialize = False AndAlso e.Node.Level = 0 Then
            e.Node.Override.NodeAppearance.BackColor = Office2003Colors.ToolbarGradientLight
            e.Node.Override.NodeAppearance.BackColor2 = Office2003Colors.ToolbarGradientDark
            e.Node.Override.NodeAppearance.BackGradientStyle = GradientStyle.Vertical
            e.Node.Override.NodeAppearance.ForeColor = Color.DarkBlue
            e.Node.Override.NodeAppearance.FontData.Bold = DefaultableBoolean.True
        End If

        '	Expand each node as it is initialized
        e.Node.Expanded = True

    End Sub

    Private Sub ultraTree1_ColumnSetGenerated(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.ColumnSetGeneratedEventArgs) Handles UltraTree1.ColumnSetGenerated
        Dim i As Integer
        For i = 0 To e.ColumnSet.Columns.Count - 1

            Dim column As UltraTreeNodeColumn = e.ColumnSet.Columns(i)

            '	Format the DateTime columns with the current culture's ShortDatePattern
            If column.DataType Is GetType(DateTime) Then
                column.Format = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
            End If

            '	Hide the integer columns, which represent record IDs
            If (column.DataType Is GetType(Integer)) Then
                column.Visible = False
            End If

            '	Set the MapToColumn for the relationship ColumnSets
            If e.ColumnSet.Key = "SpinoffToSeries" Or e.ColumnSet.Key = "SpinoffToSpinoff" Then
                ' column.MapToColumn = column.Key
            End If

            '	Set the Text property to something more descriptive
            If column.Key = "Account Name" Then

                column.Text = "Account Name"

                '	Make the name column wider
                column.LayoutInfo.PreferredLabelSize = New Size(Me.UltraTree1.Width / 2, 0)


            End If
        Next i

    End Sub

 
End Class