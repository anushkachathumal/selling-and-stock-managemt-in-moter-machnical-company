Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmSetupTime
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim _MCNo As String


    Function Search_MCNo() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from T03Machine where T03Name ='" & Trim(cboMachine.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _MCNo = M01.Tables(0).Rows(0)("T03Code")
                Search_MCNo = True
            Else
                Search_MCNo = False
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        OPR1.Enabled = True
        OPR2.Enabled = True
        OPR5.Enabled = True
        OPR6.Enabled = True
        OPR7.Enabled = True
        OPR8.Enabled = True

        ' Call Clear_Text()
        cmdAdd.Enabled = False
        cboMachine.ToggleDropdown()

    End Sub


    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select T03name as [Machine Name] from T03Machine"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMachine
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275
            End With
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub frmSetupTime_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtLoding.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtUnloading.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtNomal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPower.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCharge.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFmachine.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFLight.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDDark.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDExtra.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDMedium.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFDark.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFExtra.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFLight.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFmachine.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFMedium.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCDark.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCExtra.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCharge.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCLight.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCMedium.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPM.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtOther.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDlight.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        Call Load_Combo()


    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cboMachine_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMachine.AfterCloseUp
        Call Search_Records()
    End Sub

    Private Sub cboMachine_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboMachine.InitializeLayout

    End Sub

    Private Sub cboMachine_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMachine.KeyUp
        If e.KeyCode = 13 Then

            Call Search_Records()
            txtLoding.Focus()
        End If
    End Sub

    Private Sub txtLoding_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLoding.KeyUp
        If e.KeyCode = 13 Then
            txtUnloading.Focus()
        End If
    End Sub

    Private Sub txtLoding_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLoding.ValueChanged

    End Sub

    Private Sub txtUnloading_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnloading.KeyUp
        If e.KeyCode = 13 Then
            txtNomal.Focus()
        End If
    End Sub

    Private Sub txtUnloading_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUnloading.ValueChanged

    End Sub

    Private Sub txtNomal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNomal.KeyUp
        If e.KeyCode = 13 Then
            txtPower.Focus()
        End If
    End Sub

    Private Sub txtNomal_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNomal.ValueChanged

    End Sub

    Private Sub txtPower_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPower.KeyUp
        If e.KeyCode = 13 Then
            txtCharge.Focus()
        End If
    End Sub

    Private Sub txtPower_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPower.ValueChanged

    End Sub

    Private Sub txtCharge_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCharge.KeyUp
        If e.KeyCode = 13 Then
            txtFmachine.Focus()
        End If
    End Sub

    Private Sub txtCharge_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCharge.ValueChanged

    End Sub

    Private Sub txtFmachine_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFmachine.KeyUp
        If e.KeyCode = 13 Then
            txtDlight.Focus()
        End If
    End Sub

    Private Sub txtFmachine_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFmachine.ValueChanged

    End Sub

    Private Sub txtDlight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDlight.KeyUp
        If e.KeyCode = 13 Then
            txtDMedium.Focus()
        End If
    End Sub

    Private Sub txtDlight_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDlight.ValueChanged

    End Sub

    Private Sub txtDMedium_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDMedium.KeyUp
        If e.KeyCode = 13 Then
            txtDDark.Focus()
        End If
    End Sub

    Private Sub txtDMedium_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDMedium.ValueChanged

    End Sub

    Private Sub txtDDark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDDark.KeyUp
        If e.KeyCode = 13 Then
            txtDExtra.Focus()
        End If
    End Sub

    Private Sub txtDDark_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDDark.ValueChanged

    End Sub

    Private Sub txtDExtra_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDExtra.KeyUp
        If e.KeyCode = 13 Then
            txtFLight.Focus()
        End If
    End Sub

    Private Sub txtDExtra_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDExtra.ValueChanged

    End Sub

    Private Sub txtFLight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFLight.KeyUp
        If e.KeyCode = 13 Then
            txtFMedium.Focus()
        End If
    End Sub

    Private Sub txtFLight_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFLight.ValueChanged

    End Sub

    Private Sub txtFMedium_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFMedium.KeyUp
        If e.KeyCode = 13 Then
            txtFDark.Focus()
        End If
    End Sub

    Private Sub txtFMedium_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFMedium.ValueChanged

    End Sub

    Private Sub txtFDark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFDark.KeyUp
        If e.KeyCode = 13 Then
            txtFExtra.Focus()
        End If
    End Sub

    Private Sub txtFDark_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFDark.ValueChanged

    End Sub

    Private Sub txtFExtra_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFExtra.KeyUp
        If e.KeyCode = 13 Then
            txtCLight.Focus()
        End If
    End Sub

    Private Sub txtFExtra_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFExtra.ValueChanged

    End Sub

    Private Sub txtCLight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCLight.KeyUp
        If e.KeyCode = 13 Then
            txtCMedium.Focus()
        End If
    End Sub

    Private Sub txtCLight_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCLight.ValueChanged

    End Sub

    Private Sub txtCMedium_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCMedium.KeyUp
        If e.KeyCode = 13 Then
            txtCDark.Focus()
        End If
    End Sub

    Private Sub txtCMedium_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCMedium.ValueChanged

    End Sub

    Private Sub txtCDark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCDark.KeyUp
        If e.KeyCode = 13 Then
            txtCExtra.Focus()
        End If
    End Sub

    Private Sub txtCDark_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCDark.ValueChanged

    End Sub

    Private Sub txtCExtra_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCExtra.KeyUp
        If e.KeyCode = 13 Then
            txtPM.Focus()
        End If
    End Sub

    Private Sub txtCExtra_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCExtra.ValueChanged

    End Sub

    Private Sub txtPM_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPM.KeyUp
        If e.KeyCode = 13 Then
            txtOther.Focus()
        End If
    End Sub

    Private Sub txtPM_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPM.ValueChanged

    End Sub

    Private Sub txtOther_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOther.KeyUp
        If e.KeyCode = 13 Then
            If cmdEdit.Enabled = True Then
                cmdEdit.Focus()
            Else
                cmdSave.Focus()
            End If
        End If
    End Sub

    Private Sub txtOther_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOther.ValueChanged

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

        Dim M01 As DataSet
        Try
            If Trim(txtLoding.Text) <> "" Then
                If IsNumeric(txtLoding.Text) Then
                Else
                    MsgBox("Please enter the correct Loding value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If
            If Trim(txtUnloading.Text) <> "" Then
                If IsNumeric(txtUnloading.Text) Then
                Else
                    MsgBox("Please enter the correct Unloding value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtNomal.Text) <> "" Then
                If IsNumeric(txtNomal.Text) Then
                Else
                    MsgBox("Please enter the correct Nomal value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtPower.Text) <> "" Then
                If IsNumeric(txtPower.Text) Then
                Else
                    MsgBox("Please enter the correct Power value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCharge.Text) <> "" Then
                If IsNumeric(txtCharge.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFmachine.Text) <> "" Then
                If IsNumeric(txtFmachine.Text) Then
                Else
                    MsgBox("Please enter the correct Machine value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDlight.Text) <> "" Then
                If IsNumeric(txtDlight.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Light value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDMedium.Text) <> "" Then
                If IsNumeric(txtDMedium.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Medium value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDDark.Text) <> "" Then
                If IsNumeric(txtDDark.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDExtra.Text) <> "" Then
                If IsNumeric(txtDExtra.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Extra Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFLight.Text) <> "" Then
                If IsNumeric(txtFLight.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Light value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFMedium.Text) <> "" Then
                If IsNumeric(txtFMedium.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Medium value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFDark.Text) <> "" Then
                If IsNumeric(txtFDark.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFExtra.Text) <> "" Then
                If IsNumeric(txtFExtra.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Extra Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCLight.Text) <> "" Then
                If IsNumeric(txtCLight.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Light value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCMedium.Text) <> "" Then
                If IsNumeric(txtCMedium.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Medium value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCDark.Text) <> "" Then
                If IsNumeric(txtCDark.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCExtra.Text) <> "" Then
                If IsNumeric(txtCExtra.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Extra Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtPM.Text) <> "" Then
                If IsNumeric(txtPM.Text) Then
                Else
                    MsgBox("Please enter the correct PM Shadule Time", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtOther.Text) <> "" Then
                If IsNumeric(txtOther.Text) Then
                Else
                    MsgBox("Please enter the correct Other Shadule Time", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If
            If Search_MCNo() = True Then

            Else
                MsgBox("Please select the correct Machine ", MsgBoxStyle.Information, "Textued Jersey .........")
                cboMachine.ToggleDropdown()
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------
            'UPDATE RECORDS
            nvcFieldList1 = "Insert Into M012Setup_Time(M012MCCode,M012Loading,M012Unloading,M012Nomal,M012Power,M012Charge,M012Machine,M012DLight,M012DMedium,M012DDark,M012DExtra,M012FLight,M012FMedium,M012FDark,M012FExtra,M012CLight,M012CMedium,M012CDark,M012CExtra,M012PM,M012Other)" & _
                                                                                         " values('" & _MCNo & "'," & txtLoding.Text & "," & txtUnloading.Text & "," & txtNomal.Text & "," & Trim(txtPower.Text) & "," & txtCharge.Text & "," & txtFmachine.Text & "," & txtDlight.Text & "," & txtDMedium.Text & "," & txtDDark.Text & "," & txtDlight.Text & "," & txtFLight.Text & "," & txtFMedium.Text & "," & txtFDark.Text & "," & txtFExtra.Text & "," & txtCLight.Text & "," & txtCMedium.Text & "," & txtCDark.Text & "," & txtCExtra.Text & "," & txtPM.Text & "," & txtOther.Text & ")"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            MsgBox("Records update successfully", MsgBoxStyle.Information, "Textued Jersey .......")
            transaction.Commit()
            common.ClearAll(OPR0, OPR1, OPR2, OPR6, OPR7, OPR8, OPR5)
            Clicked = ""



            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR1, OPR2, OPR6, OPR7, OPR8, OPR5)
        Clicked = ""

        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdAdd.Focus()
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M012Setup_Time inner join T03Machine on T03code=M012MCCode where T03Name='" & Trim(cboMachine.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01.Tables(0)
                    txtLoding.Text = .Rows(0)("M012Loading")
                    txtUnloading.Text = .Rows(0)("M012Unloading")
                    txtNomal.Text = .Rows(0)("M012Nomal")
                    txtPower.Text = .Rows(0)("M012Power")
                    txtCharge.Text = .Rows(0)("M012Charge")
                    txtFmachine.Text = .Rows(0)("M012Machine")
                    txtDlight.Text = .Rows(0)("M012DLight")
                    txtDMedium.Text = .Rows(0)("M012DMedium")
                    txtDDark.Text = .Rows(0)("M012DDark")
                    txtDExtra.Text = .Rows(0)("M012DExtra")
                    txtFLight.Text = .Rows(0)("M012FLight")
                    txtFMedium.Text = .Rows(0)("M012FMedium")
                    txtFDark.Text = .Rows(0)("M012FDark")
                    txtFExtra.Text = .Rows(0)("M012FExtra")
                    txtCLight.Text = .Rows(0)("M012CLight")
                    txtCMedium.Text = .Rows(0)("M012CMedium")
                    txtCDark.Text = .Rows(0)("M012CDark")
                    txtCExtra.Text = .Rows(0)("M012CExtra")
                    txtPM.Text = .Rows(0)("M012PM")
                    txtOther.Text = .Rows(0)("M012Other")
                End With
                cmdDelete.Enabled = True
                cmdEdit.Enabled = True
                cmdSave.Enabled = False
            Else
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
                cmdSave.Enabled = True
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

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

        Dim M01 As DataSet
        Try
            If Trim(txtLoding.Text) <> "" Then
                If IsNumeric(txtLoding.Text) Then
                Else
                    MsgBox("Please enter the correct Loding value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If
            If Trim(txtUnloading.Text) <> "" Then
                If IsNumeric(txtUnloading.Text) Then
                Else
                    MsgBox("Please enter the correct Unloding value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtNomal.Text) <> "" Then
                If IsNumeric(txtNomal.Text) Then
                Else
                    MsgBox("Please enter the correct Nomal value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtPower.Text) <> "" Then
                If IsNumeric(txtPower.Text) Then
                Else
                    MsgBox("Please enter the correct Power value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCharge.Text) <> "" Then
                If IsNumeric(txtCharge.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFmachine.Text) <> "" Then
                If IsNumeric(txtFmachine.Text) Then
                Else
                    MsgBox("Please enter the correct Machine value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDlight.Text) <> "" Then
                If IsNumeric(txtDlight.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Light value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDMedium.Text) <> "" Then
                If IsNumeric(txtDMedium.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Medium value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDDark.Text) <> "" Then
                If IsNumeric(txtDDark.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtDExtra.Text) <> "" Then
                If IsNumeric(txtDExtra.Text) Then
                Else
                    MsgBox("Please enter the correct Drain Extra Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFLight.Text) <> "" Then
                If IsNumeric(txtFLight.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Light value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFMedium.Text) <> "" Then
                If IsNumeric(txtFMedium.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Medium value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFDark.Text) <> "" Then
                If IsNumeric(txtFDark.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtFExtra.Text) <> "" Then
                If IsNumeric(txtFExtra.Text) Then
                Else
                    MsgBox("Please enter the correct Filling to Machine- Extra Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCLight.Text) <> "" Then
                If IsNumeric(txtCLight.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Light value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCMedium.Text) <> "" Then
                If IsNumeric(txtCMedium.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Medium value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCDark.Text) <> "" Then
                If IsNumeric(txtCDark.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtCExtra.Text) <> "" Then
                If IsNumeric(txtCExtra.Text) Then
                Else
                    MsgBox("Please enter the correct Charge Tank Transfer- Extra Dark value", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtPM.Text) <> "" Then
                If IsNumeric(txtPM.Text) Then
                Else
                    MsgBox("Please enter the correct PM Shadule Time", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If

            If Trim(txtOther.Text) <> "" Then
                If IsNumeric(txtOther.Text) Then
                Else
                    MsgBox("Please enter the correct Other Shadule Time", MsgBoxStyle.Exclamation, "Textued Jersey .........")
                    Exit Sub
                End If
            End If
            If Search_MCNo() = True Then

            Else
                MsgBox("Please select the correct Machine ", MsgBoxStyle.Information, "Textued Jersey .........")
                cboMachine.ToggleDropdown()
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------
            'UPDATE RECORDS
            'nvcFieldList1 = "Insert Into M012Setup_Time(M012MCCode,M012Loading,M012Unloading,M012Nomal,M012Power,M012Charge,M012Machine,M012DLight,M012DMedium,M012DDark,M012DExtra,M012FLight,M012FMedium,M012FDark,M012FExtra,M012CLight,M012CMedium,M012CDark,M012CExtra,M012PM,M012Other)" & _
            '                                                                             " values('" & _MCNo & "'," & txtLoding.Text & "," & txtUnloading.Text & "," & txtNomal.Text & "," & Trim(txtPower.Text) & "," & txtCharge.Text & "," & txtFmachine.Text & "," & txtDlight.Text & "," & txtDMedium.Text & "," & txtDDark.Text & "," & txtDlight.Text & "," & txtFLight.Text & "," & txtFMedium.Text & "," & txtFDark.Text & "," & txtFExtra.Text & "," & txtCLight.Text & "," & txtCMedium.Text & "," & txtCDark.Text & "," & txtCExtra.Text & "," & txtPM.Text & "," & txtOther.Text & ")"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "UPDATE M012Setup_Time set M012Loading=" & txtLoding.Text & ",M012Unloading=" & txtUnloading.Text & ",M012Nomal=" & txtNomal.Text & ",M012Power=" & txtPower.Text & ",M012Charge=" & txtCharge.Text & ",M012Machine=" & txtFmachine.Text & ",M012DLight=" & txtDlight.Text & ",M012DMedium=" & txtDMedium.Text & ",M012DDark=" & txtDDark.Text & ",M012DExtra=" & txtDExtra.Text & ",M012FLight=" & txtFLight.Text & ",M012FMedium=" & txtFMedium.Text & ",M012FDark=" & txtFDark.Text & ",M012FExtra=" & txtFExtra.Text & ",M012CLight=" & txtCLight.Text & ",M012CMedium=" & txtCMedium.Text & ",M012CDark=" & txtCDark.Text & ",M012CExtra=" & txtCExtra.Text & ",M012PM=" & txtPM.Text & ",M012Other=" & txtOther.Text & " where M012MCCode='" & _MCNo & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            MsgBox("Records update successfully", MsgBoxStyle.Information, "Textued Jersey .......")
            transaction.Commit()
            common.ClearAll(OPR0, OPR1, OPR2, OPR6, OPR7, OPR8, OPR5)
            Clicked = ""



            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
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
        Dim M01 As DataSet
        Dim A As String
        Try
            A = MsgBox("Are you sure want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Textued Jersey .........")
            If A = vbYes Then

                Call Search_MCNo()
                nvcFieldList1 = "delete from M012Setup_Time where M012MCCode='" & _MCNo & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                MsgBox("Records update successfully", MsgBoxStyle.Information, "Textued Jersey .......")
                transaction.Commit()
                common.ClearAll(OPR0, OPR1, OPR2, OPR6, OPR7, OPR8, OPR5)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class