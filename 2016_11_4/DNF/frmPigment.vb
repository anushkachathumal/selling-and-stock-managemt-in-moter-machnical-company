
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_Distributors
Public Class frmPigment
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim T As Boolean



    Private Sub frmPigment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRef.ReadOnly = True


        Call Search_RefDoc()

        txtPigStDate.Text = Today
        txtPigStopDate.Text = Today

        txtPigStartTime.Text = TimeOfDay
        txtPigStartTime.MaskInput = "hh:mm"
        ' UltraDateTimeEditor1.FormatString = "HH:MM"
        txtPigStartTime.Enabled = True
        txtPigStartTime.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always
        txtPigStartTime.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never

        txtPigStopTime.Text = TimeOfDay
        txtPigStopTime.MaskInput = "hh:mm"
        ' UltraDateTimeEditor1.FormatString = "HH:MM"
        txtPigStopTime.Enabled = True
        txtPigStopTime.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always
        txtPigStopTime.DropDownButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Never


        Call Load_BatchNo()

        Call Load_Batchtype()

        txtTD1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD6.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD7.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD8.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTD9.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtDyeMc.ReadOnly = True
        txtDyed_Mtr.ReadOnly = True
        txtDyed_Mtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtShade.ReadOnly = True
        txtShadecode.ReadOnly = True
        txtQuality.ReadOnly = True
        txtBatch_Wgt.ReadOnly = True
        txtBatch_Wgt.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtConfactor.ReadOnly = True
        txtConfactor.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFinishQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFinish_Confact.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFinihMtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFinihMtr.ReadOnly = True

        cboBatch.ToggleDropdown()

    End Sub

    Function Load_BatchNo()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select M04Lotno as [Batch No] from M04Lot where M04programetype not in ('I','F')"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboBatch
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 175
            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Batchtype()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select m21name as [Batch Type] from M21Batch_Type"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboType
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 175
            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_RefDoc()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select * from P01Parameter where P01code='PIG'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtRef.Text = T01.Tables(0).Rows(0)("P01lastno")
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboBatch_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatch.AfterCloseUp
        Call Search_Records()
    End Sub

  
    Private Sub cboBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBatch.KeyUp
        If e.KeyCode = 13 Then
            cboType.ToggleDropdown()
            Call Search_Records()
        End If
    End Sub

    Function Search_Records() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim _Status As String

        If Trim(cboType.Text) = "Normal" Then
            _Status = "N"
        ElseIf Trim(cboType.Text) = "Other" Then
            _Status = "O"

        ElseIf Trim(cboType.Text) = "Redye" Then
            _Status = "R"

        ElseIf Trim(cboType.Text) = "Wash" Then
            _Status = "W"
        End If


        Try
            SQL = "select M09MKG,M04Batchwt,T03Name,M04Quality,M04Shade_Code,M04Shade from M04Lot inner join T03Machine on T03Code=M04Machine_No inner join M09Quality on M04Quality=M09Code where M04Lotno='" & Trim(cboBatch.Text) & "' and M04ProgrameType='" & _Status & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtDyeMc.Text = T01.Tables(0).Rows(0)("T03Name")
                txtQuality.Text = T01.Tables(0).Rows(0)("M04Quality")
                txtShadecode.Text = T01.Tables(0).Rows(0)("M04Shade_Code")
                txtShade.Text = T01.Tables(0).Rows(0)("M04Shade")
                txtBatch_Wgt.Text = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(0)("M04Batchwt"), "#.00")
                txtConfactor.Text = Microsoft.VisualBasic.Format(T01.Tables(0).Rows(0)("M09MKG"), "#.00")
                txtDyed_Mtr.Text = Microsoft.VisualBasic.Format(Val(T01.Tables(0).Rows(0)("M09MKG")) * Val(T01.Tables(0).Rows(0)("M04Batchwt")), "#.00")
            Else
                txtDyeMc.Text = ""
                txtQuality.Text = ""
                txtShadecode.Text = ""
                txtShade.Text = ""
                txtBatch_Wgt.Text = ""
                txtConfactor.Text = ""
                txtDyed_Mtr.Text = ""
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboType_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboType.AfterCloseUp
        Call Search_Records()
    End Sub

    Private Sub cboType_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboType.InitializeLayout

    End Sub

    Private Sub cboType_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboType.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
            txtTD1.Focus()

        End If
    End Sub

    Private Sub cboBatch_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboBatch.InitializeLayout

    End Sub

    Private Sub txtTD1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD2.Focus()
            End If
        End If
    End Sub

    Private Sub txtTD1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTD1.ValueChanged
        T = False
        If Trim(txtTD1.Text) <> "" Then
            If IsNumeric(txtTD1.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD3.Focus()
            End If
        End If
    End Sub

    Private Sub txtTD2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTD2.ValueChanged
        T = False
        If Trim(txtTD2.Text) <> "" Then
            If IsNumeric(txtTD2.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD3.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD4.Focus()

            End If
        End If
    End Sub

    Private Sub txtTD3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTD3.ValueChanged
        T = False
        If Trim(txtTD3.Text) <> "" Then
            If IsNumeric(txtTD3.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD4.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD5.Focus()

            End If
        End If
    End Sub

    Private Sub txtTD4_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTD4.ValueChanged
        T = False
        If Trim(txtTD4.Text) <> "" Then
            If IsNumeric(txtTD4.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD4
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD5.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD6.Focus()

            End If
        End If
    End Sub

    Private Sub txtTD5_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTD5.ValueChanged
        T = False
        If Trim(txtTD5.Text) <> "" Then
            If IsNumeric(txtTD5.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD5
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD6_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD6.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD7.Focus()

            End If
        End If
    End Sub

    Private Sub txtTD6_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTD6.ValueChanged
        T = False
        If Trim(txtTD6.Text) <> "" Then
            If IsNumeric(txtTD6.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD6
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD7_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD7.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD8.Focus()
            End If
        End If
    End Sub

    Private Sub txtTD7_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTD7.ValueChanged
        T = False
        If Trim(txtTD7.Text) <> "" Then
            If IsNumeric(txtTD7.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD7
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD8_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD8.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtTD9.Focus()

            End If
        End If
    End Sub

    Private Sub txtTD8_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTD8.ValueChanged
        T = False
        If Trim(txtTD8.Text) <> "" Then
            If IsNumeric(txtTD8.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD8
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTD9_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTD9.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtFinishQty.Focus()
            End If
        End If
    End Sub

    Private Sub txtTD9_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTD9.ValueChanged
        T = False
        If Trim(txtTD9.Text) <> "" Then
            If IsNumeric(txtTD9.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct TD Number", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtTD9
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtFinishQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFinishQty.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtFinish_Confact.Focus()
            End If
        End If
    End Sub

    Private Sub txtFinishQty_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFinishQty.ValueChanged
        T = False
        If Trim(txtFinishQty.Text) <> "" Then
            If IsNumeric(txtFinishQty.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Finish Qty", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtFinishQty
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtFinish_Confact_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFinish_Confact.ValueChanged
        T = False
        If Trim(txtFinish_Confact.Text) <> "" Then
            If IsNumeric(txtFinish_Confact.Text) Then

            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Finish Conversion Factor", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtFinish_Confact
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub UltraLabel24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel24.Click

    End Sub

    Private Sub UltraGroupBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox11.Click

    End Sub
End Class