
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_Distributors
Public Class frmEngineering
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim T As Boolean



    Private Sub frmEngineering_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRef.ReadOnly = True

        txtD1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtD2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtW3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtW1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtW2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtE.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtM1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtM2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtF1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtF2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtEDate.Text = Today
        txtFromDate.Text = Today


    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Search_RefDoc()

        Clicked = "ADD"
        OPR0.Enabled = True
        OPR2.Enabled = True
        OPR1.Enabled = True
        OPR3.Enabled = True
        OPR4.Enabled = True
        OPR10.Enabled = True

        'OPR9.Enabled = TrueK
        txtFromDate.Text = Today
        cmdAdd.Enabled = False
        cmdSave.Enabled = True
        txtF1.Focus()
    End Sub

    Function Search_RefDoc()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select * from P01Parameter where P01code='DU'"
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

    Private Sub txtF1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtF2.Focus()
            End If
        ElseIf e.KeyCode = 46 Then
            lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
        ElseIf e.KeyCode = 8 Then
            lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
        End If
    End Sub



    Private Sub txtF1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF1.ValueChanged
        T = False
        If Trim(txtF1.Text) <> "" Then
            If IsNumeric(txtF1.Text) Then
                lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Thermic Heater", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtF1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtF2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtF2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtD1.Focus()
            End If
        ElseIf e.KeyCode = 46 Then
            lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
        ElseIf e.KeyCode = 8 Then
            lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
        End If

    End Sub

    Private Sub txtF2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF2.ValueChanged
        T = False
        If Trim(txtF2.Text) <> "" Then
            If IsNumeric(txtF2.Text) Then
                lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Steam Boiler", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtF2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtD2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtD2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtE.Focus()

            End If

        ElseIf e.KeyCode = 46 Then
            lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
        ElseIf e.KeyCode = 8 Then
            lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
        End If
    End Sub

    Private Sub txtD2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtD2.ValueChanged
        T = False
        If Trim(txtD2.Text) <> "" Then
            If IsNumeric(txtD2.Text) Then
                lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Furnace oil for Generators", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtD1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtD1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtD2.Focus()
            End If
        ElseIf e.KeyCode = 46 Then
            lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
        ElseIf e.KeyCode = 8 Then
            lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
        End If
    End Sub

    Private Sub txtD1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtD1.ValueChanged
        T = False
        If Trim(txtD1.Text) <> "" Then
            If IsNumeric(txtD1.Text) Then
                lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Furnace oil for Fork Lift", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtD1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtE_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtE.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtW1.Focus()
            End If
        End If
    End Sub

    Private Sub txtE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtE.ValueChanged
        T = False
        If Trim(txtE.Text) <> "" Then
            If IsNumeric(txtE.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Electricity - KWH", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtE
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtW1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtW1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then

                txtW2.Focus()

            End If
        ElseIf e.KeyCode = 46 Then
            lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
        ElseIf e.KeyCode = 46 Then
            lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
        End If
    End Sub

    Private Sub txtW1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtW1.ValueChanged
        T = False
        If Trim(txtW1.Text) <> "" Then
            If IsNumeric(txtW1.Text) Then
                lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct water consumption -Direct BOI - (m3)", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtW1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtW2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtW2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtW3.Focus()
            End If
        ElseIf e.KeyCode = 46 Then
            lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
        ElseIf e.KeyCode = 46 Then
            lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
        End If
    End Sub

    Private Sub txtW2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtW2.ValueChanged
        T = False
        If Trim(txtW2.Text) <> "" Then
            If IsNumeric(txtW2.Text) Then
                lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct water consumption -River Water - (m3)", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtW2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtW3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtW3.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtM1.Focus()
            End If
        ElseIf e.KeyCode = 46 Then
            lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
        ElseIf e.KeyCode = 8 Then
            lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
        End If
    End Sub

    Private Sub txtW3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtW3.ValueChanged
        T = False
        If Trim(txtW3.Text) <> "" Then
            If IsNumeric(txtW3.Text) Then
                lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Water consumption for Dyeing", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtW3
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtM1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtM1.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                txtM2.Focus()
            End If
        End If
    End Sub

    Private Sub txtM2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtM2.KeyUp
        If e.KeyCode = 13 Then
            If T = False Then
                cmdSave.Focus()

            End If
        End If

    End Sub

    Private Sub txtM2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtM2.ValueChanged
        T = False
        If Trim(txtM2.Text) <> "" Then
            If IsNumeric(txtM2.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Water waste New Plant", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtM2
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtM1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtM1.ValueChanged
        T = False
        If Trim(txtM1.Text) <> "" Then
            If IsNumeric(txtM1.Text) Then
            Else
                Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Water waste Old Plant", _
                                                "Information ...", _
                                                MessageBoxButtons.OK, _
                                                MessageBoxIcon.Information, _
                                                MessageBoxDefaultButton.Button2)
                If result3 = Windows.Forms.DialogResult.OK Then
                    T = True
                    With txtM1
                        .Focus()
                        .SelectionStart = 0
                        .SelectionLength = Len(.Text)
                    End With
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim nvcFieldList1 As String
        Dim Sql As String
        Dim X1 As Integer
        '  _TotalHR = "00:00"
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Dim M04Lot As DataSet
        Dim nvcVccode As String

        Dim ncQryType As String
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim hh1 As Integer
        Dim mm1 As Integer
        Dim _TimeDifferance1 As Date
        Dim n_year As Integer

        Dim vMax As Integer
        ncQryType = "ADD"

        Dim _WEEKNO As Integer

        Try
            T = False
            If Trim(txtF1.Text) <> "" Then
                If IsNumeric(txtF1.Text) Then
                    lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Thermic Heater", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtF1
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If


            T = False
            If Trim(txtF2.Text) <> "" Then
                If IsNumeric(txtF2.Text) Then
                    lblF.Text = Val(txtF1.Text) + Val(txtF2.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Steam Boiler", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtF2
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtD1.Text) <> "" Then
                If IsNumeric(txtD1.Text) Then
                    lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Furnace oil for Fork Lift", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtD1
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtD2.Text) <> "" Then
                If IsNumeric(txtD2.Text) Then
                    lblD.Text = Val(txtD1.Text) + Val(txtD2.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Furnace oil for Generators", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtD2
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If


            T = False
            If Trim(txtE.Text) <> "" Then
                If IsNumeric(txtE.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Electricity - KWH", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtE
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtW1.Text) <> "" Then
                If IsNumeric(txtW1.Text) Then
                    lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct water consumption -Direct BOI - (m3)", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtW1
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtW2.Text) <> "" Then
                If IsNumeric(txtW2.Text) Then
                    lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct water consumption -River Water - (m3)", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtW2
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtW3.Text) <> "" Then
                If IsNumeric(txtW3.Text) Then
                    lblW.Text = Val(txtW1.Text) + Val(txtW2.Text) + Val(txtW3.Text)
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Water consumption for Dyeing", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtW3
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtM1.Text) <> "" Then
                If IsNumeric(txtM1.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Water waste Old Plant", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtM1
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            T = False
            If Trim(txtM2.Text) <> "" Then
                If IsNumeric(txtM2.Text) Then
                Else
                    Dim result3 As DialogResult = MessageBox.Show("Please enter the correct Water waste New Plant", _
                                                    "Information ...", _
                                                    MessageBoxButtons.OK, _
                                                    MessageBoxIcon.Information, _
                                                    MessageBoxDefaultButton.Button2)
                    If result3 = Windows.Forms.DialogResult.OK Then
                        T = True
                        With txtM2
                            .Focus()
                            .SelectionStart = 0
                            .SelectionLength = Len(.Text)
                        End With
                        Exit Sub
                    End If
                End If
            End If

            '-------------------------------------------------------------------------------------
            'UPDATE RECORDS

            Dim thisCulture = Globalization.CultureInfo.CurrentCulture
            Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(txtEDate.Text)
            ' dayOfWeek.ToString() would return "Sunday" but it's an enum value,
            ' the correct dayname can be retrieved via DateTimeFormat.
            ' Following returns "Sonntag" for me since i'm in germany '
            Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)


            ' MsgBox(dayName)
            If dayName = "Sunday" Then
                Dim N_Date1 As Date
                N_Date1 = CDate(txtEDate.Text).AddDays(-1)
                _WEEKNO = DatePart(DateInterval.WeekOfYear, N_Date1)
                ' _WeekDis = "Week" & CStr(_WEEKNO)
                n_year = Microsoft.VisualBasic.Year(N_Date1)
            Else

                If DatePart(DateInterval.WeekOfYear, CDate(txtEDate.Text)) = 53 Then
                    Dim N_Date1 As Date
                    N_Date1 = CDate(txtEDate.Text).AddDays(+6)
                    _WEEKNO = 1

                Else
                    _WEEKNO = DatePart(DateInterval.WeekOfYear, CDate(txtEDate.Text))
                End If
                'If txtEDate.Text = "12/31/" & Microsoft.VisualBasic.Year(txtEDate.Text) Then
                '    Dim N_Date1 As Date
                '    N_Date1 = CDate(txtEDate.Text).AddDays(+1)
                '    _WEEKNO = DatePart(DateInterval.WeekOfYear, N_Date1)
                '    ' _WeekDis = "Week" & CStr(_WEEKNO)
                '    n_year = Microsoft.VisualBasic.Year(N_Date1)
                'Else

                '    _WEEKNO = DatePart(DateInterval.WeekOfYear, CDate(txtEDate.Text))
                '    ' _WeekDis = "Week" & CStr(_WEEKNO)
                '    'n_year = Microsoft.VisualBasic.Year(_EndDate)
                'End If
            End If


            Sql = "select * from DAILY_UTILITY where Effect_Date='" & txtEDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            If isValidDataset(T01) Then
                nvcFieldList1 = "Update DAILY_UTILITY set Enter_Date='" & txtFromDate.Text & "',Thernic_Heater='" & txtF1.Text & "',Steam_Boil='" & txtF2.Text & "',D_ForkLift='" & txtD1.Text & "',D_Genarator='" & txtD2.Text & "',WC_DirectBOI='" & txtW1.Text & "',WC_River='" & txtW2.Text & "',WC_Dyeing='" & txtW3.Text & "',W_Old='" & txtM1.Text & "',W_New='" & txtM2.Text & "',WeekNo=" & _WEEKNO & ",Electracity='" & txtE.Text & "' where Effect_Date='" & txtEDate.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                vMax = Get_highestVouNumber()
                nvcFieldList1 = "(Ref_No," & "Enter_Date," & "Effect_Date," & "Thernic_Heater," & "Steam_Boil," & "D_ForkLift," & "D_Genarator," & "WC_DirectBOI," & "WC_River," & "WC_Dyeing," & "W_Old," & "W_New," & "WeekNo," & "Electracity) " & "values(" & txtRef.Text & ",'" & txtFromDate.Text & "','" & txtEDate.Text & "','" & txtF1.Text & "','" & txtF2.Text & "','" & txtD1.Text & "','" & txtD2.Text & "','" & txtW1.Text & "','" & txtW2.Text & "','" & txtW3.Text & "','" & txtM1.Text & "','" & txtM2.Text & "'," & _WEEKNO & ",'" & txtE.Text & "')"
                up_GetSetDAILY_UTILITY(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)

            End If


            MsgBox("Record updateing successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()


            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            common.ClearAll(OPR0, OPR1, OPR2, OPR3, OPR4, OPR10)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdAdd.Focus()
            Call Search_Records()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Search_Records()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet

        Try
            SQL = "select * from DAILY_UTILITY where Effect_Date='" & txtEDate.Text & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtRef.Text = T01.Tables(0).Rows(0)("Ref_No")
                txtFromDate.Text = T01.Tables(0).Rows(0)("Enter_Date")
                If IsDBNull(T01.Tables(0).Rows(0)("Electracity")) Then
                Else
                    txtE.Text = CInt(T01.Tables(0).Rows(0)("Electracity"))
                End If

                If IsDBNull(T01.Tables(0).Rows(0)("Thernic_Heater")) Then
                Else
                    txtF1.Text = CInt(T01.Tables(0).Rows(0)("Thernic_Heater"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("Steam_Boil")) Then
                Else
                    txtF2.Text = CInt(T01.Tables(0).Rows(0)("Steam_Boil"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("D_ForkLift")) Then
                Else
                    txtD1.Text = CInt(T01.Tables(0).Rows(0)("D_ForkLift"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("D_Genarator")) Then
                Else
                    txtD2.Text = CInt(T01.Tables(0).Rows(0)("D_Genarator"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("WC_DirectBOI")) Then
                Else
                    txtW1.Text = CInt(T01.Tables(0).Rows(0)("WC_DirectBOI"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("WC_River")) Then
                Else
                    txtW2.Text = CInt(T01.Tables(0).Rows(0)("WC_River"))
                End If
                'txtD1.Text = CInt(T01.Tables(0).Rows(0)("Dye_PQty_Colour"))
                'txtD2.Text = CInt(T01.Tables(0).Rows(0)("Dye_PQty_White"))
                'txtD3.Text = CInt(T01.Tables(0).Rows(0)("Dye_PQty_Mal"))
                If IsDBNull(T01.Tables(0).Rows(0)("WC_Dyeing")) Then
                Else
                    txtW3.Text = CInt(T01.Tables(0).Rows(0)("WC_Dyeing"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("W_Old")) Then
                Else
                    txtM1.Text = CInt(T01.Tables(0).Rows(0)("W_Old"))
                End If
                If IsDBNull(T01.Tables(0).Rows(0)("W_New")) Then
                Else
                    txtM2.Text = CInt(T01.Tables(0).Rows(0)("W_New"))
                End If
            Else
                Search_RefDoc()
                txtF1.Text = ""
                txtF2.Text = ""
                txtD1.Text = ""
                txtD2.Text = ""
                txtW1.Text = ""
                txtW2.Text = ""
                txtW3.Text = ""
                txtE.Text = ""
                txtM1.Text = ""
                txtM2.Text = ""

                lblD.Text = "00.00"
                lblF.Text = "00.00"
                lblW.Text = "00.00"
            End If



            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Function Get_highestVouNumber() As String
        Dim con = New SqlConnection()
        Dim vMax As String

        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetParameter", New SqlParameter("@cQryType", "UPD"), New SqlParameter("@vcCode", "DU"))
            If common.isValidDataset(dsUser) Then
                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    vMax = dsUser.Tables(0).Rows(0)("P01LastNo")
                    Return vMax
                Next
            Else
                MessageBox.Show("Record Not Found", "Textured Jersey", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
        '=========================================================================
        ' "asdasd"
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR1, OPR2, OPR3, OPR4, OPR10)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdAdd.Focus()
    End Sub

    Private Sub txtEDate_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtEDate.BeforeDropDown

    End Sub

    Private Sub txtEDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEDate.TextChanged
        Call Search_Records()
    End Sub
End Class