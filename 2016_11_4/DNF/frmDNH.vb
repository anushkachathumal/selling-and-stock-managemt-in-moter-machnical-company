Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmDNH
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim _EralCode As Integer
    Dim _MCNo As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        OPR2.Enabled = True
        OPR1.Enabled = True
        OPR6.Enabled = True
        OPR7.Enabled = True
        OPR8.Enabled = True
        OPR9.Enabled = True
        txtDate.Text = Today

        ' Call Clear_Text()
        cmdAdd.Enabled = False
        cboBatch.ToggleDropdown()
        Call Search_RefNo()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub frmDNH_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Combo()
        Call Search_RefNo()

        txtDye.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M04Lotno as [Batch No] from M04Lot group by M04Lotno"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboBatch
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_EralCode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M04Lot where M04Lotno='" & cboBatch.Text & "' and M04ProgrameType='" & cboLot.Text & "' and M04Machine_No='" & _MCNo & "' and M04Shade='" & Trim(txtShade.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_EralCode = True
                _EralCode = M01.Tables(0).Rows(0)("M04Ref")
                txtDye.Text = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(0)("M04Batchwt"), "#.00")
            Else
                Search_EralCode = False
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_MCNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from T03Machine where T03Name ='" & Trim(txtMC.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _MCNo = M01.Tables(0).Rows(0)("T03Code")
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Function Load_ComboLOT()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M04ProgrameType as [LOT Type] from M04Lot where M04Lotno='" & Trim(cboBatch.Text) & "' group by M04ProgrameType"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboLot
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Search_RefNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from P01Parameter where P01Code='DH'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtInvoice.Text = M01.Tables(0).Rows(0)("P01LastNo")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Shade()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M04Shade as [Shade] from M04Lot inner join T03Machine on M04Machine_No=T03Code where M04ProgrameType='" & cboLot.Text & "' and M04Lotno='" & cboBatch.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtShade
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 285
            End With


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Function Search_LotDetailes() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M04Lot inner join T03Machine on M04Machine_No=T03Code where M04ProgrameType='" & cboLot.Text & "' and M04Lotno='" & cboBatch.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_LotDetailes = True
                txtQuality.Text = M01.Tables(0).Rows(0)("M04Quality")
                ' txtShade.Text = M01.Tables(0).Rows(0)("M04Shade")
                'txtMC.Text = M01.Tables(0).Rows(0)("T03Name")
                'txtDye.Text = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(0)("M04Batchwt"), "#.00")
                Sql = "select t03name as [Machine] from M04Lot inner join T03Machine on M04Machine_No=T03Code where M04ProgrameType='" & cboLot.Text & "' and M04Lotno='" & cboBatch.Text & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                With txtMC
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
                cmdSave.Enabled = True


            Else
                Search_LotDetailes = False
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Sub cboBatch_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatch.AfterCloseUp
        Call Load_ComboLOT()
    End Sub


    Private Sub cboBatch_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboBatch.InitializeLayout

    End Sub

    Private Sub cboBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBatch.KeyUp
        If e.KeyCode = 13 Then
            Call Load_ComboLOT()
            cboLot.ToggleDropdown()

        End If
    End Sub

    Private Sub cboLot_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLot.AfterCloseUp
        Call Search_LotDetailes()

    End Sub

    Private Sub cboLot_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboLot.InitializeLayout

    End Sub

    Private Sub cboLot_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLot.KeyUp
        If e.KeyCode = 13 Then
            Call Search_LotDetailes()
            txtSub.Focus()

        End If
    End Sub

    Private Sub chkBulk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBulk.CheckedChanged
        If chkBulk.Checked = True Then
            chkOn.Checked = False
        End If
    End Sub

    Private Sub chkOn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOn.CheckedChanged
        If chkOn.Checked = True Then
            chkBulk.Checked = False
        End If
    End Sub

    Private Sub chkPilot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPilot.CheckedChanged
        If chkPilot.Checked = True Then
            chkPigment.Checked = False
            chkDH.Checked = False
            chkUnlevel.Checked = False
        End If

    End Sub

    Private Sub chkPigment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPigment.CheckedChanged
        If chkPigment.Checked = True Then
            chkPilot.Checked = False
            chkDH.Checked = False
            chkUnlevel.Checked = False
        End If

    End Sub

    Private Sub chkDH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDH.CheckedChanged
        If chkDH.Checked = True Then
            chkPigment.Checked = False
            chkPilot.Checked = False
            chkUnlevel.Checked = False
        End If

    End Sub

    Private Sub chkUnlevel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUnlevel.CheckedChanged
        If chkUnlevel.Checked = True Then
            chkPigment.Checked = False
            chkDH.Checked = False
            chkPigment.Checked = False
        End If

    End Sub

    Private Sub chkMCChange_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMCChange.CheckedChanged
        'If chkMCChange.Checked = True Then
        '    chkSC.Checked = False
        '    chkProcess.Checked = False
        '    chkLiquor.Checked = False
        '    chkDyeLot.Checked = False

        'End If
    End Sub

    Private Sub chkSC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSC.CheckedChanged
        'If chkSC.Checked = True Then
        '    chkMCChange.Checked = False
        '    chkProcess.Checked = False
        '    chkLiquor.Checked = False
        '    chkDyeLot.Checked = False

        'End If

    End Sub

    Private Sub chkProcess_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkProcess.CheckedChanged
        'If chkProcess.Checked = True Then
        '    chkSC.Checked = False
        '    chkMCChange.Checked = False
        '    chkLiquor.Checked = False
        '    chkDyeLot.Checked = False

        'End If

    End Sub

    Private Sub chkLiquor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLiquor.CheckedChanged
        'If chkLiquor.Checked = True Then
        '    chkSC.Checked = False
        '    chkProcess.Checked = False
        '    chkMCChange.Checked = False
        '    chkDyeLot.Checked = False

        'End If

    End Sub

    Private Sub chkDyeLot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDyeLot.CheckedChanged
        'If chkDyeLot.Checked = True Then
        '    chkSC.Checked = False
        '    chkProcess.Checked = False
        '    chkLiquor.Checked = False
        '    chkMCChange.Checked = False

        'End If

    End Sub

    Private Sub chkYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkYes.CheckedChanged
        If chkYes.Checked = True Then
            chkNo.Checked = False
           

        End If
    End Sub

    Private Sub chkNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNo.CheckedChanged
        If chkNo.Checked = True Then
            chkYes.Checked = False
          

        End If

    End Sub

    Private Sub chkOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOff.CheckedChanged
        If chkOff.Checked = True Then
            chkU.Checked = False
            chkDyeS.Checked = False
            chkDyeM.Checked = False
            chkCF.Checked = False
            chkB.Checked = False
            chkR.Checked = False
        End If
    End Sub

    Private Sub chkU_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkU.CheckedChanged
        If chkU.Checked = True Then
            chkOff.Checked = False
            chkDyeS.Checked = False
            chkDyeM.Checked = False
            chkCF.Checked = False
            chkB.Checked = False
            chkR.Checked = False
        End If

    End Sub

    Private Sub chkDyeM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDyeM.CheckedChanged
        If chkDyeM.Checked = True Then
            chkU.Checked = False
            chkDyeS.Checked = False
            chkOff.Checked = False
            chkCF.Checked = False
            chkB.Checked = False
            chkR.Checked = False
        End If
    End Sub

    Private Sub chkCF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCF.CheckedChanged
        If chkCF.Checked = True Then
            chkU.Checked = False
            chkDyeS.Checked = False
            chkDyeM.Checked = False
            chkOff.Checked = False
            chkB.Checked = False
            chkR.Checked = False
        End If
    End Sub

    Private Sub chkDyeS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDyeS.CheckedChanged
        If chkDyeS.Checked = True Then
            chkU.Checked = False
            chkOff.Checked = False
            chkDyeM.Checked = False
            chkCF.Checked = False
            chkB.Checked = False
            chkR.Checked = False
        End If

    End Sub

    Private Sub chkB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkB.CheckedChanged
        If chkB.Checked = True Then
            chkU.Checked = False
            chkDyeS.Checked = False
            chkDyeM.Checked = False
            chkCF.Checked = False
            chkOff.Checked = False
            chkR.Checked = False
        End If

    End Sub

    Private Sub txtSub_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSub.KeyUp
        If e.KeyCode = 13 Then
            txtMC.ToggleDropdown()
        End If
    End Sub

    Private Sub txtSub_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSub.ValueChanged

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR1, OPR2, OPR6, OPR7, OPR8, OPR9)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        chkBulk.Checked = False
        chkOn.Checked = False

        chkPilot.Checked = False
        chkPigment.Checked = False
        chkDH.Checked = False
        chkUnlevel.Checked = False

        chkMCChange.Checked = False
        chkSC.Checked = False
        chkDyeLot.Checked = False
        chkProcess.Checked = False
        chkLiquor.Checked = False

        chkYes.Checked = False
        chkNo.Checked = False

        chkOff.Checked = False
        chkDyeM.Checked = False
        chkB.Checked = False
        chkCF.Checked = False
        chkU.Checked = False
        chkDyeS.Checked = False
        chkR.Checked = False
        cmdAdd.Focus()
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
        Dim M02 As DataSet

        Dim strBatch As String
        Dim StrDyeHouse As String
        Dim strReceip As String
        Dim strQC As String
        Dim StrStatus As String

        Try
            If Search_EralCode() = False Then
                MsgBox("Please enter the correct data", MsgBoxStyle.Information, "Textued Jersey .......")
                Exit Sub
            End If
            If Trim(txtSub.Text) <> "" Then
            Else
                MsgBox("Please enter the Sub No", MsgBoxStyle.Information, "Information ......")
                txtSub.Focus()
                Exit Sub
            End If
            '-------------------------------------------------------------------------------------
            If Trim(txtReject.Text) <> "" Then
                If IsNumeric(txtReject.Text) Then

                Else
                    MsgBox("Please enter the correct Reject Qty", MsgBoxStyle.Information, "Information .........")
                    txtReject.Focus()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Reject Qty", MsgBoxStyle.Information, "Information ......")
                txtReject.Focus()
                Exit Sub
            End If
            '-------------------------------------------------------------------------------------
            strBatch = ""
            If chkBulk.Checked = True Then
                strBatch = "B"
            ElseIf chkOn.Checked = True Then
                strBatch = "O"
            End If
            '----------------------------------------------------
            If strBatch <> "" Then
            Else
                MsgBox("Please enter the Batch Type", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
            '----------------------------------------------------
            StrDyeHouse = ""
            If chkPilot.Checked = True Then
                StrDyeHouse = "PI"
            ElseIf chkPigment.Checked = True Then
                StrDyeHouse = "PG"
            ElseIf chkDH.Checked = True Then
                StrDyeHouse = "DH"
            ElseIf chkUnlevel.Checked = True Then
                StrDyeHouse = "UL"
            End If
            If StrDyeHouse <> "" Then
            Else
                MsgBox("Please select the Dye house  Shade comments", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
            '-------------------------------------------------------
            strReceip = ""
            Dim SC As String
            Dim DL As String
            Dim PR As String
            Dim LQ As String
            Dim MC As String

            DL = "N"
            SC = "N"
            PR = "N"
            LQ = "N"
            MC = "N"
            If chkDyeLot.Checked = True Then
                DL = "Y"
            ElseIf chkSC.Checked = True Then
                SC = "Y"
            ElseIf chkProcess.Checked = True Then
                PR = "Y"
            ElseIf chkLiquor.Checked = True Then
                LQ = "Y"
            ElseIf chkMCChange.Checked = True Then
                MC = "Y"

            End If
            '---------------------------------------------------------
            'If strReceip <> "" Then
            'Else
            '    MsgBox("Please select the Recipe Detailes", MsgBoxStyle.Information, "Information ........")
            '    Exit Sub
            'End If
            '--------------------------------------------------------
            StrStatus = ""
            If chkYes.Checked = True Then
                StrStatus = "Y"
            ElseIf chkNo.Checked = True Then
                StrStatus = "N"
            End If

            If StrStatus <> "" Then
            Else
                MsgBox("Please select the Status", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
            '-------------------------------------------------------

            strQC = ""
            If chkCF.Checked = True Then
                strQC = "QC"
            ElseIf chkU.Checked = True Then
                strQC = "U"
            ElseIf chkDyeS.Checked = True Then
                strQC = "DS"
            ElseIf chkDyeM.Checked = True Then
                strQC = "DM"
            ElseIf chkOff.Checked = True Then
                strQC = "OF"
            ElseIf chkB.Checked = True Then
                strQC = "B"
            ElseIf chkR.Checked = True Then
                strQC = "R"
            End If

            'If strQC <> "" Then
            'Else
            '    MsgBox("Please select the QC & Exam comments", MsgBoxStyle.Information, "Information ........")
            '    Exit Sub
            'End If

            'BATCH TYPE
            ' chkBulk = strBatch = "B"
            'chkOn =strBatch = "O"

            'DYE HOUSE SHADE COMMENT
            'chkPilot =StrDyeHouse = "PI"
            'chkPigment =StrDyeHouse = "PG"
            'chkDH.Checked =  StrDyeHouse = "DH"
            'chkUnlevel = StrDyeHouse = "UL"

            'RECIPE DETAILES
            'chkDyeLot strReceip = "DL"
            'chkSC strReceip = "SC"
            'chkProcess strReceip = "PR"
            'chkLiquor strReceip = "LQ"
            'chkMCChange strReceip = "MC"

           
            nvcFieldList1 = "select * from M04Lot where M04Ref='" & _EralCode & "' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then

                nvcFieldList1 = "select * from T03DNH where T03Ecode='" & _EralCode & "' "
                M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M02) Then
                    nvcFieldList1 = "update T03DNH set T03QC='" & strQC & "' where T03Ref='" & txtInvoice.Text & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    MsgBox("Records update successfully", MsgBoxStyle.Information, "Textured Jersey .........")
                    transaction.Commit()
                Else
                    Call Search_RefNo()

                    nvcFieldList1 = "Update P01Parameter set P01LastNo = " & Val(txtInvoice.Text) + 1 & " where P01Code = 'DH'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into T03DNH(T03Ref,T03Date,T03Batch,T03LotType,T03SubNo,T03Reject,T03Batchtype,T03DyeH,T03MC,T03WetOn,T03QC,T03Remark,T03Ongoin,T03User1,T03User2,T03SC,T03Pro,T03Liq,T03Dye,T03ECode)" & _
                                                                                       " values(" & Trim(txtInvoice.Text) & ",'" & txtDate.Text & "','" & cboBatch.Text & "','" & cboLot.Text & "','" & Trim(txtSub.Text) & "','" & Trim(txtReject.Text) & "','" & strBatch & "','" & StrDyeHouse & "','" & MC & "','" & StrStatus & "','" & strQC & "','" & txtReson.Text & "','" & txtOngoing.Text & "','" & strDisname & "','" & strDisname & "','" & SC & "','" & PR & "','" & LQ & "','" & DL & "','" & _EralCode & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    MsgBox("Records update successfully", MsgBoxStyle.Information, "Textured Jersey .........")
                    transaction.Commit()
                End If
            End If
            common.ClearAll(OPR0, OPR1, OPR2, OPR6, OPR7, OPR8, OPR9)
            Clicked = ""

            chkBulk.Checked = False
            chkOn.Checked = False

            chkPilot.Checked = False
            chkPigment.Checked = False
            chkDH.Checked = False
            chkUnlevel.Checked = False

            chkMCChange.Checked = False
            chkSC.Checked = False
            chkDyeLot.Checked = False
            chkProcess.Checked = False
            chkLiquor.Checked = False

            chkYes.Checked = False
            chkNo.Checked = False

            chkOff.Checked = False
            chkDyeM.Checked = False
            chkB.Checked = False
            chkCF.Checked = False
            chkU.Checked = False
            chkDyeS.Checked = False

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

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

    End Sub

    Private Sub chkR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkR.CheckedChanged
        If chkR.Checked = True Then
            chkU.Checked = False
            chkDyeS.Checked = False
            chkDyeM.Checked = False
            chkCF.Checked = False
            chkB.Checked = False
            chkOff.Checked = False
        End If
    End Sub

    Function Search_DNH()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select * from T03DNH where T03Ecode='" & _EralCode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01
                    txtInvoice.Text = .Tables(0).Rows(0)("T03Ref")
                    txtDate.Text = .Tables(0).Rows(0)("T03Date")
                    txtSub.Text = .Tables(0).Rows(0)("T03SubNo")
                    txtReject.Text = Microsoft.VisualBasic.Format(.Tables(0).Rows(0)("T03Reject"), "#.00")
                    If Trim(.Tables(0).Rows(0)("T03Batchtype")) = "B" Then
                        chkBulk.Checked = True
                    Else
                        chkOn.Checked = True
                    End If

                    If Trim(.Tables(0).Rows(0)("T03DyeH")) = "PI" Then
                        chkPilot.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T03DyeH")) = "PG" Then
                        chkPigment.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T03DyeH")) = "DH" Then
                        chkDH.Checked = True
                    ElseIf Trim(.Tables(0).Rows(0)("T03DyeH")) = "UL" Then
                        chkUnlevel.Checked = True
                    End If

                    If Trim(.Tables(0).Rows(0)("T03MC")) = "Y" Then
                        chkMCChange.Checked = True
                    End If


                    If Trim(.Tables(0).Rows(0)("T03Dye")) = "Y" Then
                        chkDyeLot.Checked = True
                    End If
                    If Trim(.Tables(0).Rows(0)("T03SC")) = "Y" Then
                        chkSC.Checked = True
                    End If
                    If Trim(.Tables(0).Rows(0)("T03Pro")) = "Y" Then
                        chkProcess.Checked = True
                    End If

                    If Trim(.Tables(0).Rows(0)("T03Liq")) = "Y" Then
                        chkLiquor.Checked = True
                    End If
                    '----------------------------------------------------------------
                    txtReson.Text = Trim(.Tables(0).Rows(0)("T03Remark"))
                    txtOngoing.Text = Trim(.Tables(0).Rows(0)("T03Ongoin"))

                    If Trim(.Tables(0).Rows(0)("T03WetOn")) = "Y" Then
                        chkYes.Checked = True
                    Else
                        chkNo.Checked = True
                    End If

                End With
            Else
                chkPilot.Checked = False
                chkPigment.Checked = False
                chkDH.Checked = False
                chkUnlevel.Checked = False

                chkMCChange.Checked = False
                chkSC.Checked = False
                chkDyeLot.Checked = False
                chkProcess.Checked = False
                chkLiquor.Checked = False

                chkYes.Checked = False
                chkNo.Checked = False

                chkOff.Checked = False
                chkDyeM.Checked = False
                chkB.Checked = False
                chkCF.Checked = False
                chkU.Checked = False
                chkDyeS.Checked = False
                chkR.Checked = False
                txtReject.Text = ""
                txtReson.Text = ""
                txtOngoing.Text = ""
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtMC_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMC.AfterCloseUp
        Call Load_Shade()
        Call Search_MCNo()
    End Sub

    Private Sub txtMC_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMC.KeyUp
        If e.KeyCode = 13 Then
            Call Load_Shade()
            Call Search_MCNo()
            txtShade.ToggleDropdown()
        End If
    End Sub

    Private Sub txtShade_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShade.AfterCloseUp
        Call Search_EralCode()
        Call Search_DNH()
    End Sub

    Private Sub txtShade_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles txtShade.InitializeLayout

    End Sub

    Private Sub txtShade_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShade.KeyUp
        If e.KeyCode = 13 Then
            Call Search_EralCode()
            Call Search_DNH()
            txtReject.Focus()
        End If
    End Sub

    Private Sub txtReson_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtReson.KeyUp
        If e.KeyCode = 13 Then
            txtOngoing.Focus()
        End If
    End Sub
End Class