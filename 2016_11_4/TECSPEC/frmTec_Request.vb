Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmTec_Request
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub

    Private Sub chk1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBrush_One.CheckedChanged
        If chkBrush_One.Checked = True Then
            chkBrush_Both.Checked = False
        End If
    End Sub

    Private Sub chk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBrush_Both.CheckedChanged
        If chkBrush_Both.Checked = True Then
            chkBrush_One.Checked = False
        End If
    End Sub

    Private Sub frmTec_Request_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtRequierd_Date.Text = Today
        txtReuest_Date.Text = Today
        txtOrder_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCuttable_Width.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTarger_Price.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTarget_Weight.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTarget_Width.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEnd_End.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYardage.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCustomer_Referance.ReadOnly = True

        Call Load_Combo_Customer() '------------------------------------>>> CUSTOMER DETAILES
        Call Load_Combo_Test_Standerd() '------------------------------->>> TEST STANDERD
        Call Load_Hangers() '--------------------------------------------->>> NO OF HANGERS

        chKApproval_Aesth.Checked = True

        cboCustomer.ToggleDropdown()
    End Sub

    Function Load_Combo_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select M01Cus_Name as [##] from M01_TEC_Customer order by M01Cus_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 280
                '   .Rows.Band.Columns(1).Width = 260


            End With
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Combo_Test_Standerd()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select M02DESCRIPTION as [##] from M02_TEC_TestStanderd "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboTest_Standard
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                '   .Rows.Band.Columns(1).Width = 260


            End With
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_Customer_RefNo() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select * from M01_TEC_Customer where M01Cus_Name='" & Trim(cboCustomer.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCustomer_Referance.Text = Trim(M01.Tables(0).Rows(0)("M01Cus_RefNo"))
                Search_Customer_RefNo = True
            End If
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cboCustomer_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomer.AfterCloseUp
        Call Search_Customer_RefNo()
    End Sub

    Private Sub cboCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCustomer.KeyUp
        If e.KeyCode = 13 Then
            txtOrder_Qty.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtOrder_Qty.Focus()
        End If
    End Sub

    Private Sub txtOrder_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrder_Qty.KeyUp
        If e.KeyCode = 13 Then
            txtTarger_Price.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtTarger_Price.Focus()
        End If
    End Sub

    Private Sub txtTarger_Price_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTarger_Price.KeyUp
        If e.KeyCode = 13 Then
            cboEnd_User.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboEnd_User.ToggleDropdown()
        End If
    End Sub

   
    Private Sub cboEnd_User_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEnd_User.KeyUp
        If e.KeyCode = 13 Then
            txtFabric_Des.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtFabric_Des.Focus()
        End If
    End Sub

    Private Sub txtCompositin_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCompositin.KeyUp
        If e.KeyCode = 13 Then
            txtTarget_Width.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtTarget_Width.Focus()
        End If
    End Sub

    Private Sub txtTarget_Width_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTarget_Width.KeyUp
        If e.KeyCode = 13 Then
            txtEnd_End.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtEnd_End.Focus()
        End If
    End Sub

    Private Sub txtEnd_End_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEnd_End.KeyUp
        If e.KeyCode = 13 Then
            txtCuttable_Width.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCuttable_Width.Focus()
        End If
    End Sub

    Private Sub txtCuttable_Width_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCuttable_Width.KeyUp
        If e.KeyCode = 13 Then
            txtTarget_Weight.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtTarget_Weight.Focus()
        End If
    End Sub

    Private Sub txtColour_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtColour.KeyUp
        If e.KeyCode = 13 Then
            txtColour_no.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtColour_no.Focus()
        End If
    End Sub

    Private Sub txtTarget_Weight_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTarget_Weight.KeyUp
        If e.KeyCode = 13 Then
            txtColour.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtColour.Focus()
        End If
    End Sub

    Private Sub txtColour_no_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtColour_no.KeyUp
        If e.KeyCode = 13 Then
            cboTest_Standard.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboTest_Standard.ToggleDropdown()
        End If
    End Sub

    Private Sub cboTest_Standard_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTest_Standard.KeyUp
        If e.KeyCode = 13 Then
            txtOther_Requierment.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtOther_Requierment.Focus()
        End If
    End Sub

    Private Sub chKApproval_Aesth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chKApproval_Aesth.CheckedChanged
        If chKApproval_Aesth.Checked = True Then
            chkApproval_Technical.Checked = False
        Else
            chkApproval_Technical.Checked = True
        End If
    End Sub

    Private Sub chkApproval_Technical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkApproval_Technical.CheckedChanged
        If chkApproval_Technical.Checked = True Then
            chKApproval_Aesth.Checked = False
        Else
            chKApproval_Aesth.Checked = True
        End If
    End Sub

    Private Sub UltraButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton17.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim _OriginalRef As String
        Dim _Submision As String
        Dim _TestReport As String
        Dim _Solid As String
        Dim _Marl As String
        Dim _Print As String
        Dim _DY As String
        Dim _PDY As String
        Dim _Brush As String
        Dim _Anti_Pill As String
        Dim _Bio_Polish As String
        Dim _Sude As String
        Dim _Ref As Integer
        Dim _REFST As String

        Try
            If Search_Customer_RefNo() = True Then
            Else
                MsgBox("Please select the correct customer", MsgBoxStyle.Information, "Information .......")
                cboCustomer.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

            If txtOrder_Qty.Text <> "" Then
                If IsNumeric(txtOrder_Qty.Text) Then

                Else
                    MsgBox("Please select the correct Order Qty", MsgBoxStyle.Information, "Information .......")
                    txtOrder_Qty.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Order Qty", MsgBoxStyle.Information, "Information .......")
                txtOrder_Qty.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtTarger_Price.Text <> "" Then
                If IsNumeric(txtTarger_Price.Text) Then

                Else
                    MsgBox("Please select the correct Target price", MsgBoxStyle.Information, "Information .......")
                    txtTarger_Price.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Target price", MsgBoxStyle.Information, "Information .......")
                txtTarger_Price.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtTarget_Weight.Text <> "" Then
                If IsNumeric(txtTarget_Weight.Text) Then

                Else
                    MsgBox("Please select the correct Target weight", MsgBoxStyle.Information, "Information .......")
                    txtTarget_Weight.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Target weight", MsgBoxStyle.Information, "Information .......")
                txtTarget_Weight.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtTarget_Width.Text <> "" Then
                If IsNumeric(txtTarget_Width.Text) Then

                Else
                    MsgBox("Please select the correct Target width", MsgBoxStyle.Information, "Information .......")
                    txtTarget_Width.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Target width", MsgBoxStyle.Information, "Information .......")
                txtTarget_Width.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtEnd_End.Text <> "" Then
                If IsNumeric(txtEnd_End.Text) Then

                Else
                    MsgBox("Please select the correct end to end width", MsgBoxStyle.Information, "Information .......")
                    txtEnd_End.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the end to end width", MsgBoxStyle.Information, "Information .......")
                txtEnd_End.Focus()
                connection.Close()
                Exit Sub
            End If


            If txtCuttable_Width.Text <> "" Then
                If IsNumeric(txtCuttable_Width.Text) Then

                Else
                    MsgBox("Please select the correct cuttable width", MsgBoxStyle.Information, "Information .......")
                    txtCuttable_Width.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the cuttable width", MsgBoxStyle.Information, "Information .......")
                txtCuttable_Width.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtCompositin.Text <> "" Then
            Else
                MsgBox("Please enter the composition", MsgBoxStyle.Information, "Information .......")
                txtCompositin.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtFabric_Des.Text <> "" Then
            Else
                MsgBox("Please enter the Fabric Description", MsgBoxStyle.Information, "Information .......")
                txtFabric_Des.Focus()
                connection.Close()
                Exit Sub
            End If

            If cboTest_Standard.Text <> "" Then
            Else
                MsgBox("Please enter the Test Standard", MsgBoxStyle.Information, "Information .......")
                cboTest_Standard.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

            If cboHanger.Text <> "" Then
            Else
                MsgBox("Please enter the No of Hangers", MsgBoxStyle.Information, "Information .......")
                cboHanger.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

            If txtYardage.Text <> "" Then
                If IsNumeric(txtYardage.Text) Then

                Else
                    MsgBox("Please select the correct Yarderge", MsgBoxStyle.Information, "Information .......")
                    txtYardage.Focus()
                    connection.Close()
                    Exit Sub
                End If
            Else
                MsgBox("Please enter the Yarderge", MsgBoxStyle.Information, "Information .......")
                txtYardage.Focus()
                connection.Close()
                Exit Sub
            End If


            'OIGINAL REFERANCE
            _OriginalRef = "No"

            If chkCounter_Sample.Checked = True Then
                _OriginalRef = "1"
            End If

            If chkCustomer_Spec.Checked = True Then
                _OriginalRef = "2"
            End If

            '-----------------------------------------------------------
            'SUBMISION FORMATE
            If chKApproval_Aesth.Checked = True Then
                _Submision = "1"
            End If

            If chkApproval_Technical.Checked = True Then
                _Submision = "2"
            End If
            '------------------------------------------------------------
            'TEST REPORT
            _TestReport = "No"
            If chkTest_Internal.Checked = True Then
                _TestReport = "1"
            End If

            If chkTest_Outside.Checked = True Then
                _TestReport = "2"
            End If
            '------------------------------------------------------------
            'SPECIAL APPLICATION
            _Solid = "No"
            _PDY = "No"
            _DY = "No"
            _Marl = "No"
            _Print = "No"

            If chkSolid.Checked = True Then
                _Solid = "YES"
            End If

            If chkMarl.Checked = True Then
                _Marl = "YES"
            End If
            If chkPrint.Checked = True Then
                _Print = "YES"
            End If
            If chkDyed_Yarn.Checked = True Then
                _DY = "YES"
            End If
            If chkPies_Dyed.Checked = True Then
                _PDY = "YES"
            End If
            '--------------------------------------------------------------
            'BRUSH
            _Brush = "No"
            If chkBrush_One.Checked = True Then
                _Brush = "1"
            End If

            If chkBrush_Both.Checked = True Then
                _Brush = "2"
            End If
            '----------------------------------------------------------------
            'ANTI PILL
            _Anti_Pill = "No"
            If chkAntipill_One.Checked = True Then
                _Anti_Pill = "1"
            End If
            If chkAntipill_Both.Checked = True Then
                _Anti_Pill = "2"
            End If

            _Sude = "No"
            _Bio_Polish = "No"

            If chkSueded.Checked = True Then
                _Sude = "YES"
            End If

            If chkBio_Polish.Checked = True Then
                _Bio_Polish = "YES"
            End If

            nvcFieldList1 = "select * from T01_TEC_Development_Request where T01Req_No=" & strFab_Req_No & " "
            dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(dsUser) Then

            Else
                nvcFieldList1 = "select * from P01PARAMETER where P01CODE='FR'"
                dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(dsUser) Then
                    _Ref = Trim(dsUser.Tables(0).Rows(0)("P01NO"))
                End If

                If _Ref < 10 Then
                    _REFST = "000" & _Ref
                ElseIf _Ref >= 10 And _Ref < 100 Then
                    _REFST = "00" & _Ref
                ElseIf _Ref >= 100 And _Ref < 1000 Then
                    _REFST = "0" & _Ref
                Else
                    _REFST = _Ref
                End If
                nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO +" & 1 & " WHERE P01CODE='FR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Insert Into T01_TEC_Development_Request(T01Req_No,T01Req_No_St,T01Customer_Ref,T01Req_Date,T01Requied_Date,T01Order_Qty,T01Target_Price,T01End_User,T01Fabric_Description,T01Composition,T01Target_Width,T01Target_Weight,T01End_End,T01Cutterble_Width,T01Color)" & _
                '                                               " values(" & _Ref & ", '" & _REFST & "','" & Trim(txtCustomer_Referance.Text) & "','" & txtReuest_Date.Text & "','" & txtRequierd_Date.Text & "','" & txtOrder_Qty.Text & "','" & txtTarger_Price.Text & "','" & Trim(cboEnd_User.Text) & "','" & Trim(txtFabric_Des.Text) & "','" & Trim(txtCompositin.Text) & "','" & txtTarget_Width.Text & "','" & txtTarget_Weight.Text & "','" & txtEnd_End.Text & "','" & txtCuttable_Width.Text & "','" & Trim(txtColour.Text) & "')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T01_TEC_Development_Request(T01Req_No,T01Req_No_St,T01Customer_Ref,T01Req_Date,T01Requied_Date,T01Order_Qty,T01Target_Price,T01End_User,T01Fabric_Description,T01Composition,T01Target_Width,T01Target_Weight,T01End_End,T01Cutterble_Width,T01Color,T01Color_No,T01Test_Std,T01Other_Req,T01Ref_Quality,T01Org_Ref,T01Approval,T01Hangers,T01Yardage,T01Test_Rpt,T01Brush,T01Anti_PIll,T01Sueded,T01Bio_Polish,T01Merchant,T01Time,T01Status,T01Development_Stage,T01User_Group,T01App_Time)" & _
                                                              " values(" & _Ref & ", '" & _REFST & "','" & Trim(txtCustomer_Referance.Text) & "','" & txtReuest_Date.Text & "','" & txtRequierd_Date.Text & "','" & txtOrder_Qty.Text & "','" & txtTarger_Price.Text & "','" & Trim(cboEnd_User.Text) & "','" & Trim(txtFabric_Des.Text) & "','" & Trim(txtCompositin.Text) & "','" & txtTarget_Width.Text & "','" & txtTarget_Weight.Text & "','" & txtEnd_End.Text & "','" & txtCuttable_Width.Text & "','" & Trim(txtColour.Text) & "','" & Trim(txtColour_no.Text) & "','" & cboTest_Standard.Text & "','" & Trim(txtOther_Requierment.Text) & "','" & cboQuality.Text & "','" & _OriginalRef & "','" & _Submision & "'," & cboHanger.Text & ",'" & txtYardage.Text & "','" & _TestReport & "','" & _Brush & "','" & _Anti_Pill & "','" & _Sude & "','" & _Bio_Polish & "','" & strDisname & "','" & Now & "','DR','NO','" & strUGroup & "','" & Now & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'UPDATE SPECIAL APPLICATION
                nvcFieldList1 = "Insert Into T02_TEC_Special_Application(T02Ref_No,T02Solid,T02Marl,T02Print,T02Dyed_Yarn,T02Pies_Dyed)" & _
                                                               " values('" & _Ref & "', '" & _Solid & "','" & _Marl & "','" & _Print & "','" & _DY & "','" & _PDY & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Record update successfuly on Ref.No" & _REFST, MsgBoxStyle.Information, "Information .....")
                transaction.Commit()

            End If

            connection.Close()
            Call Clear_Text()
            Call Load_Count()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try

    End Sub

    Function Clear_Text()
        Me.txtReuest_Date.Text = Today
        Me.txtRequierd_Date.Text = Today
        Me.txtColour.Text = ""
        Me.txtColour_no.Text = ""
        Me.txtOrder_Qty.Text = ""
        Me.txtCompositin.Text = ""
        Me.txtOrder_Qty.Text = ""
        Me.txtTarger_Price.Text = ""
        Me.txtTarget_Weight.Text = ""
        Me.txtTarget_Width.Text = ""
        Me.txtOther_Requierment.Text = ""
        Me.txtFabric_Des.Text = ""
        Me.cboCustomer.Text = ""
        Me.txtCuttable_Width.Text = ""
        Me.txtEnd_End.Text = ""
        Me.txtYardage.Text = ""
        Me.cboTest_Standard.Text = ""
        Me.cboHanger.Text = ""
        Me.cboQuality.Text = ""
        Me.cboEnd_User.Text = ""
        Me.chkCounter_Sample.Checked = False
        Me.chkCustomer_Spec.Checked = False
        Me.chKApproval_Aesth.Checked = True
        Me.chkApproval_Technical.Checked = False
        Me.chkTest_Internal.Checked = False
        Me.chkTest_Outside.Checked = False
        Me.chkSolid.Checked = False
        Me.chkMarl.Checked = False
        Me.chkPies_Dyed.Checked = False
        Me.chkPrint.Checked = False
        Me.chkDyed_Yarn.Checked = False
        Me.chkSueded.Checked = False
        Me.chkBio_Polish.Checked = False
        Me.chkBrush_Both.Checked = False
        Me.chkBrush_One.Checked = False
        Me.chkAntipill_Both.Checked = False
        Me.chkAntipill_One.Checked = False
        Me.cboCustomer.ToggleDropdown()
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
    End Sub

    Function Load_Hangers()
        Dim i As Integer
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("##", GetType(String))
        For i = 1 To 9
            dt.Rows.Add(New Object() {i})
        Next

        dt.AcceptChanges()

        Me.cboHanger.SetDataBinding(dt, Nothing)
        '  Me.UltraDropDown1.ValueMember = "ID"
        Me.cboHanger.DisplayMember = "##"
    End Function

    Function Load_Count()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M02 As DataSet
        'Load sales order to cboSO combobox

        Try
            If Trim(strUGroup) = "MERCHENT" Then
                Sql = "select count(T01Merchant) as Qty from T01_TEC_Development_Request where T01Merchant='" & strDisname & "' and T01Status='DR' group by T01Merchant"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    With MDIMain.UltraToolbarsManager1
                        .Ribbon.Tabs(0).Groups(2).Tools(0).SharedProps.Caption = "Requested" & " " & M02.Tables(0).Rows.Count
                    End With
                End If
            End If
            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
End Class