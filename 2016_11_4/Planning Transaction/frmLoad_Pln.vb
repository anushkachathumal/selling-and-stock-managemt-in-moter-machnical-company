Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration

Public Class frmLoad_Pln
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

   
    Private Sub txtUse_FG_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUse_FG.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_FG.Text) Then
                Value = txtUse_FG.Text
                txtUse_FG.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_FG.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtUse_WIP.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_FG.Text) Then
                Value = txtUse_FG.Text
                txtUse_FG.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_FG.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtUse_WIP.Focus()

        End If
    End Sub

    Private Sub txtUse_FG_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUse_FG.TextChanged
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
    End Sub

    Private Sub txtUse_FG_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUse_FG.ValueChanged

    End Sub

    Private Sub txtUse_WIP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUse_WIP.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_WIP.Text) Then
                Value = txtUse_WIP.Text
                txtUse_WIP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_WIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If

        ElseIf e.KeyCode = Keys.Tab Then
            Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
            If IsNumeric(txtUse_WIP.Text) Then
                Value = txtUse_WIP.Text
                txtUse_WIP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtUse_WIP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        End If
    End Sub

    Private Sub txtUse_WIP_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUse_WIP.TextChanged
        Call frmDelivaryQuatnew.CalculateBalance_To_Produce()
    End Sub

    Private Sub txtUse_WIP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUse_WIP.ValueChanged

    End Sub

    Private Sub chkCry1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCry1.CheckedChanged
        If chkCry1.Checked = True Then
            chkCry2.Checked = False
        End If
    End Sub

    Private Sub chkCry2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCry2.CheckedChanged
        If chkCry2.Checked = True Then
            chkCry1.Checked = False
        End If
    End Sub

    Private Sub chkLab1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLab1.CheckedChanged
        If chkLab1.Checked = True Then
            chkLab2.Checked = False
        End If
    End Sub

    Private Sub chkLab2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLab2.CheckedChanged
        If chkLab2.Checked = True Then
            chkLab1.Checked = False
        End If
    End Sub

    Private Sub chkNPL1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNPL1.CheckedChanged
        If chkNPL1.Checked = True Then
            chkNPL2.Checked = False
        End If
    End Sub

    Private Sub chkNPL2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNPL2.CheckedChanged
        If chkNPL2.Checked = True Then
            chkNPL1.Checked = False
        End If
    End Sub

    Private Sub chkPP1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPP1.CheckedChanged
        If chkPP1.Checked = True Then
            chkPP2.Checked = False
        End If
    End Sub

    Private Sub chkPP2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPP2.CheckedChanged
        If chkPP2.Checked = True Then
            chkPP1.Checked = False
        End If
    End Sub

    Private Sub txtMOQ_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMOQ.ValueChanged

    End Sub

    Private Sub UltraTextEditor21_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReq_Grg.ValueChanged

    End Sub

    Function Load_Parameter()
        Dim M01 As DataSet
        Dim M02 As DataSet


        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)

            ''vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and M22Strich_Lenth>0"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "PAR"))
            If isValidDataset(M01) Then
                GrgRef = M01.Tables(0).Rows(0)("P01NO")
            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Update_Parameter()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer

        Try
            nvcFieldList1 = "update P01PARAMETER set P01NO=P01NO +" & 1 & " where P01CODE='GP'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try

    End Function

    Function Update_Date()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim ncQryType As String

    

        Try
            vcWhere = "T11Ref_No=" & Delivary_Ref & " and T11Sales_Order='" & txtSO.Text & "' and T11Line_Item=" & txtLine_Item.Text & ""
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "LADC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                nvcFieldList1 = "update T11Lab_Dip_ConfDate set T11Date='" & txtDate.Text & "' where T11Ref_No=" & Delivary_Ref & " and T11Sales_Order='" & txtSO.Text & "' and T11Line_Item=" & txtLine_Item.Text & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                ncQryType = "LADU"
                nvcFieldList1 = "(T11Ref_No," & "T11Sales_Order," & "T11Line_Item," & "T11Date," & "T11User) " & "values(" & Delivary_Ref & ",'" & txtSO.Text & "'," & txtLine_Item.Text & ",'" & txtDate.Text & "','" & strDisname & "')"
                up_GetSetLAB_DIP_DateConf(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                '  MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function
    Private Sub chkG_Stock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkG_Stock.CheckedChanged
        Dim Value As Double
        Dim A As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m01 As DataSet
        Dim VCWHERE As String

        If chkLab1.Checked = True Then
        Else
            VCWHERE = "T11Ref_No=" & Delivary_Ref & " and T11Sales_Order='" & txtSO.Text & "' and T11Line_Item=" & txtLine_Item.Text & ""
            m01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "LADC"), New SqlParameter("@vcWhereClause1", VCWHERE))
            If isValidDataset(m01) Then
            Else
                A = MsgBox("LD is not Approved are you want to  continue.", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ....")
                If A = vbYes Then
                    txtDate.Visible = True
                    txtDate.Text = Today
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
        con.CLOSE()
            If chkG_Stock.Checked = True Then
                chkY_Orde.Checked = False
                chkY_Orde.Checked = False
                If txtReg_LIb.Text <> "" Then
                Else
                    txtReg_LIb.Text = "0"
            End If
            If txtReq_Grg.Text <> "" Then
            Else
                txtReq_Grg.Text = "0"
            End If
                Value = CDbl(txtReq_Grg.Text) + CDbl(txtReg_LIb.Text)
                With frmGriege_Stock
                    .txtGriege_Qty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtGriege_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    .txtFabric.Text = txtFabrication.Text
                    .txtLIB.Text = txtLIB.Text
                End With
                Call Search_Data()

                Call Load_Parameter()
                Call Update_Parameter()
                If Trim(txtFabric_Shade.Text) = "Yarn Dyes" Then
                    frmGriege_Stock.chkKnt_Plan.Text = "Yarn Dye Plan"
                End If
                frmGriege_Stock.Show()
            End If
    End Sub

    Function Search_Data()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
      

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)
            Dim TestString As String
            Dim TestArray() As String

            vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and M22Strich_Lenth>0"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                With frmGriege_Stock
                    .txtGauge.Text = M01.Tables(0).Rows(0)("M22Machine_Type")
                    TestString = M01.Tables(0).Rows(0)("M22Machine_Type")
                    TestArray = Split(TestString)
                    strMC_Group = TestArray(2) & TestArray(0)
                End With
            End If
            'COMMON QUALITY

            vcWhere = " M26Quality30='" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                With frmGriege_Stock
                    .txtCommon.Text = M01.Tables(0).Rows(0)("M26Quality20")
                End With
            Else
                With frmGriege_Stock
                    .txtCommon.Text = "NO"
                End With
            End If
            '------------------------------------------------------------------
            'SUTABLE GRIGE
            vcWhere = " M14Order='" & Trim(txtRcode.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "GRG"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                With frmGriege_Stock
                    .txtShade.Text = M01.Tables(0).Rows(0)("M14Grige")
                End With
            Else
                'With frmGriege_Stock
                '    .txtCommon.Text = "NO"
                'End With
            End If
            '-----------------------------------------------------------------
            'PERDAY KNITTING OUTPUT
            Dim Value As Double
            vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)='NE' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(dsUser) Then
                '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
                vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If


                vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPS3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                'End If

            End If
            '------------------------------------------------------------------------------

            vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)='NM' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(dsUser) Then
                '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
                vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '   Exit Function

                End If

                vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value

                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '   Exit Function

                End If


                vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '  Exit Function

                End If

                vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPN3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '   Exit Function

                End If

                'End If

            End If
            'DT
            vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)='DT' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(dsUser) Then
                '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
                vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '  Exit Function

                End If

                vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '  Exit Function

                End If


                vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPD3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '  Exit Function

                End If

                'End If

            End If
            '------------------------------------------------------------------------------
            'DE
            vcWhere = " M22Quality='" & Trim(txtQuality.Text) & "' and RIGHT(LEFT(M22Yarn,6),2)IN ('DE','D') "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(dsUser) Then
                '  If Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y1" Or Microsoft.VisualBasic.Left(Trim(txtQuality.Text), 2) = "Y3" Then
                vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '  Exit Function

                End If

                vcWhere = "left(M22Quality,2)in ('Y1') and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If


                vcWhere = "left(M22Quality,1) not in ('Y') and M22Yarn_Cons='100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE2"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '  Exit Function

                End If

                vcWhere = "left(M22Quality,2)  in ('Y3') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Machine_Type not like '%Auto Stripe%' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    '   Exit Function

                End If

                vcWhere = "left(M22Quality,1) NOT in ('Y') and M22Yarn_Cons<'100' and M22Fabric_Type='SINGLE JERSEY' and M22Strich_Lenth<>0 and M22Quality='" & Trim(txtQuality.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "RPE3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    Value = M01.Tables(0).Rows(0)("kgH")
                    strPer_Day = Value
                    With frmGriege_Stock
                        .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    End With
                    ' Exit Function

                End If

                'End If

            End If
            'RIB

            Dim _EFF As Double
            _EFF = 0
            Value = frmGriege_Stock.txtPer_Day.Text
            vcWhere = "left(M22Quality,2)in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Quality='" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _EFF = 0.6
            End If

            vcWhere = "M22Product_Type like '%MARL%' and left(M22Quality,2) not in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Quality='" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _EFF = 0.65
            End If

            vcWhere = "M22Product_Type not like '%MARL%' and left(M22Quality,2) not in ('Y1','Y3') and M22Fabric_Type='SINGLE JERSEY' and M22Quality='" & Trim(txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPER_DAY_KNTOUTPUT", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _EFF = 0.7
            End If
            If Value > 0 And _EFF > 0 Then
                Value = Value * _EFF * 24
                strPer_Day = Value
                With frmGriege_Stock
                    .txtPer_Day.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtPer_Day.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End With
            End If
           
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Private Sub frmLoad_Pln_AutoSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.AutoSizeChanged

    End Sub

    Private Sub frmLoad_Pln_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmGriege_Stock.Close()
    End Sub


    Private Sub frmLoad_Pln_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFabric_Shade.ReadOnly = True
    End Sub

    Private Sub txtDate_AfterDropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.AfterDropDown
        '  Call Update_Date()
    End Sub

    
    Private Sub txtDate_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDate.BeforeDropDown

    End Sub

    Private Sub txtDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.TextChanged
        Call Update_Date()
    End Sub

  
End Class