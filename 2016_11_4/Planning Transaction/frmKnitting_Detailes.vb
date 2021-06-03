
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Public Class frmKnitting_Detailes
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Rowindex As Integer
    Dim _YarnQty As Double

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmKnitting_Detailes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtQuality_Group.ReadOnly = True

        txtDate.Text = Today
        txtSales_Order.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSales_Order.ReadOnly = True
        txtLine_Item.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLine_Item.ReadOnly = True
        txtQty.ReadOnly = True
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDaily_Capacity.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDaily_Capacity.ReadOnly = True
        txtQuality.ReadOnly = True
        txtDays.ReadOnly = True
        txtDays.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMC_Group.ReadOnly = True
        txtSales_Order.Text = strSales_Order
        txtLine_Item.Text = strLine_Item
        txtQuality.Text = strQuality
        txtDaily_Capacity.Text = (strPer_Day.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtDaily_Capacity.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", strPer_Day))
        txtQty.Text = (strQty.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtQty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", strQty))
        If txtQty.Text <> "" And txtDaily_Capacity.Text <> "" Then
            txtDays.Text = CInt(txtQty.Text / txtDaily_Capacity.Text)
        End If

        txtMC_Group.Text = strMC_Group
        Call Load_Gride()
        Call Quality_Group()
    End Sub
    Function Quality_Group()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Try
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "NSL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Non Lycra"
                con.close()
                Exit Function
            End If
            '-----------------------------------------------------------------------
            'SOLID LYCRA HEVY
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "SLH"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Lycra Heavy"
                con.close()
                Exit Function
            End If
            '-----------------------------------------------------------------------
            'SOLID LYCRA SLACK
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "SLS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Solid Lycra Slack"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'MARL NON LYCRA
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "MNL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Marl Non Lycra"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'MARL LYCRA
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "MLC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Marl Lycra"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'Dye Yarn Lycra 
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "DYL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Dye Yarn Lycra"
                con.close()
                Exit Function
            End If
            '----------------------------------------------------------------------
            'Dye Yarn Non Lycra 
            vcWhere = "M22Quality='" & Trim(txtQuality.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetQuality_Group", New SqlParameter("@cQryType", "DYNL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtQuality_Group.Text = "Dye Yarn Non Lycra"
                con.close()
                Exit Function
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableKnitting_PLN
        UltraGrid4.DataSource = c_dataCustomer1
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
          
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
  

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function
    Function Search_WeekNo()
        On Error Resume Next
        Dim _Date As Date

        _Date = txtDate.Text
        If txtWeek.Text <> "" Then
            txtYear.Text = Year(txtDate.Text)
            txtWeek.Text = DatePart("WW", _Date, FirstDayOfWeek.Monday)
        Else
            txtYear.Text = Year(txtDate.Text)
            txtWeek.Text = DatePart("WW", _Date, FirstDayOfWeek.Monday)
        End If
    End Function
    Private Sub txtDate_AfterDropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.AfterDropDown
        Search_WeekNo()
    End Sub

    Private Sub TestDateAdd()
        If IsNumeric(txtWeek.Text) And IsNumeric(txtYear.Text) Then
            Dim weekStart As DateTime = GetWeekStartDate(txtWeek.Text, txtYear.Text)
            txtDate.Text = weekStart
        End If
    End Sub

    Private Function GetWeekStartDate(ByVal weekNumber As Integer, ByVal year As Integer) As Date
        Dim startDate As New DateTime(year, 1, 1)
        Dim weekDate As DateTime = DateAdd(DateInterval.WeekOfYear, weekNumber - 1, startDate)
        Return DateAdd(DateInterval.Day, (-weekDate.DayOfWeek) + 1, weekDate)
    End Function

    Private Sub txtDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.TextChanged
        Call Search_WeekNo()
    End Sub

    Private Sub txtWeek_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWeek.KeyUp
        If e.KeyCode = 13 Then
            txtYear.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtYear.Focus()
        End If
    End Sub

    Private Sub txtWeek_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWeek.ValueChanged

    End Sub

    Private Sub txtYear_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear.KeyUp
        If e.KeyCode = 13 Then
            txtDate.Focus()
            Call TestDateAdd()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDate.Focus()
            Call TestDateAdd()
        End If
    End Sub

    Private Sub txtYear_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtYear.LostFocus
        Call TestDateAdd()
    End Sub

    Private Sub txtYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtYear.ValueChanged

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If strKnitting_PlanStatus = "Yarn Booking" Then
            Call Save_Data(Today)
        End If
    End Sub

    Function Search_Available_KMCNew(ByVal strDate As Date)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet
        Dim Value As Double
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim strQty As String

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _Balance_Qty As Double
        Dim _MinQty As Double
        Dim i As Integer
        Dim _FromDate As Date
        Dim _Todate As Date
        Dim _TimeSpam As TimeSpan
        Dim _TotalTime As Double
        Dim _AllocateMC As Integer
        Dim _Knited_Time As Integer
        Dim _StartDate As Date
        Dim _EndDate As Date
        Dim _WeekNo As Integer
        Dim _UseMCNo As Integer
        Dim _McName As String
        Dim x As Integer

        If Microsoft.VisualBasic.Left(txtMC_Group.Text, 1) = "S" Then
            _McName = "SJ"
        ElseIf Microsoft.VisualBasic.Left(txtMC_Group.Text, 1) = "R" Then
            _McName = "DJ"

        End If
        Try
            _MinQty = txtDaily_Capacity.Text / (24 * 60)
            _Balance_Qty = txtQty.Text
            i = 0
            _AllocateMC = 0

            If _McName = "SJ" Then
                vcWhere = "M38MC='" & _McName & "' and M38Group='" & strGuarge & "'"
            Else
                vcWhere = "M38MC='" & _McName & "' and LEFT(M38Group,2)='" & Microsoft.VisualBasic.Left(strGuarge, 2) & "'"
            End If
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "MCNO"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _UseMCNo = M01.Tables(0).Rows(0)("M38Mc_Count")
            End If

            vcWhere = "tmpQuality='" & txtQuality.Text & "' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Knited_Time = 0
                If _AllocateMC > txtAlocate_MC.Text Or _Balance_Qty = 0 Then
                    'connection.Close()
                    Exit For
                End If
                'MsgBox(M01.Tables(0).Rows(i)("tmpEnd_Date"))

                If M01.Tables(0).Rows(i)("tmpEnd_Date") < Today Then
                    _FromDate = Today & " " & "7:30 AM"
                    _Todate = txtDate.Text & " " & "7:30 AM"
                Else
                    _FromDate = M01.Tables(0).Rows(i)("tmpEnd_Time")
                    _Todate = txtDate.Text & " " & "7:30 AM"
                End If

                _TimeSpam = _Todate.Subtract(_FromDate)
                _TotalTime = (_TimeSpam.TotalMinutes * _MinQty) '/ 1000
                If _TotalTime > 0 Then
                Else
                    i = i + 1
                    Continue For
                End If
                If _TotalTime > txtQty.Text Then
                    _Balance_Qty = 0
                    _AllocateMC = _AllocateMC + 1
                    _Knited_Time = txtQty.Text / _MinQty
                    If _Knited_Time > 0 Then
                        _Todate = _FromDate.AddMinutes(+_Knited_Time)
                    End If

                    _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                    _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))

                    _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)

                    ncQryType = "KADD"
                    nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & txtQty.Text & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','S')"
                    up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    'transaction.Commit()
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                    newRow("Start Date") = _FromDate
                    newRow("End Date") = _Todate

                    newRow("Qty") = txtQty.Text
                    newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                    newRow("Status") = "Same Quality"
                    c_dataCustomer1.Rows.Add(newRow)
                    'Exit Function
                Else
                    _Balance_Qty = _Balance_Qty - _TotalTime

                    If txtAlocate_MC.Text < _AllocateMC Then
                        MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                        vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                        M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                        'transaction.Commit()
                        'connection.Close()
                        Exit Function
                    Else
                        _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                        _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                        _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)
                        ncQryType = "KADD"
                        nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & _TotalTime & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','S')"
                        up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                        ' transaction.Commit()
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                        newRow("Start Date") = _FromDate
                        newRow("End Date") = _Todate
                        Value = _TotalTime
                        strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow("Qty") = strQty
                        newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                        newRow("Status") = "Same Quality"
                        c_dataCustomer1.Rows.Add(newRow)
                        _AllocateMC = _AllocateMC + 1
                    End If
                End If
                i = i + 1
            Next
            '-----------------------------------------------------------------------------------------
            If _Balance_Qty > 0 Then
                'Quality Change
                x = 0
                vcWhere = "M40Group_Name='" & Trim(txtQuality_Group.Text) & "' "
                M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "MGN"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    i = 0
                    vcWhere = "tmpQuality<>'" & txtQuality.Text & "' and left(tmpGroup,2)='" & _McName & "' and tmpStatus='" & Trim(M02.Tables(0).Rows(x)("M40Priority_Group")) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow4 As DataRow In M01.Tables(0).Rows
                        vcWhere = "M38Group='" & strGuarge & "' and  M39Mc_No='" & Trim(M01.Tables(0).Rows(i)("tmpMC_No")) & "'"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "CAMC"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                        Else
                            i = i + 1
                            Continue For
                        End If

                        vcWhere = "tmpRef_No=" & Delivary_Ref & " and  tmpMC_No='" & Trim(M01.Tables(0).Rows(i)("tmpMC_No")) & "'"
                        dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPCK"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            i = i + 1
                            Continue For
                        End If

                        _Knited_Time = 0
                        If _AllocateMC > txtAlocate_MC.Text Or _Balance_Qty = 0 Then
                            'connection.Close()
                            Exit For
                        End If
                        'MsgBox(M01.Tables(0).Rows(i)("tmpEnd_Date"))
                        Dim _Qstatus As String

                        If M01.Tables(0).Rows(i)("tmpEnd_Date") < strDate Then
                            _FromDate = strDate & " " & "7:30 AM"
                            _Todate = txtDate.Text & " " & "7:30 AM"

                        Else
                            _FromDate = M01.Tables(0).Rows(i)("tmpEnd_Time")
                            _Todate = txtDate.Text & " " & "7:30 AM"
                            'If txtQuality.Text = M01.Tables(0).Rows(i)("tmpQuality") Then
                            '    _Qstatus = "S"
                            'Else
                            _Qstatus = "QC"
                            _FromDate = _FromDate.AddHours(M02.Tables(0).Rows(x)("M40Mc_Change_HR"))
                            ' End If
                        End If

                        _TimeSpam = _Todate.Subtract(_FromDate)
                        _TotalTime = (_TimeSpam.TotalMinutes * _MinQty) '/ 1000
                        If _TotalTime > 0 Then
                        Else
                            i = i + 1
                            Continue For
                        End If
                        If _TotalTime > txtQty.Text Then
                            _Balance_Qty = 0
                            _AllocateMC = _AllocateMC + 1
                            _Knited_Time = txtQty.Text / _MinQty
                            If _Knited_Time > 0 Then
                                _Todate = _FromDate.AddMinutes(+_Knited_Time)
                            End If

                            _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                            _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))

                            _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)

                            ncQryType = "KADD"
                            nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & txtQty.Text & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','" & _Qstatus & "')"
                            up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                            'transaction.Commit()
                            Dim newRow As DataRow = c_dataCustomer1.NewRow

                            newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                            newRow("Start Date") = _FromDate
                            newRow("End Date") = _Todate
                            Value = txtQty.Text
                            strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            newRow("Qty") = strQty
                            newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                            If _Qstatus = "S" Then
                                newRow("Status") = "Same Quality"
                            Else

                                newRow("Status") = "Quality Change"
                            End If
                            c_dataCustomer1.Rows.Add(newRow)
                            'Exit Function
                        Else
                            _Balance_Qty = _Balance_Qty - _TotalTime

                            If txtAlocate_MC.Text < _AllocateMC Then
                                MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                                vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                                'transaction.Commit()
                                'connection.Close()
                                Exit Function
                            Else
                                _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                                _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                                _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)
                                ncQryType = "KADD"
                                nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date," & "tmpStatus," & "tmpQ_Status) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & _TotalTime & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "','" & txtQuality_Group.Text & "','QC')"
                                up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                                ' transaction.Commit()
                                Dim newRow As DataRow = c_dataCustomer1.NewRow

                                newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                                newRow("Start Date") = _FromDate
                                newRow("End Date") = _Todate
                                Value = _TotalTime
                                strQty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                strQty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                                newRow("Qty") = strQty
                                newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                                newRow("Status") = "Quality Change"
                                c_dataCustomer1.Rows.Add(newRow)
                                _AllocateMC = _AllocateMC + 1
                            End If
                        End If
                        i = i + 1
                    Next

                    x = x + 1
                Next
            End If
            transaction.Commit()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try

    End Function

    Function Search_Available_KMC()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _Balance_Qty As Double
        Dim _MinQty As Double
        Dim i As Integer
        Dim _FromDate As Date
        Dim _Todate As Date
        Dim _TimeSpam As TimeSpan
        Dim _TotalTime As Double
        Dim _AllocateMC As Integer
        Dim _Knited_Time As Integer
        Dim _StartDate As Date
        Dim _EndDate As Date
        Dim _WeekNo As Integer

        Try
            _MinQty = txtDaily_Capacity.Text / (24 * 60)
            _Balance_Qty = txtQty.Text
            i = 0
            _AllocateMC = 0
            vcWhere = "tmpQuality='" & txtQuality.Text & "' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Knited_Time = 0
                If _AllocateMC > txtAlocate_MC.Text Or _Balance_Qty = 0 Then
                    'connection.Close()
                    Exit For
                End If
                'MsgBox(M01.Tables(0).Rows(i)("tmpEnd_Date"))

                If M01.Tables(0).Rows(i)("tmpEnd_Date") < Today Then
                    _FromDate = Today & " " & "7:30 AM"
                    _Todate = txtDate.Text & " " & "7:30 AM"
                Else
                    _FromDate = M01.Tables(0).Rows(i)("tmpEnd_Time")
                    _Todate = txtDate.Text & " " & "7:30 AM"
                End If

                _TimeSpam = _Todate.Subtract(_FromDate)
                _TotalTime = (_TimeSpam.TotalMinutes * _MinQty) '/ 1000
                If _TotalTime > 0 Then
                Else
                    i = i + 1
                    Continue For
                End If
                If _TotalTime > txtQty.Text Then
                    _Balance_Qty = 0
                    _AllocateMC = _AllocateMC + 1
                    _Knited_Time = txtQty.Text / _MinQty
                    If _Knited_Time > 0 Then
                        _Todate = _FromDate.AddMinutes(+_Knited_Time)
                    End If

                    _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                    _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))

                    _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)

                    ncQryType = "KADD"
                    nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & txtQty.Text & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "')"
                    up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    'transaction.Commit()
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                    newRow("Start Date") = _FromDate
                    newRow("End Date") = _Todate

                    newRow("Qty") = txtQty.Text
                    newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                    c_dataCustomer1.Rows.Add(newRow)
                    'Exit Function
                Else
                    _Balance_Qty = _Balance_Qty - _TotalTime

                    If txtAlocate_MC.Text < _AllocateMC Then
                        MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                        vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                        M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                        'transaction.Commit()
                        'connection.Close()
                        Exit Function
                    Else
                        _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                        _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                        _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)
                        ncQryType = "KADD"
                        nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & _TotalTime & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "')"
                        up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                        ' transaction.Commit()
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                        newRow("Start Date") = _FromDate
                        newRow("End Date") = _Todate

                        newRow("Qty") = _TotalTime
                        newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                        c_dataCustomer1.Rows.Add(newRow)
                        _AllocateMC = _AllocateMC + 1
                    End If
                End If
                i = i + 1
            Next
            If _AllocateMC >= txtAlocate_MC.Text And _Balance_Qty > 0 Then
                MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))

                connection.Close()
                Exit Function
            Else
                'QuLITY CHANGE
                If _Balance_Qty > 0 Then
                    i = 0
                    vcWhere = "tmpQuality<>'" & txtQuality.Text & "' "
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KPLA"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows
                        _Knited_Time = 0
                        If _AllocateMC > txtAlocate_MC.Text Or _Balance_Qty = 0 Then
                            'connection.Close()
                            Exit For
                        End If
                        'MsgBox(M01.Tables(0).Rows(i)("tmpEnd_Date"))

                        If M01.Tables(0).Rows(i)("tmpEnd_Date") < Today Then
                            _FromDate = Today & " " & "7:30 AM"
                            _Todate = txtDate.Text & " " & "7:30 AM"
                        Else
                            _FromDate = M01.Tables(0).Rows(i)("tmpEnd_Time")
                            _FromDate = _FromDate.AddHours(4)
                            _Todate = txtDate.Text & " " & "7:30 AM"
                        End If

                        _TimeSpam = _Todate.Subtract(_FromDate)
                        _TotalTime = (_TimeSpam.TotalMinutes * _MinQty) '/ 1000
                        If _TotalTime > 0 Then
                        Else
                            i = i + 1
                            Continue For
                        End If
                        If _TotalTime > _Balance_Qty Then

                            _AllocateMC = _AllocateMC + 1
                            _Knited_Time = _Balance_Qty / _MinQty

                            If _Knited_Time > 0 Then
                                _Todate = _FromDate.AddMinutes(+_Knited_Time)
                            End If

                            _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                            _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))

                            _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)

                            ncQryType = "KADD"
                            nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & _Balance_Qty & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "')"
                            up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                            '  transaction.Commit()
                            Dim newRow As DataRow = c_dataCustomer1.NewRow

                            newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                            newRow("Start Date") = _FromDate
                            newRow("End Date") = _Todate

                            newRow("Qty") = _Balance_Qty
                            newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                            c_dataCustomer1.Rows.Add(newRow)
                            _Balance_Qty = 0
                        Else
                            _Balance_Qty = _Balance_Qty - _TotalTime

                            If txtAlocate_MC.Text < _AllocateMC Then
                                MsgBox("No Available Machine Capacity. Please add the allocate machine", MsgBoxStyle.Information, "Information ....")
                                vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))
                                'transaction.Commit()
                                'connection.Close()
                                Exit Function
                            Else
                                _StartDate = (Month(_FromDate) & "/" & Microsoft.VisualBasic.Day(_FromDate) & "/" & Year(_FromDate))
                                _EndDate = (Month(_Todate) & "/" & Microsoft.VisualBasic.Day(_Todate) & "/" & Year(_Todate))
                                _WeekNo = DatePart("WW", _EndDate, FirstDayOfWeek.Monday)
                                ncQryType = "KADD"
                                nvcFieldList1 = "(tmpRef_No," & "tmpMC_No," & "tmpGroup," & "tmpQuality," & "tmp20Class," & "tmpSales_Order," & "tmpLine_Item," & "tmpKnt_Order," & "tmpBalance," & "tmpWeek_No," & "tmpYear," & "tmpStart_Time," & "tmpEnd_Time," & "tmpStart_Date," & "tmpEnd_Date) " & "values(" & Delivary_Ref & ",'" & M01.Tables(0).Rows(i)("tmpMC_No") & "','" & txtMC_Group.Text & "','" & txtQuality.Text & "','" & str20Class & "','" & strSales_Order & "'," & strLine_Item & ",'" & txtQty.Text & "','" & _TotalTime & "'," & _WeekNo & "," & Year(_EndDate) & ",'" & _FromDate & "','" & _Todate & "','" & _StartDate & "','" & _EndDate & "')"
                                up_GetSetBlock_KntPlanning_Boad(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                                ' transaction.Commit()
                                Dim newRow As DataRow = c_dataCustomer1.NewRow

                                newRow("Machine No") = M01.Tables(0).Rows(i)("tmpMC_No")
                                newRow("Start Date") = _FromDate
                                newRow("End Date") = _Todate

                                newRow("Qty") = _TotalTime
                                newRow("No of Hour") = (_TimeSpam.Days * 24 + _TimeSpam.Hours) & "." & _TimeSpam.Minutes
                                c_dataCustomer1.Rows.Add(newRow)
                                _AllocateMC = _AllocateMC + 1
                            End If
                        End If
                        i = i + 1
                    Next
                End If
            End If
            If _Balance_Qty > 0 Then
                MsgBox("No enough machine capacity", MsgBoxStyle.Exclamation, "Technova .......")
                connection.Close()
                Exit Function
            End If
            transaction.Commit()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub txtDate_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDate.BeforeDropDown

    End Sub

    Function Save_Data(ByVal strDate As Date)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String
        Dim M02 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            Call Load_Gride()
            If txtAlocate_MC.Text <> "" Then
                vcWhere = "tmpRef_No=" & Delivary_Ref & ""
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DKP"), New SqlParameter("@vcWhereClause1", vcWhere))

                transaction.Commit()
                connection.Close()
                If txtDate.Text <= strDate Then
                    MsgBox("Can't Plan this day.Please try again", MsgBoxStyle.Information, "Information .....")
                    Exit Function
                End If
                Call Search_Available_KMCNew(strDate)
            Else
                MsgBox("Please enter the Allocated Machine", MsgBoxStyle.Information, "Information .....")
            End If
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub cmdChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChart.Click
        frmKnitting_Plan_Board.Show()
    End Sub
End Class