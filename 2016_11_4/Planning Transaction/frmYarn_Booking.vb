Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors



Public Class frmYarn_Booking
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim c_dataCustomer2 As System.Data.DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Gridewith_Data()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)

           
            Dim Z As Integer
            Z = 0
            i = 0
            vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim _shade As String
                Z = 0
                If frmGriege_Stock.txtShade.Text = "L" Or frmGriege_Stock.txtShade.Text = "WS" Then
                    _shade = "('Light','SP')"
                ElseIf frmGriege_Stock.txtShade.Text = "D" Then
                    _shade = "('Dark','SP')"
                ElseIf frmGriege_Stock.txtShade.Text = "M" Then
                    _shade = "('Marl','SP')"

                End If
                If frmGriege_Stock.txtShade.Text <> "" Then
                    If i = 0 Then
                        If txtCom1.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        End If
                    ElseIf i = 1 Then
                        If txtCom2.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        End If
                    ElseIf i = 2 Then
                        If txtCom3.Text > 15 Then
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "'  and M34Shade in " & _shade & " and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        Else
                            vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                        End If
                    End If

                Else
                    vcWhere = "dis='" & Microsoft.VisualBasic.Left(Trim(M01.Tables(0).Rows(i)("M22Yarn")), 4) & "' and M33Yarn_Location in ('2020','2005','2116','2009','2110')"
                End If
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarn_Booking", New SqlParameter("@cQryType", "CLS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow4 As DataRow In M02.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("10Class") = M02.Tables(0).Rows(Z)("M3310Class")
                    newRow("Description") = M02.Tables(0).Rows(Z)("M33Description")
                    newRow("Stock Code") = M02.Tables(0).Rows(Z)("M33Stock_Code")

                    newRow("Location") = M02.Tables(0).Rows(Z)("M33Yarn_Location")
                    _To = Month(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(Z)("M33Date")) & "/" & Year(M02.Tables(0).Rows(Z)("M33Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Yarn Aging") = Diff.Days & " days"
                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = M02.Tables(0).Rows(Z)("M33Qty")
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)

                    Z = Z + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                c_dataCustomer1.Rows.Add(newRow1)
                i = i + 1
            Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableYarn_Booking
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 60
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = True

            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub frmYarn_Booking_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmGriege_Stock.chkKnt_Plan.Checked = False
        frmGriege_Stock.Close()
        frmKnitting_Detailes.Close()
    End Sub

    Function Load_GrideStock()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTablePreVious_Stock
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 40
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 40
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False


            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '  .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Grid_SockCode()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim i As Integer
        Dim Value As Double
        Dim _VString As String

        Try
            i = 0
            vcWhere = "m42Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "STC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow

                newRow("Stock Code") = M01.Tables(0).Rows(i)("M42Stock_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M24Customer")
                newRow("Week No") = M01.Tables(0).Rows(i)("M24Week")
                newRow("Year") = M01.Tables(0).Rows(i)("M24Year")
                Value = M01.Tables(0).Rows(i)("Qty")
                _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Used Qty(Kg)") = _VString
                c_dataCustomer2.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub frmYarn_Booking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_GrideStock()
        Call Load_Grid_SockCode()
        Call Load_Gridewith_Data()
        txtRequest_Date.Text = Today
        txtGriege_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtGriege_Qty.ReadOnly = True
        txtK_Wastage.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        ' txtK_Wastage.ReadOnly = True

        txtCom1.ReadOnly = True
        txtCom2.ReadOnly = True
        txtCom3.ReadOnly = True

        txtCom1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCom2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCom3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtYarn1.ReadOnly = True
        txtYarn2.ReadOnly = True
        txtYarn3.ReadOnly = True

        txtYarn1.Appearance.TextHAlign = Infragistics.Win.HAlign.Left
        txtYarn2.Appearance.TextHAlign = Infragistics.Win.HAlign.Left
        txtYarn3.Appearance.TextHAlign = Infragistics.Win.HAlign.Left

        txtReq1.ReadOnly = True
        txtReq2.ReadOnly = True
        txtReq3.ReadOnly = True

        txtReq1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtReq2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtReq3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        lblBalance.Text = txtGriege_Qty.Text
    End Sub

    Private Sub UltraLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel2.Click

    End Sub

    Private Sub UltraGrid1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles UltraGrid1.AfterRowUpdate
        Call Calculation_Balance()
    End Sub

    Function Calculation_Balance()
        Dim I As Integer
        Dim Value As Double
        Dim _Vstring As String
        Dim _RowIndex As Integer
        Dim _Value1 As Integer
        Dim _Value2 As Integer
        Dim _Value3 As Integer


        Try
            '  MsgBox(UltraGrid1.Selected.Rows.Item(0).ToString)
            I = 0
            Value = 0
            _Value1 = 0
            _Value2 = 0
            _Value3 = 0

            For Each uRow As UltraGridRow In UltraGrid1.Rows

                With UltraGrid1
                    If IsNumeric(.Rows(I).Cells(6).Text) Then

                        If CDbl((.Rows(I).Cells(6).Text)) <= CDbl((.Rows(I).Cells(5).Text)) Then
                            If Microsoft.VisualBasic.Left((.Rows(I).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                                If IsNumeric(.Rows(I).Cells(6).Text) Then
                                    _Value1 = _Value1 + .Rows(I).Cells(6).Text
                                End If
                            ElseIf Microsoft.VisualBasic.Left((.Rows(I).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                                If IsNumeric(.Rows(I).Cells(6).Text) Then
                                    _Value2 = _Value2 + .Rows(I).Cells(6).Text
                                End If
                            ElseIf Microsoft.VisualBasic.Left((.Rows(I).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn3.Text, 4) Then
                                If IsNumeric(.Rows(I).Cells(6).Text) Then
                                    _Value3 = _Value3 + .Rows(I).Cells(6).Text
                                End If
                            End If
                            Value = Value + CDbl((.Rows(I).Cells(6).Text))
                            .Rows(I).Cells(0).Appearance.BackColor = Color.White
                            .Rows(I).Cells(1).Appearance.BackColor = Color.White
                            .Rows(I).Cells(2).Appearance.BackColor = Color.White
                            .Rows(I).Cells(3).Appearance.BackColor = Color.White
                            .Rows(I).Cells(4).Appearance.BackColor = Color.White
                            .Rows(I).Cells(5).Appearance.BackColor = Color.White
                            .Rows(I).Cells(6).Appearance.BackColor = Color.White

                        Else
                            MsgBox("Qty grater than to stock", MsgBoxStyle.Information, "Information ....")
                            .Rows(I).Cells(0).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(1).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(2).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(3).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(4).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(5).Appearance.BackColor = Color.Red
                            .Rows(I).Cells(6).Appearance.BackColor = Color.Red
                            .Rows(I).Selected = True
                            Exit For
                        End If
                        '_Vstring = Value
                        '_Vstring = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        '_Vstring = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        'UltraGrid1.Rows(I).Cells(6).Value = _Vstring
                    End If
                End With
                I = I + 1
            Next
            Value = CDbl(txtGriege_Qty.Text) - Value
            lblBalance.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            lblBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            'pbCount1.Value = _Value1
            'pbCount2.Value = _Value2
            'pbCount3.Value = _Value3


            If CDbl(lblBalance.Text) < 0 Then
                MsgBox("Qty grater than to stock", MsgBoxStyle.Information, "Information ....")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

  
    Private Sub cmdKnt_Chart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKnt_Chart.Click
        frmKnitting_Plan_Board.Show()
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub cmdKnt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKnt.Click
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
        Dim i As Integer
        Try
            strKnitting_PlanStatus = "Yarn Booking"
            If txtReq1.Text <> "" And txtReq2.Text <> "" And txtReq3.Text <> "" Then
                strQty = CDbl(txtReq1.Text) + CDbl(txtReq2.Text) + CDbl(txtReq3.Text)
            ElseIf txtReq1.Text <> "" And txtReq2.Text <> "" Then
                strQty = CDbl(txtReq1.Text) + CDbl(txtReq2.Text)
            ElseIf txtReq1.Text <> "" Then
                strQty = CDbl(txtReq1.Text)
            End If

            i = 0
            Dim _status As Boolean
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BKD"))
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("tmpUser")) = strDisname Then
                    If CDbl(lblBalance.Text) = 0 Then
                        frmKnitting_Detailes.Show()
                    Else
                        '1st Yarn Checking
                        For Each uRow As UltraGridRow In UltraGrid1.Rows
                            If Microsoft.VisualBasic.Left(Trim(UltraGrid1.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                                _status = True
                                Exit For
                            End If
                            i = i + 1
                        Next

                        If _status = False Then
                            MsgBox("Please request the Yarn", MsgBoxStyle.Information, "Information .....")
                            OPR8.Visible = True
                            Exit Sub
                        End If

                        _status = False
                        i = 0
                        For Each uRow As UltraGridRow In UltraGrid1.Rows
                            If Microsoft.VisualBasic.Left(Trim(UltraGrid1.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                                _status = True
                                Exit For
                            End If
                            i = i + 1
                        Next


                        If _status = False Then
                            MsgBox("Please request the Yarn", MsgBoxStyle.Information, "Information .....")
                            OPR8.Visible = True
                            Exit Sub
                        End If

                        MsgBox("Can't Plan Knitting Machine yet", MsgBoxStyle.Information, "Information .......")
                        Exit Sub
                    End If

                Else
                    MsgBox("System used by " & M01.Tables(0).Rows(0)("tmpUser") & " Please try again", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    Exit Sub
                End If
            Else

                If CDbl(lblBalance.Text) = 0 Then
                    ncQryType = "BADD"
                    nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpUser) " & "values('" & strSales_Order & "'," & strLine_Item & ",'" & strDisname & "')"
                    up_GetSetBlock_KntMC(ncQryType, nvcFieldList1, vcWhere, connection, transaction)

                    transaction.Commit()
                    connection.Close()

                    frmKnitting_Detailes.Show()
                Else
                    MsgBox("Can't Plan Knitting Machine yet", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Sub
                End If

            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Call Update_Records()
    End Sub

    Function Update_Records()
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
        Dim i As Integer
        Dim _Balance As Double
        i = 0
        '1st Yarn
        Try
            If txtYarn1.Text <> "" Then
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If Microsoft.VisualBasic.Left(Trim(UltraGrid1.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn1.Text, 4) Then
                        If Trim(UltraGrid1.Rows(i).Cells(6).Text) <> "" Then
                            _Balance = _Balance + Trim(UltraGrid1.Rows(i).Cells(6).Value)
                        End If
                    End If

                    i = i + 1
                Next
            End If

            If (CDbl(txtReq1.Text) - _Balance) > 0 Then
                vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn1.Text & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then

                Else
                    ncQryType = "AYR"
                    nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Req_Date," & "T14Status," & "T14Qty," & "T14Time) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn1.Text & "','" & strDisname & "','" & txtRequest_Date.Text & "','N','" & CDbl(txtReq1.Text) - _Balance & "','" & Now & "')"
                    up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                End If
            End If
            '2nd Yarn
            i = 0
            _Balance = 0
            If txtYarn2.Text <> "" Then
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If Microsoft.VisualBasic.Left(Trim(UltraGrid1.Rows(i).Cells(1).Text), 4) = Microsoft.VisualBasic.Left(txtYarn2.Text, 4) Then
                        If Trim(UltraGrid1.Rows(i).Cells(6).Text) <> "" Then
                            _Balance = _Balance + Trim(UltraGrid1.Rows(i).Cells(6).Value)
                        End If
                    End If

                    i = i + 1
                Next
            End If

            If (CDbl(txtReq2.Text) - _Balance) > 0 Then
                vcWhere = "T14Ref_no=" & Delivary_Ref & " and T14Sales_order='" & strSales_Order & "' and T14Line_Item=" & strLine_Item & " and T14Yarn='" & txtYarn2.Text & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YRQ"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then

                Else
                    ncQryType = "AYR"
                    nvcFieldList1 = "(T14Ref_no," & "T14Sales_order," & "T14Line_Item," & "T14Yarn," & "T14Req_By," & "T14Req_Date," & "T14Status," & "T14Qty," & "T14Time) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & txtYarn2.Text & "','" & strDisname & "','" & txtRequest_Date.Text & "','N','" & CDbl(txtReq2.Text) - _Balance & "','" & Now & "')"
                    up_GetSetYarn_Request(ncQryType, nvcFieldList1, vcWhere, connection, transaction)


                End If
            End If

            transaction.Commit()
            connection.Close()
            Me.Close()
         
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click

    End Sub
End Class