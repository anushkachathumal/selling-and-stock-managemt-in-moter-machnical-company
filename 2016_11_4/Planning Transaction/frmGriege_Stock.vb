Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors

Public Class frmGriege_Stock
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim c_dataCustomer2 As System.Data.DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGrige_Stock
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 130
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = True
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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
    Private Sub frmGriege_Stock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_GrideStock()
        txtGriege_Qty.ReadOnly = True
        lblBalance.Text = txtGriege_Qty.Text
        Call Load_Grid_SockCode()
        Call Load_WithData()

        ' lblBalance.Text = txtGriege_Qty.Text
    End Sub

    Function Load_WithData()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        Dim X As Integer
        Dim _Date As Date
        Dim _BatchNo As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)
            Dim _qty1 As Double
            i = 0
            vcWhere = "M21Material='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "' and left(M21Sales_Order,2)='20' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _qty1 = 0
                vcWhere = "T12Stock_Code='" & M01.Tables(0).Rows(i)("M21Batch_No") & "' and T12Time>='" & M01.Tables(0).Rows(i)("M21Update_Time") & "'  "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    _qty1 = M02.Tables(0).Rows(0)("T12Qty")
                End If
                If M01.Tables(0).Rows(i)("M21Qty") > _qty1 Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("20Class") = M01.Tables(0).Rows(i)("M2120Class")
                    newRow("Grige Order No") = M01.Tables(0).Rows(i)("M21Sales_Order")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M21Batch_No")
                    If i = 0 Then
                        _BatchNo = "" & M01.Tables(0).Rows(i)("M21Batch_No")
                    Else
                        _BatchNo = _BatchNo & "','" & M01.Tables(0).Rows(i)("M21Batch_No")
                    End If
                    _To = Month(M01.Tables(0).Rows(i)("M21Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M21Date")) & "/" & Year(M01.Tables(0).Rows(i)("M21Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Age") = Diff.Days & " days"

                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = CDbl(M01.Tables(0).Rows(i)("M21Qty")) - _qty1
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)
                End If

                i = i + 1
            Next
            i = 0
            vcWhere = "M21Material='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "' and left(M21Sales_Order,2)='20' and M21Batch_No not in ('" & _BatchNo & "') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _qty1 = 0
                vcWhere = "T12Stock_Code='" & M01.Tables(0).Rows(i)("M21Batch_No") & "' and T12Time>='" & M01.Tables(0).Rows(i)("M21Update_Time") & "'  "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    _qty1 = M02.Tables(0).Rows(0)("T12Qty")
                End If
                If M01.Tables(0).Rows(i)("M21Qty") > _qty1 Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("20Class") = M01.Tables(0).Rows(i)("M2120Class")
                    newRow("Grige Order No") = M01.Tables(0).Rows(i)("M21Sales_Order")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M21Batch_No")
                    _To = Month(M01.Tables(0).Rows(i)("M21Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M21Date")) & "/" & Year(M01.Tables(0).Rows(i)("M21Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Age") = Diff.Days & " days"

                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = CDbl(M01.Tables(0).Rows(i)("M21Qty")) - _qty1
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)
                End If
                i = i + 1
            Next

            con.close()
            X = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                i = 0
                'If lblBalance.Text = "0.00" Then
                '    Exit Function
                'End If
                For Each uRow1 As UltraGridRow In UltraGrid1.Rows
                    Dim _Qty As Double
                    _Qty = 0
                    With UltraGrid1
                        If Trim(.Rows(i).Cells(2).Value) = Trim(UltraGrid2.Rows(X).Cells(0).Value) Then
                            _Qty = .Rows(i).Cells(4).Value
                            If CDbl(lblBalance.Text) >= _Qty Then
                                If UltraGrid1.Rows(i).Cells(6).Value = True Then
                                Else
                                    lblBalance.Text = CDbl(lblBalance.Text) - _Qty
                                    .Rows(i).Cells(0).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(1).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(2).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(3).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(4).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(5).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(6).Appearance.BackColor = Color.Blue
                                    ' .Rows(i).Cells(5).Text = lblBalance.Text
                                    .Rows(i).Cells(5).Value = _Qty
                                    .Rows(i).Cells(6).Value = True
                                End If
                            Else
                                If UltraGrid1.Rows(i).Cells(6).Value = True Then
                                Else
                                    If CDbl(lblBalance.Text) = "0.00" Then
                                    Else

                                        .Rows(i).Cells(0).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(1).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(2).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(3).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(4).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(5).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(6).Appearance.BackColor = Color.Blue
                                        ' .Rows(i).Cells(5).Text = lblBalance.Text
                                        .Rows(i).Cells(5).Value = lblBalance.Text
                                        .Rows(i).Cells(6).Value = True
                                        lblBalance.Text = "0.00"
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If


                    End With
                    i = i + 1
                Next
                X = X + 1
            Next
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_WithData1()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        Dim X As Integer
        Dim _Date As Date
        Dim _BatchNo As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)
            Dim _qty1 As Double
            i = 0
            vcWhere = "M21Material='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "' and left(M21Sales_Order,2)='20' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _qty1 = 0
                vcWhere = "T12Stock_Code='" & M01.Tables(0).Rows(i)("M21Batch_No") & "' and T12Time>='" & M01.Tables(0).Rows(i)("M21Update_Time") & "'  "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    _qty1 = M02.Tables(0).Rows(0)("T12Qty")
                End If
                If M01.Tables(0).Rows(i)("M21Qty") > _qty1 Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("20Class") = M01.Tables(0).Rows(i)("M2120Class")
                    newRow("Grige Order No") = M01.Tables(0).Rows(i)("M21Sales_Order")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M21Batch_No")
                    If i = 0 Then
                        _BatchNo = "" & M01.Tables(0).Rows(i)("M21Batch_No")
                    Else
                        _BatchNo = _BatchNo & "','" & M01.Tables(0).Rows(i)("M21Batch_No")
                    End If
                    _To = Month(M01.Tables(0).Rows(i)("M21Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M21Date")) & "/" & Year(M01.Tables(0).Rows(i)("M21Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Age") = Diff.Days & " days"

                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = CDbl(M01.Tables(0).Rows(i)("M21Qty")) - _qty1
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)
                End If

                i = i + 1
            Next
            i = 0
            vcWhere = "M21Material='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M23Shade,1)='" & Trim(txtShade.Text) & "' and left(M21Sales_Order,2)='20' and M21Batch_No not in ('" & _BatchNo & "') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "UGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _qty1 = 0
                vcWhere = "T12Stock_Code='" & M01.Tables(0).Rows(i)("M21Batch_No") & "' and T12Time>='" & M01.Tables(0).Rows(i)("M21Update_Time") & "'  "
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    _qty1 = M02.Tables(0).Rows(0)("T12Qty")
                End If
                If M01.Tables(0).Rows(i)("M21Qty") > _qty1 Then
                    Dim newRow As DataRow = c_dataCustomer1.NewRow

                    newRow("20Class") = M01.Tables(0).Rows(i)("M2120Class")
                    newRow("Grige Order No") = M01.Tables(0).Rows(i)("M21Sales_Order")
                    newRow("Stock Code") = M01.Tables(0).Rows(i)("M21Batch_No")
                    _To = Month(M01.Tables(0).Rows(i)("M21Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M21Date")) & "/" & Year(M01.Tables(0).Rows(i)("M21Date"))
                    Diff = Today.Subtract(_To)
                    newRow("Age") = Diff.Days & " days"

                    'If Diff.Days < 30 Then
                    '    newRow("Age") = "Below 1 Month"
                    'ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                    '    newRow("Age") = "Below 2 Month"
                    'ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                    '    newRow("Age") = "Below 3 Month"
                    'Else
                    '    newRow("Age") = "above 3 Month"
                    'End If
                    Value = CDbl(M01.Tables(0).Rows(i)("M21Qty")) - _qty1
                    _VString = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _VString = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Available Qty(Kg)") = _VString
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)
                End If
                i = i + 1
            Next

            con.close()
            X = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                i = 0
                'If lblBalance.Text = "0.00" Then
                '    Exit Function
                'End If
                For Each uRow1 As UltraGridRow In UltraGrid1.Rows
                    Dim _Qty As Double
                    _Qty = 0
                    With UltraGrid1
                        If Trim(.Rows(i).Cells(2).Value) = Trim(UltraGrid2.Rows(X).Cells(0).Value) Then
                            _Qty = .Rows(i).Cells(4).Value
                            If CDbl(lblBalance.Text) >= _Qty Then
                                If UltraGrid1.Rows(i).Cells(6).Value = True Then
                                Else
                                    lblBalance.Text = CDbl(lblBalance.Text) - _Qty
                                    .Rows(i).Cells(0).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(1).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(2).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(3).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(4).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(5).Appearance.BackColor = Color.Blue
                                    .Rows(i).Cells(6).Appearance.BackColor = Color.Blue
                                    ' .Rows(i).Cells(5).Text = lblBalance.Text
                                    .Rows(i).Cells(5).Value = _Qty
                                    .Rows(i).Cells(6).Value = True
                                End If
                            Else
                                If UltraGrid1.Rows(i).Cells(6).Value = True Then
                                Else
                                    If CDbl(lblBalance.Text) = "0.00" Then
                                    Else

                                        .Rows(i).Cells(0).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(1).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(2).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(3).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(4).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(5).Appearance.BackColor = Color.Blue
                                        .Rows(i).Cells(6).Appearance.BackColor = Color.Blue
                                        ' .Rows(i).Cells(5).Text = lblBalance.Text
                                        .Rows(i).Cells(5).Value = lblBalance.Text
                                        .Rows(i).Cells(6).Value = True
                                        lblBalance.Text = "0.00"
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If


                    End With
                    i = i + 1
                Next
                X = X + 1
            Next
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Calculation_Balance()
        Dim I As Integer
        Dim Value As Double
        Dim _Vstring As String
        Try
            I = 0
            Value = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows

                With UltraGrid1
                    If IsNumeric(.Rows(I).Cells(5).Text) Then

                        If CDbl((.Rows(I).Cells(5).Text)) <= CDbl((.Rows(I).Cells(4).Text)) Then
                            Value = Value + CDbl((.Rows(I).Cells(5).Text))
                            If (.Rows(I).Cells(6).Value) = True Then
                            Else
                                .Rows(I).Cells(0).Appearance.BackColor = Color.White
                                .Rows(I).Cells(1).Appearance.BackColor = Color.White
                                .Rows(I).Cells(2).Appearance.BackColor = Color.White
                                .Rows(I).Cells(3).Appearance.BackColor = Color.White
                                .Rows(I).Cells(4).Appearance.BackColor = Color.White
                                .Rows(I).Cells(5).Appearance.BackColor = Color.White
                                .Rows(I).Cells(6).Appearance.BackColor = Color.White
                            End If
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


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Update_Records()
        Dim nvcFieldList1 As String
        Dim M01 As DataSet

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer

        For Each uRow As UltraGridRow In UltraGrid1.Rows

            With UltraGrid1
                If IsNumeric(.Rows(i).Cells(5).Text) Then

                    If CDbl((.Rows(i).Cells(5).Text)) <= CDbl((.Rows(i).Cells(4).Text)) Then

                    Else
                        MsgBox("Qty grater than to stock", MsgBoxStyle.Information, "Information ....")
                        .Rows(i).Cells(0).Appearance.BackColor = Color.Red
                        .Rows(i).Cells(1).Appearance.BackColor = Color.Red
                        .Rows(i).Cells(2).Appearance.BackColor = Color.Red
                        .Rows(i).Cells(3).Appearance.BackColor = Color.Red
                        .Rows(i).Cells(4).Appearance.BackColor = Color.Red
                        .Rows(i).Cells(5).Appearance.BackColor = Color.Red

                        .Rows(i).Selected = True
                        Exit For
                    End If
                End If
            End With
            i = i + 1
        Next
    End Function
  

    Private Sub UltraGrid1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles UltraGrid1.AfterRowUpdate
        '  MsgBox("")
        Calculation_Balance()
    End Sub


    Function Search_Tec_Spec()
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

            i = 0
            vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                With frmYarn_Booking
                    If i = 0 Then
                        .txtYarn1.Text = M01.Tables(0).Rows(i)("M22Yarn")
                        .txtCom1.Text = CInt(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                        Value = .txtGriege_Qty.Text
                        Value = Value * (.txtCom1.Text / 100)
                        If IsNumeric(.txtK_Wastage.Text) Then
                            Value = Value / ((100 - .txtK_Wastage.Text) / 100)
                        End If
                        .txtReq1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtReq1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        .pbCount1.Maximum = Value
                    ElseIf i = 1 Then
                        .txtYarn2.Text = M01.Tables(0).Rows(i)("M22Yarn")
                        .txtCom2.Text = CInt(M01.Tables(0).Rows(i)("M22Yarn_Cons"))

                        Value = .txtGriege_Qty.Text
                        Value = Value * (.txtCom2.Text / 100)
                        If IsNumeric(.txtK_Wastage.Text) Then
                            Value = Value / ((100 - .txtK_Wastage.Text) / 100)
                        End If
                        .txtReq2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtReq2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        .pbCount2.Maximum = Value
                    ElseIf i = 2 Then
                        .txtYarn3.Text = M01.Tables(0).Rows(i)("M22Yarn")
                        .txtCom3.Text = CInt(M01.Tables(0).Rows(i)("M22Yarn_Cons"))

                        Value = .txtGriege_Qty.Text
                        Value = Value * (.txtCom3.Text / 100)
                        If IsNumeric(.txtK_Wastage.Text) Then
                            Value = Value / ((100 - .txtK_Wastage.Text) / 100)
                        End If
                        .txtReq3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        .txtReq3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        .pbCount3.Maximum = Value
                    End If

                End With
               
                
                i = i + 1
            Next

            '----------------------------------------------------------------
            'Dim Z As Integer
            'Z = 0
            'i = 0
            'vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' "
            'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            'For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '    Z = 0
            '    vcWhere = "M33Description='" & Trim(M01.Tables(0).Rows(i)("M22Yarn")) & "'"
            '    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            '    For Each DTRow4 As DataRow In M02.Tables(0).Rows
            '        Dim newRow As DataRow = c_dataCustomer1.NewRow

            '        newRow("10Class") = M02.Tables(0).Rows(Z)("M3310Class")
            '        newRow("Description") = M02.Tables(0).Rows(Z)("M33Description")
            '        newRow("Stock Code") = M02.Tables(0).Rows(Z)("M33Stock_Code")

            '        c_dataCustomer1.Rows.Add(newRow)

            '        Z = Z + 1
            '    Next
            '    i = i + 1
            'Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Private Sub chkKnt_Plan_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkKnt_Plan.CheckedChanged
        On Error Resume Next
        Dim Value As Double

        If chkKnt_Plan.Checked = True Then
            Call Update_Records_Grige1()
            ' Value = CDbl(txtGriege_Qty.Text) + CDbl(txtLIB.Text)
            If chkKnt_Plan.Text = "Yarn Dye Plan" Then
                With frmDyed_Yarn1
                    .txtFabric_type.Text = frmLoad_Pln.txtFabrication.Text
                End With
                frmDyed_Yarn1.Show()
            Else
                Value = lblBalance.Text
                With frmYarn_Booking
                    .txtGriege_Qty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    .txtGriege_Qty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    .txtK_Wastage.Text = "2"
                End With
                Call Search_Tec_Spec()

                Dim TestString As String = Trim(txtGauge.Text)
                Dim TestArray() As String = Split(TestString)

                ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""} 
                Dim LastNonEmpty As Integer = -1
                For z1 As Integer = 0 To TestArray.Length - 1
                    If TestArray(z1) <> "" Then
                        LastNonEmpty += 1
                        TestArray(LastNonEmpty) = TestArray(z1)
                        ' If z1 = 2 Then
                        ''_Quality = TestArray(LastNonEmpty)
                        ''Exit For
                        'End If
                    End If
                Next
                strGuarge = Microsoft.VisualBasic.Left(TestArray(0), 4) & "-" & TestArray(3)
                frmYarn_Booking.Show()
            End If
            'frmKnt_Plan.Show()
        End If
    End Sub

    Private Sub UltraGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.Click
        On Error Resume Next
        Dim _RowIndex As Integer
        _RowIndex = UltraGrid1.ActiveRow.Index
        If Trim(UltraGrid1.Rows(_RowIndex).Cells(6).Value) = True Then
            UltraGrid1.Rows(_RowIndex).Cells(5).Value = UltraGrid1.Rows(_RowIndex).Cells(4).Value
        End If
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub txtGauge_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGauge.ValueChanged

    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Call Update_Records_Grige()
    End Sub

    Function Update_Records_Grige()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String

        Dim M01 As DataSet
        Dim i As Integer
        Dim ncQryType As String
        Try
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If IsNumeric(UltraGrid1.Rows(i).Cells(5).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next

            If lblBalance.Text = "0.00" Then
                nvcFieldList1 = "update M01Sales_Order_SAP set M01Status='I' where M01Sales_Order='" & strSales_Order & "' and M01Line_Item=" & strLine_Item & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            nvcFieldList1 = "delete from tmpBlock_SalesOrder where tmpSales_Order='" & strSales_Order & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            connection.Close()
            Me.Close()
            frmLoad_Pln.Close()
            frmDelivaryQuatnew.Load_Gride_SalesOrder()
            frmDelivaryQuatnew.Load_SalesOrder()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Records_Grige1()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim vcWhere As String

        Dim M01 As DataSet
        Dim i As Integer
        Dim ncQryType As String
        Try
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If IsNumeric(UltraGrid1.Rows(i).Cells(5).Value) Then
                    vcWhere = "T12Ref_No=" & Delivary_Ref & " and T12Sales_Order='" & strSales_Order & "' and T12Line_Item=" & strLine_Item & " and T12Stock_Code='" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetUse_Griege_Qty", New SqlParameter("@cQryType", "CGS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                    Else
                        ncQryType = "GADD"
                        nvcFieldList1 = "(T12Ref_No," & "T12Sales_Order," & "T12Line_Item," & "T12Date," & "T12Time," & "T12Stock_Code," & "T12Qty," & "T12Status," & "T12Confirm_By) " & "values(" & Delivary_Ref & ",'" & strSales_Order & "'," & strLine_Item & ",'" & Today & "','" & Now & "','" & Trim(UltraGrid1.Rows(i).Cells(2).Value) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Value) & "','N','-')"
                        up_GetSetConsume_Grige(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                    End If
                End If


                i = i + 1
            Next
            If lblBalance.Text = "0.00" Then
                nvcFieldList1 = "update M01Sales_Order_SAP set M01Status='I' where M01Sales_Order='" & strSales_Order & "' and M01Line_Item=" & strLine_Item & ""
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            'nvcFieldList1 = "delete from tmpBlock_SalesOrder where tmpSales_Order='" & strSales_Order & "'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            connection.Close()
            ' Me.Close()
            'frmLoad_Pln.Close()
            'frmDelivaryQuatnew.Load_Gride_SalesOrder()
            'frmDelivaryQuatnew.Load_SalesOrder()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Gride()

    End Sub

    Private Sub chkLIB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLIB.CheckedChanged

    End Sub
End Class