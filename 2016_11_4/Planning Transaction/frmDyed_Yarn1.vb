
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Public Class frmDyed_Yarn1
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Rowindex As Integer
    Dim _YarnQty As Double


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub frmDyed_Yarn1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmGriege_Stock.chkKnt_Plan.Checked = False
    End Sub

    Private Sub frmDyed_Yarn1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFabric_type.ReadOnly = True
        txtBasic_Yarn.ReadOnly = True
        txtNo_Colour.ReadOnly = True
        txtSpe_Yarn.ReadOnly = True
        ' txtShade.ReadOnly = True
        txtReq_Grg.ReadOnly = True
        txtDate.Text = Today

        Dim ToolTip1 As New ToolTip()
        ToolTip1.AutomaticDelay = 5000
        ToolTip1.InitialDelay = 1000
        ToolTip1.ReshowDelay = 500
        ToolTip1.ShowAlways = True
        Dim strTT As String
        ToolTip1.SetToolTip(cmdChart, cmdChart.Text & ControlChars.NewLine & "Graphical Yarn Plan")

        Dim ToolTip2 As New ToolTip()
        ToolTip2.AutomaticDelay = 5000
        ToolTip2.InitialDelay = 1000
        ToolTip2.ReshowDelay = 500
        ToolTip2.ShowAlways = True
        Dim strTT1 As String
        ToolTip2.SetToolTip(cmdDye_Yarn, cmdDye_Yarn.Text & ControlChars.NewLine & "Dyed Yarn Plan")

        Dim ToolTip3 As New ToolTip()
        ToolTip3.AutomaticDelay = 5000
        ToolTip3.InitialDelay = 1000
        ToolTip3.ReshowDelay = 500
        ToolTip3.ShowAlways = True
        Dim strTT3 As String
        ToolTip3.SetToolTip(cmdYarn_Request, cmdYarn_Request.Text & ControlChars.NewLine & "Yarn Request")

        Dim ToolTip4 As New ToolTip()
        ToolTip4.AutomaticDelay = 5000
        ToolTip4.InitialDelay = 1000
        ToolTip4.ReshowDelay = 500
        ToolTip4.ShowAlways = True
        Dim strTT4 As String
        ToolTip4.SetToolTip(cmdWinding, cmdWinding.Text & ControlChars.NewLine & "Soft Winding Plan")

        txt15Class.ReadOnly = True
        txtQty.ReadOnly = True

        txtQuality.ReadOnly = True
        txtQuality.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtReq_Grg.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Detailes()
        Call Load_Gride()
        Call Load_Gridewith_Data()
        ' Call Load_DataGD()
        Call Load_Gride_StockCode()
        Call Load_Gride_YarnStock()
        Call Delete_Transaction()
        Call Load_GrideDye_Plan()

    End Sub

    Function Load_Detailes()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        Dim MyText As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)

            txtReq_Grg.Text = frmGriege_Stock.txtGriege_Qty.Text
            txtQuality.Text = frmLoad_Pln.txtQuality.Text
            i = 0
            vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M22M_Class,2)='15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtNo_Colour.Text = M01.Tables(0).Rows.Count
                MyText = M01.Tables(0).Rows(0)("M22Yarn")
                Dim myIndex = MyText.IndexOf("(")
                txtBasic_Yarn.Text = Microsoft.VisualBasic.Left(M01.Tables(0).Rows(0)("M22Yarn"), myIndex)
            End If

            vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M22M_Class,2)<>'15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("M22Strich_Lenth") < 0 Then
                    chkNPL1.Checked = True
                    txtSpe_Yarn.Text = M01.Tables(0).Rows(0)("M22Yarn")

                Else
                    chkNPL2.Checked = True
                End If
            End If

            txtWastage.Text = frmLoad_Pln.txtDye_Wast.Text
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "DWS"))
            If isValidDataset(M01) Then
                txtDye_Wast.Text = M01.Tables(0).Rows(0)("M35D_WST")
                txtYarn_Wst.Text = M01.Tables(0).Rows(0)("M35Y_WST")

            End If

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
        c_dataCustomer1 = CustomerDataClass.MakeDataTableYarn_Dyeing
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Style = ColumnStyle.EditButton
            .DisplayLayout.Bands(0).Columns(6).Width = 110
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Style = ColumnStyle.EditButton
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_GrideDye_Plan()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableDye_Plan
        UltraGrid4.DataSource = c_dataCustomer1
        With UltraGrid4
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110

        End With
    End Function

    Function Load_Gride_StockCode()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableYarn_Stock
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(2).Width = 160
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_YarnStock()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableYarn
        UltraGrid3.DataSource = c_dataCustomer1
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 40
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(2).Width = 160
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = True
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Function Load_Gride_DataStock()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        i = 0
        Try

            If UltraGrid1.Rows(_Rowindex).Cells(0).Text <> "" Then
                vcWhere = "M3310Class='" & Trim(UltraGrid1.Rows(_Rowindex).Cells(0).Value) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YST"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Value = 0
                    Value = M01.Tables(0).Rows(i)("M33Qty")
                    'T10Dyed_Yarn Table

                    vcWhere = "T1015Class='" & txt15Class.Text & "' and T10Stock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DYN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Value = Value - M02.Tables(0).Rows(0)("Qty")
                    End If

                    'tmpBlock_YarnStock
                    vcWhere = "tmp15Class='" & txt15Class.Text & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmpUser<>'" & strDisname & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        newRow("Date") = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))

                        newRow("Qty") = _STValue
                        newRow("Log User") = M02.Tables(0).Rows(0)("tmpUser")

                        c_dataCustomer1.Rows.Add(newRow)



                    Else

                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        newRow("Date") = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        newRow("Qty") = _STValue
                        newRow("Log User") = "-"

                        c_dataCustomer1.Rows.Add(newRow)


                    End If

                    i = i + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                c_dataCustomer1.Rows.Add(newRow1)

            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Con.close()
            End If
        End Try
    End Function

    Function Load_GrideData_YarnStock()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer
        Dim _Date As Date

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        i = 0
        Try
            txtSearch.Text = ""

            If UltraGrid1.Rows(_Rowindex).Cells(0).Text <> "" Then
                vcWhere = "M33Yarn_Location='2020' and left(M33Description,4)='" & Microsoft.VisualBasic.Left(txtBasic_Yarn.Text, 4) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YST"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Value = 0
                    Value = M01.Tables(0).Rows(i)("M33Qty")
                    'T10Dyed_Yarn Table

                    vcWhere = "T1015Class='" & txt15Class.Text & "' and T10Stock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DYN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Value = Value - M02.Tables(0).Rows(0)("Qty")
                    End If

                    vcWhere = "tmp10Class='" & M01.Tables(0).Rows(i)("M3310Class") & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmp15Class<>'" & UltraGrid1.Rows(_Rowindex).Cells(0).Value & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        ' MsgBox("")
                        If IsDBNull(M02.Tables(0).Rows(0)("Qty")) Then
                        Else
                            Value = Value - M02.Tables(0).Rows(0)("Qty")
                        End If
                    End If

                    'tmpBlock_YarnStock
                    vcWhere = "tmp10Class='" & M01.Tables(0).Rows(i)("M3310Class") & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmpUser<>'" & strDisname & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        Diff = Today.Subtract(_Date)
                        newRow("Age") = Diff.Days
                        newRow("Qty") = _STValue
                        newRow("Log User") = M02.Tables(0).Rows(0)("tmpUser")

                        c_dataCustomer1.Rows.Add(newRow)



                    Else

                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        Diff = Today.Subtract(_Date)
                        newRow("Age") = Diff.Days
                        newRow("Qty") = _STValue
                        newRow("Log User") = "-"

                        c_dataCustomer1.Rows.Add(newRow)


                    End If

                    i = i + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                c_dataCustomer1.Rows.Add(newRow1)

            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GrideData_YarnStockLike()
        Dim i As Integer
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim _VString As String
        Dim Diff As TimeSpan
        Dim _To As Date
        'Dim Value As Double
        Dim _Rowcount As Integer
        Dim _Date As Date

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        i = 0
        Try

            If UltraGrid1.Rows(_Rowindex).Cells(0).Text <> "" Then
                vcWhere = "M33Yarn_Location='2020' and M33Description like '%" & txtSearch.Text & "%' and left(M33Description,4)='" & Microsoft.VisualBasic.Left(txtBasic_Yarn.Text, 4) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YST"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Value = 0
                    Value = M01.Tables(0).Rows(i)("M33Qty")
                    'T10Dyed_Yarn Table

                    vcWhere = "T1015Class='" & txt15Class.Text & "' and T10Stock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DYN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Value = Value - M02.Tables(0).Rows(0)("Qty")
                    End If

                    'tmpBlock_YarnStock
                    vcWhere = "tmp15Class='" & M01.Tables(0).Rows(i)("M3310Class") & "' and tmpStock_Code='" & M01.Tables(0).Rows(i)("M33Stock_Code") & "' and tmpUser<>'" & strDisname & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "BTY"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        Diff = Today.Subtract(_Date)
                        newRow("Age") = Diff.Days
                        newRow("Qty") = _STValue
                        newRow("Log User") = M02.Tables(0).Rows(0)("tmpUser")

                        c_dataCustomer1.Rows.Add(newRow)



                    Else

                        Dim newRow As DataRow = c_dataCustomer1.NewRow

                        Dim _STValue As String

                        _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        newRow("##") = False
                        newRow("10 Class") = M01.Tables(0).Rows(i)("M3310Class")
                        newRow("Stock Code") = M01.Tables(0).Rows(i)("M33Stock_Code")
                        newRow("Description") = M01.Tables(0).Rows(i)("M33Description")
                        _Date = Month(M01.Tables(0).Rows(i)("M33Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M33Date")) & "/" & Year(M01.Tables(0).Rows(i)("M33Date"))
                        Diff = Today.Subtract(_Date)
                        newRow("Age") = Diff.Days
                        newRow("Qty") = _STValue
                        newRow("Log User") = "-"

                        c_dataCustomer1.Rows.Add(newRow)


                    End If

                    i = i + 1
                Next
                Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                c_dataCustomer1.Rows.Add(newRow1)

            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

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
        Dim _Rowcount As Integer
        Dim characterToRemove As String

        Try
            Dim con = New SqlConnection()
            con = DBEngin.GetConnection(True)


            Dim Z As Integer
            Z = 0
            i = 0
            vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M22M_Class,2)='15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                Dim _STValue As String
                Dim _Rcode As String

                newRow("15Class") = M01.Tables(0).Rows(i)("M22M_Class")
                newRow("Description") = M01.Tables(0).Rows(i)("M22Yarn")
                newRow("Composition") = CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                _Rcode = ""
                _Rcode = Microsoft.VisualBasic.Right(M01.Tables(0).Rows(i)("M22Yarn"), 6)
                _Rcode = Microsoft.VisualBasic.Left(_Rcode, 5)

                characterToRemove = "Y"
                _Rcode = (Replace(_Rcode, characterToRemove, ""))
                _Rcode = Trim(_Rcode)
                vcWhere = "M14Order='" & _Rcode & "' and M14Type='Y'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCDE1"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M02) Then
                    If Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "D" Then
                        newRow("Shade") = "DARK"
                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "L" Then
                        newRow("Shade") = "LIGHT"
                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "M" Then
                        newRow("Shade") = "MARL"
                    ElseIf Trim(M02.Tables(0).Rows(0)("M14Shade_Cat")) = "MARL" Then
                        newRow("Shade") = "MARL"
                    End If
                End If
                Value = 0
                If IsNumeric(txtReq_Grg.Text) Then
                    Value = CDbl(txtReq_Grg.Text)
                    If IsNumeric(txtDye_Wast.Text) Then
                        Value = Value / ((100 - txtDye_Wast.Text) / 100)
                    End If
                    Value = Value * CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                    Value = Value / 100

                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                newRow("Dyed Yarn Req.Knit") = _STValue
                _STValue = ""
                If IsNumeric(txtYarn_Wst.Text) Then
                    Value = Value / ((100 - txtYarn_Wst.Text) / 100)
                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                newRow("Yarn Req - Dyed Yarn") = _STValue
                newRow("Balance Qty for Allocate") = _STValue
                newRow("No of Cons") = CInt(Value / 1.05)
                c_dataCustomer1.Rows.Add(newRow)


                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            c_dataCustomer1.Rows.Add(newRow1)
            _Rowcount = UltraGrid1.Rows.Count
            UltraGrid1.Rows(_Rowcount - 1).Cells(0).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(1).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(2).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(3).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(4).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(5).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(6).Appearance.BackColor = Color.DeepSkyBlue
            UltraGrid1.Rows(_Rowcount - 1).Cells(7).Appearance.BackColor = Color.DeepSkyBlue
            i = 0
            vcWhere = "M22Quality='" & Trim(frmLoad_Pln.txtQuality.Text) & "' and left(M22M_Class,2)<>'15'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                Dim _STValue As String

                _STValue = ""
                ' newRow("15Class") = M01.Tables(0).Rows(i)("M22M_Class")
                newRow("Description") = M01.Tables(0).Rows(i)("M22Yarn")
                newRow("Composition") = CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                If IsNumeric(txtReq_Grg.Text) Then
                    Value = CDbl(txtReq_Grg.Text)
                    If IsNumeric(txtDye_Wast.Text) Then
                        Value = Value / 0.98
                    End If
                    Value = Value * CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
                    Value = Value / 100

                    _STValue = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _STValue = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                newRow("Dyed Yarn Req.Knit") = _STValue
                newRow("Yarn Req - Dyed Yarn") = _STValue
                c_dataCustomer1.Rows.Add(newRow)


                i = i + 1
            Next

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            c_dataCustomer1.Rows.Add(newRow2)
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function

    Function Load_Gridewith_DyePlan()
        Dim i As Integer

        Try
          

            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If Trim(UltraGrid1.Rows(i).Cells(0).Text) <> "" Then
                    If Trim(UltraGrid1.Rows(i).Cells(9).Text) <> "" Then
                        Dim newRow As DataRow = c_dataCustomer1.NewRow


                        newRow("15Class") = Trim(UltraGrid1.Rows(i).Cells(0).Value)
                        newRow("Description") = Trim(UltraGrid1.Rows(i).Cells(1).Value)
                        newRow("Shade") = Trim(UltraGrid1.Rows(i).Cells(2).Value)
                        newRow("Allocate Yarn") = Trim(UltraGrid1.Rows(i).Cells(9).Value)
                        newRow("Allocate Con") = CInt(CDbl(Trim(UltraGrid1.Rows(i).Cells(9).Value)) / 1.05)
                        c_dataCustomer1.Rows.Add(newRow)
                    End If
                End If
                i = i + 1
            Next
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'Con.close()
            End If
        End Try
    End Function


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

    Private Sub UltraGrid1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid1.ClickCellButton
        On Error Resume Next
        Dim _ColumIndex As Integer


        If UltraGroupBox6.Visible = True Then
            UltraGroupBox6.Visible = False
        Else
            _Rowindex = UltraGrid1.ActiveRow.Index
            _ColumIndex = UltraGrid1.ActiveCell.Column.Index

            If Trim(UltraGrid1.Rows(_Rowindex).Cells(0).Text) <> "" Then
                If _ColumIndex = 5 Then
                    txt15Class.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Value
                    txtQty.Text = UltraGrid1.Rows(_Rowindex).Cells(5).Value
                    If txt15Class.Text <> "" Then
                        UltraGroupBox6.Visible = True
                    End If
                    Call Load_Gride_StockCode()
                    Call Load_Gride_DataStock()
                ElseIf _ColumIndex = 9 Then
                    If UltraGroupBox7.Visible = True Then
                        UltraGroupBox7.Visible = False
                    Else
                        txtBalance.Text = UltraGrid1.Rows(_Rowindex).Cells(7).Value
                        UltraGroupBox6.Visible = False
                        UltraGroupBox7.Visible = True

                        Call Load_Gride_YarnStock()
                        Call Load_GrideData_YarnStock()
                    End If
                End If

            End If
        End If
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub UltraGrid1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid1.KeyUp

    End Sub

    Private Sub UltraGrid2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid2.AfterCellUpdate

    End Sub

    Private Sub UltraGrid2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles UltraGrid2.AfterRowUpdate

        Dim _Index As Integer
        Dim i As Integer
        Dim Value As Double
        Dim _stBalance As String

        Try
            _Index = UltraGrid2.ActiveRow.Index
            If UltraGrid2.Rows(_Index).Cells(1).Text <> "" Then
                ' MsgBox(Trim(UltraGrid2.Rows(_Index).Cells(0).Text))
                If UltraGrid2.Rows(_Index).Cells(0).Text = True Then

                    If CDbl(UltraGrid2.Rows(_Index).Cells(4).Value) <= txtQty.Text And Trim(UltraGrid2.Rows(_Index).Cells(6).Text) <> strDisname Then
                        UltraGrid2.Rows(_Index).Cells(5).Value = UltraGrid2.Rows(_Index).Cells(4).Value
                    End If
                Else
                    ' UltraGrid2.Rows(_Index).Cells(5).Value = ""
                End If
                If Trim(UltraGrid2.Rows(_Index).Cells(6).Value) <> "" Then

                    If Trim(UltraGrid2.Rows(_Index).Cells(6).Value) = "-" Then
                        '  Call Update_Transaction(Trim(txt15Class.Text), UltraGrid2.Rows(_Index).Cells(1).Value, UltraGrid2.Rows(_Index).Cells(5).Value)

                        Value = 0
                        _stBalance = ""
                        i = 0
                        For Each uRow As UltraGridRow In UltraGrid2.Rows
                            If IsNumeric(UltraGrid2.Rows(i).Cells(5).Value) Then
                                Value = Value + UltraGrid2.Rows(i).Cells(5).Value
                            End If
                            i = i + 1
                        Next

                        If Value <= txtQty.Text Then
                            Call Update_Transaction(Trim(txt15Class.Text), UltraGrid2.Rows(_Index).Cells(1).Value, UltraGrid2.Rows(_Index).Cells(5).Value)

                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            UltraGrid1.Rows(_Rowindex).Cells(4).Value = _stBalance

                            Value = CDbl(txtQty.Text) - Value
                            Value = Value + (Value * (txtYarn_Wst.Text / 100))
                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            UltraGrid1.Rows(_Rowindex).Cells(6).Value = _stBalance

                            UltraGrid1.Rows(_Rowindex).Cells(7).Value = CInt(Value / 1.05)
                        Else
                            MsgBox("Stock Quantity miss match please try again", MsgBoxStyle.Exclamation, "Technova ......")
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Sub

    Function Load_DataGD()
        Dim _Index As Integer
        Dim i As Integer
        Dim Value As Double
        Dim _stBalance As String

        Try
            _Index = 0
            For Each uRow1 As UltraGridRow In UltraGrid2.Rows

                If UltraGrid2.Rows(_Index).Cells(1).Text <> "" Then
                    ' MsgBox(Trim(UltraGrid2.Rows(_Index).Cells(0).Text))
                    If UltraGrid2.Rows(_Index).Cells(0).Text = True Then

                        If UltraGrid2.Rows(_Index).Cells(4).Value <= txtQty.Text And Trim(UltraGrid2.Rows(_Index).Cells(7).Text) <> strDisname Then
                            UltraGrid2.Rows(_Index).Cells(6).Value = UltraGrid2.Rows(_Index).Cells(4).Value
                        End If
                    Else
                        ' UltraGrid2.Rows(_Index).Cells(5).Value = ""
                    End If

                    Call Update_Transaction(Trim(txt15Class.Text), UltraGrid2.Rows(_Index).Cells(1).Value, UltraGrid2.Rows(_Index).Cells(5).Value)

                    Value = 0
                    _stBalance = ""
                    i = 0
                    For Each uRow As UltraGridRow In UltraGrid2.Rows
                        If IsNumeric(UltraGrid2.Rows(i).Cells(5).Value) Then
                            Value = Value + UltraGrid2.Rows(i).Cells(5).Value
                        End If
                        i = i + 1
                    Next

                    If Value <= txtQty.Text Then
                        _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        UltraGrid1.Rows(_Rowindex).Cells(4).Value = _stBalance

                        Value = CDbl(txtQty.Text) - Value
                        Value = Value + (Value * (txtYarn_Wst.Text / 100))
                        _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        UltraGrid1.Rows(_Rowindex).Cells(6).Value = _stBalance

                        UltraGrid1.Rows(_Rowindex).Cells(7).Value = CInt(Value / 1.05)
                    Else
                        ' MsgBox("Stock Quantity miss match please try again", MsgBoxStyle.Exclamation, "Technova ......")
                        Exit Function
                    End If
                End If
                _Index = _Index + 1
            Next
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Function

    Private Sub UltraGrid2_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid2.CellChange
        'Dim _Index As Integer
        '_Index = UltraGrid2.ActiveRow.Index
        'If UltraGrid2.Rows(_Index).Cells(1).Text <> "" Then
        '    ' MsgBox(Trim(UltraGrid2.Rows(_Index).Cells(0).Text))
        '    If UltraGrid2.Rows(_Index).Cells(0).Text = True Then

        '        If UltraGrid2.Rows(_Index).Cells(4).Value <= txtQty.Text And Trim(UltraGrid2.Rows(_Index).Cells(6).Text) <> strDisname Then
        '            UltraGrid2.Rows(_Index).Cells(5).Value = UltraGrid2.Rows(_Index).Cells(4).Value
        '        End If
        '    Else
        '        ' UltraGrid2.Rows(_Index).Cells(5).Value = ""
        '    End If
        'End If
    End Sub

    Function Delete_Transaction()
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try

            nvcFieldList1 = "delete from tmpBlock_Yarn_Stock_Code where tmpUser='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Function

    Function Update_Transaction(ByVal str15 As String, ByVal strStock As String, ByVal strQty As Double)
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim vcFieldList As String
        Dim ncQryType As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try

            vcWhere = "tmp15Class='" & UltraGrid1.Rows(_Rowindex).Cells(0).Value & "' and tmpStock_Code='" & strStock & "' and tmpUser='" & strDisname & "' and tmp10Class='" & str15 & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YSS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If strDisname = Trim(M01.Tables(0).Rows(0)("tmpUser")) Then
                    nvcFieldList1 = "update tmpBlock_Yarn_Stock_Code set tmpQty='" & strQty & "' where tmp15Class='" & UltraGrid1.Rows(_Rowindex).Cells(0).Value & "' and tmpStock_Code='" & strStock & "' and tmpUser='" & strDisname & "' and tmp10Class='" & str15 & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            Else

                ncQryType = "ADD1"
                nvcFieldList1 = "(tmpRefNo," & "tmp15Class," & "tmpStock_Code," & "tmpQty," & "tmpUser," & "tmp10Class) " & "values(" & Delivary_Ref & ",'" & UltraGrid1.Rows(_Rowindex).Cells(0).Value & "','" & strStock & "','" & strQty & "','" & strDisname & "','" & str15 & "')"
                up_GetSetDelivary_Planning(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If
            transaction.Commit()
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try


    End Function

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub


    Private Sub UltraGrid3_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles UltraGrid3.AfterRowUpdate

        Dim _Index As Integer
        Dim i As Integer
        Dim Value As Double
        Dim _stBalance As String
        Dim _String As String

        Try
            _Index = UltraGrid3.ActiveRow.Index
            If UltraGrid3.Rows(_Index).Cells(1).Text <> "" Then
                ' MsgBox(Trim(UltraGrid2.Rows(_Index).Cells(0).Text))
                If UltraGrid3.Rows(_Index).Cells(0).Text = True Then

                    If UltraGrid3.Rows(_Index).Cells(4).Value <= txtBalance.Text And Trim(UltraGrid3.Rows(_Index).Cells(7).Text) <> strDisname Then
                        UltraGrid3.Rows(_Index).Cells(6).Value = UltraGrid3.Rows(_Index).Cells(4).Value
                    End If
                Else
                    ' UltraGrid2.Rows(_Index).Cells(5).Value = ""
                End If

                If Trim(UltraGrid3.Rows(_Index).Cells(6).Text) <> "" Then
                    Dim _10Class As String
                    _10Class = Trim(UltraGrid3.Rows(_Index).Cells(1).Value)
                    If Trim(UltraGrid3.Rows(_Index).Cells(7).Value) = "-" Then
                        '  Call Update_Transaction(Trim(_10Class), UltraGrid3.Rows(_Index).Cells(3).Value, UltraGrid3.Rows(_Index).Cells(6).Value)

                        Value = 0
                        _stBalance = ""
                        i = 0
                        For Each uRow As UltraGridRow In UltraGrid3.Rows
                            If IsNumeric(UltraGrid3.Rows(i).Cells(6).Value) Then
                                Value = Value + UltraGrid3.Rows(i).Cells(6).Value
                            End If
                            i = i + 1
                        Next

                        If Value <= txtBalance.Text Then
                            Call Update_Transaction(Trim(_10Class), UltraGrid3.Rows(_Index).Cells(3).Value, UltraGrid3.Rows(_Index).Cells(6).Value)
                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            UltraGrid1.Rows(_Rowindex).Cells(9).Value = _stBalance

                            Value = _stBalance
                            '  Value = Value + (Value * (txtYarn_Wst.Text / 100))
                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            UltraGrid1.Rows(_Rowindex).Cells(9).Value = _stBalance

                            
                            If txtDis.Text <> "" Then

                                'txtDis.Focus()

                                'SendKeys.Send("{ENTER}")
                                _String = txtDis.Text & ";" & UltraGrid1.Rows(_Rowindex).Cells(0).Value & "- Allocate Cons(" & CInt(CDbl(_stBalance) / 1.05) & ")"
                                txtDis.Text = _String
                                'Dim Words As String() = _String.Split(New Char() {";"c})
                                'txtDis.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Value & "- Allocate Cons(" & CInt(CDbl(_stBalance) / 1.05) & ")"

                                '                                txtDis.Text = Words(0)

                                'SendKeys.Send("{ENTER}")
                                '                                txtDis.Text = Words(1)
                            Else
                                txtDis.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Value & "- Allocate Cons(" & CInt(CDbl(_stBalance) / 1.05) & ")"
                                _String = txtDis.Text
                            End If
                            '  UltraGrid1.Rows(_Rowindex).Cells(7).Value = CInt(Value / 1.05)
                        Else
                            MsgBox("Stock Quantity miss match please try again", MsgBoxStyle.Exclamation, "Technova ......")
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Sub



    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        Call Load_Gride_YarnStock()
        Call Load_GrideData_YarnStockLike()
    End Sub

    Private Sub UltraGrid3_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid3.InitializeLayout

    End Sub

    Function Search_WeekNo()
        On Error Resume Next
        Dim _Date As Date

        _Date = txtDate.Text
        If txtWeek.Text <> "" Then
        Else
            txtYear.Text = Year(txtDate.Text)
            txtWeek.Text = DatePart("WW", _Date, FirstDayOfWeek.Monday)
        End If
    End Function
    Private Sub cmdDye_Yarn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDye_Yarn.Click
        If OPR12.Visible = True Then
            OPR12.Visible = False
        Else
            OPR12.Visible = True
            Call Load_GrideDye_Plan()
            Call Load_Gridewith_DyePlan()

        End If
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
       
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
        Dim _MC As String
        Dim _string As String
        Dim result1 As DialogResult
        Dim _Date As Date
        Dim i As Integer
        Dim _DyeStartDate As Date
        Dim _DyeEnd_Date As Date
        Dim _Balance_Qty As Double
        Dim X As Integer
        Dim _Status As Boolean
        Dim _No_Of_Batch As Integer

        Try
            _Date = txtDate.Text
            If txtWeek.Text <> "" Then
            Else
                txtYear.Text = Year(txtDate.Text)
                txtWeek.Text = DatePart("WW", _Date, FirstDayOfWeek.Monday)
            End If

            If IsNumeric(txtWeek.Text) Then
            Else
                MsgBox("Please enter the correct Week No", MsgBoxStyle.Information, "Information ......")
                txtWeek.Focus()
                Exit Sub
            End If

            If IsNumeric(txtYear.Text) Then
            Else
                MsgBox("Please enter the correct Year", MsgBoxStyle.Information, "Information ......")
                txtYear.Focus()
                Exit Sub
            End If
            Dim _WeekDel_Date As Date

            If Trim(txtWeek.Text) <> "" Then
                If Trim(txtYear.Text) <> "" Then
                Else
                    MsgBox("Please enter the Year", MsgBoxStyle.Information, "Information .......")
                    Exit Sub
                End If
                Dim StartDate As New DateTime(txtYear.Text, 1, 1)
                _WeekDel_Date = DateAdd(DateInterval.WeekOfYear, txtWeek.Text - 1, StartDate)
                ' MsgBox(WeekdayName(Weekday(_WeekDel_Date)))
                If (WeekdayName(Weekday(_WeekDel_Date))) = "Sunday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+4)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Monday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+3)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Tuesday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+2)


                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Wednesday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(+1)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Friday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(-1)
                ElseIf (WeekdayName(Weekday(_WeekDel_Date))) = "Saturday" Then
                    _WeekDel_Date = _WeekDel_Date.AddDays(-1)
                End If

            Else
                If txtDate.Text > Today Then
                    _WeekDel_Date = txtDate.Text
                Else
                    MsgBox("Please check the delivary date", MsgBoxStyle.Information, "Information .....")
                    txtDate.Focus()
                    Exit Sub
                End If


            End If

            'Check Dye MC Block
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "DMC"))
            If isValidDataset(M01) Then
                If Trim(M01.Tables(0).Rows(0)("M37User")) = strDisname Then
                Else
                    result1 = MessageBox.Show(M01.Tables(0).Rows(0)("M37User") & " used this dye machine.", _
                                    "Error ....", _
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        Exit Sub
                    End If
                End If
            Else
                ncQryType = "ADD"
                nvcFieldList1 = "(M37Date," & "M37User) " & "values('" & Today & "','" & strDisname & "')"
                up_GetSetBlock_DyeMC(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
            End If

            transaction.Commit()
            connection.Close()
            '-------------------------------------------------------------------------------------------------
            Call Load_GrideDye_Plan()
            i = 0
            _DyeStartDate = _Date.AddDays(-14)

            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If UltraGrid1.Rows(i).Cells(0).Text <> "" Then
                    Dim newRow1 As DataRow = c_dataCustomer1.NewRow



                    newRow1("15Class") = UltraGrid1.Rows(i).Cells(0).Value
                    newRow1("Description") = UltraGrid1.Rows(i).Cells(1).Value

                    c_dataCustomer1.Rows.Add(newRow1)
                Else
                    i = i + 1
                    Continue For
                End If
                i = i + 1
            Next

            Dim newRow As DataRow = c_dataCustomer1.NewRow
            newRow("15Class") = ""
            newRow("Description") = ""
            c_dataCustomer1.Rows.Add(newRow)
            '----------------------------------------------------------------------------
            'Check the Suterble Machine
            _Balance_Qty = 0
            i = 0
            _Status = False

            connection = DBEngin.GetConnection(True)
            connectionCreated = True
            transaction = connection.BeginTransaction()
            transactionCreated = True

            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If UltraGrid1.Rows(i).Cells(0).Text <> "" Then
                    _Balance_Qty = CInt(UltraGrid1.Rows(i).Cells(7).Value) + 5
                    vcFieldList = "left(m36mc_no,1)='Y' and m36Max_Qty>='" & _Balance_Qty & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "MCC"), New SqlParameter("@vcWhereClause1", vcFieldList))
                    X = 0
                    _No_Of_Batch = 0
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows
                        _DyeStartDate = _Date.AddDays(-14)
                        If M01.Tables(0).Rows(X)("m36min_qty") <= _Balance_Qty And M01.Tables(0).Rows(X)("m36max_qty") > _Balance_Qty Then
                            vcFieldList = "tmpMC_No='" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "'"
                            M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "DYM"), New SqlParameter("@vcWhereClause1", vcFieldList))
                            If isValidDataset(M02) Then
                                ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                If CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")) <= _DyeStartDate Then
                                    _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                    ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                    _DyeEnd_Date = CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")).AddHours(+10 * _No_Of_Batch)
                                    _DyeStartDate = M02.Tables(0).Rows(X)("tmpEnd_Time")
                                    If CDate(_DyeEnd_Date) > txtDate.Text Then

                                    Else
                                        _Status = True
                                        UltraGrid4.Rows(i).Cells(2).Value = _No_Of_Batch
                                        UltraGrid4.Rows(i).Cells(3).Value = _Balance_Qty
                                        UltraGrid4.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                        UltraGrid4.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                        UltraGrid1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                        ncQryType = "ADD1"
                                        nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & UltraGrid1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & UltraGrid1.Rows(i).Cells(3).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                        up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                        'transaction.Commit()
                                        Exit For
                                    End If
                                End If
                            Else
                                _Status = True
                                _DyeStartDate = _DyeStartDate & " 7:30AM"
                                _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                _DyeEnd_Date = _DyeStartDate.AddHours(+10)

                                UltraGrid4.Rows(i).Cells(2).Value = _No_Of_Batch
                                UltraGrid4.Rows(i).Cells(3).Value = _Balance_Qty
                                UltraGrid4.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                UltraGrid4.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                UltraGrid1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                ncQryType = "ADD1"
                                nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & UltraGrid1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & UltraGrid1.Rows(i).Cells(1).Value & "','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                ' transaction.Commit()
                                Exit For
                            End If
                        End If
                        X = X + 1
                    Next
                Else
                    Dim _Machine_No As String
                    _Machine_No = ""
                    If _Status = False Then
                        vcFieldList = "left(m36mc_no,1)='Y' and m36Max_Qty<='" & _Balance_Qty & "'"
                        M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "MCC"), New SqlParameter("@vcWhereClause1", vcFieldList))
                        X = 0
                        _No_Of_Batch = 0
                        For Each DTRow3 As DataRow In M01.Tables(0).Rows
                            _DyeStartDate = _Date.AddDays(-14)
                            If M01.Tables(0).Rows(X)("m36min_qty") <= _Balance_Qty Then
                                vcFieldList = "tmpMC_No='" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "'"
                                M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDyeMC_Process", New SqlParameter("@cQryType", "DYM"), New SqlParameter("@vcWhereClause1", vcFieldList))
                                If isValidDataset(M02) Then
                                    ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                    If CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")) <= _DyeStartDate Then
                                        _No_Of_Batch = _No_Of_Batch + (M01.Tables(0).Rows(X)("m36max_qty") / _Balance_Qty)
                                        ' MsgBox(CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")))
                                        _DyeEnd_Date = CDate(M02.Tables(0).Rows(X)("tmpEnd_Time")).AddHours(+10 * _No_Of_Batch)
                                        _DyeStartDate = M02.Tables(0).Rows(X)("tmpEnd_Time")
                                        If CDate(_DyeEnd_Date) > txtDate.Text Then

                                        Else
                                            _Status = True
                                            UltraGrid4.Rows(i).Cells(2).Value = _No_Of_Batch
                                            UltraGrid4.Rows(i).Cells(3).Value = _Balance_Qty
                                            UltraGrid4.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                            UltraGrid4.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                            UltraGrid1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                            ncQryType = "ADD1"
                                            nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & UltraGrid1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & UltraGrid1.Rows(i).Cells(3).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                            up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                            'transaction.Commit()
                                            Exit For
                                        End If
                                    End If
                                Else
                                    _Status = True
                                    _DyeStartDate = _DyeStartDate & " 7:30AM"
                                    _No_Of_Batch = _No_Of_Batch + (_Balance_Qty / M01.Tables(0).Rows(X)("m36max_qty"))
                                    _DyeEnd_Date = _DyeStartDate.AddHours(+10)

                                    UltraGrid4.Rows(i).Cells(2).Value = _No_Of_Batch
                                    UltraGrid4.Rows(i).Cells(3).Value = _Balance_Qty
                                    UltraGrid4.Rows(i).Cells(4).Value = CInt(_Balance_Qty / 1.05)
                                    UltraGrid4.Rows(i).Cells(5).Value = M01.Tables(0).Rows(X)("M36MC_No")
                                    If _Machine_No <> "" Then
                                        _Machine_No = _Machine_No & "/" & M01.Tables(0).Rows(X)("M36MC_No")
                                    Else
                                        _Machine_No = M01.Tables(0).Rows(X)("M36MC_No")
                                    End If
                                    UltraGrid1.Rows(i).Cells(7).Value = UltraGrid1.Rows(i).Cells(7).Value + (5 * _No_Of_Batch)

                                    ncQryType = "ADD1"
                                    nvcFieldList1 = "(tmpRefNo," & "tmpMC_No," & "tmp15Class," & "tmpDate," & "tmpSTTime," & "tmpEnd_Time," & "tmpQty," & "tmpStatus," & "tmpShade," & "tmpCatagary," & "tmpCon) " & "values('" & Delivary_Ref & "','" & Trim(M01.Tables(0).Rows(X)("M36MC_No")) & "','" & UltraGrid1.Rows(i).Cells(0).Value & "','" & CDate(_DyeEnd_Date) & "','" & _DyeStartDate & "','" & _DyeEnd_Date & "','" & _Balance_Qty & "','I','" & UltraGrid1.Rows(i).Cells(1).Value & "','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & CInt(_Balance_Qty / 1.05) & "')"
                                    up_GetSetYarn_DyeMCPln(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                    ' transaction.Commit()
                                    Exit For
                                End If
                            End If
                            X = X + 1
                        Next
                    End If
                End If
                i = i + 1
            Next

            transaction.Commit()
            connection.Close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub txtDate_AfterDropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.AfterDropDown
        Call Search_WeekNo()
    End Sub


    Private Sub txtDate_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDate.BeforeDropDown

    End Sub

    Private Sub txtDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.LostFocus
        Call Search_WeekNo()
    End Sub

    Private Sub cmdChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChart.Click
        frmYrn_Dye_Pln.Show()
    End Sub

    Private Sub cmdKnt_Chart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKnt_Chart.Click
        frmKnitting_Plan_Board.Show()
    End Sub

    Private Sub cmdExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub txtWeek_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWeek.ValueChanged

    End Sub

    Private Sub txtYear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtYear.ValueChanged

    End Sub

    Private Sub cmdYarn_Request_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYarn_Request.Click

    End Sub

    Private Sub cmdWinding_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWinding.Click

    End Sub

    Private Sub UltraGroupBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox5.Click

    End Sub
End Class