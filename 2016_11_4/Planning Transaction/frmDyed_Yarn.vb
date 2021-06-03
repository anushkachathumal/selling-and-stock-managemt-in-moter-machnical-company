
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors

Public Class frmDyed_Yarn
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Rowindex As Integer


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub frmDyed_Yarn_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmGriege_Stock.chkKnt_Plan.Checked = False
    End Sub

    Private Sub frmDyed_Yarn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFabric_type.ReadOnly = True
        txtBasic_Yarn.ReadOnly = True
        txtNo_Colour.ReadOnly = True
        txtSpe_Yarn.ReadOnly = True
        ' txtShade.ReadOnly = True
        txtReq_Grg.ReadOnly = True

        'Dim ToolTip1 As New ToolTip()
        'ToolTip1.AutomaticDelay = 5000
        'ToolTip1.InitialDelay = 1000
        'ToolTip1.ReshowDelay = 500
        'ToolTip1.ShowAlways = True
        'Dim strTT As String
        'ToolTip1.SetToolTip(cmdGraphic, cmdGraphic.Text & ControlChars.NewLine & "Graphical Yarn Plan")

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
            .DisplayLayout.Bands(0).Columns(2).Width = 100
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Style = ColumnStyle.EditButton
            .DisplayLayout.Bands(0).Columns(6).Width = 110
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Style = ColumnStyle.EditButton
            '   .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(3).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            ''  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
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
                vcWhere = "M33Yarn_Location='2020' "
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
                vcWhere = "M33Yarn_Location='2020' and M33Description like '%" & txtSearch.Text & "%' "
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

                newRow("15Class") = M01.Tables(0).Rows(i)("M22M_Class")
                newRow("Description") = M01.Tables(0).Rows(i)("M22Yarn")
                newRow("Composition") = CDbl(M01.Tables(0).Rows(i)("M22Yarn_Cons"))
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
                If _ColumIndex = 4 Then
                    txt15Class.Text = UltraGrid1.Rows(_Rowindex).Cells(0).Value
                    txtQty.Text = UltraGrid1.Rows(_Rowindex).Cells(3).Value
                    If txt15Class.Text <> "" Then
                        UltraGroupBox6.Visible = True
                    End If
                    Call Load_Gride_StockCode()
                    Call Load_Gride_DataStock()
                ElseIf _ColumIndex = 8 Then
                    If UltraGroupBox7.Visible = True Then
                        UltraGroupBox7.Visible = False
                    Else
                        txtBalance.Text = UltraGrid1.Rows(_Rowindex).Cells(6).Value
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

            vcWhere = "tmp15Class='" & str15 & "' and tmpStock_Code='" & strStock & "' and tmpUser='" & strDisname & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "YSS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If strDisname = Trim(M01.Tables(0).Rows(0)("tmpUser")) Then
                    nvcFieldList1 = "update tmpBlock_Yarn_Stock_Code set tmpQty='" & strQty & "' where tmp15Class='" & str15 & "' and tmpStock_Code='" & strStock & "' and tmpUser='" & strDisname & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            Else

                ncQryType = "ADD1"
                nvcFieldList1 = "(tmpRefNo," & "tmp15Class," & "tmpStock_Code," & "tmpQty," & "tmpUser) " & "values(" & Delivary_Ref & ",'" & str15 & "','" & strStock & "','" & strQty & "','" & strDisname & "')"
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

                            UltraGrid1.Rows(_Rowindex).Cells(8).Value = _stBalance

                            Value = _stBalance
                            '  Value = Value + (Value * (txtYarn_Wst.Text / 100))
                            _stBalance = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _stBalance = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            UltraGrid1.Rows(_Rowindex).Cells(8).Value = _stBalance

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

  
    Private Sub UltraGroupBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox2.Click

    End Sub

    Private Sub cmdKnt_Chart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKnt_Chart.Click
        frmKnitting_Plan_Board.Show()
    End Sub

    Private Sub frmDyed_Yarn_Fill_Panel_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles frmDyed_Yarn_Fill_Panel.Paint

    End Sub
End Class