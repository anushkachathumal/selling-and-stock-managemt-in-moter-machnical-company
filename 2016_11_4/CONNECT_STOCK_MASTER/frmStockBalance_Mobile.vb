Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Infragistics.Win
Imports Infragistics.Win.UltraWinToolTip
Imports Infragistics.Win.FormattedLinkLabel
Imports Infragistics.Win.Misc
Public Class frmStockBalance_Mobile
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _LocCode As String

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M06Vehi_No from M06Vehicle_Master where M06Status='A' and M06Vehi_No='" & Trim(cboLocation.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Location = True
                _LocCode = Trim(M01.Tables(0).Rows(0)("M06Vehi_No"))
            End If


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Stock_Adjestment_1
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 250
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70


            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center



            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            ' .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit



        End With
    End Function

    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M06Vehi_No as [##],'Mobile Location' as [Loc Type] from M06Vehicle_Master where M06Status='A'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 124
                .Rows.Band.Columns(1).Width = 242


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

  
    Private Sub frmStockBalance_Mobile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCurrent.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCurrent.ReadOnly = True
        txtNew.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLast_Date.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLast_Date.ReadOnly = True
        txtLast_OB.ReadOnly = True
        txtLast_OB.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLoading.ReadOnly = True
        txtLoading.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtunloading.ReadOnly = True
        txtunloading.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGrn.ReadOnly = True
        txtGrn.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMk_Return.ReadOnly = True
        txtMk_Return.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtWastage.ReadOnly = True
        txtWastage.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtStock_IN.ReadOnly = True
        txtStock_IN.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTransfer.ReadOnly = True
        txtTransfer.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride2()

        Call Load_Location()
        Call Load_Item_Code()
        Call Load_Item_Name()

        cboLocation.ToggleDropdown()
    End Sub

    Function Search_Itemcode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim m02 As DataSet

        Try
            Sql = "select M11Part_Name  from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM' and M11Part_No='" & Trim(cboCode.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Itemcode = True
                cboItemName.Text = Trim(M01.Tables(0).Rows(0)("M11Part_Name"))
            End If
            '========================================================================


            Sql = "SELECT * FROM M06Vehicle_Master WHERE M06Vehi_No='" & Trim(cboLocation.Text) & "' AND M06Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                'current balance
                Sql = "SELECT sum(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A'  AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' group by S04Part_No"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then
                    txtCurrent.Text = Trim(m02.Tables(0).Rows(0)("S04Qty"))
                Else
                    txtCurrent.Text = "0"
                End If

                'LAST O/B
                Sql = "SELECT S04Date,S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='OB' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "'"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then
                    txtLast_Date.Text = Microsoft.VisualBasic.Day(m02.Tables(0).Rows(0)("S04Date")) & "/" & Month(m02.Tables(0).Rows(0)("S04Date")) & "/" & Year(m02.Tables(0).Rows(0)("S04Date"))
                    txtLast_OB.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))

                Else
                    txtLast_Date.Text = "-"
                    txtLast_OB.Text = "0"
                End If
                '=======================================================================
                'GRN
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='GRN' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtGrn.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtGrn.Text = "0"
                End If
                '=======================================================================
                'STOCK IN
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='STOCK IN' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtStock_IN.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtStock_IN.Text = "0"
                End If
                '=================================================================================================================================================
                'WASTAGE
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='WST' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtWastage.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtWastage.Text = "0"
                End If
                '=====================================================================================================================
                'LOADING
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='LOADING' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtLoading.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtLoading.Text = "0"
                End If
                '=====================================================================================================================
                'UNLOADING
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='UNLOADING' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtunloading.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtunloading.Text = "0"
                End If
                '=====================================================================================================================
                'TRANSFER
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='TR' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtTransfer.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtTransfer.Text = "0"
                End If
                '=====================================================================================================================
                'Market Return
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='MK_RETURN' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtMk_Return.Text = CInt(m02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtMk_Return.Text = "0"
                End If
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_ItemName()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet

        Try
            Sql = "select M11Part_No  from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM' and M11Part_Name='" & Trim(cboItemName.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                cboCode.Text = Trim(M01.Tables(0).Rows(0)("M11Part_No"))
            End If
            '========================================================================


            Sql = "SELECT * FROM M06Vehicle_Master WHERE M06Vehi_No='" & Trim(cboLocation.Text) & "' AND M06Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M02) Then
                'current balance
                Sql = "SELECT sum(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A'  AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' group by S04Part_No"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    txtCurrent.Text = Trim(M02.Tables(0).Rows(0)("S04Qty"))
                Else
                    txtCurrent.Text = "0"
                End If

                'LAST O/B
                Sql = "SELECT S04Date,S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='OB' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    txtLast_Date.Text = Microsoft.VisualBasic.Day(M02.Tables(0).Rows(0)("S04Date")) & "/" & Month(M02.Tables(0).Rows(0)("S04Date")) & "/" & Year(M02.Tables(0).Rows(0)("S04Date"))
                    txtLast_OB.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))

                Else
                    txtLast_Date.Text = "-"
                    txtLast_OB.Text = "0"
                End If
                '=======================================================================
                'GRN
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='GRN' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtGrn.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtGrn.Text = "0"
                End If
                '=======================================================================
                'STOCK IN
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='STOCK IN' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtStock_IN.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtStock_IN.Text = "0"
                End If
                '=================================================================================================================================================
                'WASTAGE
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='WST' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtWastage.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtWastage.Text = "0"
                End If
                '=====================================================================================================================
                'LOADING
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='LOADING' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtLoading.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtLoading.Text = "0"
                End If
                '=====================================================================================================================
                'UNLOADING
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='UNLOADING' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtunloading.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtunloading.Text = "0"
                End If
                '=====================================================================================================================
                'TRANSFER
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='TR' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtTransfer.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtTransfer.Text = "0"
                End If
                '=====================================================================================================================
                'Market Return
                Sql = "SELECT SUM(S04Qty) as S04Qty FROM S04Stock_Mobile WHERE S04Part_No='" & Trim(cboCode.Text) & "' AND S04Status='A' AND S04Tr_Type='MK_RETURN' AND S04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("M06Vehi_No")) & "' GROUP BY S04Part_No,S04Tr_Type,S04Loc_Code"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then

                    txtMk_Return.Text = CInt(M02.Tables(0).Rows(0)("S04Qty"))
                Else

                    txtMk_Return.Text = "0"
                End If
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_No as [##] from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 124
                '  .Rows.Band.Columns(1).Width = 160


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Item_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_Name as [##] from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItemName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 384
                '  .Rows.Band.Columns(1).Width = 160


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function clear_text()
        Me.cboCode.Text = ""
        Me.cboItemName.Text = ""
        Me.cboLocation.Text = ""
        Me.txtCurrent.Text = ""
        Me.txtNew.Text = ""
        Me.txtStock_IN.Text = ""
        Me.txtLast_Date.Text = ""
        Me.txtLast_OB.Text = ""
        Me.txtLoading.Text = ""
        Me.txtWastage.Text = ""
        Me.txtTransfer.Text = ""
        Me.txtMk_Return.Text = ""
        Me.txtGrn.Text = ""
        Me.txtunloading.Text = ""
        cboLocation.ToggleDropdown()
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call clear_text()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Dim A As String
        A = MsgBox("Are you sure you want to exit this", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Exit .......")
        If A = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.AfterCloseUp
        Call Search_Itemcode()
    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = 13 Then
            If Trim(cboCode.Text) <> "" Then
                txtNew.Focus()
            Else
                cboItemName.ToggleDropdown()
            End If

        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        ElseIf e.KeyCode = Keys.Escape Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        End If
    End Sub

    Private Sub cboItemName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItemName.AfterCloseUp
        Call Search_ItemName()
    End Sub

    Private Sub cboItemName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItemName.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemName()
            txtNew.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        ElseIf e.KeyCode = Keys.Escape Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        End If
    End Sub

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            cboCode.ToggleDropdown()
        End If
    End Sub

    Private Sub UltraButton1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraButton1.MouseHover
        Dim image As String = FormattedLinkEditor.EncodeImage(Me.ImageList1.Images(0))
        Dim TipInfo As New UltraToolTipInfo()
        TipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraButton1, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = ToolTipDisplayStyle.BalloonTip

        TipInfo.ToolTipTextFormatted = "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "font-weight:bold; " + _
     "text-smoothing-mode:AntiAlias;'> " + _
    "Techno Help</p> " + _
     "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "text-Click the button to add data to gride;'> " + _
    "<img data='" + image + "' " + _
     "align='left' " + _
     "HSpace='5'/> " + _
    "Click the button to add data to gride.</p>"
    End Sub

    Private Sub cboCode_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.MouseHover
        Dim image As String = FormattedLinkEditor.EncodeImage(Me.ImageList1.Images(0))
        Dim TipInfo As New UltraToolTipInfo()
        TipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
        Me.UltraToolTipManager1.SetUltraToolTip(Me.cboCode, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = ToolTipDisplayStyle.BalloonTip

        TipInfo.ToolTipTextFormatted = "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "font-weight:bold; " + _
     "text-smoothing-mode:AntiAlias;'> " + _
    "Techno Help</p> " + _
     "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "text-Click the button to add data to gride;'> " + _
    "<img data='" + image + "' " + _
     "align='left' " + _
     "HSpace='5'/> " + _
    "Press F1 or Esc button for Item search</p>"
    End Sub



    Private Sub cboItemName_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItemName.MouseHover
        Dim image As String = FormattedLinkEditor.EncodeImage(Me.ImageList1.Images(0))
        Dim TipInfo As New UltraToolTipInfo()
        TipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
        Me.UltraToolTipManager1.SetUltraToolTip(Me.cboItemName, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = ToolTipDisplayStyle.BalloonTip

        TipInfo.ToolTipTextFormatted = "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "font-weight:bold; " + _
     "text-smoothing-mode:AntiAlias;'> " + _
    "Techno Help</p> " + _
     "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "text-Click the button to add data to gride;'> " + _
    "<img data='" + image + "' " + _
     "align='left' " + _
     "HSpace='5'/> " + _
    "Press F1 or Esc button for Item search</p>"
    End Sub



    Private Sub txtNew_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew.KeyUp
        Dim i As Integer

        If e.KeyCode = 13 Then
            UltraButton1.Focus()
        End If
    End Sub
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Search_Location() = True Then
        Else
            MsgBox("Please select the Location", MsgBoxStyle.Information, "Information .......")
            cboLocation.ToggleDropdown()
            Exit Sub
        End If

        If UltraGrid2.Rows.Count > 0 Then
        Else
            MsgBox("Please enter the Item Details", MsgBoxStyle.Information, "Information ......")
            cboCode.ToggleDropdown()
            Exit Sub
        End If

        Call Save_Data()
    End Sub

    Function Save_Data()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim i As Integer
        Dim result1 As String
        Try




            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows

                nvcFieldList1 = "UPDATE  S04Stock_Mobile SET S04Status='CLOSE' where S04Part_No='" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "' AND S04Loc_Code='" & _LocCode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Mobile(S04Part_No,S04Loc_Code,S04Date,S04Tr_Type,S04Ref_No,S04Qty,S04Status)" & _
                                                                    " values('" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "', '" & _LocCode & "','" & Today & "','OB','-','" & Trim(UltraGrid2.Rows(i).Cells(2).Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                     " values('NEW_O/BALANCE','SAVE', '" & Now & "','" & strDisname & "','" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = i + 1
            Next
            MsgBox("Stock Adjestment creation successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.Close()
            Call Load_Gride2()
            Call clear_text()
            cboCode.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub cboCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.TextChanged
        Call Search_Itemcode()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim i As Integer

        If txtNew.Text <> "" Then
            If Search_Itemcode() = True Then
            Else
                MsgBox("Please enter the correct Item code", MsgBoxStyle.Information, "Information ........")
                cboCode.ToggleDropdown()
                Exit Sub
            End If

            If IsNumeric(txtNew.Text) Then
            Else
                MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .......")
                txtNew.Focus()
                Exit Sub
            End If
            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If (Trim(UltraGrid2.Rows(i).Cells(0).Text)) = Trim(cboCode.Text) Then
                    UltraGrid2.Rows(i).Cells(2).Value = CInt(UltraGrid2.Rows(i).Cells(2).Text) + CInt(txtNew.Text)
                    Me.txtNew.Text = ""
                    Me.cboCode.Text = ""
                    Me.cboItemName.Text = ""
                    cboCode.ToggleDropdown()
                    Exit Sub
                End If
                i = i + 1
            Next
            Dim newRow As DataRow = c_dataCustomer1.NewRow
            newRow("Item Code") = Trim(cboCode.Text)
            newRow("Item Name") = cboItemName.Text
            newRow("Qty") = txtNew.Text
            c_dataCustomer1.Rows.Add(newRow)
            Me.txtNew.Text = ""
            Me.cboCode.Text = ""
            Me.cboItemName.Text = ""
            cboCode.ToggleDropdown()
        End If
    End Sub
End Class