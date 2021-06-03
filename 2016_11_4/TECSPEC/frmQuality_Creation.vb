Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
'Imports Infragistics.Win.UltraWinGrid.RowLayoutStyle.GroupLayout
'Imports Infragistics.Win.UltraWinToolTip
'Imports Infragistics.Win.FormattedLinkLabel
'Imports Infragistics.Win.FormattedLinkLabel
'Imports Infragistics.Win.Misc
'Imports System.Diagnostics
'Imports Microsoft.Office.Interop.Excel
Imports System.Globalization
Public Class frmQuality_Creation
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer
    Dim c_dataCustomer2 As System.Data.DataTable
    Dim c_dataCustomer3 As System.Data.DataTable
    Dim c_dataCustomer4 As System.Data.DataTable

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim vcWhere As String
        Dim _Code As Integer
        Dim T01 As DataSet
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup

        UltraGrid4.DisplayLayout.Bands(0).Groups.Clear()
        UltraGrid4.DisplayLayout.Bands(0).Columns.Dispose()

        'If UltraGrid3.DisplayLayout.Bands(0).GroupHeadersVisible = True Then
        'Else
        '  agroup1.Key = ""
        '  agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove("GroupH")
        agroup1 = UltraGrid4.DisplayLayout.Bands(0).Groups.Add("GroupH")


        _Code = _Code - 1
        agroup1.Header.Caption = "Special"

        agroup1.Width = 110
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        Dim colWork As New DataColumn("##", GetType(Boolean))
        dt.Columns.Add(colWork)
        'colWork.ReadOnly = True
        'colWork.ReadOnly = True
        colWork = New DataColumn("Code", GetType(String))
        colWork.MaxLength = 250
        dt.Columns.Add(colWork)
        ' colWork.ReadOnly = True
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dt.Columns.Add(colWork)
        '  colWork.ReadOnly = True '

        For I = 1 To 16
            dt.Rows.Add(False, "")

        Next



        Me.UltraGrid4.SetDataBinding(dt, Nothing)
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(0).Group = agroup1
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(1).Group = agroup1
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(2).Group = agroup1
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(0).Width = 50
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(1).Width = 70
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns(2).Width = 110


        agroup2 = UltraGrid4.DisplayLayout.Bands(0).Groups.Add("Group1")

        agroup2.Header.Caption = "Structure"
        agroup2.Width = 220
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns.Add("G1", "##")
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G1").Group = agroup2
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G1").Width = 50
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        Me.UltraGrid4.DisplayLayout.Bands(0).Columns.Add("G2", "Code")
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G2").Group = agroup2
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G2").Width = 70
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        Me.UltraGrid4.DisplayLayout.Bands(0).Columns.Add("G3", "Description")
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G3").Group = agroup2
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G3").Width = 110


        agroup3 = UltraGrid4.DisplayLayout.Bands(0).Groups.Add("Group3")

        agroup3.Header.Caption = "Fiber Composition"
        agroup3.Width = 220
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns.Add("G4", "##")
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G4").Group = agroup3
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G4").Width = 50
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G4").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        Me.UltraGrid4.DisplayLayout.Bands(0).Columns.Add("G5", "Code")
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G5").Group = agroup3
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G5").Width = 70
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G5").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        Me.UltraGrid4.DisplayLayout.Bands(0).Columns.Add("G6", "Description")
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G6").Group = agroup3
        Me.UltraGrid4.DisplayLayout.Bands(0).Columns("G6").Width = 110


        Sql = "select * from M54Special where M54Status='1'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        I = 0
        For Each DTRow3 As DataRow In M01.Tables(0).Rows
            UltraGrid4.Rows(I).Cells(1).Value = M01.Tables(0).Rows(I)("M54Code")
            UltraGrid4.Rows(I).Cells(2).Value = M01.Tables(0).Rows(I)("M54Des")
            I = I + 1
        Next

        Sql = "select * from M54Special where M54Status='2'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        I = 0
        For Each DTRow3 As DataRow In M01.Tables(0).Rows
            UltraGrid4.Rows(I).Cells(3).Style = ColumnStyle.CheckBox
            UltraGrid4.Rows(I).Cells(3).Value = False
            UltraGrid4.Rows(I).Cells(4).Value = M01.Tables(0).Rows(I)("M54Code")
            UltraGrid4.Rows(I).Cells(5).Value = M01.Tables(0).Rows(I)("M54Des")
            I = I + 1
        Next


        Sql = "select * from M54Special where M54Status='3'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        I = 0
        For Each DTRow3 As DataRow In M01.Tables(0).Rows
            UltraGrid4.Rows(I).Cells(6).Style = ColumnStyle.CheckBox
            UltraGrid4.Rows(I).Cells(6).Value = False
            UltraGrid4.Rows(I).Cells(7).Value = M01.Tables(0).Rows(I)("M54Code")
            UltraGrid4.Rows(I).Cells(8).Value = M01.Tables(0).Rows(I)("M54Des")
            I = I + 1
        Next
        lbl1.Text = ""
    End Function

    Private Sub frmQuality_Creation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Private Sub UltraGrid4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid4.AfterCellUpdate
       
    End Sub

    Private Sub UltraGrid4_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid4.CellChange
        On Error Resume Next
        Dim I As Integer
        Dim characterToRemove As String
        I = UltraGrid4.ActiveRow.Index
        If UltraGrid4.Rows(I).Cells(0).Text = True Then
            If UltraGrid4.Rows(I).Cells(1).Text <> "" Then
                lbl1.Text = lbl1.Text & Trim(UltraGrid4.Rows(I).Cells(1).Text)
            End If
        Else
            characterToRemove = Trim(UltraGrid4.Rows(I).Cells(1).Text)

            'MsgBox(Trim(fields(9)))
            lbl1.Text = (Replace(lbl1.Text, characterToRemove, ""))

        End If

        If UltraGrid4.Rows(I).Cells(3).Text = True Then
            If UltraGrid4.Rows(I).Cells(4).Text <> "" Then
                lbl1.Text = lbl1.Text & Trim(UltraGrid4.Rows(I).Cells(4).Text)
            End If
        Else
            characterToRemove = Trim(UltraGrid4.Rows(I).Cells(4).Text)

            'MsgBox(Trim(fields(9)))
            lbl1.Text = (Replace(lbl1.Text, characterToRemove, ""))

        End If


        If UltraGrid4.Rows(I).Cells(6).Text = True Then
            If UltraGrid4.Rows(I).Cells(7).Text <> "" Then
                lbl1.Text = lbl1.Text & Trim(UltraGrid4.Rows(I).Cells(7).Text)
            End If
        Else
            characterToRemove = Trim(UltraGrid4.Rows(I).Cells(7).Text)

            'MsgBox(Trim(fields(9)))
            lbl1.Text = (Replace(lbl1.Text, characterToRemove, ""))

        End If
    End Sub

    Private Sub UltraGrid4_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid4.InitializeLayout

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Me.Close()
    End Sub
End Class