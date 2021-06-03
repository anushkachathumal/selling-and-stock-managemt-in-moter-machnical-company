Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.Net.Mail
'Imports Microsoft.Office.Interop.Excel
'Imports System.Drawing
'Imports Spire.XlS
'Imports System.IO.File
'Imports System.IO.StreamWriter
Public Class frmViewDH

    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim vMax As Integer
    'Dim oFile As System.IO.File
    'Dim oWrite As System.IO.StreamWriter
    'Dim exc As New Application

    'Dim workbooks As Workbooks = exc.Workbooks
    'Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    'Dim sheets As Sheets = Workbook.Worksheets
    'Dim worksheet1 As _Worksheet = CType(Sheets.Item(1), _Worksheet)

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableDNH
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub frmViewDH_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        cmdAdd.Enabled = False
        txtFromDate.Text = Today
        txtTodate.Text = Today
        ' cmdSave.Enabled = True
        cmdEdit.Enabled = True
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Call Load_Gride()
        Sql = "select T03Name,M04Quality,M04Shade,M04Batchwt,T03Batch,T03LotType,T03SubNo,T03Reject,T03Batchtype,T03DyeH,T03MC,T03SC,T03Pro,T03Liq,T03Dye,T03WetOn,T03QC,T03Remark,T03Ongoin from T03DNH inner join M04Lot on M04Ref=T03Ecode inner join T03Machine on M04Machine_No=T03code where T03Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            Dim newRow As DataRow = c_dataCustomer1.NewRow

            'For Each DTRow1 As DataRow In M01.Tables(0).Rows
            newRow("Batch No") = M01.Tables(0).Rows(i)("T03Batch")
            newRow("Lot No") = M01.Tables(0).Rows(i)("T03LotType")
            newRow("Sub No") = M01.Tables(0).Rows(i)("T03SubNo")
            newRow("Machine") = M01.Tables(0).Rows(i)("T03Name")
            newRow("Quality") = M01.Tables(0).Rows(i)("M04Quality")
            newRow("Shade") = M01.Tables(0).Rows(i)("M04Shade")
            newRow("Dyed Quantity (Kg)") = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(i)("M04Batchwt"), "#.00")
            newRow("Reject Qty") = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(i)("T03Reject"), "#.00")
            If Trim(M01.Tables(0).Rows(i)("T03Batchtype")) = "B" Then
                newRow("Batch type") = "1st Bulk"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03Batchtype")) = "O" Then
                newRow("Batch type") = "ON Going"
            End If
            If Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "PI" Then
                newRow("Dye House Shade") = "Pilot"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "PG" Then
                newRow("Dye House Shade") = "Pigment"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "DH" Then
                newRow("Dye House Shade") = "D&H"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "UL" Then
                newRow("Dye House Shade") = "Unlevel"
            End If
           
            If Trim(M01.Tables(0).Rows(i)("T03MC")) = "Y" Then
                newRow("M/C Change") = True
            Else
                newRow("M/C Change") = False
            End If

            If Trim(M01.Tables(0).Rows(i)("T03SC")) = "Y" Then
                newRow("S/C Change") = True
            Else
                newRow("S/C Change") = False
            End If

            If Trim(M01.Tables(0).Rows(i)("T03Pro")) = "Y" Then
                newRow("Process Change") = True
            Else
                newRow("Process Change") = False
            End If

            If Trim(M01.Tables(0).Rows(i)("T03Liq")) = "Y" Then
                newRow("Liquor ratio Change") = True
            Else
                newRow("Liquor ratio Change") = False
            End If

            If Trim(M01.Tables(0).Rows(i)("T03Dye")) = "Y" Then
                newRow("Dye Lot Change") = True
            Else
                newRow("Dye Lot Change") = False
            End If

            newRow("Reason") = Trim(M01.Tables(0).Rows(i)("T03Remark"))
            newRow("Bulk Shaid Status") = Trim(M01.Tables(0).Rows(i)("T03Ongoin"))

            If Trim(M01.Tables(0).Rows(i)("T03WetOn")) = "Y" Then
                newRow("Wet on Wet comment") = "Yes"
            Else
                newRow("Wet on Wet comment") = "No"
            End If


            If Trim(M01.Tables(0).Rows(i)("T03QC")) = "QC" Then
                newRow("QC & Exam comments") = "C/F  or Yellowing issue"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "U" Then
                newRow("QC & Exam comments") = "Unlevel"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "DS" Then
                newRow("QC & Exam comments") = "Dye stricky"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "DM" Then
                newRow("QC & Exam comments") = "Dye marks"
                newRow("QC & Exam comments") = "Dye stricky"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "OF" Then
                newRow("QC & Exam comments") = "Off shade"
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "B" Then
                newRow("QC & Exam comments") = "Bursting issue"
            End If

            c_dataCustomer1.Rows.Add(newRow)
            i = i + 1
        Next
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        'exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        ' Dim range1 As Range
        Dim _NETTOTAL As Integer
        Dim T04 As DataSet
        Dim n_per As Double
        Dim Y As Integer
        Dim _cOUNT As Integer

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        'Workbooks.Application.Sheets.Add()
        'Dim sheets1 As Sheets = Workbook.Worksheets
        'Dim worksheet11 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        'worksheet11.Name = "Dye & Hold Batches "

        'worksheet11.Columns("A").ColumnWidth = 45



    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Load_Gride()
        txtFromDate.Text = Today
        txtTodate.Text = Today
        cmdEdit.Enabled = False
        cmdAdd.Enabled = True
        cmdAdd.Focus()

    End Sub
End Class