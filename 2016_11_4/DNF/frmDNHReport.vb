Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.Quarrys
Imports System.IO.File
Imports System.IO.StreamWriter
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS
Public Class frmDNHReport
    Dim strLine As String
    Dim strLineflu As String
    Dim strDash As String
    Dim StrDisCode As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    Dim exc As New Application

    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet1 As _Worksheet = CType(Sheets.Item(1), _Worksheet)

    Private Sub frmDNHReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtFromDate.Text = Today
        txtTodate.Text = Today
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        OPR1.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False

        txtFromDate.Text = Today
        txtTodate.Text = Today

        cmdSave.Enabled = True
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
        exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer
        Dim T04 As DataSet
        Dim n_per As Double
        Dim Y As Integer
        Dim _cOUNT As Integer
        Dim X As Integer
        Dim M01 As DataSet

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet1.Name = "Daily Refinish Batches"

        worksheet1.Columns("A").ColumnWidth = 10
        worksheet1.Columns("B").ColumnWidth = 10
        worksheet1.Columns("c").ColumnWidth = 10
        worksheet1.Columns("d").ColumnWidth = 10
        worksheet1.Columns("e").ColumnWidth = 10
        worksheet1.Columns("f").ColumnWidth = 10
        worksheet1.Columns("g").ColumnWidth = 16
        worksheet1.Columns("s").ColumnWidth = 28
        worksheet1.Columns("t").ColumnWidth = 20

        worksheet1.Cells(4, 1) = "Batch No"
        worksheet1.Cells(4, 1).WrapText = True
        worksheet1.Rows(4).Font.Bold = True
        worksheet1.Rows(4).Font.size = 9

        worksheet1.Rows(5).Font.Bold = True
        worksheet1.Rows(5).Font.size = 8

        worksheet1.Cells(4, 2) = "Lot Type"
        worksheet1.Cells(4, 2).WrapText = True
        worksheet1.Cells(4, 3) = "Sub No"
        worksheet1.Cells(4, 3).WrapText = True
        worksheet1.Cells(4, 4) = "Machine"
        worksheet1.Cells(4, 4).WrapText = True
        worksheet1.Cells(4, 5) = "Quality"
        worksheet1.Cells(4, 5).WrapText = True
        worksheet1.Cells(4, 6) = "Shade"
        worksheet1.Cells(4, 6).WrapText = True
        worksheet1.Cells(4, 7) = "Dyed Quantity (Kg)"
        worksheet1.Cells(4, 7).WrapText = True
        worksheet1.Cells(4, 8) = "Batch type"
        worksheet1.Cells(4, 8).WrapText = True
        worksheet1.Range(worksheet1.Cells(4, 8), worksheet1.Cells(4, 9)).Merge()
        worksheet1.Range(worksheet1.Cells(4, 8), worksheet1.Cells(4, 9)).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(4, 8)

        worksheet1.Cells(4, 10) = "Dye house  Shade comments"
        worksheet1.Cells(4, 10).WrapText = True
        worksheet1.Range(worksheet1.Cells(4, 10), worksheet1.Cells(4, 13)).Merge()
        worksheet1.Range(worksheet1.Cells(4, 10), worksheet1.Cells(4, 13)).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(4, 10)

        worksheet1.Cells(4, 14) = "Recipe Detailes"
        worksheet1.Cells(4, 14).WrapText = True
        worksheet1.Range(worksheet1.Cells(4, 14), worksheet1.Cells(4, 18)).Merge()
        worksheet1.Range(worksheet1.Cells(4, 14), worksheet1.Cells(4, 18)).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(4, 14)
        worksheet1.Cells(4, 19) = "Reason for Off shade"
        worksheet1.Cells(4, 20) = "On-going bulk shade status"

        worksheet1.Cells(4, 21) = "Wet on Wet comment"
        worksheet1.Cells(4, 21).WrapText = True
        worksheet1.Range(worksheet1.Cells(4, 21), worksheet1.Cells(4, 22)).Merge()
        worksheet1.Range(worksheet1.Cells(4, 21), worksheet1.Cells(4, 22)).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(4, 22)

        worksheet1.Cells(4, 23) = "QC & Exam comments"
        worksheet1.Cells(4, 23).WrapText = True
        worksheet1.Range(worksheet1.Cells(4, 23), worksheet1.Cells(4, 29)).Merge()
        worksheet1.Range(worksheet1.Cells(4, 23), worksheet1.Cells(4, 29)).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(4, 23)

        worksheet1.Range("A4:A4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A4:A4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A5:A5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A5:A5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("b4:b4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b4:b4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b5:b5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b5:b5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("c4:c4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c4:c4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c5:c5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c5:c5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("d4:d4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d4:d4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5:d5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5:d5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("e4:e4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e4:e4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e5:e5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e5:e5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("f4:f4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f4:f4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f5:f5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f5:f5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("c4:g4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g4:g4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g5:g5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g5:g5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(4, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("h4:i4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i4:i4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h4:i4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i5:i5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h5:h5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h5:i5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Cells(5, 8) = "1stBulk"
        worksheet1.Cells(5, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 9) = "On Going"
        worksheet1.Cells(5, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("j4:m4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m4:m4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j4:m4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m5:m5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j5:j5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k5:k5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l5:l5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m5:m5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j5:m5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Cells(5, 10) = "Pilot"
        worksheet1.Cells(5, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 11) = "Pigment"
        worksheet1.Cells(5, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 12) = "D&H"
        worksheet1.Cells(5, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 13) = "Unlevel"
        worksheet1.Cells(5, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
    
        worksheet1.Range("n4:r4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r4:r4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r5:r5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n4:r4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n5:r5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n5:n5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o5:o5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p5:p5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q5:q5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        'worksheet1.Range("r5:r5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Cells(5, 14) = "M/C Change"
        worksheet1.Cells(5, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 15) = "S/C Change"
        worksheet1.Cells(5, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 16) = "Pro Change "
        worksheet1.Cells(5, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 17) = "Liquor Ratio"
        worksheet1.Cells(5, 18) = "Dye Lot "
        worksheet1.Cells(5, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter

        'Const CheckMark As Char = ChrW(&H2713)
        Const HeavyCheckMark As Char = ChrW(&H2714)
        'worksheet1.Cells(6, 1) = (HeavyCheckMark)
        'worksheet1.OLEObjects.Add(ClassType:="Forms.CheckBox.1", Link:=False, _
        'DisplayAsIcon:=False, Left:=65.25, Top:=24, Width:=108, Height:=21). _
        'Select()
        worksheet1.Range("S4:S4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("S4:S5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s5:s5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("t4:t4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t4:t5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t5:t5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("u4:v4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v4:v5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u4:v4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u5:u5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u5:u5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v5:v5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous


        worksheet1.Cells(5, 21) = "YES"
        worksheet1.Cells(5, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 22) = "NO"
        worksheet1.Cells(5, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("w4:ac4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("ac4:ac4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("ac5:ac5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w5:ac5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w5:w5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w5:w5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x5:x5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x5:x5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y5:y5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y5:y5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("z5:z5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z5:z5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("aa5:aa5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("aa5:aa5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("ab5:ab5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("ab5:ab5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("ac5:ac5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Cells(5, 23) = "Off Shade "
        worksheet1.Cells(5, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 24) = "Unlevel "
        worksheet1.Cells(5, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 25) = "Dye Strickly "
        worksheet1.Cells(5, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 26) = "Dye Marks "
        worksheet1.Cells(5, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 27) = "C/F or Y/I "
        worksheet1.Cells(5, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 28) = "Bursting"
        worksheet1.Cells(5, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(5, 29) = "R2R Change"
        worksheet1.Cells(5, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("A4:A4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("A5:A5").Interior.Color = RGB(255, 192, 0)

        worksheet1.Range("b4:b4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("b5:b5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("c4:c4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("c5:c5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("d4:d4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("d5:d5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("e4:e4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("e5:e5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("f4:f4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("f5:f5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("g4:g4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("g5:g5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("h4:h4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("h5:h5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("i5:i5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("j4:j4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("j5:j5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("k5:k5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("l5:l5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("m5:m5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("n4:n4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("n5:n5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("o5:o5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("p5:p5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("q5:q5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("r5:r5").Interior.Color = RGB(255, 192, 0)

        worksheet1.Range("s4:s4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("s5:s5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("t4:t4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("t5:t5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("u4:u4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("u5:u5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("v5:v5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("w4:w4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("w5:w5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("x5:x5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("y5:y5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("z5:z5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("aa5:aa5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("ab5:ab5").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("ac5:ac5").Interior.Color = RGB(255, 192, 0)


        'worksheet1.Cells(5, 24).Font.Name = "Wingdings 2"
        'worksheet1.Cells(5, 24).Font.Size = 10
        'worksheet1.Cells(5, 24) = "=CHAR(80)"
        ''worksheet1.Cells(5, 24).FormulaR1C1 = "P"

        X = 1
        Y = 6
        SQL = "select T03Name,M04Quality,M04Shade,M04Batchwt,T03Batch,T03LotType,T03SubNo,T03Reject,T03Batchtype,T03DyeH,T03MC,T03SC,T03Pro,T03Liq,T03Dye,T03WetOn,T03QC,T03Remark,T03Ongoin from T03DNH inner join M04Lot on M04Ref=T03Ecode inner join T03Machine on T03code=M04Machine_No where T03Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            X = 1
            worksheet1.Rows(Y).Font.size = 9
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("T03Batch")
            worksheet1.Cells(5, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = X + 1
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("T03LotType")
            worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = X + 1
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("T03SubNo")
            worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = X + 1
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("T03Name")
            worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = X + 1
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("M04Quality")
            worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = X + 1
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("M04Shade")
            worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = X + 1
            worksheet1.Cells(Y, X) = Microsoft.VisualBasic.Format(M01.Tables(0).Rows(i)("M04Batchwt"), "#.00")
            worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(Y, X)
            range1.NumberFormat = "0"
            X = X + 1
            If Trim(M01.Tables(0).Rows(i)("T03Batchtype")) = "B" Then
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 2
            Else
                X = X + 1
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            End If
            If Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "PI" Then
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 4
            ElseIf Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "PG" Then
                X = X + 1
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 3
            ElseIf Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "DH" Then
                X = X + 2
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 2
            ElseIf Trim(M01.Tables(0).Rows(i)("T03DyeH")) = "UL" Then
                X = X + 3
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            End If

            'If Trim(M01.Tables(0).Rows(i)("T03Batchtype")) = "B" Then
            '    worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
            '    worksheet1.Cells(Y, X).Font.Size = 10
            '    worksheet1.Cells(Y, X) = "=CHAR(80)"
            '    worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '    X = X + 2
            'Else
            '    X = X + 1
            '    worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
            '    worksheet1.Cells(Y, X).Font.Size = 10
            '    worksheet1.Cells(Y, X) = "=CHAR(80)"
            '    worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '    X = X + 1
            'End If
            If Trim(M01.Tables(0).Rows(i)("T03MC")) = "Y" Then
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            Else
                X = X + 1
            End If
            If Trim(M01.Tables(0).Rows(i)("T03SC")) = "Y" Then

                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            Else
                X = X + 1
            End If

            If Trim(M01.Tables(0).Rows(i)("T03Pro")) = "Y" Then

                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            Else
                X = X + 1
            End If
            If Trim(M01.Tables(0).Rows(i)("T03Liq")) = "Y" Then

                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            Else
                X = X + 1
            End If

            If Trim(M01.Tables(0).Rows(i)("T03Dye")) = "Y" Then

                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            Else
                X = X + 1
            End If
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("T03Remark")
            X = X + 1
            worksheet1.Cells(Y, X) = M01.Tables(0).Rows(i)("T03Ongoin")
            X = X + 1
            If Trim(M01.Tables(0).Rows(i)("T03WetOn")) = "Y" Then

                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 2
            Else
                X = X + 1
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
                X = X + 1
            End If

            If Trim(M01.Tables(0).Rows(i)("T03QC")) = "OF" Then

                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter

            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "U" Then
                X = X + 1
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "DS" Then
                X = X + 2
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "DM" Then
                X = X + 3
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "CF" Then
                X = X + 4
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "B" Then
                X = X + 5
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf Trim(M01.Tables(0).Rows(i)("T03QC")) = "R" Then
                X = X + 6
                worksheet1.Cells(Y, X).Font.Name = "Wingdings 2"
                worksheet1.Cells(Y, X).Font.Size = 10
                worksheet1.Cells(Y, X) = "=CHAR(80)"
                worksheet1.Cells(Y, X).HorizontalAlignment = XlHAlign.xlHAlignCenter
            End If

            worksheet1.Range("A" & Y & ":a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y & ":a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y & ":b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y & ":b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y & ":c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y & ":c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y & ":d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y & ":d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y & ":e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y & ":e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y & ":f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y & ":f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y & ":g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y & ":g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y & ":h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y & ":h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y & ":i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y & ":i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y & ":j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y & ":j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y & ":k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y & ":k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y & ":l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y & ":l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y & ":m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y & ":m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y & ":n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y & ":n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y & ":o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y & ":o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y & ":p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y & ":p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y & ":q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y & ":q" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y & ":r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y & ":r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s" & Y & ":s" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s" & Y & ":s" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t" & Y & ":t" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t" & Y & ":t" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u" & Y & ":u" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u" & Y & ":u" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v" & Y & ":v" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v" & Y & ":v" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w" & Y & ":w" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w" & Y & ":w" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x" & Y & ":x" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x" & Y & ":x" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y" & Y & ":y" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y" & Y & ":y" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("z" & Y & ":z" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("z" & Y & ":z" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("Aa" & Y & ":aa" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("Aa" & Y & ":aa" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("Ab" & Y & ":ab" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("Ab" & Y & ":ab" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("Ac" & Y & ":ac" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("Ac" & Y & ":ac" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            Y = Y + 1
            i = i + 1
        Next
    End Sub
End Class