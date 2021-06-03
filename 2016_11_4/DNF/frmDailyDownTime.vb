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
Public Class frmDailyDownTime
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
    Dim worksheet As _Worksheet = CType(sheets.Item(1), _Worksheet)
    ' Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
    'Dim cdo2Massege As CDO.Message
    'Dim Cdo2Configuration As CDO.Configuration
    'Dim Cdo2Fields



    Function Print_Daily_DownTime()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet

        Dim FileName As String
        exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer

        '  Dim worksheet1 As _Worksheet = CType(sheets.Item(2), _Worksheet)
        'workbooks.Application.Sheets.Add()
        worksheet.Name = "Daily Down Time %"
        worksheet.Cells(2, 3) = "Textured Jersey PLC"
        worksheet.Rows(2).Font.Bold = True
        worksheet.Rows(2).Font.size = 26

        worksheet.Range("A2:J2").MergeCells = True
        worksheet.Range("A2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


        worksheet.Cells(4, 1) = "Daily Down Time Report on " & txtFromDate.Text
        worksheet.Rows(4).Font.Bold = True
        worksheet.Rows(4).Font.size = 10

        worksheet.Rows(6).rowheight = 20.25

        worksheet.Rows(6).Font.Bold = True
        worksheet.Rows(6).Font.size = 10
        ' worksheet.Rows(7).Font.Bold = True
        ' worksheet.Rows(7).Font.size = 10

        ' worksheet.Columns(1).Format = Text

        'UPGRADE_WARNING: Couldn't resolve default property of object worksheet.Cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'range1 = worksheet.Cells(6, 1)
        'range1.Borders.LineStyle = XlLineStyle.xlDouble
        'range1 = worksheet.Cells(6, 2)
        'range1.Borders.LineStyle = XlLineStyle.xlDouble
        worksheet.Cells(6, 1) = "Machine Name"
        range1 = worksheet.Cells(6, 1)
        range1.Borders.LineStyle = XlLineStyle.xlContinuous
        worksheet.Columns(1).columnwidth = 17
        range1.Interior.Color = RGB(192, 192, 255)
        Dim X1 As Integer
        Dim n_Date As Date
        Dim n_Date1 As Date

        n_Date = txtFromDate.Text & " " & "7:30AM"
        If txtFromDate.Text = txtTodate.Text Then
            n_Date1 = txtFromDate.Text & " " & "7:30PM"
        Else
            n_Date1 = txtTodate.Text & " " & "7:30AM"
        End If
        _NETTOTAL = 0
        X1 = 1
        SQL = "select T01Down_Time from T01Down_Time where T01Date between '" & n_Date & "' and '" & n_Date1 & "' group by T01Down_Time"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For i = 0 To T01.Tables(0).Rows.Count - 1

            'If T01.Tables(0).Rows.Count + 1 > i Then

            worksheet.Cells(6, X1 + 1) = T01.Tables(0).Rows(i)("T01Down_Time")
            range1 = worksheet.Cells(6, X1 + 1)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            range1.Interior.Color = RGB(192, 192, 255)
            ' range1.XlCellType.xlCellTypeAllFormatConditions
            X1 = X1 + 1
            '  range1 = worksheet.Cells(6, 2)
            ' range1.Borders.LineStyle = XlLineStyle.xlDouble
            '  i = i + 1
        Next
        worksheet.Cells(6, X1 + 1) = "Grand Total"
        worksheet.Columns(X1 + 1).columnwidth = 12
        range1 = worksheet.Cells(6, X1 + 1)
        range1.Borders.LineStyle = XlLineStyle.xlContinuous
        range1.Interior.Color = RGB(192, 192, 255)
        '-----------------------------------------------------------------------------

        Dim X As Integer
        SQL = "select max(T03Name) as T03Name,T03Code from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01Date between '" & n_Date & "' and '" & n_Date1 & "' and M01Code in ('01','02') group by  T03Code"
        dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow2 As DataRow In dsUser.Tables(0).Rows
            worksheet.Cells(X + 7, 1) = dsUser.Tables(0).Rows(X)("T03Name")
            SQL = "select T01Down_Time from T01Down_Time where T01Date between '" & n_Date & "' and '" & n_Date1 & "' group by T01Down_Time"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            _GrandTotal = 0
            For Each DTRow1 As DataRow In T01.Tables(0).Rows
                Dim _STTime As String
                SQL = "select T03Name,sum(T01Taken) as T01Taken from  T03Machine  inner join T01Down_Time on T01Machine=T03Code where T01Date between '" & n_Date & "' and '" & n_Date1 & "' and T01Down_Time='" & Trim(T01.Tables(0).Rows(i)("T01Down_Time")) & "' and T03Code='" & Trim(dsUser.Tables(0).Rows(X)("T03Code")) & "' group by T03Name"
                T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T03) Then
                    _STTime = (Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60) Mod 60, "0#") & _
                               ":" & Format(T03.Tables(0).Rows(0)("T01Taken") Mod 60, "0#"))

                    'If Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60) Mod 60, "0#") >= 24 Then
                    '    _STTime = "1/1/1900" & " " & "12:19 AM"
                    'End If
                    _GrandTotal = _GrandTotal + Val(T03.Tables(0).Rows(0)("T01Taken"))
                    worksheet.Cells(X + 7, i + 2) = _STTime

                Else

                End If
                i = i + 1
            Next
            _STGrand = (Format(Int(_GrandTotal / 60) Mod 60, "0#") & _
                            ":" & Format(_GrandTotal Mod 60, "0#"))

            ' MsgBox(Format(Int(_GrandTotal / 60), "0#"))
            _STGrand = (Format(Int(_GrandTotal / 60), "0#")) & ":" & Format(_GrandTotal - (Int(_GrandTotal / 60) * 60), "0#")
            worksheet.Cells(X + 7, i + 2) = _STGrand
            ' oSheet.Cells.NumberFormat = "@"; // @ means saving as "text" format
            worksheet.Cells.NumberFormat = "[hh]:mm"
            X = X + 1
        Next
        worksheet.Cells(X + 7, 1) = "Grand Total"
        range1 = worksheet.Cells(X + 7, 1)
        range1.Borders.LineStyle = XlLineStyle.xlContinuous
        range1.Interior.Color = RGB(192, 192, 255)
        worksheet.Rows(X + 7).Font.Bold = True
        worksheet.Rows(X + 7).Font.size = 10

        i = 0

        SQL = "select T01Down_Time from T01Down_Time where T01Date between '" & n_Date & "' and '" & n_Date1 & "' group by T01Down_Time"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow1 As DataRow In T01.Tables(0).Rows
            ' SQL = "select sum(T01Taken) as T01Taken from T01Down_Time where T01Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and T01Down_Time='" & T01.Tables(0).Rows(i)("T01Down_Time") & "' "

            SQL = "select sum(T01Taken) as T01Taken from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01Date between '" & n_Date & "' and '" & n_Date1 & "' and M01Code in ('01','02') and T01Down_Time='" & T01.Tables(0).Rows(i)("T01Down_Time") & "' group by  T01Down_Time"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T03) Then
                _STGrand = (Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60), "0#")) & ":" & Format(T03.Tables(0).Rows(0)("T01Taken") - (Int(T03.Tables(0).Rows(0)("T01Taken") / 60) * 60), "0#")
                worksheet.Cells(X + 7, i + 2) = _STGrand
                ' oSheet.Cells.NumberFormat = "@"; // @ means saving as "text" format
                worksheet.Cells.NumberFormat = "[hh]:mm"
                range1 = worksheet.Cells(X + 7, i + 2)
                range1.Borders.LineStyle = XlLineStyle.xlContinuous
                range1.Interior.Color = RGB(192, 192, 255)
            Else
                range1 = worksheet.Cells(X + 7, i + 2)
                range1.Borders.LineStyle = XlLineStyle.xlContinuous
                range1.Interior.Color = RGB(192, 192, 255)

            End If
            i = i + 1
        Next

        '-----------------------SUM OF GRAND TOTAL
        'SQL = "select sum(T01Taken) as T01Taken from T01Down_Time where T01Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' GROUP BY T01Date"
        SQL = "select sum(T01Taken) as T01Taken from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01Date between '" & n_Date & "' and '" & n_Date1 & "' and M01Code in ('01','02')  group by  T03Code"
        T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T03) Then
            _NETTOTAL = T03.Tables(0).Rows(0)("T01Taken")
            _STGrand = (Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60), "0#")) & ":" & Format(T03.Tables(0).Rows(0)("T01Taken") - (Int(T03.Tables(0).Rows(0)("T01Taken") / 60) * 60), "0#")
            worksheet.Cells(X + 7, i + 2) = _STGrand
            ' oSheet.Cells.NumberFormat = "@"; // @ means saving as "text" format
            worksheet.Cells.NumberFormat = "[hh]:mm"

            range1 = worksheet.Cells(X + 7, i + 2)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            range1.Interior.Color = RGB(192, 192, 255)

        End If
        '-----------------------------------------------------------------------
        'AVARAGE 

        worksheet.Cells(X + 10, 1) = "Total %"
        range1 = worksheet.Cells(X + 10, 1)
        worksheet.Rows(X + 10).Font.Bold = True
        worksheet.Rows(X + 10).Font.size = 10
        range1.Font.Color = RGB(0, 0, 205)
        i = 0
        SQL = "select T01Down_Time from T01Down_Time where T01Date between '" & n_Date & "' and '" & n_Date1 & "' group by T01Down_Time"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow1 As DataRow In T01.Tables(0).Rows
            SQL = "select sum(T01Taken) as T01Taken from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01Date between '" & n_Date & "' and '" & n_Date1 & "' and M01Code in ('01','02') and T01Down_Time='" & T01.Tables(0).Rows(i)("T01Down_Time") & "' group by  T01Down_Time"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T03) Then
                Dim Y As Double

                Y = Val(T03.Tables(0).Rows(0)("T01Taken")) / _NETTOTAL

                worksheet.Cells(X + 10, i + 2) = Microsoft.VisualBasic.Format(Y * 100, "#.0")
                ' worksheet.Range("B" & X + 10).NumberFormat = "@"
                range1 = worksheet.Cells(X + 10, i + 2)
                worksheet.Rows(X + 10).Font.Bold = True
                worksheet.Rows(X + 10).Font.size = 10
                range1.Font.Color = RGB(0, 0, 205)
            End If
            i = i + 1
        Next
        '---------------------------------------------------------------------------------



    End Function

    Function Weekly_DownTime()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet

        Dim FileName As String
        exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer

        '  Dim worksheet1 As _Worksheet = CType(sheets.Item(2), _Worksheet)
        workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        ' worksheet1 = workbooks.Application.Sheets(2)
        worksheet1.Name = "Weekly Down Time Summery %"
        worksheet1.Cells(2, 3) = "Textured Jersey PLC"
        worksheet1.Rows(2).Font.Bold = True
        worksheet1.Rows(2).Font.size = 26

        worksheet1.Range("A2:J2").MergeCells = True
        worksheet1.Range("A2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


        worksheet1.Cells(4, 1) = "Weekly Down Time Report on Week :" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text))
        worksheet1.Rows(4).Font.Bold = True
        worksheet1.Rows(4).Font.size = 10

        worksheet1.Rows(6).rowheight = 20.25

        worksheet1.Rows(6).Font.Bold = True
        worksheet1.Rows(6).Font.size = 10

        worksheet1.Cells(6, 1) = "Machine Type"
        range1 = worksheet1.Cells(6, 1)
        range1.Borders.LineStyle = XlLineStyle.xlContinuous
        worksheet1.Columns(1).columnwidth = 17
        range1.Interior.Color = RGB(192, 192, 255)
        Dim X1 As Integer
        _NETTOTAL = 0
        X1 = 1
        SQL = "select T01Down_Time from T01Down_Time where T01WeekNo=" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text)) & " and year(T01date)='" & Year(txtFromDate.Text) & "' group by T01Down_Time"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For i = 0 To T01.Tables(0).Rows.Count - 1

            'If T01.Tables(0).Rows.Count + 1 > i Then

            worksheet1.Cells(6, X1 + 1) = T01.Tables(0).Rows(i)("T01Down_Time")
            range1 = worksheet1.Cells(6, X1 + 1)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            range1.Interior.Color = RGB(192, 192, 255)
            ' range1.XlCellType.xlCellTypeAllFormatConditions

            worksheet1.Columns(X1 + 1).columnwidth = 12
            range1 = worksheet1.Cells(6, X1 + 1)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            range1.Interior.Color = RGB(192, 192, 255)

            X1 = X1 + 1
            '  range1 = worksheet.Cells(6, 2)
            ' range1.Borders.LineStyle = XlLineStyle.xlDouble
            '  i = i + 1
        Next

        Dim X As Integer
        X = 0
        SQL = "select max(M01Description) as T03Name,M01Code from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01WeekNo=" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text)) & " and year(T01date)='" & Year(txtFromDate.Text) & "' and M01Code in ('01','02') group by  M01Code"
        dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow2 As DataRow In dsUser.Tables(0).Rows
            worksheet1.Cells(X + 7, 1) = dsUser.Tables(0).Rows(X)("T03Name")
            SQL = "select T01Down_Time from T01Down_Time where T01weekNo=" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text)) & " and year(T01date)='" & Year(txtFromDate.Text) & "' group by T01Down_Time"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            _GrandTotal = 0
            For Each DTRow1 As DataRow In T01.Tables(0).Rows
                Dim _STTime As String
                SQL = "select sum(T01Taken) as T01Taken from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01WeekNo=" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text)) & " and year(T01date)='" & Year(txtFromDate.Text) & "' and T01Down_Time='" & Trim(T01.Tables(0).Rows(i)("T01Down_Time")) & "' and M01Code='" & Trim(dsUser.Tables(0).Rows(X)("M01Code")) & "' group by M01Code "
                T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T03) Then
                    'MsgBox(CInt(T03.Tables(0).Rows(0)("T01Taken")))
                    ' _STTime = (Format(CInt(T03.Tables(0).Rows(0)("T01Taken") / 60) Mod 60, "0#") & _
                    '            ":" & Format(T03.Tables(0).Rows(0)("T01Taken") Mod 60, "0#"))
                    _STTime = (Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60), "0#")) & ":" & Format(T03.Tables(0).Rows(0)("T01Taken") - (Int(T03.Tables(0).Rows(0)("T01Taken") / 60) * 60), "0#")
                    'If Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60) Mod 60, "0#") >= 24 Then
                    '    _STTime = "1/1/1900" & " " & "12:19 AM"
                    'End If
                    _GrandTotal = _GrandTotal + Val(T03.Tables(0).Rows(0)("T01Taken"))
                    worksheet1.Cells(X + 7, i + 2) = _STTime

                Else

                End If
                i = i + 1
            Next
            '_STGrand = (Format(Int(_GrandTotal / 60) Mod 60, "0#") & _
            '                ":" & Format(_GrandTotal Mod 60, "0#"))

            '' MsgBox(Format(Int(_GrandTotal / 60), "0#"))
            '_STGrand = (Format(Int(_GrandTotal / 60), "0#")) & ":" & Format(_GrandTotal - (Int(_GrandTotal / 60) * 60), "0#")
            'worksheet1.Cells(X + 7, i + 2) = _STGrand
            '' oSheet.Cells.NumberFormat = "@"; // @ means saving as "text" format
            worksheet1.Cells.NumberFormat = "[hh]:mm"
            X = X + 1
        Next

        worksheet1.Cells(X + 7, 1) = "Grand Total"
        range1 = worksheet1.Cells(X + 7, 1)
        range1.Borders.LineStyle = XlLineStyle.xlContinuous
        range1.Interior.Color = RGB(192, 192, 255)
        worksheet1.Rows(X + 7).Font.Bold = True
        worksheet1.Rows(X + 7).Font.size = 10

        i = 0

        SQL = "select T01Down_Time from T01Down_Time where T01WeekNo=" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text)) & " and year(T01Date)='" & Year(txtFromDate.Text) & "' group by T01Down_Time"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow1 As DataRow In T01.Tables(0).Rows
            ' SQL = "select sum(T01Taken) as T01Taken from T01Down_Time where T01Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and T01Down_Time='" & T01.Tables(0).Rows(i)("T01Down_Time") & "' "

            SQL = "select sum(T01Taken) as T01Taken from  T03Machine  inner join T01Down_Time on T01Machine=T03Code inner join M01Dyeing_MC_Type on M01Code=T03Type where T01WeekNo=" & DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text)) & " and year(T01date)='" & Year(txtFromDate.Text) & "' and M01Code in ('01','02','03') and T01Down_Time='" & T01.Tables(0).Rows(i)("T01Down_Time") & "' group by  T01Down_Time"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T03) Then
                _STGrand = (Format(Int(T03.Tables(0).Rows(0)("T01Taken") / 60), "0#")) & ":" & Format(T03.Tables(0).Rows(0)("T01Taken") - (Int(T03.Tables(0).Rows(0)("T01Taken") / 60) * 60), "0#")
                worksheet1.Cells(X + 7, i + 2) = _STGrand
                ' oSheet.Cells.NumberFormat = "@"; // @ means saving as "text" format
                worksheet1.Cells.NumberFormat = "[hh]:mm"
                range1 = worksheet1.Cells(X + 7, i + 2)
                range1.Borders.LineStyle = XlLineStyle.xlContinuous
                range1.Interior.Color = RGB(192, 192, 255)
            Else
                range1 = worksheet1.Cells(X + 7, i + 2)
                range1.Borders.LineStyle = XlLineStyle.xlContinuous
                range1.Interior.Color = RGB(192, 192, 255)

            End If
            i = i + 1
        Next


    End Function
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        '  Call Weekly_DownTime()
        Call Print_Daily_DownTime()
        Call Weekly_DownTime()
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

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub
End Class