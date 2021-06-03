﻿'// Sales ordering module for the Planner
'// Development Date - 08.05.2014
'// Developed by - Suranga Wijesinghe
'// Audit by     - Amila Priyankara - TJL
'// Referance Table - M01Sales_Order_SAP (Master Table)
'//                 - P01PARAMETER (For add referance No)
'//                 - T01Delivary_Request
'//                 - USERS     
'//---------------------------------------------------------->>>
'Automate the Email  send by merchant to  Planner & Excell data migration
'once merchant enter the S order 2 system.system will gather all requred infor by callling SAP sales order file & fill the e mail requrment

Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports xl = Microsoft.Office.Interop.Excel
Imports System.Globalization
'Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Interop.Outlook
Imports System.Drawing
Imports Spire.XlS
Imports System.Xml
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class frmDelivary_Quatation

    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _EPF As String
    Dim _Email As String
    Dim _Delivary_Qut_No As Integer
    Dim _status As Boolean
    Dim _status1 As Boolean
    Dim _Parameter As Integer
    Dim strMerchent As String
    Dim c_dataCustomer As DataTable
    Dim advancedSearchTag As String = ""
    Dim _LeadTime As String '--------------------------->>> Lead time variable

    Function Search_Parameter()

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select * from P01PARAMETER where P01code='DQ'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Parameter = Trim(M01.Tables(0).Rows(0)("P01no"))
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Lead_Time() As Boolean

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select * from M02Lead_Time_Master where M02Dis='" & Trim(cboLeadTime.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _LeadTime = Trim(M01.Tables(0).Rows(0)("M02Code"))
                Search_Lead_Time = True
            Else
                Search_Lead_Time = False
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Function GetWeekStartDate(ByVal weekNumber As Integer, ByVal year As Integer) As Date
        Dim startDate As New DateTime(year, 1, 1)
        Dim weekDate As DateTime = DateAdd(DateInterval.WeekOfYear, weekNumber - 1, startDate)
        Return DateAdd(DateInterval.Day, (-weekDate.DayOfWeek) + 1, weekDate)
    End Function

    Function MakeDataTable_Delivary_Quatation() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Line Item", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Material", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Req Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("P4P", GetType(Boolean))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        colWork = New DataColumn("Liability", GetType(Boolean))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)

        Return dataTable
    End Function
    Function MakeDataTable_Delivary_Quatation1() As DataTable
        Dim I As Integer
        Dim X As Integer
        Dim _Lastweek As Integer


        ' MsgBox(DatePart("ww", Today))
        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Line Item", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        ' Dim Keys(0) As DataColumn

        ' Keys(0) = colWork
        colWork.ReadOnly = True
        'dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Material", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Quality", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Qty", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("Req Date", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        colWork.ReadOnly = True

        colWork = New DataColumn("P4P", GetType(Boolean))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        '  colWork.ReadOnly = True
        colWork = New DataColumn("Liability", GetType(Boolean))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        '---------------------------------------------------------------------------------
        ' Call Load_Week()



        If chkFOB.Checked = True Then
            Dim _TimeSpan As TimeSpan
            If Trim(txtD_Splite.Text) <> "" Then
                Dim weekStart As DateTime
                If IsNumeric(txtD_Splite.Text) Then
                    _TimeSpan = CDate(txtTo_Date.Text).Subtract(CDate(txtDate.Text))
                    X = _TimeSpan.Days / Val(txtD_Splite.Text)
                    For I = 0 To X
                        If I = 0 Then
                            weekStart = txtDate.Text
                            colWork = New DataColumn(weekStart, GetType(String))
                            colWork.MaxLength = 250
                            dataTable.Columns.Add(colWork)
                        Else
                            weekStart = weekStart.AddDays(+Val(txtD_Splite.Text))
                            colWork = New DataColumn(weekStart, GetType(String))
                            colWork.MaxLength = 250
                            dataTable.Columns.Add(colWork)
                        End If
                    Next

                Else
                    MsgBox("Please enter the Splite by Days", MsgBoxStyle.Information, "Information ......")
                    Exit Function
                End If
            End If
        Else
            Dim weekStart As DateTime
            If Trim(txtFrom.Text) <> "" Then
                If Val(txtTo.Text) >= Val(txtFrom.Text) Then
                    X = Val(txtTo.Text) - Val(txtFrom.Text)
                    ' X = X + 1
                    If X = 0 Then
                    Else
                        For I = 0 To X
                            Dim _String As String
                            If I = 0 Then

                                weekStart = GetWeekStartDate(Val(txtFrom.Text), Year(Today))
                                '  MsgBox(weekStart)
                                weekStart = (weekStart.AddDays(+3))
                                _String = "Wk" & txtFrom.Text & "-" & weekStart
                                colWork = New DataColumn(_String, GetType(String))
                                colWork.MaxLength = 250
                                dataTable.Columns.Add(colWork)
                                ' colWork.ReadOnly = True
                            Else
                                weekStart = GetWeekStartDate(Val(txtFrom.Text) + I, Year(Today))
                                weekStart = (weekStart.AddDays(+3))
                                _String = "Wk" & Val(txtFrom.Text) + I & "-" & weekStart
                                ' _String = "Week" & Val(txtFrom.Text) + I
                                colWork = New DataColumn(_String, GetType(String))
                                colWork.MaxLength = 250
                                dataTable.Columns.Add(colWork)
                            End If
                        Next
                        colWork = New DataColumn("LIB", GetType(String))
                        '  colWork.MaxLength = 70
                        dataTable.Columns.Add(colWork)
                    End If
                Else

                    X = Val(txtTo.Text) - 1
                    _Lastweek = (DatePart("ww", "12/31/" & Year(Today)))
                    X = X + (_Lastweek - Val(txtFrom.Text))
                    X = X + 1

                    If X = 0 Then
                    Else
                        For I = 0 To X
                            Dim _String As String
                            If I = 0 Then

                                weekStart = GetWeekStartDate(Val(txtFrom.Text), Year(Today))
                                '  MsgBox(weekStart)
                                weekStart = (weekStart.AddDays(+3))
                                _String = "Wk" & txtFrom.Text & "-" & weekStart
                                colWork = New DataColumn(_String, GetType(String))
                                colWork.MaxLength = 250
                                dataTable.Columns.Add(colWork)
                                ' colWork.ReadOnly = True
                                weekStart = (weekStart.AddDays(+7))
                            Else
                                _Lastweek = (DatePart(DateInterval.WeekOfYear, weekStart, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays))
                                weekStart = GetWeekStartDate(_Lastweek, Year(weekStart))
                                ' MsgBox(weekStart.Date)
                                weekStart = (weekStart.AddDays(+3))
                                _String = "Wk" & _Lastweek & "-" & weekStart
                                colWork = New DataColumn(_String, GetType(String))
                                colWork.MaxLength = 250
                                dataTable.Columns.Add(colWork)
                                ' colWork.ReadOnly = True
                                weekStart = (weekStart.AddDays(+7))
                            End If
                        Next
                        colWork = New DataColumn("LIB", GetType(String))
                        '  colWork.MaxLength = 70
                        dataTable.Columns.Add(colWork)
                    End If
                End If
            End If
        End If
        colWork = New DataColumn("#", GetType(String))
        '  colWork.MaxLength = 70
        dataTable.Columns.Add(colWork)
        Return dataTable
    End Function

    Function Load_Gride_SalesOrder()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = MakeDataTable_Delivary_Quatation()
        UltraGrid1.DataSource = c_dataCustomer
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 210
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Load_Gride_SalesOrder1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer = MakeDataTable_Delivary_Quatation1()
        UltraGrid1.DataSource = c_dataCustomer
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 210
            '.DisplayLayout.Bands(0).Columns(3).Width = 60
            '.DisplayLayout.Bands(0).Columns(5).Width = 60
            '.DisplayLayout.Bands(0).Columns(8).Width = 60
            '.DisplayLayout.Bands(0).Columns(7).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).Width = 60

        End With
    End Function

    Function Send_Email()

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet


        Dim OutlookApp As Microsoft.Office.Interop.Outlook._Application
        Dim wordInSubject As String
        OutlookApp = New Microsoft.Office.Interop.Outlook.Application
        Dim _RefNo As String
        If Microsoft.VisualBasic.Len(cboRef.Text) = 1 Then
            _RefNo = "000" & Trim(cboRef.Text)
        ElseIf Microsoft.VisualBasic.Len(cboRef.Text) = 2 Then
            _RefNo = "00" & Trim(cboRef.Text)
        ElseIf Microsoft.VisualBasic.Len(cboRef.Text) = 3 Then
            _RefNo = "0" & Trim(cboRef.Text)
        Else
            _RefNo = Trim(cboRef.Text)
        End If

        wordInSubject = Trim(cboSO.Text) & "-" & _RefNo
        Dim scope As String = "Inbox"
        Dim filter As String = "urn:schemas:mailheader:subject LIKE '%" + wordInSubject + "%'"
        Dim advancedSearch As Microsoft.Office.Interop.Outlook.Search = Nothing
        Dim folderInbox As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
        Dim folderSentMail As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
        Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Nothing
        Dim oFolders As Microsoft.Office.Interop.Outlook.Folders
        Dim RootFolder As Microsoft.Office.Interop.Outlook.MAPIFolder
        Dim i As Integer

        Try
            ns = OutlookApp.GetNamespace("MAPI")
            oFolders = ns.Folders
            'MsgBox(oFolders.Count)
            i = 1
            Dim oMsg As Microsoft.Office.Interop.Outlook.MailItem
            Dim oMsg1 As Microsoft.Office.Interop.Outlook.MailItem
            Dim olFormat As OlBodyFormat

            Dim receivetime As Date
            receivetime = "1900/1/1 12:00AM"
            For i = 1 To oFolders.Count
                RootFolder = oFolders.Item(i)
                scope = "'" & RootFolder.FolderPath & "\" & "Inbox'"

                advancedSearch = OutlookApp.AdvancedSearch(scope, filter, True, advancedSearchTag)
                advancedSearch.Results.Sort("[ReceivedTime]", True)
                If advancedSearch.Results.Count > 0 Then
                    oMsg = advancedSearch.Results.GetFirst()
                    If receivetime = "1900/1/1 12:00AM" Then
                        receivetime = oMsg.ReceivedTime
                        oMsg1 = oMsg
                    Else
                        If receivetime > oMsg.ReceivedTime Then
                        Else
                            oMsg1 = oMsg
                        End If
                    End If
                End If
             

            Next
            'olFormat = Microsoft.Office.Interop.Outlook.OlFormatText
            'Dim oResponse As MailItem
            'oResponse = oMsg1.ReplyAll
            'oResponse.BodyFormat = OlBodyFormat.olFormatPlain

            'oResponse.Display()
            'oMsg1.Display()
            'SendKeys.SendWait("^+R")
            '------------------------------------------------------------
            'FINDING WOORK BOOK RANGE
            'DEVELOPED BY SURANGA WIJESINGHE
            Dim _FilePath As String
            _FilePath = ConfigurationManager.AppSettings("FilePath") + "\" & Trim(cboSO.Text) & "-" & Trim(cboRef.Text) & ".xlsx"
            Dim app As New Microsoft.Office.Interop.Excel.Application
            Dim WB As Microsoft.Office.Interop.Excel.Workbook = app.Workbooks.Open(_FilePath)

        If WB IsNot Nothing Then

            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = WB.Worksheets.Item(1)
            Dim lCol As Long = 0
                Dim lRow As Long = 0
                Dim _Range As String

            With ws
                '~~> Check if there is any data in the sheet
                If app.WorksheetFunction.CountA(.Cells) <> 0 Then
                    lCol = .Cells.Find(What:="*", _
                                  After:=.Range("A1"), _
                                  LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, _
                                  LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, _
                                  SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, _
                                  SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, _
                                  MatchCase:=False).Column
                Else
                    lCol = 1
                End If

                If app.WorksheetFunction.CountA(.Cells) <> 0 Then
                    lRow = .Cells.Find(What:="*", _
                                  After:=.Range("A1"), _
                                  LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, _
                                  LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, _
                                  SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, _
                                  SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, _
                                  MatchCase:=False).Row
                Else
                    lRow = 1
                End If
            End With
            Dim A As String
            Dim X As Integer


            If lCol <= 26 Then
                X = 96 + lCol
                A = (UCase(Chr(X)))
            ElseIf lCol > 26 And lCol <= 52 Then
                X = 96 + (lCol - 26)
                A = "A" & (UCase(Chr(X)))
            ElseIf lCol > 52 And lCol <= 78 Then
                X = 96 + (lCol - 52)
                A = "B" & (UCase(Chr(X)))
            End If

                _Range = "A1:" & A & lRow

                Dim xlRn As Microsoft.Office.Interop.Excel.Range
                Dim Connect As String
                Dim strbody As String

                'strBody = "This is a test " & vbCrLf & vbCrLf & "Thanks Michael"
                '  RangetoHTML(xlRn)

                Connect = ws.Range(_Range).Copy()
                'SendKeys.SendWait("^V")
                xlRn = ws.Range(_Range)
                xlRn.Copy()
                Dim strNewText As String
                If Trim(_LeadTime) = "01" Then
                    strNewText = "Dear " & strMerchent & ",<br>Please find the possible delivery in below "
                Else
                    strNewText = "Dear " & strMerchent & ",<br>Please find the possible short Lead time delivery in below "
                End If
                Dim oResponse As MailItem
                oResponse = oMsg1.ReplyAll

                oResponse.BodyFormat = OlBodyFormat.olFormatHTML
                oResponse.HTMLBody = (strNewText & RangetoHTML(xlRn) & oResponse.HTMLBody)

                Sql = "select * from T03Delivary_Quat_Flutter where T03RefNo=" & _Parameter & ""
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    Dim _LineItems As String
                    Dim M03 As DataSet

                    _LineItems = ""
                    i = 0
                    Sql = "select T03Line_Item from T03Delivary_Quat_Flutter where T03RefNo=" & _Parameter & " and T03P4P='Y' group by T03Line_Item"
                    M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    For Each DTRow2 As DataRow In M03.Tables(0).Rows
                        If i = 0 Then
                            _LineItems = Trim(M03.Tables(0).Rows(i)("T03Line_Item"))
                        Else
                            _LineItems = _LineItems & "," & Trim(M03.Tables(0).Rows(i)("T03Line_Item"))

                        End If
                        i = i + 1
                    Next

                    Sql = "select * from users where Designation='PRINT PLANNER'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        strNewText = "Dear " & M02.Tables(0).Rows(0)("Name") & ",<br>Please quote best possible print delivary for Line Items " & _LineItems
                        'Dim oResponse As MailItem
                        'oResponse = oMsg1.ReplyAll
                        oResponse.CC = Trim(M02.Tables(0).Rows(0)("email"))
                        oResponse.BodyFormat = OlBodyFormat.olFormatHTML
                        oResponse.HTMLBody = (strNewText & oResponse.HTMLBody)
                    End If
                End If
                oResponse.Display()
                'SendKeys.SendWait("^+R")

                '  WB.Close(False)
                'app.Quit()

                '~~> Clean Up
                releaseObject(ws)
                releaseObject(WB)
                releaseObject(app)

            End If

            'Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            'Dim xlWb As Microsoft.Office.Interop.Excel.Workbook
            'Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
            'Dim lRow As Long = 0

            '            With xlApp
            '               .Visible = True

            '~~> Open workbook
            '              xlWb = .Workbooks.Open(wordInSubject)

            '~~> Set it to the relevant sheet
            '             xlsheet = xlWb.Sheets("Sheet1")

            '            With xlsheet
            '               lRow = .Range("A" & .Rows.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            '          End With

            '         MessageBox.Show("The last row in Col A of Sheet1 which has data is " & lRow)

            '~~> Close workbook and quit Excel
            '        xlWb.Close(False)
            '       xlApp.Quit()

            '~~> Clean Up
            '      releaseObject(xlsheet)
            '     releaseObject(xlWb)
            '    releaseObject(xlApp)

            'End With

            'RootFolder = oFolders.Item(4)
            'folderInbox = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)
            'folderSentMail = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail)
            'scope = "'" + folderInbox.FolderPath + "','" + folderInbox.FolderPath + "'"
            'scope = "'" + folderInbox.FolderPath + "'"
            'advancedSearch = OutlookApp.AdvancedSearch(scope, filter, True, advancedSearchTag)
            'Dim objItems As Microsoft.Office.Interop.Outlook._Items = folderInbox.Items
            'objItems.Sort("[ReceivedTime]", True)
            'If advancedSearch.Results.Count > 0 Then

            '   Dim objMessage1 As Microsoft.Office.Interop.Outlook._MailItem = advancedSearch.Results.GetLast() 'objItems.Item(wordInSubject)
            '  objMessage1.Display()
            ' SendKeys.SendWait("^+R")

            'Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            'Dim xlWb As Microsoft.Office.Interop.Excel.Workbook
            'Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
            'Dim lRow As Long = 0

            'With xlApp
            '    .Visible = True

            '    '~~> Open workbook
            '    xlWb = .Workbooks.Open(wordInSubject)

            '    '~~> Set it to the relevant sheet
            '    xlsheet = xlWb.Sheets("Sheet1")

            '    With xlsheet
            '        lRow = .Range("A" & .Rows.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            '    End With

            '    MessageBox.Show("The last row in Col A of Sheet1 which has data is " & lRow)

            '    '~~> Close workbook and quit Excel
            '    xlWb.Close(False)
            '    xlApp.Quit()

            '    '~~> Clean Up
            '    releaseObject(xlsheet)
            '    releaseObject(xlWb)
            '    releaseObject(xlApp)

            'End With



            'Else
            'MsgBox("1")
            'End If
        Catch ex As EvaluateException
            MessageBox.Show(ex.Message, "An eexception is thrown")
        Finally
            If Not IsNothing(advancedSearch) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(advancedSearch)
            If Not IsNothing(folderSentMail) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(folderSentMail)
            If Not IsNothing(folderInbox) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(folderInbox)
            If Not IsNothing(ns) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(ns)
        End Try
    End Function

    Private Sub frmDelivary_Quatation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call Load_Gride_SalesOrder()
        Call Search_Parameter()

        Call Load_Sales_Order()
        Call Load_Combo_Lead_Time()

        txtFrom.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTo.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtD_Splite.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtDate.Text = Today
        txtTo_Date.Text = Today

    End Sub

    Function RangetoHTML(ByVal rng As Microsoft.Office.Interop.Excel.Range)
        ' Changed by Ron de Bruin 28-Oct-2006
        ' Working in Office 2000-2010
        Dim fso As Object
        Dim ts As Object
        Dim TempFile As String
        ' Dim TempWB As Microsoft.Office.Interop.Excel.Workbook

        Dim exc As New Microsoft.Office.Interop.Excel.Application
        Dim TempWB1 As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
        Dim TempWB As Microsoft.Office.Interop.Excel._Workbook = TempWB1.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)

        TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

        'Copy the range and create a new workbook to past the data in
        rng.Copy()
        'TempWB = Microsoft.Office.Interop.Excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)




        With TempWB.Sheets(1)
            .Cells(1).PasteSpecial(Paste:=8)
            ' Microsoft.Office.Interop.Excel.XlPastef
            '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, , False, False)
            '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, , False, False)
            '.Cells(1).Select()
            'Application.CutCopyMode = False
            On Error Resume Next
            .DrawingObjects.Visible = True
            .DrawingObjects.Delete()
            On Error GoTo 0
        End With


        'Publish the sheet to a htm file
        With TempWB.PublishObjects.Add( _
             SourceType:=Microsoft.Office.Interop.Excel.XlSourceType.xlSourceRange, _
             Filename:=TempFile, _
             Sheet:=TempWB.Sheets(1).Name, _
             Source:=TempWB.Sheets(1).UsedRange.Address, _
             HtmlType:=Microsoft.Office.Interop.Excel.XlHtmlType.xlHtmlStatic)
            .Publish(True)
        End With

        'Read all data from the htm file into RangetoHTML
        fso = CreateObject("Scripting.FileSystemObject")
        ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
        RangetoHTML = ts.ReadAll
        ts.Close()
        RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                              "align=left x:publishsource=")

        'Close TempWB
        TempWB.Close(savechanges:=False)

        'Delete the htm file we used in this function
        Kill(TempFile)

        ts = Nothing
        fso = Nothing
        TempWB = Nothing
    End Function

    Function Load_Gride_With_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer
        Dim T01 As DataSet
        Try
            Call Load_Gride_SalesOrder1()
            Sql = "select * from T01Delivary_Request inner join M01Sales_Order_SAP on T01Sales_Order=CONVERT(INT,M01Sales_Order) where T01Sales_Order='" & Trim(cboSO.Text) & "' and T01Planner='" & strDisname & "' and T01Status='A' order by T01Line_Item"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Delivary_Qut_No = M01.Tables(0).Rows(0)("T01RefNo")
                txtPO.Text = M01.Tables(0).Rows(0)("M01PO")
                cboRef.Text = M01.Tables(0).Rows(0)("T01RefNo")
            End If

            i = 0

            Sql = "select * from T01Delivary_Request  where T01Sales_Order='" & Trim(cboSO.Text) & "' and T01Planner='" & strDisname & "' and T01Status='A' order by T01Line_Item"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow4 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer.NewRow
                newRow("Line Item") = M01.Tables(0).Rows(i)("T01Line_Item")
                Sql = "select * from M01Sales_Order_SAP where CONVERT(INT,M01Sales_Order)='" & Trim(cboSO.Text) & "' and M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Line_Item")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    newRow("Material") = T01.Tables(0).Rows(0)("M01Material_No")
                    newRow("Quality") = T01.Tables(0).Rows(0)("M01Quality")

                End If
                newRow("Qty") = M01.Tables(0).Rows(i)("T01Qty")
                newRow("Req Date") = Month(M01.Tables(0).Rows(i)("T01RQD")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01RQD")) & "/" & Year(M01.Tables(0).Rows(i)("T01RQD"))
                newRow("P4P") = False
                newRow("Liability") = False


                c_dataCustomer.Rows.Add(newRow)


                i = i + 1
            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

            '-------------------------------------------------------------


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try


    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click

        Call Load_Gride_SalesOrder()
        Call Load_Gride_With_Data()
        _status = False
        _status1 = False
        Call Load_Week()
        '  Call Upload_NotePad_File()
        'SendKeys.SendWait("%B")
    End Sub

    Function Create_Excel_File()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim _RefNo As String
        Dim T01 As DataSet
        Dim T02 As DataSet

        Try
            Dim exc As New Microsoft.Office.Interop.Excel.Application
            Dim workbooks As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
            Dim workbook As Microsoft.Office.Interop.Excel._Workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Microsoft.Office.Interop.Excel.Sheets = workbook.Worksheets
            Dim worksheet1 As Microsoft.Office.Interop.Excel._Worksheet = CType(sheets.Item(1), Microsoft.Office.Interop.Excel._Worksheet)
            Dim range1 As Microsoft.Office.Interop.Excel.Range

            Dim objApp As Object
            Dim objEmail As Object
            If Microsoft.VisualBasic.Len(_Parameter) = 1 Then
                _RefNo = "000" & Trim(_Parameter)
            ElseIf Microsoft.VisualBasic.Len(_Parameter) = 2 Then
                _RefNo = "00" & Trim(_Parameter)
            ElseIf Microsoft.VisualBasic.Len(_Parameter) = 3 Then
                _RefNo = "0" & Trim(_Parameter)
            Else
                _RefNo = Trim(_Parameter)
            End If


            If exc.Visible = True Then
                exc.Visible = False
                exc.Visible = True
            Else
                ' exc.Visible = False
                exc.Visible = True
            End If
            worksheet1.Rows(5).Font.size = 10
            worksheet1.Rows(5).Font.bold = True
            'worksheet1.Rows(5).width = 23

            worksheet1.Cells(5, 1) = "S/O"
            worksheet1.Cells(5, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(5).Font.size = 10
            worksheet1.Columns("A").ColumnWidth = 10

            worksheet1.Cells(5, 2) = "Line Item"
            worksheet1.Cells(5, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("B").ColumnWidth = 12

            worksheet1.Cells(5, 3) = "Material"
            worksheet1.Cells(5, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("C").ColumnWidth = 20

            worksheet1.Cells(5, 4) = "Quality"
            worksheet1.Cells(5, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("D").ColumnWidth = 30

            worksheet1.Cells(5, 5) = "Quantity"
            worksheet1.Cells(5, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("E").ColumnWidth = 8

            worksheet1.Cells(5, 6) = "Matching"
            worksheet1.Cells(5, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("F").ColumnWidth = 10

            worksheet1.Cells(5, 7) = "Retailer"
            worksheet1.Cells(5, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("G").ColumnWidth = 15

            worksheet1.Cells(5, 8) = "1st Bulk"
            worksheet1.Cells(5, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("H").ColumnWidth = 15

            worksheet1.Cells(5, 9) = "Lab dye"
            worksheet1.Cells(5, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("I").ColumnWidth = 10

            worksheet1.Cells(5, 10) = "P/App Date"
            worksheet1.Cells(5, 10).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("J").ColumnWidth = 10

            worksheet1.Cells(5, 11) = "NPL"
            worksheet1.Cells(5, 11).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("K").ColumnWidth = 8

            worksheet1.Cells(5, 12) = "PP"
            worksheet1.Cells(5, 12).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("L").ColumnWidth = 10

            worksheet1.Cells(5, 13) = "Reg.Del.Date"
            worksheet1.Cells(5, 13).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("M").ColumnWidth = 10
            Dim A As Integer
            Dim i As Integer

            A = 97

            For i = 1 To 13
                worksheet1.Range(Chr(A) & "5:" & Chr(A) & "5").Interior.Color = RGB(0, 112, 192)
                A = A + 1
            Next

            A = 97
            i = 0
            For i = 1 To 13
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next
            '------------------------------------------------------------------------------------------------------------------------
            Sql = "select * from T01Delivary_Request  where T01Sales_Order=" & Trim(cboSO.Text) & "  and  T01Planner='" & strDisname & "' order by T01Line_Item"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            Dim X As Integer
            Dim Y As Integer

            X = 6
            Y = 1
            For Each DTRow As DataRow In M01.Tables(0).Rows
                Y = 1
                worksheet1.Rows(X).Font.size = 8
                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Sales_Order")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Line_Item")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                Sql = "SELECT * FROM M01Sales_Order_SAP WHERE convert(int,M01Sales_Order)='" & Trim(M01.Tables(0).Rows(i)("T01Sales_Order")) & "' AND M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Line_Item")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("M01Material_No")
                    ' worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1


                    worksheet1.Cells(X, Y + 3) = T01.Tables(0).Rows(0)("M01Department")
                    worksheet1.Cells(X, Y + 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("M01Quality")
                    ' worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1
                End If
                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Qty")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                If Trim(M01.Tables(0).Rows(i)("T01Maching")) <> "" Then
                    worksheet1.Cells(X, Y) = "Yes"
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1



                worksheet1.Cells(X, Y + 1) = M01.Tables(0).Rows(i)("T01Bulk")
                worksheet1.Cells(X, Y + 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 2

                If Trim(M01.Tables(0).Rows(i)("T01Lab_Dye")) <> "" Then

                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Lab_Dye")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Else

                    worksheet1.Cells(X, Y) = "NOT APPROVED"
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1

                If Trim(M01.Tables(0).Rows(i)("T01Bulk")) = "1st BULK" Then
                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01POD")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01NPL")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01PP")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01RQD")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                ' Dim Z As Integer
                A = 97
                ' i = 0
                Dim Z As Integer
                For Z = 1 To 13
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    A = A + 1
                Next

                X = X + 1
                i = i + 1
            Next

            str_ExcelRow = X - 6

            X = X + 2
            Y = 5
            worksheet1.Cells(X, Y) = "Body Order Info"
            worksheet1.Range(worksheet1.Cells(X, Y), worksheet1.Cells(X, Y + 2)).Merge()
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.BOLD = True

            worksheet1.Cells(X, Y + 4) = "Trim Order Info"
            worksheet1.Range(worksheet1.Cells(X, Y + 4), worksheet1.Cells(X, Y + 6)).Merge()
            worksheet1.Cells(X, Y + 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, Y + 8) = "T/B Ratio"
            A = 101
            i = 0
            Dim Z1 As Integer

            For Z1 = 5 To 13
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next

            A = 101

            For Z1 = 5 To 13
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
                A = A + 1
            Next

            X = X + 1
            Y = 5
            worksheet1.Cells(X, Y) = "L/Item"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.BOLD = True
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Quality"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Qty"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 2
            worksheet1.Cells(X, Y) = "L/Item"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.BOLD = True
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Quality"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Qty"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 2
            A = 101
            For Z1 = 5 To 13
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next
            A = 104
            worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
            A = 108
            worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)

            Sql = "select * from T01Delivary_Request where T01RefNo=" & _Delivary_Qut_No & "  and T01Maching<>''"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0

            For Each DTRow As DataRow In M01.Tables(0).Rows
                X = X + 1
                Y = 5
                worksheet1.Rows(X).Font.size = 8

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Line_Item")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1
                Sql = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(M01.Tables(0).Rows(i)("T01Sales_Order")) & "' and M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Line_Item")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("M01Quality")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1
                End If

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Qty")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 2

                Sql = "select * from T01Delivary_Request where T01Sales_Order='" & Trim(M01.Tables(0).Rows(i)("T01Sales_Order")) & "' and T01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Maching")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Maching")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1

                    Sql = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(T01.Tables(0).Rows(0)("T01Sales_Order")) & "' and M01Line_Item='" & Trim(T01.Tables(0).Rows(0)("T01Line_Item")) & "'"
                    T02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(T02) Then
                        worksheet1.Cells(X, Y) = T02.Tables(0).Rows(0)("M01Quality")
                        worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        Y = Y + 1
                    End If
                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("T01Qty")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 2

                    worksheet1.Cells(X, Y) = "=K" & X & "/G" & X
                    range1 = worksheet1.Cells(X, Y)
                    range1.NumberFormat = "0.00"
                End If

                A = 101
                For Z1 = 5 To 13
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    A = A + 1
                Next

                A = 104
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
                A = 108
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)

                i = i + 1
            Next
            '------------------------------------------------------------------------------------------------------------------------
            Dim MyPassword As String
            Dim _Path As String

            MyPassword = strDisname
            _Path = "D:\SO\" & Trim(cboSO.Text) & "-" & _Delivary_Qut_No & ".xlsx"
            'worksheet1.SaveAs(Filename:=_Path, FileFormat:=51, Password:=MyPassword, _
            'WriteResPassword:=MyPassword, ReadOnlyRecommended:=True, CreateBackup:=False)
            'worksheet1.SaveAs(Filename:=_Path)
            workbook.SaveAs(_Path)
            '    workbook.Close()
            ' xlWb.Close(False)
            '   exc.Quit()
            ' workbooks.SaveAs(_Path)
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            workbook = workbooks.Open(_Path)
            exc.Visible = True
            releaseObject(sheets)
            releaseObject(workbook)
            releaseObject(exc)
            UltraButton1.Enabled = True
            ' sheets = workbook.Sheets
            ' Shell("Notepad.exe " & _Path, vbNormalFocus)
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Sales_Order()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select T01Sales_Order as [Sales Order],max(M01Cuatomer_Name) as [Customer] from T01Delivary_Request inner join M01Sales_Order_SAP on T01Sales_Order=CONVERT(INT,M01Sales_Order) where T01Planner='" & strDisname & "' and T01Status='A' group by T01Sales_Order"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                .Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Combo_Lead_Time()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select M02Dis as [Lead Time] from M02Lead_Time_Master where M02Code in ('01','02')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLeadTime
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                '   .Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Serch_Recode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            ' Call Load_Gride_SalesOrder()
            Sql = "select * from T01Delivary_Request inner join M01Sales_Order_SAP on T01Sales_Order=CONVERT(INT,M01Sales_Order) where T01Sales_Order='" & Trim(cboSO.Text) & "' and T01Planner='" & strDisname & "' and T01Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Delivary_Qut_No = M01.Tables(0).Rows(0)("T01RefNo")
                cboRef.Text = M01.Tables(0).Rows(0)("T01RefNo")
                txtPO.Text = M01.Tables(0).Rows(0)("M01PO")
                strMerchent = M01.Tables(0).Rows(0)("T01User")
                cmdSave.Enabled = True
            End If
            '----------------------------------------------------------------------------------
            'LOAD DATA TO CBOREF COMBOBOX
            Sql = "select T01RefNo AS [Ref.Doc.No] from T01Delivary_Request  where T01Sales_Order='" & Trim(cboSO.Text) & "' and T01Planner='" & strDisname & "' and T01Status='A' group by T01RefNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRef
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Private Sub cboSO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.AfterCloseUp
        Call Serch_Recode()
    End Sub

    Function Search_Salrs_Order() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            'SERCH SALES ORDER
            Sql = "select T01Sales_Order  from T01Delivary_Request  where T01RefNo='" & Trim(cboRef.Text) & "' and T01Planner='" & strDisname & "' and T01Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                cboSO.Text = M01.Tables(0).Rows(0)("T01Sales_Order")
                Search_Salrs_Order = True

                cmdSave.Enabled = True
            Else
                Search_Salrs_Order = False
            End If
            '----------------------------------------------------------------------------------


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As EvaluateException
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim xlApp As New xl.Application
        Dim xlWb As xl.Workbook
        Dim xlsheet As xl.Worksheet
        Dim lCol As Long = 0
        Dim _Path As String
        _Path = "D:\SO\" & Trim(cboSO.Text) & "-" & _Delivary_Qut_No & ".xlsx"
        With xlApp
            .Visible = True

            '~~> Open workbook
            xlWb = .Workbooks.Open(_Path)

            '~~> Set it to the relevant sheet
            xlsheet = xlWb.Sheets("Sheet1")

            With xlsheet
                lCol = .Cells(1, .Columns.Count).End(xl.XlDirection.xlToLeft).Column
            End With

            MessageBox.Show("The last column which has data in Row 1 of Sheet1 is " & lCol)

            '~~> Close workbook and quit Excel
            xlWb.Close(False)
            xlApp.Quit()

            '~~> Clean Up
            releaseObject(xlsheet)
            releaseObject(xlWb)
            releaseObject(xlApp)

        End With
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_Excel_File()
    End Sub

    Function Load_Week()
        Dim maxCell As Microsoft.Office.Interop.Excel.Range
        Dim ws As xl.Worksheet
        Dim myxl As xl.Application
        myxl = GetObject(, "Excel.application")
        ws = myxl.ActiveSheet
        Dim _Address As String
        Dim _FilePath As String

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        _FilePath = ConfigurationManager.AppSettings("FilePath") + "\" & Trim(cboSO.Text) & "-" & Trim(cboRef.Text) & ".xlsx"
        'Dim app As New Microsoft.Office.Interop.Excel.Application
        ' Dim WB As Microsoft.Office.Interop.Excel.Workbook = app.Workbooks.Open(_FilePath)
        'Dim WB As Microsoft.Office.Interop.Excel.Workbook = myxl.Workbooks.Open(_FilePath)
        ws = myxl.ActiveSheet
        Dim NumRows As Long
        Dim NumCols As Long
        ' Dim R1 As Microsoft.Office.Interop.Excel.Range

        'ws = WB.ActiveSheet
        SendKeys.SendWait("^S")
        ' ActName = ws.Name
        maxCell = Nothing
        Try
            maxCell = DirectCast(ws.Cells(ws.Cells.Find("*", _
            DirectCast(ws.Cells(1, 1), Microsoft.Office.Interop.Excel.Range), _
            Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, _
            Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, False, False).Row, _
            ws.Cells.Find("*", DirectCast(ws.Cells(1, 1), Microsoft.Office.Interop.Excel.Range), _
            Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, _
            Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, False, False).Column), Microsoft.Office.Interop.Excel.Range)

            Dim _FirstColum As Integer
            _FirstColum = (Val(txtTo.Text) - Val(txtFrom.Text)) + 2

            Dim A As String
            Dim _Chrcount As Integer
            Dim _ENDChrcount As Integer
            Dim _LastRow As Integer

            _Chrcount = 97 + (maxCell.Column - _FirstColum)
            If maxCell.Column <= 26 Then
                A = UCase(Chr(97 + (maxCell.Column - 1)))
            ElseIf maxCell.Column > 26 And maxCell.Column <= 52 Then
                ' A="A" & 
            End If

            Sql = "select * from T01Delivary_Request where T01RefNo=" & Trim(cboRef.Text) & " and T01Sales_Order='" & Trim(cboSO.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            _LastRow = M01.Tables(0).Rows.Count + 5
            'MsgBox(UCase(Chr(_Chrcount)))
            'MsgBox(maxCell.Address)
            _Address = maxCell.Address
            ' R1 = maxCell.Resize(IIf(NumRows = 0, maxCell.Rows.Count, NumRows), IIf(NumCols = 0, maxCell.Columns.Count, NumCols)).Address(External:=True)
            ' Dim maxCell As Microsoft.Office.Interop.Excel.Range = Nothing
            'maxCell = DirectCast(ws.Cells(ws.Cells.Find("*", _
            'DirectCast(ws.Cells(1, 1), Microsoft.Office.Interop.Excel.Range), _
            'Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, _
            'Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, False, False).Row, _
            'ws.Cells.Find("*", DirectCast(ws.Cells(1, 1), Microsoft.Office.Interop.Excel.Range), _
            'Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, _
            'Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, False, False).Column), Microsoft.Office.Interop.Excel.Range)

            'MsgBox(maxCell.Address)
            maxCell = myxl.Selection
            maxCell = myxl.Range(UCase(Chr(_Chrcount)) & "6:" & A & _LastRow)
            Dim Connect As String

            'ws.Range(maxCell).Select()
            'ws.Range(maxCell).Copy()
            '---------------------------------------------------------

            Dim currentFind As Microsoft.Office.Interop.Excel.Range = Nothing

            Dim firstFind As Microsoft.Office.Interop.Excel.Range = Nothing
            _Address = "$A$1:" & _Address
            currentFind = ws.Range(_Address).Find(Trim(cboSO.Text), , Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, False)

            If currentFind IsNot Nothing Then

                '  MessageBox.Show("Text found, position is Row-" & currentFind.Row & " and column-" & currentFind.Column)

            Else

                MessageBox.Show("Sales Order no match", "Information .....", _
          MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Exit Function
            End If

            Dim myArray As Object(,) '<-- declared as 2D Array
            Dim Z As Integer
            Dim y As Integer
            myArray = maxCell.Value

            y = 0
            ' MsgBox(myArray.Length)
            If myArray.Length <= 1 Then
                MsgBox("Please select the week", MsgBoxStyle.Information, "Information ........")
                Exit Function
            Else
                For r As Integer = 1 To myArray.GetUpperBound(0)
                    Z = 7
                    For c As Integer = 1 To myArray.GetUpperBound(1)

                        Dim myValue As Object = myArray(r, c)
                        'If myValue = Nothing Then
                        UltraGrid1.Rows(y).Cells(Z).Value = myValue
                        ' End If
                        Z = Z + 1
                    Next c
                    y = y + 1
                Next r
            End If
            'ws.Range(maxCell.Address).Copy()
            'Dim Path As String
            'Path = "D:\" & Trim(cboSO.Text) & ".txt"
            'If File.Exists(Path) = False Then
            '    ' Create a file to write to. 
            '    Dim sw As StreamWriter = File.CreateText(Path)
            '    sw.Close()
            'End If
            '' Shell("notepad.exe", vbMaximizedFocus)
            'Shell("Notepad.exe " & Path, vbNormalFocus)
            'SendKeys.SendWait("^v")

            ''SendKeys.SendWait("^S")
            ''SendKeys.Send("^a")
            ''SendKeys.Send("^c")
            ''SendKeys.Send("%{F4}{TAB}{ENTER}")
            ''SendKeys("^v")
            'Dim strWindowName As String
            '' SendKeys.SendWait("^S")
            'strWindowName = Trim(cboSO.Text) & ".txt - Notepad"
            'CloseNotepad(strWindowName)
            '' SendKeys.SendWait("^S")
            '_status = True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Function Upload_NotePad_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim X As Integer
        Dim i As Integer
        Dim Y As Integer
        Dim Z As Integer
        Dim _Lastweek As Integer

        'System.IO.File.Open()
        Try
            X = 0
            If Val(txtTo.Text) >= Val(txtFrom.Text) Then
                i = Val(txtTo.Text) - Val(txtFrom.Text)
                ' i = i + 1
            Else
                i = Val(txtTo.Text) - 1
                _Lastweek = (DatePart("ww", "12/31/" & Year(Today)))
                i = i + (_Lastweek - Val(txtFrom.Text))
                i = i + 1


            End If
            ' strFileName = ConfigurationManager.AppSettings("FilePath") + "\quality.txt"
            strFileName = "D:\" & Trim(cboSO.Text) & ".txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                Dim _NotePadVariable(i) As Double

                Z = 6
                For Y = 0 To i + 1
                    UltraGrid1.Rows(X).Cells(Z).Value = Trim(fields(Y))
                    'If Z = i + 1 Then
                    '    If Trim(fields(Y)) <> "" Then
                    '        UltraGrid1.Rows(X).Cells(4).Value = True
                    '    End If
                    'End If
                    Z = Z + 1
                Next
                If Trim(UltraGrid1.Rows(X).Cells(Z - 1).Value) <> "" Then
                    UltraGrid1.Rows(X).Cells(5).Value = True
                End If
                X = X + 1
            Next
            Kill(strFileName)
            _status1 = True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Public Sub CloseNotepad(ByVal strWindowName As String)
        Dim oWdApp As Object
        Dim oTask As Object
        Dim bCreate As Boolean

        On Error Resume Next
        oWdApp = GetObject(, "Word.Application")
        ' oWdApp = GetObject(, "notepad.Application")
        On Error GoTo 0
        If oWdApp Is Nothing Then
            oWdApp = CreateObject("Word.Application")
            ' oWdApp = CreateObject("notepad.Application")
            oWdApp.Visible = False
            bCreate = True
        End If

        For Each oTask In oWdApp.Tasks
            'Provide a condition here which will match the notepad instances created by your code!
            If InStr(oTask.Name, strWindowName) > 0 Then
                'oTask.Save()
                oTask.Close()
                SendKeys.SendWait("%S")
                'oTask.close()
            End If
        Next

        If bCreate Then oWdApp.Quit()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Call Load_Week()
        Call Upload_NotePad_File()
    End Sub

    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Week()
    End Sub

    Private Sub UltraGrid1_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles UltraGrid1.BeforeCellUpdate
        If _status = False Then
            'Call Load_Week()
            'Call Upload_NotePad_File()
        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Private Sub chkFOB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFOB.CheckedChanged
        txtD_Splite.Text = ""
        If chkFOB.Checked = True Then
            lblFrom.Text = "From Date"
            lblTo.Text = "To Date"
            lblSplite.Visible = True
            txtD_Splite.Visible = True
            txtTo.Visible = False
            txtFrom.Visible = False
            txtDate.Visible = True
            txtTo_Date.Visible = True
        Else
            lblFrom.Text = "From Week"
            lblTo.Text = "To Week"
            lblSplite.Visible = False
            txtD_Splite.Visible = False
            txtTo.Visible = True
            txtFrom.Visible = True
            txtDate.Visible = False
            txtTo_Date.Visible = False
        End If
    End Sub

    Private Sub cboRef_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRef.AfterCloseUp
        Call Search_Salrs_Order()
    End Sub

    Private Sub cboRef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRef.LostFocus
        Call Search_Salrs_Order()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer
        Dim Z1 As Integer
        Dim _Qty As Integer
        Dim Z2 As Integer
        Dim Z3 As Integer
        Dim Z4 As Integer
        Dim _GrideStatus As Boolean
        Dim A As String

        Try
            For i = 0 To UltraGrid1.Rows.Count - 1
                _Qty = 0
                Z1 = 7
                Z4 = 7
                Z2 = UltraGrid1.DisplayLayout.Bands(0).Columns.Count
                _GrideStatus = True
                For Z3 = 1 To (Z2 - Z1) - 1
                    'MsgBox((Trim(UltraGrid1.Rows(i).Cells(Z4).Value)))
                    If IsNumeric((Trim(UltraGrid1.Rows(i).Cells(Z4).Text))) Then
                        _Qty = _Qty + Val(UltraGrid1.Rows(i).Cells(Z4).Text)
                    Else
                        If Trim(UltraGrid1.Rows(i).Cells(Z4).Text) <> "" Then
                            _GrideStatus = False
                        End If
                    End If
                    Z4 = Z4 + 1
                Next
                If _GrideStatus = True Then
                    ' MsgBox(UltraGrid1.Rows(i).Cells(3).Text)
                    If _Qty = Val(Val(UltraGrid1.Rows(i).Cells(3).Text)) Then
                    Else
                        MessageBox.Show("Line Item Qty and Splite Qty not match", "Information .....", _
         MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                End If
            Next

            If Search_Salrs_Order() = True Then
            Else
                MsgBox("Please select the correct Sales Order", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If

            If Search_Lead_Time() = True Then
            Else
                Dim result1 As String
                result1 = MessageBox.Show("Please select the Lead Time ", "Information .....", _
MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLeadTime.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '-----------------------------------------------------------------------------
            Call Search_Parameter()
            'UPDATE PARAMETER
            nvcFieldList1 = "update P01PARAMETER set P01NO=P01NO +" & 1 & " where P01CODE='DQ'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-----------------------------------------------------------------------------
            'INSERT T02Delivary_Quat_Header
            nvcFieldList1 = "Insert Into T02Delivary_Quat_Header(T02RefNo,T02Del_Req_No,T02OrderNo,T02Entry_Date,T02Entry_Time,T02User,T02Status,T02Lead_Time,T02SAP_Tran)" & _
                                                            " values(" & _Parameter & ", '" & Trim(cboRef.Text) & "','" & Trim(cboSO.Text) & "','" & Today & "','" & Now & "','" & strDisname & "','A','" & _LeadTime & "','N')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-----------------------------------------------------------------------------
            ' Insert T03Delivary_Quat_Flutter
            Dim _WEEKNO As Integer
            Dim _Lastweek As Integer
            Dim X As Integer
            Dim _P4P As String
            Dim _Liability As String
            Dim Z As Integer
            Dim _Qty_INT As Integer

            _WEEKNO = 0

            If Val(txtTo.Text) >= Val(txtFrom.Text) Then
                _WEEKNO = Val(txtTo.Text) - Val(txtFrom.Text)
            Else
                _WEEKNO = Val(txtTo.Text) - 1
                _Lastweek = (DatePart("ww", "12/31/" & Year(Today)))
                _WEEKNO = _WEEKNO + (_Lastweek - Val(txtFrom.Text))

            End If

            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                X = 0
                _Qty_INT = 0
                _Qty = 0
                'P4P STATUS
                If Trim(UltraGrid1.Rows(i).Cells(5).Text) = True Then
                    _P4P = "Y"
                Else
                    _P4P = "N"
                End If
                '------------------------------------------------------
                'LIABILITY
                If Trim(UltraGrid1.Rows(i).Cells(6).Text) = True Then
                    _Liability = "Y"
                Else
                    _Liability = "N"
                End If
                Z = 7

                Dim strWeek_No As Integer
                Dim weekStart As Date
                Dim dfi = DateTimeFormatInfo.CurrentInfo
                Dim calendar = dfi.Calendar

                _GrideStatus = True
                For X = 0 To _WEEKNO
                    'INSERT T02Delivary_Quat_Header
                    _Qty_INT = 0
                    If chkFOB.Checked = True Then
                        If X = 0 Then
                            weekStart = txtDate.Text
                        Else
                            weekStart = weekStart.AddDays(+1)
                        End If
                        strWeek_No = calendar.GetWeekOfYear(weekStart, dfi.CalendarWeekRule, DayOfWeek.Thursday)()
                    Else
                        If X = 0 Then
                            strWeek_No = txtFrom.Text
                            weekStart = GetWeekStartDate(strWeek_No, Year(Today))
                            weekStart = (weekStart.AddDays(+3))
                        Else
                            If strWeek_No > 52 Then
                                If strWeek_No = 53 Then
                                    strWeek_No = 1
                                End If
                                strWeek_No = strWeek_No + 1
                                weekStart = GetWeekStartDate(strWeek_No, (Year(Today)) + 1)
                                weekStart = (weekStart.AddDays(+3))
                            Else
                                strWeek_No = strWeek_No + 1
                                weekStart = GetWeekStartDate(strWeek_No, Year(Today))
                                weekStart = (weekStart.AddDays(+3))
                            End If
                        End If



                    End If
                    If IsNumeric(Trim(UltraGrid1.Rows(i).Cells(Z).Text)) Then
                        _Qty_INT = Trim(UltraGrid1.Rows(i).Cells(Z).Text)
                    Else
                    End If
                    nvcFieldList1 = "Insert Into T03Delivary_Quat_Flutter(T03RefNo,T03Line_Item,T03P4P,T03Liability,T03Qty,T03Qty_Int,T03WeekNo,T03Date,T03P4PConfirm,T03Status,T03FD_Status)" & _
                                                                    " values(" & _Parameter & ", '" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','" & _P4P & "','" & _Liability & "','" & Trim(UltraGrid1.Rows(i).Cells(Z).Text) & "'," & _Qty_INT & ",'" & strWeek_No & "','" & weekStart & "','N','A','N')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    '--------------------------------------------------------------------------------------------

                    If IsNumeric(Val(UltraGrid1.Rows(i).Cells(Z).Text)) Then
                        _Qty = _Qty_INT + _Qty
                    Else
                        If Trim(UltraGrid1.Rows(i).Cells(Z).Value) <> "" Then
                            _GrideStatus = False
                        End If
                    End If

                    Z = Z + 1
                Next

                _Qty = Val(UltraGrid1.Rows(i).Cells(3).Text) - _Qty

                nvcFieldList1 = "update T03Delivary_Quat_Flutter set T03Liability_Qty='" & Trim(UltraGrid1.Rows(i).Cells(Z).Text) & "' where T03RefNo='" & _Parameter & "' and T03Line_Item='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '-----------------------------------------------------------------------------------------------
                'UPDATE TBC RECORDS
                Dim strTBC As Date  '-------------------------------->>> Decluair the TBC Date
                Dim _TBC As Date
                strTBC = "9/1/" & Year(Today)
                If Today > strTBC Then
                    _TBC = "12/31/" & Year(Today)
                Else
                    _TBC = "12/31/" & CInt(Year(Today)) + 1
                End If
                If _GrideStatus = False Then

                    nvcFieldList1 = "select * from T04TBC_Records where T04Del_Ref='" & Trim(cboSO.Text) & "' and T04Ref=" & _Parameter & " and T04Line_Item='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' and T04Status='A'"
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(dsUser) Then
                        nvcFieldList1 = "update T04TBC_Records set T03Qty='" & _Qty & "' where T04Del_Ref='" & Trim(cboSO.Text) & "' and T04Ref=" & _Parameter & " and T04Line_Item='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' and T04Status='A'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into T04TBC_Records(T04Del_Ref,T04Ref,T04Line_Item,T03Qty,T04Date,T04Status)" & _
                                                                   " values(" & Trim(cboRef.Text) & "," & _Parameter & ", '" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','" & _Qty & "','" & _TBC & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                    End If
                Else

                    nvcFieldList1 = "update T01Delivary_Request set T01Status='C' where T01RefNo='" & Trim(cboSO.Text) & "' and T01Sales_Order='" & Trim(cboSO.Text) & "' and T01Line_Item='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' and T01Status='A'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                i = i + 1
            Next

            transaction.Commit()
            A = MsgBox("Are you sure you want to send e-mail to Merchant", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .........")
            If A = vbYes Then
                Call Send_Email() '------------------SENDING EMAIL

            End If
            common.ClearAll(OPR0)
            Clicked = ""
            OPR0.Enabled = True
            cmdSave.Enabled = False



            Call Load_Gride_SalesOrder()
            Call Search_Parameter()
            Call Load_Sales_Order()

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try

    End Sub

    Private Sub cboSO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSO.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Salrs_Order()
            cboLeadTime.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_Salrs_Order()
            cboLeadTime.ToggleDropdown()
        End If
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim SheetName = "SheetName"
        'Dim fileFullName As String = "C:\MyExcel.xlsx"
        'Dim SqlFilterExpression As String


        'Using xslDoc As SpreadsheetDocument = SpreadsheetDocument.Open(fileFullName, False)
        '    Dim theSheet As Sheet = xslDoc.WorkbookPart.Workbook.Sheets.FirstOrDefault(Function(SheetEl) CType(SheetEl, Sheet).Name = SheetName)
        '    Dim WorkSheetPart As WorksheetPart = xslDoc.WorkbookPart.GetPartById(theSheet.Id)

        '    Dim Filter = WorkSheetPart.Worksheet.Elements(Of AutoFilter).First

        '    If Not IsNothing(Filter) Then
        '        For Each FilterPerColumn In Filter.Elements(Of FilterColumn)()

        '            Dim CustomFilters = FilterPerColumn.Elements(Of CustomFilters).FirstOrDefault
        '            If Not IsNothing(CustomFilters) Then


        '                Dim filters = CustomFilters.Elements(Of CustomFilter)()

        '                'check if filter contient tow condition
        '                If filters.Count = 1 Then
        '                    Select Case CType(filters(0).Operator, FilterOperatorValues)
        '                        Case FilterOperatorValues.LessThan
        '                            '...Ect.
        '                    End Select

        '                Else
        '                    'check if condition joined by or/and
        '                    If CustomFilters.And Then
        '                        'and joined
        '                    Else
        '                        'or joined
        '                    End If
        '                End If

        '            End If
        '        Next
        '    End If
        'End Using
    End Sub

  
   
    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub cboSO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboSO.InitializeLayout

    End Sub
End Class