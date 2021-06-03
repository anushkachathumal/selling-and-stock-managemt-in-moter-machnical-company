'Imports Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Interop.Outlook
'Imports System.Data.SqlClient
Module Email

    'Function Send_Email_To_Projection(ByVal strProcument As String, ByVal strEmail As String)
    '    Dim Sql As String
    '    Dim con = New SqlConnection()
    '    con = DBEngin.GetConnection(True)
    '    ' con1 = DBEngin1.GetConnection(True)
    '    Dim M01 As DataSet
    '    Try
    '        Dim exc1 As New Microsoft.Office.Interop.Excel.Application
    '        Dim objApp As Object
    '        Dim objEmail As Object
    '        Dim xlRn As Microsoft.Office.Interop.Excel.Range
    '        Dim Connect As String
    '        Dim strbody As String

    '        Dim workbooks1 As Workbooks = exc1.Workbooks
    '        Dim workbook1 As _Workbook = workbooks1.Add(XlWBATemplate.xlWBATWorksheet)
    '        Dim sheets1 As Sheets = workbook1.Worksheets
    '        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
    '        Dim range1 As Range
    '        Dim A As String
    '        Dim I As Integer
    '        Dim x As Integer
    '        Dim Z As Integer

    '        objApp = CreateObject("Outlook.Application")
    '        objEmail = objApp.CreateItem(0)


    '        With objEmail

    '            .To = strEmail
    '            .Subject = "Yarn Request for /" & Trim(strSales_Order) & "-" & strLine_Item
    '            exc1.Visible = True

    '            Dim sheetsM As Sheets = workbook1.Worksheets
    '            Dim worksheet_Pro As _Worksheet = CType(sheetsM.Item(1), _Worksheet)
    '            worksheet_Pro.Rows(2).Font.size = 12
    '            worksheet_Pro.Rows(2).Font.Bold = True
    '            worksheet_Pro.Rows(2).rowheight = 50
    '            With worksheet_Pro
    '                .Cells(2, 1) = "Special Yarn"
    '                .Cells(2, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                ' worksheet1.Columns("A").ColumnWidth = 12
    '                .Range("A2:H2").MergeCells = True
    '                .Range("A2:H2").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '                .Range("A2:H2").Interior.Color = RGB(0, 112, 192)

    '                .Columns("A").ColumnWidth = 10
    '                .Columns("B").ColumnWidth = 10
    '                .Columns("C").ColumnWidth = 55
    '                .Columns("D").ColumnWidth = 10
    '                .Columns("E").ColumnWidth = 10
    '                .Columns("F").ColumnWidth = 10
    '                .Columns("G").ColumnWidth = 10
    '                .Columns("H").ColumnWidth = 18

    '            End With

    '            A = 97
    '            I = 0
    '            For I = 1 To 8
    '                With worksheet_Pro
    '                    .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "2", Chr(A) & "2").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

    '                End With
    '                A = A + 1
    '            Next
    '            x = 3
    '            worksheet_Pro.Rows(x).Font.size = 10
    '            worksheet_Pro.Rows(x).Font.Bold = True
    '            worksheet_Pro.Rows(x).rowheight = 30
    '            With worksheet_Pro
    '                .Cells(x, 1) = "Quality"
    '                .Cells(x, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                .Cells(x, 2) = "10 Class"
    '                .Cells(x, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                .Cells(x, 3) = "Yarn Description"
    '                .Cells(x, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

    '                .Cells(x, 4) = "Qty(Kg)"
    '                .Cells(x, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                .Cells(x, 5) = "Sales Order"
    '                .Cells(x, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                .Cells(x, 6) = "Line Item"
    '                .Cells(x, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                .Cells(x, 7) = "Capacity"
    '                .Cells(x, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                .Cells(x, 7).WrapText = True
    '                .Cells(x, 8) = "Weekly Consumption"
    '                .Cells(x, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                '.Cells(x, 9) = "Weekly Consumption"
    '                '.Cells(x, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '            End With

    '            A = 97
    '            I = 0
    '            For I = 1 To 8
    '                With worksheet_Pro
    '                    .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "3", Chr(A) & "3").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                    .Range(Chr(A) & "3", Chr(A) & "3").Interior.Color = RGB(169, 208, 142)

    '                    .Range(Chr(A) & "3", Chr(A) & "3").MergeCells = True
    '                    .Range(Chr(A) & "3", Chr(A) & "3").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '                End With
    '                A = A + 1
    '            Next

    '            Sql = "T14Sales_order='" & strSales_Order & "' AND T14Line_Item=" & strLine_Item & " AND T14Ref_no=" & Delivary_Ref & ""
    '            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarn_Booking", New SqlParameter("@cQryType", "YRNR"), New SqlParameter("@vcWhereClause1", Sql))
    '            Z = 0
    '            x = x + 1
    '            For Each DTRow3 As DataRow In M01.Tables(0).Rows
    '                worksheet_Pro.Rows(x).Font.size = 10
    '                worksheet_Pro.Rows(x).rowheight = 20
    '                With worksheet_Pro
    '                    .Cells(x, 1) = frmKnitting_Plan_WithTab.txtQuality.Text
    '                    .Cells(x, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                    .Cells(x, 2) = M01.Tables(0).Rows(Z)("T14Class")
    '                    .Cells(x, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


    '                    .Cells(x, 3) = M01.Tables(0).Rows(Z)("T14Yarn")
    '                    .Cells(x, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft


    '                    .Cells(x, 4) = M01.Tables(0).Rows(Z)("T14Qty")
    '                    .Cells(x, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                    range1 = .Cells(x, 4)
    '                    range1.NumberFormat = "0.00"

    '                    .Cells(x, 5) = M01.Tables(0).Rows(Z)("T14Sales_order")
    '                    .Cells(x, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                    .Cells(x, 6) = M01.Tables(0).Rows(Z)("T14Line_Item")
    '                    .Cells(x, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


    '                    .Cells(x, 7) = "Week " & " " & Trim(M01.Tables(0).Rows(Z)("T14Week"))
    '                    .Cells(x, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

    '                    .Cells(x, 8) = M01.Tables(0).Rows(Z)("T14Qty")
    '                    .Cells(x, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    '                    range1 = .Cells(x, 8)
    '                    range1.NumberFormat = "0.00"


    '                    A = 97
    '                    I = 0
    '                    For I = 1 To 8
    '                        With worksheet_Pro
    '                            .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                            .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                            .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                            .Range(Chr(A) & x, Chr(A) & x).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    '                            '.Range(Chr(A) & "3", Chr(A) & "3").Interior.Color = RGB(169, 208, 142)

    '                            .Range(Chr(A) & x, Chr(A) & x).MergeCells = True
    '                            .Range(Chr(A) & x, Chr(A) & x).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    '                        End With
    '                        A = A + 1
    '                    Next
    '                End With

    '                x = x + 1
    '                Z = Z + 1
    '            Next






    '            'strBody = "This is a test " & vbCrLf & vbCrLf & "Thanks Michael"
    '            '  RangetoHTML(xlRn)

    '            Connect = worksheet_Pro.Range("A1:H" & x - 1).Copy()
    '            xlRn = worksheet_Pro.Range("A1:H" & x + 1)
    '            'xlRn.Copy()

    '            '.HTMLBody = " Dear" & Trim(cboPlaner.Text) & "," & vbNewLine & _
    '            '              "Please Quote best possible delivery for below" & Chr(10) _
    '            '                                   & RangetoHTML(xlRn)

    '            .HTMLBody = "Dear " & Trim(strProcument) & ",<br>Please order below yarns " & RangetoHTML1(xlRn)

    '            .display()

    '        End With
    '        objEmail = Nothing
    '        objApp = Nothing

    '        DBEngin.CloseConnection(con)
    '        con.ConnectionString = ""
    '        con.close()
    '        ' worksheet1.Cells(4, 5) = _Fail_Batch
    '        'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
    '        ' MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
    '        ' MsgBox(_Fail_Batch)
    '    Catch returnMessage As EvaluateException
    '        If returnMessage.Message <> Nothing Then
    '            MessageBox.Show(returnMessage.Message)
    '            DBEngin.CloseConnection(con)
    '            con.ConnectionString = ""
    '            con.close()
    '        End If
    '    End Try
    'End Function


    'Function RangetoHTML1(ByVal rng As Microsoft.Office.Interop.Excel.Range)
    '    ' Changed by Ron de Bruin 28-Oct-2006
    '    ' Working in Office 2000-2010
    '    Dim fso As Object
    '    Dim ts As Object
    '    Dim TempFile As String
    '    ' Dim TempWB As Microsoft.Office.Interop.Excel.Workbook

    '    Dim exc As New Microsoft.Office.Interop.Excel.Application
    '    Dim TempWB1 As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
    '    Dim TempWB As Microsoft.Office.Interop.Excel._Workbook = TempWB1.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)

    '    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    '    'Copy the range and create a new workbook to past the data in
    '    rng.Copy()
    '    'TempWB = Microsoft.Office.Interop.Excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)




    '    With TempWB.Sheets(1)
    '        .Cells(1).PasteSpecial(Paste:=8)
    '        ' Microsoft.Office.Interop.Excel.XlPastef
    '        '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, , False, False)
    '        '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, , False, False)
    '        '.Cells(1).Select()
    '        'Application.CutCopyMode = False
    '        On Error Resume Next
    '        .DrawingObjects.Visible = True
    '        .DrawingObjects.Delete()
    '        On Error GoTo 0
    '    End With


    '    'Publish the sheet to a htm file
    '    With TempWB.PublishObjects.Add( _
    '         SourceType:=Microsoft.Office.Interop.Excel.XlSourceType.xlSourceRange, _
    '         Filename:=TempFile, _
    '         Sheet:=TempWB.Sheets(1).Name, _
    '         Source:=TempWB.Sheets(1).UsedRange.Address, _
    '         HtmlType:=Microsoft.Office.Interop.Excel.XlHtmlType.xlHtmlStatic)
    '        .Publish(True)
    '    End With

    '    'Read all data from the htm file into RangetoHTML
    '    fso = CreateObject("Scripting.FileSystemObject")
    '    ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    '    RangetoHTML1 = ts.ReadAll
    '    ts.Close()
    '    RangetoHTML1 = Replace(RangetoHTML1, "align=center x:publishsource=", _
    '                          "align=left x:publishsource=")

    '    'Close TempWB
    '    TempWB.Close(SaveChanges:=False)

    '    'Delete the htm file we used in this function
    '    Kill(TempFile)

    '    ts = Nothing
    '    fso = Nothing
    '    TempWB = Nothing
    'End Function

End Module
