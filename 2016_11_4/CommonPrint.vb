Imports System.Data.SqlClient
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.MDIMain

Public Class CommonPrint
    'Public Shared Function PrintWinnerVoucher(ByVal MaxVouno As String)
    '    Dim dsUser As DataSet
    '    Dim SQL As String
    '    Dim con = New SqlConnection()
    '    Dim cSqlStr As String
    '    Dim cDetailLine() As Object
    '    Dim i, j, k As Integer
    '    Dim VSaveTime As Date
    '    Dim VSaveUser As String
    '    Dim VWorkStID As String
    '    Dim nLineCntr As Integer
    '    Dim RetVal As Long
    '    Dim TicType As String
    '    Dim cvMax As String
    '    Dim vMax As String
    '    Dim PrintAgentInfo As Boolean = False


    '    vMax = "2595198"
    '    cvMax = "0"

    '    ReDim cDetailLine(0 To 18)
    '    For i = 0 To 18
    '        cDetailLine(i) = Space(87)
    '    Next i

    '    cDetailLine(1) = Space(26) & "Development Lotteries Board"
    '    cDetailLine(2) = Space(29) & "Prize Payment Voucher"

    '    FileClose(1)

    '    lMoreThanOnepage = False
    '    cDetailLine(0) = Chr(27) & Chr(67) + Chr(32) + Chr(27) & "P" & Chr(13) & Chr(27) & Chr(18) & Chr(13) & Chr(27) & Chr(48) & Chr(13) & Chr(27) & "M"
    '    cDetailLine(1) = Space(26) & "Development Lotteries Board"
    '    cDetailLine(2) = Space(29) & "Prize Payment Voucher"
    '    cDetailLine(18) = Chr(12)
    '    'imti
    '    con = DBEngin.GetConnection()
    '    SQL = "SELECT T37TicketType FROM T37CreditVoucher WHERE T37PayVoucher='" & vMax & "'"
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)

    '    If dsUser.Tables(0).Rows.Count > 0 Then

    '        If dsUser.Tables(0).Rows(0)("T37TicketType") = "I" Then
    '            TicType = "INSTANT"
    '        Else
    '            TicType = "DRAW"
    '        End If
    '    End If
    '    'imti end

    '    SQL = "SELECT T20Date,T20PaymentNo,T20Method,T20Amount,T20WorkStationId,T20User,T20Time FROM T20PaymentHeader WHERE T20PaymentNo = " & vMax
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)

    '    If dsUser.Tables(0).Rows.Count > 0 Then
    '        === For Print Footer ======================
    '        VSaveTime = dsUser.Tables(0).Rows(0)("T20Time")
    '        VSaveUser = dsUser.Tables(0).Rows(0)("T20User")
    '        VWorkStID = dsUser.Tables(0).Rows(0)("T20WorkStationId")
    '        ===========================================
    '        Mid(cDetailLine(3), 1, 29) = "Date       : " & dsUser.Tables(0).Rows(0)("T20Date")

    '        Mid(cDetailLine(3), 53, 21) = "Voucher No : " & Str(dsUser.Tables(0).Rows(0)("T20PaymentNo"))

    '        Mid(cDetailLine(7), 1, 29) = "Amount     : " & FormatCurrency(dsUser.Tables(0).Rows(0)("T20Amount"))
    '        Mid(cDetailLine(7), 35, 21) = "Ticket (S)     : " & "123"
    '        Mid(cDetailLine(4), 68, 21) = "------------------"

    '        Mid(cDetailLine(4), 53, 7) = TicType
    '        Mid(cDetailLine(5), 65, 21) = "    Credit Slip      "
    '        Mid(cDetailLine(8), 1, 85) = "------------------------------------------------------------------------------------"
    '        Mid(cDetailLine(17), 1, 85) = "------------------------------------------------------------------------------------"
    '        Mid(cDetailLine(9), 65, 22) = "       Amount        "
    '        Mid(cDetailLine(12), 68, 21) = "--------------------"
    '        Mid(cDetailLine(16), 68, 21) = "--------------------"

    '        For i = 1 To 3
    '            Mid(cDetailLine(4 + i), 67, 1) = "|"
    '            Mid(cDetailLine(4 + i), 85, 1) = "|"
    '        Next i
    '        ***************************************
    '        SQL = dsUser.Tables(0).Rows(0)("T20User") & ":" & dsUser.Tables(0).Rows(0)("T20WorkStationId") & ":" & dsUser.Tables(0).Rows(0)("T20Time") & " >>PRINT ON " & Now & " BY:" & "ADMIN" '<--------- MUST CALL USER NAME
    '        If PrintAgentInfo Then
    '            Dim strAgInf As String
    '            strAgInf = "AGENT : " & Trim(dcAgent.Text) & " " & Trim(txtAgent_Name.Text) & "   " & "kURUMEGALA" '<---------------MUST CALL THE DISTRICAT NAME
    '            Mid(cDetailLine(15), 1, Len(strAgInf)) = strAgInf
    '        End If
    '        =========================================================================================
    '        Add01()
    '        ====================================
    '        If dsUser.Tables(0).Rows(0)("T20Method") = "003" Then Mid(cDetailLine(7), 68, 16) = (Space(16) & FormatCurrency(dsUser.Tables(0).Rows(0)("T20Amount")))
    '        If dsUser.Tables(0).Rows(0)("T20Method") = "003" Then
    '            SQL = "SELECT T37CreditVoucher FROM T37CreditVoucher WHERE T37PayVoucher =" & vMax
    '            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '            If dsUser.Tables(0).Rows.Count > 0 Then
    '                Mid(cDetailLine(6), 69, 7) = dsUser.Tables(0).Rows(0)("T37creditvoucher")
    '                MsgBox(dsUser.Tables(0).Rows(0)("T37creditvoucher"))
    '            End If
    '        End If
    '    End If
    '    ***************************************SET TOTAL TICKETS
    '    SQL = "SELECT COUNT(T21PaymentNo) AS TOTALTKT FROM T21Prizedetail WHERE T21PaymentNo='" & vMax & "' GROUP BY T21PaymentNo"
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '    If dsUser.Tables(0).Rows.Count > 0 Then
    '        Mid(cDetailLine(7), 35, 21) = "Ticket (S)     : " & dsUser.Tables(0).Rows(0)("TOTALTKT")
    '    End If
    '    ************************************************* SET WINNERS NAME
    '    SQL = "SELECT M03AgentNo,M03Name FROM T20PaymentHeader INNER JOIN M03Agent ON M03AgentNo=T20Agent_NonAg WHERE T20PaymentNo ='" & vMax & "'"
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '    If dsUser.Tables(0).Rows.Count > 0 Then

    '        If dsUser.Tables(0).Rows(0)("M03AgentNo") = "000000" Then
    '            SQL = "SELECT * FROM T20PaymentHeader INNER JOIN T35Winners ON T20PaymentNo = T35PaymentNo WHERE T20PaymentNo =" & vMax
    '            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '            If dsUser.Tables(0).Rows.Count > 0 Then
    '                Mid(cDetailLine(5), 1, 65) = "Name : " & (dsUser.Tables(0).Rows(0)("T35Winner") & Space(58))
    '                Mid(cDetailLine(5), 1, 65) = "Name : " & ToString.PadRight(dsUser.Tables(0).Rows(0)("T35Winner") & Space(58))
    '                Mid(cDetailLine(6), 1, 65) = "NIC NO :" & CStr(dsUser.Tables(0).Rows(0)("T35NicNo"))
    '            End If
    '        Else
    '            Mid(cDetailLine(5), 1, 65) = "Name : " & CStr(dsUser.Tables(0).Rows(0)("M03AgentNo")) & " " & (dsUser.Tables(0).Rows(0)("M03Name") & Space(50))
    '        End If

    '        Mid(cDetailLine(6), 1, 65) = "Address :" &

    '    End If
    '    '********************************************************
    '    SQL = "SELECT SUM(T21Amount / M20Amount) AS nTkt FROM T21Prizedetail " & _
    '              "INNER JOIN M20Prize ON T21TierCode=M20PriceCode " & _
    '              "WHERE Left(T21TierCode,1)<>'X' AND T21PaymentNo = " & vMax
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '    If Not IsDBNull(dsUser.Tables(0).Rows(0)("nTkt")) Then Mid(cDetailLine(7), 30, 20) = "Ticket (s): " & IIf(Int(dsUser.Tables(0).Rows(0)("nTkt")) = 0, "1", Str(Int(dsUser.Tables(0).Rows(0)("nTkt"))))

    '    If rsTemp.state = 1 Then rsTemp.Close()
    '    SQL = "SELECT * FROM T21Prizedetail WHERE T21PaymentNo = " & vMax & " and T21Amount > 99 "
    '    nLineCntr = 9
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '    If dsUser.Tables(0).Rows.Count > 0 Then
    '        Mid(cDetailLine(4), 25, 20) = "Game  : " & dsUser.Tables(0).Rows(0)("T21ItemNo")

    '    Else
    '        If T20PaymentHeader.state = 1 Then T20PaymentHeader.Close()
    '        SQL = "SELECT * FROM T20PaymentHeader " & _
    '                                        " INNER JOIN M33PaymentMethod ON T20Method = M33MethodNo " & _
    '                                        " INNER JOIN M06CodeMaintenance ON T20TransactCode=M06Code " & _
    '                                        " WHERE T20PaymentNo = '" & vMax & "'"
    '        dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '        If dsUser.Tables(0).Rows.Count > 0 Then
    '            cDetailLine(10) = "Transaction    : " & RTrim(dsUser.Tables(0).Rows(0)("M33Description"))
    '            cDetailLine(11) = "Payment Method : " & RTrim(dsUser.Tables(0).Rows(0)("M33Description"))
    '            cDetailLine(12) = "Referance      : " & RTrim(dsUser.Tables(0).Rows(0)("T20Chq_VouchNo"))
    '        End If
    '    End If
    '    '**********************************************************8
    '    SQL = "SELECT sum(T21Amount)AS TOT,(T21TierCode),count(T21TierCode)AS Tcount FROM T21Prizedetail WHERE T21PaymentNo='" & vMax & "' GROUP BY T21TierCode"
    '    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
    '    i = 0
    '    Dim X As Integer
    '    For i = 0 To dsUser.Tables(0).Rows.Count + 1
    '        j = 0
    '        For j = 1 To 3

    '            If i >= dsUser.Tables(0).Rows.Count Then Exit For
    '            Mid(cDetailLine(nLineCntr), 26 * j - 24, 26) = (dsUser.Tables(0).Rows(i)("T21TierCode")) & "X" & (Space(2) & dsUser.Tables(0).Rows(i)("Tcount"))
    '            Mid(cDetailLine(nLineCntr), 28 * j - 15, 12) = FormatCurrency(dsUser.Tables(0).Rows(i)("Tot"))
    '            i = i + 1

    '        Next
    '        nLineCntr = nLineCntr + 1
    '        If i >= dsUser.Tables(0).Rows.Count Then Exit For
    '        i = i - 1

    '    Next

    '    cDetailLine(17) = Space(0) & VSaveUser & " : " & VWorkStID & " " & VSaveTime & " >> PRINT ON " & Now & " BY : " & LoggedUser


    '    cSqlStr = "TYPE LMXPRINT.PMT > PRN"
    '    FileOpen(1, Application.StartupPath + "\LMXPRINT.PMT", OpenMode.Output)
    '    FileOpen(1, Application.StartupPath & "\print.bat", OpenMode.Output)
    '    For i = 0 To 18
    '        PrintLine(1, cDetailLine(i))
    '    Next i
    '    FileClose(1)


    '    cSqlStr = "TYPE LMXPRINT.PMT > PRN"
    '    FileOpen(1, Application.StartupPath & "\print.bat", OpenMode.Output)
    '    PrintLine(1, cSqlStr)
    '    FileClose(1)
    '     ===========================================
    '    RetVal = Shell(Application.StartupPath & "\print.bat")


    ' End Function
End Class
