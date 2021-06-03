
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO

Public Class frmUpload
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableEral
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

    Private Sub frmUpload_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Upload_EralFile()
        Dim strFileName As String
        Dim _LotNo As String
        Dim _Batchwhight As Double
        Dim _MCNo As String
        Dim _ProgrammeNo As String
        Dim _ProgrameName As String
        Dim _StratDate As Date
        Dim _StartTime As String
        Dim _EndDate As Date
        Dim _EndTime As String
        Dim _Quality As String
        Dim _ShadeCode As String
        Dim _ShadeType As String
        Dim _Shade As String
        Dim _StandedTime As String
        Dim _LotType As String
        Dim _TotalHR As Integer
        Dim _QltyGroup As String
        Dim _Rating As String
        Dim _Delete As String
        Dim _Weight As Double
        Dim _WeekNo As Integer
        Dim _Shift As Integer
        Dim _WeekDis As String
        Dim _StDate As Date
        Dim _EDDate As Date
        Dim _Taken As String
        Dim _Status As String
        Dim hh As Integer
        Dim mm As Integer


        Dim nvcFieldList1 As String
        Dim Sql As String
        Dim X1 As Integer
        '  _TotalHR = "00:00"
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Dim M04Lot As DataSet
        Dim nvcVccode As String

        Dim ncQryType As String
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim hh1 As Integer
        Dim mm1 As Integer
        Dim _TimeDifferance1 As Date
        Dim n_year As Integer


        strFileName = ConfigurationManager.AppSettings("FilePath") + "\download1.csv"
        Dim theTextFieldParser As FileIO.TextFieldParser

        theTextFieldParser = My.Computer.FileSystem.OpenTextFieldParser(strFileName)
        theTextFieldParser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
        theTextFieldParser.Delimiters = New String() {","}
        Dim X11 As Integer
        ' Declare a variable named currentRow of type string array.
        Dim currentRow() As String
        Try

            X11 = 0
            While Not theTextFieldParser.EndOfData
                ' Try
                ' Read the fields on the current line
                ' and assign them to the currentRow array variable.
                currentRow = theTextFieldParser.ReadFields()

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                ' Use the currentField variable to loop
                ' through fields in the currentRow.
                i = 0

                'If X11 = 335 Then
                '    MsgBox("")
                'End If
                For Each currentField In currentRow
                    ' Add the the currentField (a string)
                    ' to the demoLstBox items.
                    If i = 0 Then
                        _LotNo = currentField
                    ElseIf i = 1 Then
                        _MCNo = currentField
                    ElseIf i = 2 Then
                        _ProgrammeNo = currentField
                    ElseIf i = 3 Then
                        _ProgrameName = currentField
                    ElseIf i = 4 Then
                        _LotType = currentField
                    ElseIf i = 5 Then
                        _StandedTime = currentField
                    ElseIf i = 6 Then
                        '  MsgBox(Format(CDate(currentField), "mm/dd/yyyy"))
                        'MsgBox(VB6.Format(currentField, "mm/dd/yyyy"))
                        Dim D_U As String

                        'If Microsoft.VisualBasic.Left(currentField, 2) = "04" Then
                        '    Dim d As DateTime
                        'd = currentField
                        'MsgBox(VB6.Format((currentField), "mm/DD/yyyy"))
                        ' D_U = Microsoft.VisualBasic.Left(currentField, 5)
                        D_U = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(currentField, 5), 2)
                        _StratDate = D_U & "/" & Microsoft.VisualBasic.Left(currentField, 2) & "/" & Microsoft.VisualBasic.Right(currentField, 4)
                        '_StratDate = (VB6.Format(currentField, "mm/dd/yyyy"))
                    ElseIf i = 7 Then
                        _StartTime = currentField
                    ElseIf i = 8 Then
                        Dim D_U As String

                        'If Microsoft.VisualBasic.Left(currentField, 2) = "04" Then
                        '    Dim d As DateTime
                        'd = currentField
                        'MsgBox(VB6.Format((currentField), "mm/DD/yyyy"))
                        D_U = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(currentField, 5), 2)
                        _EndDate = D_U & "/" & Microsoft.VisualBasic.Left(currentField, 2) & "/" & Microsoft.VisualBasic.Right(currentField, 4)
                        ' _EndDate = (VB6.Format(currentField, "mm/dd/yyyy"))
                    ElseIf i = 9 Then
                        _EndTime = currentField
                    ElseIf i = 10 Then

                        If Len(currentField) = 5 Then
                            hh = Microsoft.VisualBasic.Left(currentField, 2)
                            mm = Microsoft.VisualBasic.Right(currentField, 2)
                        ElseIf Len(currentField) = 4 Then
                            hh = Microsoft.VisualBasic.Left(currentField, 1)
                            mm = Microsoft.VisualBasic.Right(currentField, 2)
                        End If
                        'hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Hour
                        'mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Minute

                        _TotalHR = (hh * 60) + mm
                    ElseIf i = 11 Then
                        _Quality = currentField
                    ElseIf i = 12 Then
                        _QltyGroup = currentField
                    ElseIf i = 13 Then
                        _ShadeCode = currentField
                    ElseIf i = 14 Then
                        _Shade = currentField
                    ElseIf i = 15 Then
                        _ShadeType = currentField
                    ElseIf i = 16 Then
                        '   _Rating = currentField
                        '  _ShadeType = currentField
                    ElseIf i = 17 Then
                        _Weight = currentField
                        'ElseIf i = 18 Then
                        '    _Weight = currentField
                    End If


                   




                    '------------------------------------------------------------------------------------------------

                    i = i + 1
                Next

                'If _TotalHR = "24:23" Then
                '    MsgBox("")
                'End If
                'MsgBox(DateTime.Parse(_EndDate).DayOfWeek())
                Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_EndDate)
                ' dayOfWeek.ToString() would return "Sunday" but it's an enum value,
                ' the correct dayname can be retrieved via DateTimeFormat.
                ' Following returns "Sonntag" for me since i'm in germany '
                Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)


                ' MsgBox(dayName)
                If dayName = "Sunday" Then
                    Dim N_Date1 As Date
                    N_Date1 = CDate(_EndDate).AddDays(-1)
                    _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                    _WeekDis = "Week" & CStr(_WeekNo)
                    n_year = Microsoft.VisualBasic.Year(N_Date1)
                Else
                    If _EndDate = "12/31/" & Microsoft.VisualBasic.Year(_EndDate) Then
                        Dim N_Date1 As Date
                        N_Date1 = CDate(_EndDate).AddDays(+1)
                        _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_year = Microsoft.VisualBasic.Year(N_Date1)
                    Else

                        _WeekNo = DatePart(DateInterval.WeekOfYear, _EndDate)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_year = Microsoft.VisualBasic.Year(_EndDate)
                    End If
                End If
                If DateAndTime.TimeValue(_StartTime) <= "07:30:00 AM" And DateAndTime.TimeValue(_EndTime) >= "07:30:00 PM" Then
                    _Shift = 1

                Else
                    _Shift = 2
                End If


                _StDate = (_StratDate) & " " & (_StartTime)
                _EDDate = (_EndDate) & " " & (_EndTime)
                ' _Taken = System.DateTime.From(CDate(_EDDate).ToOADate) - System.DateTime.FromOADate(CDate(_StDate).ToOADate)

                '  hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Hour
                ' mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Minute

                ' _Taken = DateAndTime.TimeValue(System.DateTime.FromOADate(CDate(_EDDate).ToOADate - CDate(_StDate).ToOADate))

                _Taken = _TotalHR

                '  hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(_StandedTime).Subtract(_Taken).Hours
                ' mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(_StandedTime).Subtract(_Taken).Minutes
                '  _TimeDifferance = hh.ToString.PadLeft(2, CChar("0")) & ":" & mm.ToString.PadLeft(2, CChar("0"))
                '_TimeDifferance1 = hh1.ToString.PadLeft(2, CChar("0")) & ":" & mm1.ToString.PadLeft(2, CChar("0"))
                If IsNumeric(_LotNo) Then
                    If Microsoft.VisualBasic.Len(_LotNo) = 6 Then
                        _Status = "L"
                    Else
                        _Status = "A"
                    End If
                Else
                    Sql = "select M05Code from M05DownTime where M05Code='" & Trim(_LotNo) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M01) Then
                        _Status = "D"
                    Else
                        _Status = "I"
                    End If
                End If

                ncQryType = "ADD"
                ' vMax = Get_highestVouNumber()

                'If _ProgrameName <> "" Then
                _ProgrameName = myReplace.do(_ProgrameName, types.one)
                ' End If
                'nvcFieldList1 = "(M04Ref," & "M04Week," & "M04WeekNo," & "M04Lotno," & "M04Batchwt," & "M04Machine_No," & "M04Program," & "M04Type," & "M04STD," & "M04DateIn," & "M04TimeIn," & "M04Date_Out," & "M04Time_Out," & "M04Taken," & "M04Quality," & "M04Shade_Code," & "M04STD2," & "M04Shift," & "M04Status," & "M04Year," & "M04Month," & "M04ProgrameType," & "M04ETime," & "M04Shade_Type) " & "values('" & vMax & "','" & _WeekNo & "','" & _WeekDis & "','" & _LotNo & "','" & _Weight & "','" & _MCNo & "','" & _ProgrammeNo & "','" & _ProgrameName & "','" & _StandedTime & "','" & _StratDate & "','" & _StartTime & "','" & _EndDate & "','" & _EndTime & "'," & _Taken & ",'" & _Quality & "','" & _ShadeCode & "','" & _StDate & "','" & _Shift & "','" & _Status & "'," & n_year & ",'" & Microsoft.VisualBasic.Month(_EndDate) & "','" & _LotType & "','" & _EDDate & "','" & _ShadeType & "')"
                'up_GetSetM04Lot(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                If _Status = "D" Then
                    Dim _Taken_Min As Integer

                    'mm1 = 0
                    'hh1 = 0
                    '' MsgBox(Len(_Taken))
                    'If Len(_Taken) = 10 Then
                    '    mm1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), (Len(CInt(_Taken)) - 2)), 2))

                    '    hh1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), Len(CInt(_Taken))), (Len(CInt(_Taken)) - 9)))
                    'Else
                    '    mm1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), (Len(CInt(_Taken)) - 3)), 2))
                    '    hh1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), Len(CInt(_Taken))), (Len(CInt(_Taken)) - 9)))
                    'End If
                    '_Taken_Min = 0
                    '_Taken_Min = (hh1 * 60)
                    '_Taken_Min = _Taken_Min + mm1
                    _Taken_Min = _Taken

                    Dim n_Time As Date

                    n_Time = _EndDate & " " & _EndTime

                    'Sql = "select T01Taken from T01Down_Time where T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                    'T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    'If isValidDataset(T01) Then

                    '    nvcFieldList1 = "UPDATE T01Down_Time SET T01Taken=T01Taken +'" & _Taken_Min & "' WHERE T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                    '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'Else

                    '    nvcFieldList1 = "Insert Into T01Down_Time(T01Date,T01Week,T01WeekNo,T01Down_Time,T01Machine,T01Taken,T01Month,T01Year)" & _
                    '                             " values('" & n_Time & "', '" & _WeekDis & "'," & _WeekNo & ",'" & _LotNo & "','" & _MCNo & "','" & _Taken_Min & "','" & Microsoft.VisualBasic.Month(_EndDate) & "','" & n_year & "')"
                    '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    'End If
                End If
                '-----------------------------------------------------------------------------Clear Variable
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Lot No") = _LotNo
                newRow("Machine No") = _MCNo
                newRow("Programe No") = _ProgrammeNo
                newRow("Programe Name") = _ProgrameName
                newRow("Lot Type") = _LotType
                newRow("Standed Time") = _StandedTime
                newRow("Start Date") = _StratDate
                newRow("Start Time") = _StartTime
                newRow("End Date") = _EndDate
                newRow("End Time") = _EndTime
                newRow("Total Hour") = _TotalHR
                newRow("Quality") = _Quality
                newRow("Quality Group") = _QltyGroup
                newRow("Shade Code") = _ShadeCode
                newRow("Shade") = _Shade
                newRow("Shade Type") = _ShadeType
                newRow("Weight") = _Weight

                c_dataCustomer1.Rows.Add(newRow)

                pbCount.Value = pbCount.Value + 1
                lblDis.Text = _LotNo

                _LotNo = ""
                _MCNo = ""
                _ProgrammeNo = ""
                _ProgrameName = ""
                _Quality = ""
                _ShadeCode = ""
                _ShadeType = ""
                _StandedTime = ""
                _LotType = ""
                _TotalHR = 0
                _QltyGroup = ""
                _Rating = ""
                _Delete = ""
                _Weight = 0
                mm = 0
                hh = 0
                X11 = X11 + 1
            End While

            'MsgBox("Record update Successfully")
            'transaction.Commit()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            '  MsgBox(X11)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            theTextFieldParser.Close()
        End Try


    End Function

    Function FileUpload()
        Me.Cursor = Cursors.WaitCursor
        Dim count As Integer
        count = 0
        Dim obj As IO.StreamReader
        Dim A As String
        A = ConfigurationManager.AppSettings("FilePath") + "\download1.csv"
        obj = New IO.StreamReader(A)
        ''loop through the file until the end
        Do Until obj.ReadLine Is Nothing
            count = count + 1
        Loop
        ''close file and show count
        obj.Close()
        A = ConfigurationManager.AppSettings("FilePath") + "\download2.csv"
        obj = New IO.StreamReader(A)
        ''loop through the file until the end
        Do Until obj.ReadLine Is Nothing
            count = count + 1
        Loop
        ''close file and show count
        obj.Close()
        ' MsgBox(count)
        pbCount.Minimum = 0
        lblDis.Text = ""
        pbCount.Value = pbCount.Minimum
        pbCount.Maximum = count

        Upload_EralFile()
        Upload_EralFile2()
        Me.Cursor = Cursors.Arrow
        cmdEdit.Enabled = True
    End Function

    Function Upload_DyePrograme_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim X As Integer


        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M09Quality As DataSet
        Dim StringWithQuotes As String

        Try
            X = 0
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\dyephdr.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X = 5727 Then
                    '      MsgBox("")
                End If
                If Trim(fields(0)) <> "" Then
                    StringWithQuotes = Trim(fields(2))
                    StringWithQuotes = Replace(StringWithQuotes, Chr(34), "")
                    StringWithQuotes = Replace(StringWithQuotes, Chr(39), "")
                    nvcFieldList1 = "select * from M22Dye_Material where M22Group='" & Trim(fields(0)) & "' and M22MachineNo='" & Trim(fields(1)) & "'"
                    M09Quality = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M09Quality) Then
                        nvcFieldList1 = "update M22Dye_Material set M22Programe_Name='" & StringWithQuotes & "' where M22Group='" & Trim(fields(0)) & "' and M22MachineNo='" & Trim(fields(1)) & "'"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into M22Dye_Material(M22Group,M22MachineNo,M22Programe_Name)" & _
                                                                " values('" & Trim(fields(0)) & "', '" & Trim(fields(1)) & "','" & StringWithQuotes & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                End If
                X = X + 1
            Next

            transaction.Commit()
            DBEngin.CloseConnection(connection)

        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)
            MsgBox(X)
        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Function

    Function Upload_Quality_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim X As Integer


        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M09Quality As DataSet
        Dim StringWithQuotes As String

        Try
            X = 0
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\quality.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X = 5727 Then
                    '      MsgBox("")
                End If
                If Trim(fields(0)) <> "" Then
                    StringWithQuotes = Trim(fields(1))
                    StringWithQuotes = Replace(StringWithQuotes, Chr(34), "")
                    StringWithQuotes = Replace(StringWithQuotes, Chr(39), "")
                    nvcFieldList1 = "select * from M09Quality where M09Code='" & Trim(fields(0)) & "'"
                    M09Quality = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M09Quality) Then

                    Else
                        nvcFieldList1 = "Insert Into M09Quality(M09Code,M09Dis,M09Type,M09RTR,M09MKG)" & _
                                                                " values('" & Trim(fields(0)) & "', '" & StringWithQuotes & "','" & Trim(fields(2)) & "','" & Trim(fields(3)) & "','" & Trim(fields(4)) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                End If
                X = X + 1
            Next

            transaction.Commit()
            DBEngin.CloseConnection(connection)

        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)
            MsgBox(X)
        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Function

    Function Upload_Lot_txtfile()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Dim X11 As Integer
        Dim _BatchNo As String
        Dim _Weight As Double
        Dim _Machine As String
        Dim _ProGrameNo As String
        Dim _LotType As String
        Dim _StartDate As Date
        Dim _StartTime As Date
        Dim _EndDate As Date
        Dim _EndTime As Date
        Dim _Quality As String
        Dim _ShadeCode As String
        Dim _ShadeType As String
        Dim _SDTime As Integer
        Dim _Shift As Integer
        Dim _StDate As Date
        Dim _EDDate As Date
        Dim _Status As String
        Dim Sql As String
        Dim _QltyGroup As String

        Dim M01 As DataSet
        Dim _ProgrameName As String
        Dim _TotalHR As Double
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T03 As DataSet
        Dim _Shade As String


        Dim X As Integer
        Dim _Month As String

        Try
            X = 0
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\lots.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                If Trim(fields(0)) <> "" Then
                    If Trim(fields(18)) = "" Or Trim(fields(18)) = "0" Then
                        Exit For
                    End If

                    'If X = 120 Then
                    '    '  MsgBox("")
                    'End If
                    _BatchNo = Trim(fields(0))
                    _Weight = Trim(fields(1))
                    _Machine = Trim(fields(9))
                    _ProGrameNo = Trim(fields(10))
                    _LotType = Trim(fields(12))
                    _Month = (Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(15)), 5), 2))
                    _StartDate = _Month & "/" & Microsoft.VisualBasic.Left(Trim(fields(15)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(15)), 4)
                    _StartTime = Trim(fields(16))
                    _Month = (Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(17)), 5), 2))
                    _EndDate = _Month & "/" & Microsoft.VisualBasic.Left(Trim(fields(17)), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(17)), 4)
                    _EndTime = Trim(fields(18))
                    _Quality = Trim(fields(50))
                    _ShadeCode = Trim(fields(51))
                    _ShadeType = Trim(fields(52))
                    _SDTime = Trim(fields(58))

                    If _LotType <> "" Then
                    Else
                        _LotType = "I"
                    End If
                    Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                    Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_EndDate)
                    ' dayOfWeek.ToString() would return "Sunday" but it's an enum value,
                    ' the correct dayname can be retrieved via DateTimeFormat.
                    ' Following returns "Sonntag" for me since i'm in germany '
                    Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)
                    Dim _WeekNo As Integer
                    Dim _WeekDis As String
                    Dim n_year As Integer


                    ' MsgBox(dayName)
                    If dayName = "Sunday" Then
                        Dim N_Date1 As Date
                        N_Date1 = CDate(_EndDate).AddDays(-1)
                        _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_year = Microsoft.VisualBasic.Year(N_Date1)
                    Else
                        If _EndDate = "12/31/" & Microsoft.VisualBasic.Year(_EndDate) Then
                            Dim N_Date1 As Date
                            N_Date1 = CDate(_EndDate).AddDays(+1)
                            _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                            _WeekDis = "Week" & CStr(_WeekNo)
                            n_year = Microsoft.VisualBasic.Year(N_Date1)
                        Else

                            _WeekNo = DatePart(DateInterval.WeekOfYear, _EndDate)
                            _WeekDis = "Week" & CStr(_WeekNo)
                            n_year = Microsoft.VisualBasic.Year(_EndDate)
                        End If
                    End If
                    If DateAndTime.TimeValue(_EndTime) <= "07:30:00 PM" Then
                        _Shift = 1

                    Else
                        _Shift = 2
                    End If


                    _StDate = (_StartDate) & " " & (_StartTime)
                    _EDDate = (_EndDate) & " " & (_EndTime)

                    If IsNumeric(_BatchNo) Then
                        If Microsoft.VisualBasic.Len(_BatchNo) = 6 Then
                            _Status = "L"
                        Else
                            _Status = "A"
                        End If
                    Else
                        Sql = "select M05Code from M05DownTime where M05Code='" & Trim(_BatchNo) & "'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(M01) Then
                            _Status = "D"
                        Else
                            _Status = "I"
                        End If
                    End If
                    'Chech the Programe Name
                    'Dim ReadRecord As StreamReader
                    'Dim strFileName1 As String
                    'Dim X1 As Integer
                    'Dim _MCGroup As String

                    'X1 = 0
                    'strFileName1 = ConfigurationManager.AppSettings("FilePath") + "\dyephdr.txt"
                    'ReadRecord = File.OpenText(strFileName1)
                    'For Each line1 As String In System.IO.File.ReadAllLines(strFileName1)
                    '    Dim fields1() As String = line1.Split(vbTab)
                    '    _MCGroup = Trim(fields1(0))
                    '    Sql = "select * from T03Machine where T03Code='" & Trim(_Machine) & "' and T03Group='" & _MCGroup & "'"
                    '    T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    '    If isValidDataset(T03) Then
                    '        _ProgrameName = Trim(fields1(3))
                    '        Exit For
                    '    End If
                    '    _MCGroup = ""
                    '    X1 = X1 + 1
                    'Next

                    Dim startTime As New DateTime '(Year(_EDDate) & "," & Month(_EDDate) & "," & Microsoft.VisualBasic.Day(_EDDate) & "," & Hour(_EDDate) & "," & Minute(_EDDate))     ' 10:30 AM today
                    Dim endTime As New DateTime '(Year(_StDate) & "," & Month(_StDate) & "," & Microsoft.VisualBasic.Day(_StDate) & "," & Hour(_StDate) & "," & Minute(_StDate))

                    startTime = _StDate
                    endTime = _EDDate
                    Dim duration As TimeSpan = endTime - startTime
                    _TotalHR = (duration.Days * 24) * 60
                    _TotalHR = _TotalHR + (duration.Hours * 60)
                    _TotalHR = _TotalHR + duration.Minutes
                    If _TotalHR > 0 Then
                        _TotalHR = _TotalHR / 60
                    End If
                    newRow("Lot No") = _BatchNo
                    newRow("Machine No") = _Machine
                    newRow("Programe No") = _ProGrameNo
                    newRow("Programe Name") = _ProgrameName
                    newRow("Lot Type") = _LotType
                    newRow("Standed Time") = _SDTime
                    newRow("Start Date") = Year(_StartDate) & "-" & Month(_StartDate) & "-" & Microsoft.VisualBasic.Day(_StartDate)
                    newRow("Start Time") = _StartTime
                    newRow("End Date") = Year(_EndDate) & "-" & Month(_EndDate) & "-" & Microsoft.VisualBasic.Day(_EndDate)
                    newRow("End Time") = _EndTime
                    newRow("Total Hour") = _TotalHR
                    newRow("Quality") = _Quality
                    newRow("Quality Group") = _QltyGroup
                    newRow("Shade Code") = _ShadeCode
                    newRow("Shade") = _Shade
                    newRow("Shade Type") = _ShadeType
                    newRow("Weight") = _Weight

                    c_dataCustomer1.Rows.Add(newRow)

                    _ProgrameName = ""
                    _Weight = 0
                    _Shade = ""
                    _ShadeCode = ""
                    _ShadeType = ""
                    _Shift = 0
                    _QltyGroup = ""
                    _Quality = ""
                    _TotalHR = 0
                    _ProGrameNo = ""
                    _Machine = ""
                    '_MCGroup = ""
                    _BatchNo = ""

                    'pbCount.Value = pbCount.Value + 1
                    lblDis.Text = _BatchNo
                End If
                ' UltraGrid1.Refresh()
                X = X + 1
            Next


        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox("Error Record in txt File-Line -" & X, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Call Upload_Quality_File()
        Call Upload_DyePrograme_File()
        Upload_Lot_txtfile()
    End Sub

    Function Upload_EralFile2()
        Dim strFileName As String
        Dim _LotNo As String
        Dim _Batchwhight As Double
        Dim _MCNo As String
        Dim _ProgrammeNo As String
        Dim _ProgrameName As String
        Dim _StratDate As Date
        Dim _StartTime As String
        Dim _EndDate As Date
        Dim _EndTime As String
        Dim _Quality As String
        Dim _ShadeCode As String
        Dim _ShadeType As String
        Dim _Shade As String
        Dim _StandedTime As String
        Dim _LotType As String
        Dim _TotalHR As Integer
        Dim _QltyGroup As String
        Dim _Rating As String
        Dim _Delete As String
        Dim _Weight As Double
        Dim _WeekNo As Integer
        Dim _Shift As Integer
        Dim _WeekDis As String
        Dim _StDate As Date
        Dim _EDDate As Date
        Dim _Taken As String
        Dim _Status As String
        Dim hh As Integer
        Dim mm As Integer


        Dim nvcFieldList1 As String
        Dim Sql As String
        Dim X1 As Integer
        '  _TotalHR = "00:00"
        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Dim M04Lot As DataSet
        Dim nvcVccode As String

        Dim ncQryType As String
        Dim M01 As DataSet
        Dim T01 As DataSet
        Dim hh1 As Integer
        Dim mm1 As Integer
        Dim _TimeDifferance1 As Date
        Dim n_year As Integer


        strFileName = ConfigurationManager.AppSettings("FilePath") + "\download2.csv"
        Dim theTextFieldParser As FileIO.TextFieldParser

        theTextFieldParser = My.Computer.FileSystem.OpenTextFieldParser(strFileName)
        theTextFieldParser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
        theTextFieldParser.Delimiters = New String() {","}
        Dim X11 As Integer
        ' Declare a variable named currentRow of type string array.
        Dim currentRow() As String
        Try

            X11 = 0
            While Not theTextFieldParser.EndOfData
                ' Try
                ' Read the fields on the current line
                ' and assign them to the currentRow array variable.
                currentRow = theTextFieldParser.ReadFields()

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                ' Use the currentField variable to loop
                ' through fields in the currentRow.
                i = 0

                'If X11 = 335 Then
                '    MsgBox("")
                'End If
                For Each currentField In currentRow
                    ' Add the the currentField (a string)
                    ' to the demoLstBox items.
                    If i = 0 Then
                        _LotNo = currentField
                    ElseIf i = 1 Then
                        _MCNo = currentField
                    ElseIf i = 2 Then
                        _ProgrammeNo = currentField
                    ElseIf i = 3 Then
                        _ProgrameName = currentField
                    ElseIf i = 4 Then
                        _LotType = currentField
                    ElseIf i = 5 Then
                        _StandedTime = currentField
                    ElseIf i = 6 Then
                        '  MsgBox(Format(CDate(currentField), "mm/dd/yyyy"))
                        'MsgBox(VB6.Format(currentField, "mm/dd/yyyy"))
                        Dim D_U As String

                        'If Microsoft.VisualBasic.Left(currentField, 2) = "04" Then
                        '    Dim d As DateTime
                        'd = currentField
                        'MsgBox(VB6.Format((currentField), "mm/DD/yyyy"))
                        D_U = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(currentField, 5), 2)
                        _StratDate = D_U & "/" & Microsoft.VisualBasic.Left(currentField, 2) & "/" & Microsoft.VisualBasic.Right(currentField, 4)
                        '_StratDate = (VB6.Format(currentField, "mm/dd/yyyy"))
                    ElseIf i = 7 Then
                        _StartTime = currentField
                    ElseIf i = 8 Then
                        Dim D_U As String

                        'If Microsoft.VisualBasic.Left(currentField, 2) = "04" Then
                        '    Dim d As DateTime
                        'd = currentField
                        'MsgBox(VB6.Format((currentField), "mm/DD/yyyy"))
                        D_U = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(currentField, 5), 2)
                        _EndDate = D_U & "/" & Microsoft.VisualBasic.Left(currentField, 2) & "/" & Microsoft.VisualBasic.Right(currentField, 4)
                        ' _EndDate = (VB6.Format(currentField, "mm/dd/yyyy"))
                    ElseIf i = 9 Then
                        _EndTime = currentField
                    ElseIf i = 10 Then

                        If Len(currentField) = 5 Then
                            hh = Microsoft.VisualBasic.Left(currentField, 2)
                            mm = Microsoft.VisualBasic.Right(currentField, 2)
                        ElseIf Len(currentField) = 4 Then
                            hh = Microsoft.VisualBasic.Left(currentField, 1)
                            mm = Microsoft.VisualBasic.Right(currentField, 2)
                        End If
                        'hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Hour
                        'mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Minute

                        _TotalHR = (hh * 60) + mm
                    ElseIf i = 11 Then
                        _Quality = currentField
                    ElseIf i = 12 Then
                        _QltyGroup = currentField
                    ElseIf i = 13 Then
                        _ShadeCode = currentField
                    ElseIf i = 14 Then
                        _Shade = currentField
                    ElseIf i = 15 Then
                        _ShadeType = currentField
                    ElseIf i = 16 Then
                        '   _Rating = currentField
                        '  _ShadeType = currentField
                    ElseIf i = 17 Then
                        _Weight = currentField
                        'ElseIf i = 18 Then
                        '    _Weight = currentField
                    End If







                    '------------------------------------------------------------------------------------------------

                    i = i + 1
                Next

                'If _TotalHR = "24:23" Then
                '    MsgBox("")
                'End If
                'MsgBox(DateTime.Parse(_EndDate).DayOfWeek())
                Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_EndDate)
                ' dayOfWeek.ToString() would return "Sunday" but it's an enum value,
                ' the correct dayname can be retrieved via DateTimeFormat.
                ' Following returns "Sonntag" for me since i'm in germany '
                Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)


                ' MsgBox(dayName)
                If dayName = "Sunday" Then
                    Dim N_Date1 As Date
                    N_Date1 = CDate(_EndDate).AddDays(-1)
                    _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                    _WeekDis = "Week" & CStr(_WeekNo)
                    n_year = Microsoft.VisualBasic.Year(N_Date1)
                Else
                    If _EndDate = "12/31/" & Microsoft.VisualBasic.Year(_EndDate) Then
                        Dim N_Date1 As Date
                        N_Date1 = CDate(_EndDate).AddDays(+1)
                        _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_year = Microsoft.VisualBasic.Year(N_Date1)
                    Else

                        _WeekNo = DatePart(DateInterval.WeekOfYear, _EndDate)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_year = Microsoft.VisualBasic.Year(_EndDate)
                    End If
                End If
                If DateAndTime.TimeValue(_StartTime) <= "07:30:00 AM" And DateAndTime.TimeValue(_EndTime) >= "07:30:00 PM" Then
                    _Shift = 1

                Else
                    _Shift = 2
                End If


                _StDate = (_StratDate) & " " & (_StartTime)
                _EDDate = (_EndDate) & " " & (_EndTime)
                ' _Taken = System.DateTime.From(CDate(_EDDate).ToOADate) - System.DateTime.FromOADate(CDate(_StDate).ToOADate)

                '  hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Hour
                ' mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(currentField).Minute

                ' _Taken = DateAndTime.TimeValue(System.DateTime.FromOADate(CDate(_EDDate).ToOADate - CDate(_StDate).ToOADate))

                _Taken = _TotalHR

                '  hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(_StandedTime).Subtract(_Taken).Hours
                ' mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(_StandedTime).Subtract(_Taken).Minutes
                '  _TimeDifferance = hh.ToString.PadLeft(2, CChar("0")) & ":" & mm.ToString.PadLeft(2, CChar("0"))
                '_TimeDifferance1 = hh1.ToString.PadLeft(2, CChar("0")) & ":" & mm1.ToString.PadLeft(2, CChar("0"))
                If IsNumeric(_LotNo) Then
                    If Microsoft.VisualBasic.Len(_LotNo) = 6 Then
                        _Status = "L"
                    Else
                        _Status = "A"
                    End If
                Else
                    Sql = "select M05Code from M05DownTime where M05Code='" & Trim(_LotNo) & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M01) Then
                        _Status = "D"
                    Else
                        _Status = "I"
                    End If
                End If

                ncQryType = "ADD"
                ' vMax = Get_highestVouNumber()

                'If _ProgrameName <> "" Then
                _ProgrameName = myReplace.do(_ProgrameName, types.one)
                ' End If
                'nvcFieldList1 = "(M04Ref," & "M04Week," & "M04WeekNo," & "M04Lotno," & "M04Batchwt," & "M04Machine_No," & "M04Program," & "M04Type," & "M04STD," & "M04DateIn," & "M04TimeIn," & "M04Date_Out," & "M04Time_Out," & "M04Taken," & "M04Quality," & "M04Shade_Code," & "M04STD2," & "M04Shift," & "M04Status," & "M04Year," & "M04Month," & "M04ProgrameType," & "M04ETime," & "M04Shade_Type) " & "values('" & vMax & "','" & _WeekNo & "','" & _WeekDis & "','" & _LotNo & "','" & _Weight & "','" & _MCNo & "','" & _ProgrammeNo & "','" & _ProgrameName & "','" & _StandedTime & "','" & _StratDate & "','" & _StartTime & "','" & _EndDate & "','" & _EndTime & "'," & _Taken & ",'" & _Quality & "','" & _ShadeCode & "','" & _StDate & "','" & _Shift & "','" & _Status & "'," & n_year & ",'" & Microsoft.VisualBasic.Month(_EndDate) & "','" & _LotType & "','" & _EDDate & "','" & _ShadeType & "')"
                'up_GetSetM04Lot(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                If _Status = "D" Then
                    Dim _Taken_Min As Integer

                    'mm1 = 0
                    'hh1 = 0
                    '' MsgBox(Len(_Taken))
                    'If Len(_Taken) = 10 Then
                    '    mm1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), (Len(CInt(_Taken)) - 2)), 2))

                    '    hh1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), Len(CInt(_Taken))), (Len(CInt(_Taken)) - 9)))
                    'Else
                    '    mm1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), (Len(CInt(_Taken)) - 3)), 2))
                    '    hh1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(CInt(_Taken), Len(CInt(_Taken))), (Len(CInt(_Taken)) - 9)))
                    'End If
                    '_Taken_Min = 0
                    '_Taken_Min = (hh1 * 60)
                    '_Taken_Min = _Taken_Min + mm1
                    _Taken_Min = _Taken

                    Dim n_Time As Date

                    n_Time = _EndDate & " " & _EndTime

                    'Sql = "select T01Taken from T01Down_Time where T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                    'T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    'If isValidDataset(T01) Then

                    '    nvcFieldList1 = "UPDATE T01Down_Time SET T01Taken=T01Taken +'" & _Taken_Min & "' WHERE T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                    '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'Else

                    '    nvcFieldList1 = "Insert Into T01Down_Time(T01Date,T01Week,T01WeekNo,T01Down_Time,T01Machine,T01Taken,T01Month,T01Year)" & _
                    '                             " values('" & n_Time & "', '" & _WeekDis & "'," & _WeekNo & ",'" & _LotNo & "','" & _MCNo & "','" & _Taken_Min & "','" & Microsoft.VisualBasic.Month(_EndDate) & "','" & n_year & "')"
                    '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    'End If
                End If
                '-----------------------------------------------------------------------------Clear Variable
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                newRow("Lot No") = _LotNo
                newRow("Machine No") = _MCNo
                newRow("Programe No") = _ProgrammeNo
                newRow("Programe Name") = _ProgrameName
                newRow("Lot Type") = _LotType
                newRow("Standed Time") = _StandedTime
                newRow("Start Date") = _StratDate
                newRow("Start Time") = _StartTime
                newRow("End Date") = _EndDate
                newRow("End Time") = _EndTime
                newRow("Total Hour") = _TotalHR
                newRow("Quality") = _Quality
                newRow("Quality Group") = _QltyGroup
                newRow("Shade Code") = _ShadeCode
                newRow("Shade") = _Shade
                newRow("Shade Type") = _ShadeType
                newRow("Weight") = _Weight

                c_dataCustomer1.Rows.Add(newRow)

                pbCount.Value = pbCount.Value + 1
                lblDis.Text = _LotNo

                _LotNo = ""
                _MCNo = ""
                _ProgrammeNo = ""
                _ProgrameName = ""
                _Quality = ""
                _ShadeCode = ""
                _ShadeType = ""
                _StandedTime = ""
                _LotType = ""
                _TotalHR = 0
                _QltyGroup = ""
                _Rating = ""
                _Delete = ""
                _Weight = 0
                mm = 0
                hh = 0
                X11 = X11 + 1
            End While

            'MsgBox("Record update Successfully")
            'transaction.Commit()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            ' MsgBox(X)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            theTextFieldParser.Close()
        End Try


    End Function

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Me.Cursor = Cursors.WaitCursor
        frmMachine.Upload_EralFile()
        frmMachine.Upload_EralFile2()
        cmdEdit.Enabled = False
        Me.Cursor = Cursors.Arrow
    End Sub
End Class