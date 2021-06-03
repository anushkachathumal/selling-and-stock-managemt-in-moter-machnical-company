Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO

Public Class frmMachine
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim _MType As String
    Const MAX_SERIALS = 156000
    Dim vMax As Double

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtCode.Focus()
        ' MsgBox(DatePart(DateInterval.WeekOfYear, Today))

        ' Call Upload_EralFile()
        'Call Upload_EralFile2()
    End Sub

    Function Upload_EralFile2()
        Dim strFileName As String
        Dim _LotNo As String
        Dim _Batchwhight As Double
        Dim _MCNo As String
        Dim _ProgrammeNo As String
        Dim _ProgrameName As String
        Dim _StratDate As Date
        Dim _StartTime As Date
        Dim _EndDate As Date
        Dim _EndTime As Date
        Dim n_Year As Integer

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

                        'If Microsoft.VisualBasic.Right(currentField, 4) = "2013" Then
                        '    ' MsgBox("")
                        '    If Microsoft.VisualBasic.Left(currentField, 2) = "04" Then
                        '        MsgBox("")
                        '    End If

                        'End If
                        '    Dim d As DateTime
                        'd = currentField
                        'MsgBox(VB6.Format((currentField), "mm/DD/yyyy"))
                        D_U = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(currentField, 5), 2)
                        _StratDate = D_U & "/" & Microsoft.VisualBasic.Left(currentField, 2) & "/" & Microsoft.VisualBasic.Right(currentField, 4)

                        '    MsgBox(d.ToString("yyyy/MM/dd"))
                        'End If
                        '  MsgBox(Microsoft.VisualBasic.Right(currentField, 4))

                        '  _StratDate = (VB6.Format(currentField, "mm/dd/yyyy"))

                        ElseIf i = 7 Then
                            _StartTime = currentField
                        ElseIf i = 8 Then
                            Dim D_U As String
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
                        ' _ShadeType = currentField
                        ElseIf i = 17 Then
                            _Weight = currentField
                            'ElseIf i = 18 Then
                            '    _Weight = currentField
                        End If


                        '_LotNo = ""
                        '_MCNo = ""
                        '_ProgrammeNo = ""
                        '_ProgrameName = ""
                        '_Quality = ""
                        '_ShadeCode = ""
                        '_ShadeType = ""
                        '_StandedTime = ""
                        '_LotType = ""
                        '_TotalHR = 0
                        'hh = 0
                        'mm = 0
                        '_TotalHR = 0
                        '_QltyGroup = ""
                        '_Rating = ""
                        '_Delete = ""
                        '_Weight = 0




                        '------------------------------------------------------------------------------------------------

                        i = i + 1
                Next

                'If _TotalHR = "24:23" Then
                '    MsgBox("")
                'End If
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
                    n_Year = Microsoft.VisualBasic.Year(N_Date1)
                Else
                    If _EndDate = "12/31/" & Microsoft.VisualBasic.Year(_EndDate) Then
                        Dim N_Date1 As Date
                        N_Date1 = CDate(_EndDate).AddDays(+1)
                        _WeekNo = DatePart(DateInterval.WeekOfYear, N_Date1)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_Year = Microsoft.VisualBasic.Year(N_Date1)
                    Else
                        _WeekNo = DatePart(DateInterval.WeekOfYear, _EndDate)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        n_Year = Microsoft.VisualBasic.Year(_EndDate)
                    End If
                    End If
                    '_WeekNo = DatePart(DateInterval.WeekOfYear, _EndDate)
                    '_WeekDis = "Week" & CStr(_WeekNo)
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
                    vMax = Get_highestVouNumber()

                    'If _ProgrameName <> "" Then
                _ProgrameName = myReplace.do(_ProgrameName, types.one)

                nvcFieldList1 = "select * from M04Lot where M04Lotno='" & _LotNo & "' and M04ProgrameType='" & _LotType & "' and M04Machine_No='" & _MCNo & "' and M04Etime='" & _EDDate & "'"
                dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(dsUser) Then

                Else
                    ' End If
                    nvcFieldList1 = "(M04Ref," & "M04Week," & "M04WeekNo," & "M04Lotno," & "M04Batchwt," & "M04Machine_No," & "M04Program," & "M04Type," & "M04STD," & "M04DateIn," & "M04TimeIn," & "M04Date_Out," & "M04Time_Out," & "M04Taken," & "M04Quality," & "M04Shade_Code," & "M04STD2," & "M04Shift," & "M04Status," & "M04Year," & "M04Month," & "M04ProgrameType," & "M04ETime," & "M04Shade_Type," & "M04Shade," & "M04Qtype) " & "values('" & vMax & "','" & _WeekNo & "','" & _WeekDis & "','" & _LotNo & "','" & _Weight & "','" & _MCNo & "','" & _ProgrammeNo & "','" & _ProgrameName & "','" & _StandedTime & "','" & _StratDate & "','" & _StartTime & "','" & _EndDate & "','" & _EndTime & "'," & _Taken & ",'" & _Quality & "','" & _ShadeCode & "','" & _StDate & "','" & _Shift & "','" & _Status & "'," & n_Year & ",'" & Microsoft.VisualBasic.Month(_EndDate) & "','" & _LotType & "','" & _EDDate & "','" & _ShadeType & "','" & _Shade & "','" & _QltyGroup & "')"
                    up_GetSetM04Lot(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


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
                        Sql = "select T01Taken from T01Down_Time where T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                        T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(T01) Then

                            nvcFieldList1 = "UPDATE T01Down_Time SET T01Taken=T01Taken +'" & _Taken_Min & "' WHERE T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        Else

                            nvcFieldList1 = "Insert Into T01Down_Time(T01Date,T01Week,T01WeekNo,T01Down_Time,T01Machine,T01Taken,T01Month,T01Year)" & _
                                                     " values('" & n_Time & "', '" & _WeekDis & "'," & _WeekNo & ",'" & _LotNo & "','" & _MCNo & "','" & _Taken_Min & "','" & Microsoft.VisualBasic.Month(_EndDate) & "','" & n_Year & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        End If
                    End If
                    End If
                    '-----------------------------------------------------------------------------Clear Variable
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

            MsgBox("Record update Successfully")
            transaction.Commit()

            'If File.Exists("download2.txt") Then
            'Else
            '    File.Delete("download2.txt")
            'End If

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

    Function Upload_EralFile()
        Dim strFileName As String
        Dim _LotNo As String
        Dim _Batchwhight As Double
        Dim _MCNo As String
        Dim _ProgrammeNo As String
        Dim _ProgrameName As String
        Dim _StratDate As Date
        Dim _StartTime As Date
        Dim _EndDate As Date
        Dim _EndTime As Date
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
        Dim hk As Integer

        hk = 0

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


                    '_LotNo = ""
                    '_MCNo = ""
                    '_ProgrammeNo = ""
                    '_ProgrameName = ""
                    '_Quality = ""
                    '_ShadeCode = ""
                    '_ShadeType = ""
                    '_StandedTime = ""
                    '_LotType = ""
                    '_TotalHR = 0
                    'hh = 0
                    'mm = 0
                    '_TotalHR = 0
                    '_QltyGroup = ""
                    '_Rating = ""
                    '_Delete = ""
                    '_Weight = 0




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
                    vMax = Get_highestVouNumber()

                    'If _ProgrameName <> "" Then
                    _ProgrameName = myReplace.do(_ProgrameName, types.one)
                ' End If
                If X11 = 3237 Then
                    '  MsgBox("")
                End If
                nvcFieldList1 = "select * from M04Lot where M04Lotno='" & _LotNo & "' and M04ProgrameType='" & _LotType & "' and M04Machine_No='" & _MCNo & "' and M04Etime='" & _EDDate & "'"
                dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(dsUser) Then
                    ' MsgBox("")
                    'hk = hk + 1
                Else
                    nvcFieldList1 = "(M04Ref," & "M04Week," & "M04WeekNo," & "M04Lotno," & "M04Batchwt," & "M04Machine_No," & "M04Program," & "M04Type," & "M04STD," & "M04DateIn," & "M04TimeIn," & "M04Date_Out," & "M04Time_Out," & "M04Taken," & "M04Quality," & "M04Shade_Code," & "M04STD2," & "M04Shift," & "M04Status," & "M04Year," & "M04Month," & "M04ProgrameType," & "M04ETime," & "M04Shade_Type," & "M04Shade," & "M04Qtype) " & "values('" & vMax & "','" & _WeekNo & "','" & _WeekDis & "','" & _LotNo & "','" & _Weight & "','" & _MCNo & "','" & _ProgrammeNo & "','" & _ProgrameName & "','" & _StandedTime & "','" & _StratDate & "','" & _StartTime & "','" & _EndDate & "','" & _EndTime & "'," & _Taken & ",'" & _Quality & "','" & _ShadeCode & "','" & _StDate & "','" & _Shift & "','" & _Status & "'," & n_year & ",'" & Microsoft.VisualBasic.Month(_EndDate) & "','" & _LotType & "','" & _EDDate & "','" & _ShadeType & "','" & _Shade & "','" & _QltyGroup & "')"
                    up_GetSetM04Lot(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                    If _Status = "D" Then
                        Dim _Taken_Min As Integer

                        _Taken_Min = _Taken

                        Dim n_Time As Date

                        n_Time = _EndDate & " " & _EndTime
                        Sql = "select T01Taken from T01Down_Time where T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                        T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(T01) Then

                            nvcFieldList1 = "UPDATE T01Down_Time SET T01Taken=T01Taken +'" & _Taken_Min & "' WHERE T01Date='" & n_Time & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        Else

                            nvcFieldList1 = "Insert Into T01Down_Time(T01Date,T01Week,T01WeekNo,T01Down_Time,T01Machine,T01Taken,T01Month,T01Year)" & _
                                                     " values('" & n_Time & "', '" & _WeekDis & "'," & _WeekNo & ",'" & _LotNo & "','" & _MCNo & "','" & _Taken_Min & "','" & Microsoft.VisualBasic.Month(_EndDate) & "','" & n_year & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        End If
                    End If
                End If
                '-----------------------------------------------------------------------------Clear Variable
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

            ' MsgBox("Record update Successfully")
            transaction.Commit()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            theTextFieldParser.Close()
        End Try

      
    End Function


    Function Upload_EralData()
        Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\lot.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim _LotNo As String
        Dim _Batchwhight As Double
        Dim _MCNo As String
        Dim _ProgrammeNo As String
        Dim _ProgrameType As String
        Dim _StratDate As Date
        Dim _StartTime As Date
        Dim _EndDate As Date
        Dim _EndTime As Date
        Dim _Quality As String
        Dim _ShadeCode As String
        Dim _ShadeType As String
        Dim _StandedTime As String
        Dim _WeekNo As Integer
        Dim _Shift As Integer
        Dim hh As Integer
        Dim mm As Integer


        Dim nvcFieldList1 As String
        Dim Sql As String

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
        'Dim nvcVccode As String
        Try
            Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))
            Dim M01 As DataSet


            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)
                'Dim _WeekNo As Integer
                Dim _WeekDis As String
                Dim _StDate As Date
                Dim _EDDate As Date
                '' Dim _Shift As Integer
                Dim _STTime As Date
                Dim _cvetStandettime As Date
                Dim X As Double
                Dim Y As Double
                Dim _Taken As String
                Dim _TimeDifferance As String
                '  Dim hh As Integer
                ' Dim mm As Integer

                Dim hh1 As Integer
                Dim mm1 As Integer
                Dim _TimeDifferance1 As Date
                Dim _Status As String
                Dim T01 As DataSet

                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)
                If Trim(strRow) <> "" Then

                    If InStr(1, strRow, vbTab) > 0 Then
                        '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))

                        '_StDate = (Trim(Split(strRow, vbTab)(15))) & " " & (Trim(Split(strRow, vbTab)(16)))
                        '_EDDate = (Trim(Split(strRow, vbTab)(17))) & " " & (Trim(Split(strRow, vbTab)(18)))
                        '' _Taken = System.DateTime.From(CDate(_EDDate).ToOADate) - System.DateTime.FromOADate(CDate(_StDate).ToOADate)
                        '_Taken = DateAndTime.TimeValue(System.DateTime.FromOADate(CDate(_EDDate).ToOADate - CDate(_StDate).ToOADate))
                        
                        _LotNo = (Trim(Split(strRow, vbTab)(0)))
                        _Batchwhight = (Trim(Split(strRow, vbTab)(1)))
                        _MCNo = (Trim(Split(strRow, vbTab)(9)))
                        _ProgrammeNo = (Trim(Split(strRow, vbTab)(10)))
                        _ProgrameType = (Trim(Split(strRow, vbTab)(12)))
                        _StratDate = VB6.Format(Trim(Split(strRow, vbTab)(15)), "MM/DD/YYYY")
                        _StartTime = (Trim(Split(strRow, vbTab)(16)))
                        If Trim(Split(strRow, vbTab)(17)) = "00/00/0000" Then
                        Else
                            _EndDate = VB6.Format(Trim(Split(strRow, vbTab)(17)), "MM/DD/YYYY")
                            _EndTime = (Trim(Split(strRow, vbTab)(18)))
                        End If
                        _Quality = (Trim(Split(strRow, vbTab)(50)))
                        _ShadeCode = (Trim(Split(strRow, vbTab)(51)))
                        _ShadeType = (Trim(Split(strRow, vbTab)(52)))
                        _StandedTime = (Trim(Split(strRow, vbTab)(57)))
                        'strRolls = (Trim(Split(strRow, vbTab)(13)))
                        'strStatus = (Trim(Split(strRow, vbTab)(14)))


                        _WeekNo = DatePart(DateInterval.WeekOfYear, _EndDate)
                        _WeekDis = "Week" & CStr(_WeekNo)
                        If DateAndTime.TimeValue(_StDate) >= "07:30:00 AM" And DateAndTime.TimeValue(_StDate) <= "07:30:00 PM" Then
                            _Shift = 1

                        Else
                            _Shift = 2
                        End If

                        If Trim(Split(strRow, vbTab)(17)) = "00/00/0000" Then

                        Else
                            _StDate = (_StratDate) & " " & (_StartTime)
                            _EDDate = (_EndDate) & " " & (_EndTime)
                            ' _Taken = System.DateTime.From(CDate(_EDDate).ToOADate) - System.DateTime.FromOADate(CDate(_StDate).ToOADate)
                            _Taken = DateAndTime.TimeValue(System.DateTime.FromOADate(CDate(_EDDate).ToOADate - CDate(_StDate).ToOADate))


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


                            'If lLineNo = 8865 Then
                            '    MsgBox("")
                            'End If
                            If _Status = "D" Then
                                _StandedTime = "0"
                            End If
                            _STTime = (Format(Int(_StandedTime / 60) Mod 60, "0#") & _
                                 ":" & Format(_StandedTime Mod 60, "0#"))

                            ncQryType = "ADD"
                            'hh = _STTime.Subtract(_Taken).Hours
                            'mm = _STTime.Subtract(_Taken).Minutes

                            'hh1 = _TimeDifferance1.AddHours(_Taken).Hour
                            'mm1 = _TimeDifferance1.AddMinutes(_Taken).Minute

                            hh = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(_STTime).Subtract(_Taken).Hours
                            mm = Microsoft.VisualBasic.CompilerServices.Conversions.ToDate(_STTime).Subtract(_Taken).Minutes
                            _TimeDifferance = hh.ToString.PadLeft(2, CChar("0")) & ":" & mm.ToString.PadLeft(2, CChar("0"))
                            '_TimeDifferance1 = hh1.ToString.PadLeft(2, CChar("0")) & ":" & mm1.ToString.PadLeft(2, CChar("0"))


                            ncQryType = "ADD"
                            vMax = Get_highestVouNumber()

                            nvcFieldList1 = "(M04Ref," & "M04Week," & "M04WeekNo," & "M04Lotno," & "M04Batchwt," & "M04Machine_No," & "M04Program," & "M04Type," & "M04STD," & "M04DateIn," & "M04TimeIn," & "M04Date_Out," & "M04Time_Out," & "M04Taken," & "M04T_Difference," & "M04Quality," & "M04Shade_Code," & "M04STD2," & "M04Shift," & "M04Status," & "M04Year," & "M04Month," & "M04ProgrameType) " & "values('" & vMax & "','" & _WeekNo & "','" & _WeekDis & "','" & _LotNo & "','" & _Batchwhight & "','" & _MCNo & "','" & _ProgrammeNo & "','" & _ProgrameType & "','" & _StandedTime & "','" & _StratDate & "','" & _StartTime & "','" & _EndDate & "','" & _EndTime & "','" & _Taken & "','" & _TimeDifferance & "','" & _Quality & "','" & _ShadeCode & "','" & _STTime & "','" & _Shift & "','" & _Status & "'," & Microsoft.VisualBasic.Year(_StratDate) & ",'" & Microsoft.VisualBasic.Month(_StratDate) & "','" & _LotNo & "')"
                            ' nvcFieldList1 = "(M04Ref," & "M04Machine_No," & "M04STD," & "M04Week," & "M04WeekNo," & "M04Lotno," & "M04Batchwt," & "M04Program," & "M04Type," & "M04DateIn," & "M04TimeIn," & "M04Date_Out) " & "values(" & vMax & ",'" & _MCNo & "','" & _STTime & "'," & _WeekNo & ",'" & _WeekDis & "','" & _LotNo & "'," & _Batchwhight & ",'" & _ProgrammeNo & "','" & _ProgrameType & "','" & _StratDate & "','" & _StartTime & "','" & _EndDate & "')"
                            up_GetSetM04Lot(ncQryType, nvcFieldList1, nvcVccode, connection, transaction)


                            If _Status = "D" Then
                                Dim _Taken_Min As Integer

                                mm1 = 0
                                hh1 = 0
                                ' MsgBox(Len(_Taken))
                                If Len(_Taken) = 10 Then
                                    mm1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(_Taken, (Len(_Taken) - 2)), 2))

                                    hh1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(_Taken, Len(_Taken)), (Len(_Taken) - 9)))
                                Else
                                    mm1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(_Taken, (Len(_Taken) - 3)), 2))
                                    hh1 = (Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(_Taken, Len(_Taken)), (Len(_Taken) - 9)))
                                End If
                                _Taken_Min = 0
                                _Taken_Min = (hh1 * 60)
                                _Taken_Min = _Taken_Min + mm1
                                Sql = "select T01Taken from T01Down_Time where T01Date='" & _StratDate & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                                T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                                If isValidDataset(T01) Then

                                    nvcFieldList1 = "UPDATE T01Down_Time SET T01Taken=T01Taken +'" & _Taken_Min & "' WHERE T01Date='" & _StratDate & "' and T01Down_Time='" & _LotNo & "' and T01Machine='" & _MCNo & "'"
                                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                                Else

                                    nvcFieldList1 = "Insert Into T01Down_Time(T01Date,T01Week,T01WeekNo,T01Down_Time,T01Machine,T01Taken,T01Month)" & _
                                                             " values('" & _StratDate & "', '" & _WeekDis & "'," & _WeekNo & ",'" & _LotNo & "','" & _MCNo & "','" & _Taken_Min & "','" & Microsoft.VisualBasic.Month(_StratDate) & "')"
                                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                                End If

                            End If
                            End If

                    End If
                End If

                _Batchwhight = 0
                ' _EndDate = ""
                ' _EndTime = ""
                _LotNo = ""
                _MCNo = ""
                _MType = ""
                _ProgrameType = ""
                _ProgrammeNo = ""
                _Quality = ""
                _ShadeCode = ""
                _ShadeType = ""
                _StandedTime = ""
                '_StartTime = ""
                '_StratDate = ""
                _WeekDis = ""
                _WeekNo = 0


                lLineNo = lLineNo + 1

            Loop
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            MsgBox("Records update sucessfully", MsgBoxStyle.Information, "Textued Jersey ......")
            FileClose()

        Catch ex As Exception
            If transactionCreated Then transaction.Rollback()
            'Throw ex
            MessageBox.Show(Me, ex.ToString)
            FileClose()
        Finally

            If connectionCreated Then DBEngin.CloseConnection(connection)
            ' MDIMain.UltraStatusBar1.Panels(1).Text = "Transaction successfully updated ....."
        End Try
    End Function

    Private Function Get_highestVouNumber() As String
        Dim con = New SqlConnection()
        Dim vMax As String

        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetParameter", New SqlParameter("@cQryType", "UPD"), New SqlParameter("@vcCode", "LO"))
            If common.isValidDataset(dsUser) Then
                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    vMax = dsUser.Tables(0).Rows(0)("P01LastNo")
                    Return vMax
                Next
            Else
                MessageBox.Show("Record Not Found", "Textured Jersey", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
        '=========================================================================
        ' "asdasd"
    End Function
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmMachine_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M02Code as [Group] from M02Group "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboGroup.DataSource = M01
                cboGroup.Rows.Band.Columns(0).Width = 125
                ' cboGroup.Rows.Band.Columns(1).Width = 270
                'cboGroup.Rows.Band.Columns(2).Width = 170
                'cboGroup.Rows.Band.Columns(3).Width = 130
            End If

            Sql = "select M01Description as [Machine Type] from M01Dyeing_MC_Type "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboType.DataSource = M01
                cboType.Rows.Band.Columns(0).Width = 175
                ' cboGroup.Rows.Band.Columns(1).Width = 270
                'cboGroup.Rows.Band.Columns(2).Width = 170
                'cboGroup.Rows.Band.Columns(3).Width = 130
            End If

            Call LoadGride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function LoadGride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select T03Code as [Machine Code],T03Name as [Description],T03Group as [Group],M01Description as [Machine Type] from T03Machine inner join M01Dyeing_MC_Type on M01Code=T03Type where T03Status='A'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 270
            UltraGrid1.Rows.Band.Columns(2).Width = 110
            UltraGrid1.Rows.Band.Columns(3).Width = 140
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Records()
        'SEARCH MACHINE DETAILES
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T03 As DataSet

        Try
            Sql = "SELECT T03Name,T03Group,M01Description FROM T03Machine INNER JOIN M01Dyeing_MC_Type ON T03Type=M01Code WHERE T03CODE='" & Trim(txtCode.Text) & "' AND T03STATUS='A'"
            T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(T03) Then
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                cmdSave.Enabled = False
                txtDescription.Text = T03.Tables(0).Rows(0)("T03Name")
                cboGroup.Text = T03.Tables(0).Rows(0)("T03Group")
                cboType.Text = T03.Tables(0).Rows(0)("M01Description")
            Else
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False


            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            If txtCode.Text <> "" Then
                Call Search_Records()
                cboGroup.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub txtCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.ValueChanged

    End Sub

    Private Sub cboGroup_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboGroup.InitializeLayout

    End Sub

    Private Sub cboGroup_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboGroup.KeyUp
        If e.KeyCode = 13 Then
            cboType.ToggleDropdown()
        End If
    End Sub

    Function Search_MType() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T03 As DataSet
        Try

            Sql = "select * from M01Dyeing_MC_Type where M01Description='" & cboType.Text & "'"
            T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(T03) Then
                Search_MType = True
                _MType = T03.Tables(0).Rows(0)("m01code")
            Else
                Search_MType = False
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboType_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboType.InitializeLayout

    End Sub

    Private Sub cboType_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboType.KeyUp
        If e.KeyCode = 13 Then
            txtDescription.Focus()
        End If
    End Sub

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = 13 Then
            If cmdSave.Enabled = True Then
                cmdSave.Focus()
            Else
                cmdEdit.Focus()
            End If
        End If
    End Sub

    Private Sub txtDescription_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.ValueChanged
        If cmdEdit.Enabled = True Then
        Else
            cmdSave.Enabled = True
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
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


            If Trim(txtCode.Text) <> "" And Trim(txtDescription.Text) <> "" And cboGroup.Text <> "" Then
                If Search_MType() = True Then
                    nvcFieldList1 = "Insert Into T03Machine(T03Code,T03Name,T03Group,T03Type,T03Status)" & _
                                                             " values('" & Trim(txtCode.Text) & "', '" & Trim(txtDescription.Text) & "','" & cboGroup.Text & "','" & _MType & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    MsgBox("Please enter the correct machine type", MsgBoxStyle.Information, "Textured Jersey ........")
                End If
            Else
                MsgBox("Please enter the complete records", MsgBoxStyle.Information, "Textured Jersey ........")
                Exit Sub
            End If
            MsgBox("Record Update Sucessfully", MsgBoxStyle.Information, "Textured Jersey .........")
            transaction.Commit()
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()

        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
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


            If Trim(txtCode.Text) <> "" And Trim(txtDescription.Text) <> "" And cboGroup.Text <> "" Then
                If Search_MType() = True Then
                    nvcFieldList1 = "Update T03Machine set T03Name='" & Trim(txtDescription.Text) & "',T03Group='" & cboGroup.Text & "',T03Type='" & _MType & "' where T03Code='" & Trim(txtCode.Text) & "' and T03Status='A'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    MsgBox("Please enter the correct machine type", MsgBoxStyle.Information, "Textured Jersey ........")
                End If
            Else
                MsgBox("Please enter the complete records", MsgBoxStyle.Information, "Textured Jersey ........")
            End If
            MsgBox("Record Update Sucessfully", MsgBoxStyle.Information, "Textured Jersey .........")
            transaction.Commit()
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim A As String
        Dim nvcFieldList As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Try
            A = MsgBox("Are you sure you want to Delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Textured Jersey .........")
            If A = vbYes Then


                nvcFieldList = "delete from T03Machine where  t03Code = '" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, "Textured Jersey ........")
                transaction.Commit()


            End If

            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        Call LoadGride()
    End Sub

    Private Sub UltraGroupBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox1.Click

    End Sub
End Class