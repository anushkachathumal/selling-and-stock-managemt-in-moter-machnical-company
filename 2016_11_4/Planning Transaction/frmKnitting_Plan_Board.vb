Imports System.Drawing
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports System.Globalization

Public Class frmKnitting_Plan_Board


    Function Load_Header()
        Dim _SizeX As Integer
        Dim _SizeY As Integer
        Dim B As New Label
        Dim B1 As New Label
        Dim B2 As New Label
        Dim B3 As New Label '
        Dim B4 As New Label '
        Dim B5 As New Label


        UltraGroupBox6.Width = Me.Width - 50
        Panel1.Width = Me.Width - 50
        Panel1.Height = Me.Height - 250
        '  Panel2.Height = Me.Height - 250
        '_SizeX = Me.Height - 250
        '_SizeY = Me.Width - 50

        ''  HScrollBar1.Location = New HScrollBar(HScrollBar1.Location, _SizeX, _SizeY)
        Panel1.Controls.Add(B)

        _SizeX = 5
        _SizeY = 15
        'UltraGroupBox6.Width = Me.Width - 50
        B.Font = New Font(B.Font, FontStyle.Bold)
        B.AutoSize = False
        B.Width = 138
        B.Height = 32
        B.TextAlign = ContentAlignment.MiddleCenter
        B.Text = "Group No"
        B.Location = New Point(_SizeX, _SizeY)

        B.BackColor = Color.Black
        B.ForeColor = Color.White
        B.BorderStyle = BorderStyle.FixedSingle

        Panel1.Controls.Add(B3)
        'UltraGroupBox6.Width = Me.Width - 50
        ' B.Font = New Font(B.Font, FontStyle.Bold)
        B3.AutoSize = False
        B3.Width = 138
        B3.Height = 15
        B3.TextAlign = ContentAlignment.MiddleCenter
        'B.Text = _From
        B3.Location = New Point(_SizeX, _SizeY + 32)

        B3.BackColor = Color.LightGray
        ' B.ForeColor = Color.White
        B3.BorderStyle = BorderStyle.FixedSingle
        '---------------------------------------------------------------------------------
        Panel1.Controls.Add(B1)

        _SizeX = 144
        _SizeY = 15
        'UltraGroupBox6.Width = Me.Width - 50
        B1.Font = New Font(B1.Font, FontStyle.Bold)
        B1.AutoSize = False
        B1.Width = 100
        B1.Height = 32
        B1.TextAlign = ContentAlignment.MiddleCenter
        B1.Text = "Machine No"
        B1.Location = New Point(_SizeX, _SizeY)

        B1.BackColor = Color.Black
        B1.ForeColor = Color.White
        B1.BorderStyle = BorderStyle.FixedSingle

        Panel1.Controls.Add(B4)
        'UltraGroupBox6.Width = Me.Width - 50
        ' B.Font = New Font(B.Font, FontStyle.Bold)
        B4.AutoSize = False
        B4.Width = 100
        B4.Height = 15
        B4.TextAlign = ContentAlignment.MiddleCenter
        'B.Text = _From
        B4.Location = New Point(_SizeX, _SizeY + 32)

        B4.BackColor = Color.LightGray
        ' B.ForeColor = Color.White
        B4.BorderStyle = BorderStyle.FixedSingle
        '-------------------------------------------------------------------
        'Panel1.Controls.Add(B2)

        '_SizeX = 185
        '_SizeY = 15
        ''UltraGroupBox6.Width = Me.Width - 50
        'B2.Font = New Font(B2.Font, FontStyle.Bold)
        'B2.AutoSize = False
        'B2.Width = 40
        'B2.Height = 32
        'B2.TextAlign = ContentAlignment.MiddleCenter
        'B2.Text = "Max"
        'B2.Location = New Point(_SizeX, _SizeY)

        'B2.BackColor = Color.Black
        'B2.ForeColor = Color.White
        'B2.BorderStyle = BorderStyle.FixedSingle

        'Panel1.Controls.Add(B5)
        ''UltraGroupBox6.Width = Me.Width - 50
        '' B.Font = New Font(B.Font, FontStyle.Bold)
        'B5.AutoSize = False
        'B5.Width = 40
        'B5.Height = 15
        'B5.TextAlign = ContentAlignment.MiddleCenter
        ''B.Text = _From
        'B5.Location = New Point(_SizeX, _SizeY + 32)

        'B5.BackColor = Color.LightGray
        '' B.ForeColor = Color.White
        'B5.BorderStyle = BorderStyle.FixedSingle
        _SizeX = 245
        Create_Controls(_SizeX, _SizeY)
    End Function

    Private Sub frmKnitting_Plan_Board_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Header()
    End Sub

    Function Create_Controls(ByVal strX As Integer, ByVal strY As Integer)
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim T02 As DataSet

        Dim _SizeX As Integer
        Dim _SizeY As Integer
        Dim _TimeSpan As TimeSpan
        Dim _From As Date
        Dim _To As Date
        Dim _DateCount As Integer
        Dim vcWhere As String
        Dim _SizeX1 As Integer
        Dim X As Integer
        Dim strY1 As Integer
        Dim Z As Integer
        Dim M02 As DataSet
        Dim _FromWeek As Integer
        Dim _ToWeek As Integer
        Dim _weekCount As Integer
        Dim _Weekof_Year As Integer
        Dim _StartX As Integer
        Dim _StratY As Integer
        Dim _EndX As Integer
        Dim Y As Integer
        '  Dim Z As Integer

        strY1 = Me.Height - 250
        _From = Today
        _SizeX = strX
        _SizeY = strY
        _SizeX1 = strX
        _StratY = strY
        _SizeX = strX
        i = 0

        Dim dateNow = DateTime.Now
        Dim dfi = DateTimeFormatInfo.CurrentInfo
        Dim calendar = dfi.Calendar

        _FromWeek = DatePart("ww", Today, FirstDayOfWeek.Monday)
        _weekCount = _FromWeek
        vcWhere = "tmpWeek_No>='" & _FromWeek & "' and tmpYear>='" & Year(Today) & "'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarnDye_PLN", New SqlParameter("@cQryType", "KNP"), New SqlParameter("@vcWhereClause1", vcWhere))
        If isValidDataset(M01) Then
            _ToWeek = M01.Tables(0).Rows(0)("tmpWeek_No")
        End If

        _Weekof_Year = calendar.GetWeekOfYear(dateNow, dfi.CalendarWeekRule, DayOfWeek.Thursday)


        For i = _FromWeek To _ToWeek
            Dim B As New Label
            Panel1.Controls.Add(B)
            'UltraGroupBox6.Width = Me.Width - 50
            B.Font = New Font(B.Font, FontStyle.Bold)
            B.AutoSize = False
            B.Width = 220
            B.Height = 32
            B.TextAlign = ContentAlignment.MiddleCenter
            B.Text = "Week " & _weekCount
            B.Location = New Point(_SizeX, _SizeY)

            B.BackColor = Color.Black
            B.ForeColor = Color.White
            B.BorderStyle = BorderStyle.FixedSingle

            _From = _From.AddDays(+1)
            ' _SizeX = _SizeX + 120

            ' For X = 1 To 2
            Dim B1 As New Label
            Panel1.Controls.Add(B1)
            'UltraGroupBox6.Width = Me.Width - 50
            ' B.Font = New Font(B.Font, FontStyle.Bold)
            B1.AutoSize = False
            B1.Width = 220
            B1.Height = 15
            B1.TextAlign = ContentAlignment.MiddleCenter
            'B.Text = _From
            B1.Location = New Point(_SizeX, _SizeY + 32)

            B1.BackColor = Color.LightGray
            ' B.ForeColor = Color.White
            B1.BorderStyle = BorderStyle.FixedSingle

            '    _SizeX1 = _SizeX1 + 60
            'Next


            'Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Black)
            'Dim formGraphics As System.Drawing.Graphics
            'Dim dashValues As Single() = {2, 2, 2, 2}
            'formGraphics = B1.CreateGraphics()
            'myPen.DashPattern = dashValues
            '_SizeX1 = _SizeX / 2
            'formGraphics.DrawLine(myPen, 60, 1, 60, 55)
            'myPen.Dispose()
            'formGraphics.Dispose()
            _SizeX = _SizeX + 220
            _EndX = _EndX + 220
            _weekCount = _weekCount + 1
            ' i = i + 1
        Next
        ' _EndX = strX
        '===================================================================================================
        _SizeX = 5
        _SizeY = 60
        i = 0
        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "KNG"))
        For Each DTRow3 As DataRow In M01.Tables(0).Rows
            Dim B As New Label


            Panel1.Controls.Add(B)
            'UltraGroupBox6.Width = Me.Width - 50
            B.Font = New Font(B.Font, FontStyle.Bold)
            B.AutoSize = False
            B.Width = 138
            B.Height = 30
            B.TextAlign = ContentAlignment.MiddleCenter
            B.Text = M01.Tables(0).Rows(i)("tmpGroup")
            B.Location = New Point(_SizeX, _SizeY)

            B.BackColor = Color.BurlyWood
            B.ForeColor = Color.Black
            B.BorderStyle = BorderStyle.FixedSingle

            _SizeX1 = 144
            X = 0
            vcWhere = "tmpgroup='" & Trim(M01.Tables(0).Rows(i)("tmpGroup")) & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "MCK"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow4 As DataRow In M02.Tables(0).Rows
                Dim B1 As New Label
                Dim B2 As New Label

                Panel1.Controls.Add(B1)
                'UltraGroupBox6.Width = Me.Width - 50
                B1.Font = New Font(B1.Font, FontStyle.Bold)
                B1.AutoSize = False
                B1.Width = 100
                B1.Height = 30
                B1.TextAlign = ContentAlignment.MiddleCenter
                B1.Text = M02.Tables(0).Rows(X)("tmpMC_No")
                B1.Location = New Point(_SizeX1, _SizeY)

                B1.BackColor = Color.CadetBlue
                B1.ForeColor = Color.Black
                B1.BorderStyle = BorderStyle.FixedSingle


                Panel1.Controls.Add(B2)
                'UltraGroupBox6.Width = Me.Width - 50
                ' B1.Font = New Font(B1.Font, FontStyle.Bold)
                B2.AutoSize = False
                B2.Width = _EndX
                B2.Height = 30
                '   B1.TextAlign = ContentAlignment.MiddleCenter
                '  B1.Text = M02.Tables(0).Rows(X)("tmpMC_No")
                B2.Location = New Point(245, _SizeY)

                B2.BackColor = Color.White
                B2.ForeColor = Color.Black
                B2.BorderStyle = BorderStyle.None
                '--------------------------------------------------------------------------------
                Dim WeekStartdate1 As Date
                Dim C1 As Integer
                C1 = 0
                WeekStartdate1 = Today
                For Y = _FromWeek To _ToWeek
                    Dim StartdateofYear As Date
                    Dim Dateofweek As Integer
                    Dim WeekEnddate1 As Date

                    StartdateofYear = "1/1/" & Year(WeekStartdate1)

                    Dateofweek = (7 * Y) - 7
                    WeekStartdate1 = CDate(StartdateofYear).AddDays(+Dateofweek)
                    Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                    Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(WeekStartdate1)
                    Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

                    If dayName = "Sunday" Then
                        WeekStartdate1 = CDate(WeekStartdate1).AddDays(+1)
                    ElseIf dayName = "Tuesday" Then
                        WeekStartdate1 = CDate(WeekStartdate1).AddDays(-1)
                    ElseIf dayName = "Wednesday" Then
                        WeekStartdate1 = CDate(WeekStartdate1).AddDays(-2)
                    ElseIf dayName = "Thursday" Then
                        WeekStartdate1 = CDate(WeekStartdate1).AddDays(-3)
                    ElseIf dayName = "Friday" Then
                        WeekStartdate1 = CDate(WeekStartdate1).AddDays(-4)
                    ElseIf dayName = "Saturday" Then
                        WeekStartdate1 = CDate(WeekStartdate1).AddDays(-5)
                    End If
                    WeekStartdate1 = WeekStartdate1 & " " & "7:30 AM"
                    Z = 0
                    Dim _LabaleSize As Integer
                    _LabaleSize = 0
                    vcWhere = "tmpWeek_No=" & Y & " and tmpYear='" & Year(WeekStartdate1) & "' and tmpMC_No='" & M02.Tables(0).Rows(X)("tmpMC_No") & "'"
                    T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Planning", New SqlParameter("@cQryType", "PLN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In T02.Tables(0).Rows
                        Dim _LableStart As Integer

                        Dim _Value As Double
                        Dim diff As TimeSpan
                        Dim _ToTime As Date
                        Dim _FromTime As Date
                        Dim _LocX As Integer
                        Dim B7 As New Label
                        Dim Rand As New Random
                        Dim _Value1 As Double

                        _Value = 7 * 24 * 60
                        _Value1 = 220 / _Value
                        _Value = _EndX / _Value


                        _FromTime = T02.Tables(0).Rows(Z)("tmpStart_time")
                        _ToTime = T02.Tables(0).Rows(Z)("tmpEnd_Time")
                        diff = _FromTime.Subtract(WeekStartdate1)
                        _LocX = diff.Days * 24 * 60
                        _LocX = _LocX + (diff.Hours * 60)
                        _LocX = _LocX + diff.Minutes
                        _LocX = _LocX * _Value1
                        C1 = C1 + (_LabaleSize / 2)

                        If Z = 0 Then
                            C1 = _LocX + C1
                        Else

                        End If
                        _LocX = C1
                        'SET LABLE SIZE
                        ' _LabaleSize = 0
                        diff = _ToTime.Subtract(_FromTime)
                        _LabaleSize = diff.Days * 24 * 60
                        _LabaleSize = _LabaleSize + (diff.Hours * 60)
                        _LabaleSize = _LabaleSize + diff.Minutes
                        _LabaleSize = _LabaleSize * _Value

                        B2.Controls.Add(B7)
                        'UltraGroupBox6.Width = Me.Width - 50
                        B7.Font = New Font(B7.Font, FontStyle.Bold)
                        B7.AutoSize = False
                        B7.Width = _LabaleSize
                        B7.Height = 18
                        B7.TextAlign = ContentAlignment.MiddleCenter
                        B7.Text = T02.Tables(0).Rows(Z)("tmp20Class")
                        B7.Location = New Point(_LocX, 5)
                        B7.Name = "lbl" & T02.Tables(0).Rows(Z)("tmpQuality")

                        B7.BackColor = Color.FromArgb(Rand.Next(0, 255), Rand.Next(0, 256), Rand.Next(0, 256))
                        B7.ForeColor = Color.Black
                        B7.BorderStyle = BorderStyle.FixedSingle
                        Dim myFont As New Font("Sans Serif", 6, FontStyle.Bold)
                        B7.Font = myFont
                        Dim ToolTip1 As New ToolTip()
                        ToolTip1.AutomaticDelay = 5000
                        ToolTip1.InitialDelay = 1000
                        ToolTip1.ReshowDelay = 500
                        ToolTip1.ShowAlways = True
                        Dim strTT As String
                        ' strTT = B7.Text & vbTab & "," & "Shade :" & M02.Tables(0).Rows(Z)("tmpShade")
                        ToolTip1.SetToolTip(B7, B7.Text & ControlChars.NewLine & "Quality :" & T02.Tables(0).Rows(Z)("tmpQuality") & ControlChars.NewLine & "Start Time :" & T02.Tables(0).Rows(Z)("tmpStart_time") & ControlChars.NewLine & "End Time :" & T02.Tables(0).Rows(Z)("tmpEnd_Time") & ControlChars.NewLine & "Quantity :" & T02.Tables(0).Rows(Z)("tmpKnt_Order"))
                        '  ToolTip1.SetToolTip(B7, "Shade :" & M02.Tables(0).Rows(Z)("tmpShade"))
                        B7.Cursor = Cursors.Hand

                        Z = Z + 1
                    Next


                Next
                _SizeY = _SizeY + 30
                X = X + 1
            Next
            If isValidDataset(M02) Then
            Else
                _SizeY = _SizeY + 30
            End If

            i = i + 1
        Next

    End Function
End Class