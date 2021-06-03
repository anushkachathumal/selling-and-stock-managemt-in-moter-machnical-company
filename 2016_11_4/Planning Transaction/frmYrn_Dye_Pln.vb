Imports System.Drawing
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation

Public Class frmYrn_Dye_Pln


    Private Sub frmYrn_Dye_Pln_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Header()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Me.Close()
    End Sub

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
        Panel2.Width = (Me.Width - Panel1.Width) - 50
        Panel1.Height = Me.Height - 250
        Panel2.Height = Me.Height - 250
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
        B.Text = "Machine No"
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
        B1.Width = 40
        B1.Height = 32
        B1.TextAlign = ContentAlignment.MiddleCenter
        B1.Text = "Min"
        B1.Location = New Point(_SizeX, _SizeY)

        B1.BackColor = Color.Black
        B1.ForeColor = Color.White
        B1.BorderStyle = BorderStyle.FixedSingle

        Panel1.Controls.Add(B4)
        'UltraGroupBox6.Width = Me.Width - 50
        ' B.Font = New Font(B.Font, FontStyle.Bold)
        B4.AutoSize = False
        B4.Width = 40
        B4.Height = 15
        B4.TextAlign = ContentAlignment.MiddleCenter
        'B.Text = _From
        B4.Location = New Point(_SizeX, _SizeY + 32)

        B4.BackColor = Color.LightGray
        ' B.ForeColor = Color.White
        B4.BorderStyle = BorderStyle.FixedSingle
        '-------------------------------------------------------------------
        Panel1.Controls.Add(B2)

        _SizeX = 185
        _SizeY = 15
        'UltraGroupBox6.Width = Me.Width - 50
        B2.Font = New Font(B2.Font, FontStyle.Bold)
        B2.AutoSize = False
        B2.Width = 40
        B2.Height = 32
        B2.TextAlign = ContentAlignment.MiddleCenter
        B2.Text = "Max"
        B2.Location = New Point(_SizeX, _SizeY)

        B2.BackColor = Color.Black
        B2.ForeColor = Color.White
        B2.BorderStyle = BorderStyle.FixedSingle

        Panel1.Controls.Add(B5)
        'UltraGroupBox6.Width = Me.Width - 50
        ' B.Font = New Font(B.Font, FontStyle.Bold)
        B5.AutoSize = False
        B5.Width = 40
        B5.Height = 15
        B5.TextAlign = ContentAlignment.MiddleCenter
        'B.Text = _From
        B5.Location = New Point(_SizeX, _SizeY + 32)

        B5.BackColor = Color.LightGray
        ' B.ForeColor = Color.White
        B5.BorderStyle = BorderStyle.FixedSingle

        Create_Controls(2, 15)
    End Function

    Function Create_Controls(ByVal strX As Integer, ByVal strY As Integer)
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
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

        strY1 = Me.Height - 250
        _From = Today
        _SizeX = strX
        _SizeY = strY
        _SizeX1 = strX
        i = 0

     

        vcWhere = "tmpDate>='" & Today & "' "
        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarnDye_PLN", New SqlParameter("@cQryType", "MXD"), New SqlParameter("@vcWhereClause1", vcWhere))
        If isValidDataset(M01) Then
            _To = M01.Tables(0).Rows(0)("tmpDate")
            _TimeSpan = _To.Subtract(_From)
            _DateCount = _TimeSpan.Days
        End If

     

        For i = 0 To _DateCount + 1
            Dim B As New Label
            Panel2.Controls.Add(B)
            'UltraGroupBox6.Width = Me.Width - 50
            B.Font = New Font(B.Font, FontStyle.Bold)
            B.AutoSize = False
            B.Width = 120
            B.Height = 32
            B.TextAlign = ContentAlignment.MiddleCenter
            B.Text = _From
            B.Location = New Point(_SizeX, _SizeY)

            B.BackColor = Color.Black
            B.ForeColor = Color.White
            B.BorderStyle = BorderStyle.FixedSingle

            _From = _From.AddDays(+1)
            ' _SizeX = _SizeX + 120

            ' For X = 1 To 2
            Dim B1 As New Label
            Panel2.Controls.Add(B1)
            'UltraGroupBox6.Width = Me.Width - 50
            ' B.Font = New Font(B.Font, FontStyle.Bold)
            B1.AutoSize = False
            B1.Width = 120
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
            _SizeX = _SizeX + 120
            ' i = i + 1
        Next


        _SizeX = 5
        _SizeX1 = 5
        _SizeY = 60
        strY1 = 65
        Dim Rand As New Random

        i = 0
        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarnDye_PLN", New SqlParameter("@cQryType", "YDP"))
        For Each DTRow3 As DataRow In M01.Tables(0).Rows
            Dim B As New Label
            Dim B1 As New Label
            Dim B2 As New Label

            Panel1.Controls.Add(B)
            'UltraGroupBox6.Width = Me.Width - 50
            B.Font = New Font(B.Font, FontStyle.Bold)
            B.AutoSize = False
            B.Width = 138
            B.Height = 30
            B.TextAlign = ContentAlignment.MiddleCenter
            B.Text = M01.Tables(0).Rows(i)("M36MC_No")
            B.Location = New Point(_SizeX, _SizeY)

            B.BackColor = Color.LightBlue
            B.ForeColor = Color.Black
            B.BorderStyle = BorderStyle.FixedSingle

            _SizeX = _SizeX + 139
            Panel1.Controls.Add(B1)
            'UltraGroupBox6.Width = Me.Width - 50
            B1.Font = New Font(B1.Font, FontStyle.Bold)
            B1.AutoSize = False
            B1.Width = 40
            B1.Height = 30
            B1.TextAlign = ContentAlignment.MiddleCenter
            B1.Text = M01.Tables(0).Rows(i)("M36Min_Qty")
            B1.Location = New Point(_SizeX, _SizeY)

            B1.BackColor = Color.LightGreen
            B1.ForeColor = Color.Black
            B1.BorderStyle = BorderStyle.FixedSingle

            _SizeX = _SizeX + 41
            Panel1.Controls.Add(B2)
            'UltraGroupBox6.Width = Me.Width - 50
            B2.Font = New Font(B2.Font, FontStyle.Bold)
            B2.AutoSize = False
            B2.Width = 40
            B2.Height = 30
            B2.TextAlign = ContentAlignment.MiddleCenter
            B2.Text = M01.Tables(0).Rows(i)("M36Max_Qty")
            B2.Location = New Point(_SizeX, _SizeY)

            B2.BackColor = Color.LightGreen
            B2.ForeColor = Color.Black
            B2.BorderStyle = BorderStyle.FixedSingle


            _From = Today
            vcWhere = "tmpDate>='" & Today & "' "
            M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarnDye_PLN", New SqlParameter("@cQryType", "MXD"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                _To = M02.Tables(0).Rows(0)("tmpDate")
                _TimeSpan = _To.Subtract(_From)
                _DateCount = _TimeSpan.Days
            End If
            _SizeX = 2
            _SizeX1 = _SizeX

            For X = 0 To _DateCount + 1
                Dim B6 As New Label

                Panel2.Controls.Add(B6)
                'UltraGroupBox6.Width = Me.Width - 50
                B6.Font = New Font(B6.Font, FontStyle.Bold)
                B6.AutoSize = False
                B6.Width = 120
                B6.Height = 32
                B6.TextAlign = ContentAlignment.MiddleCenter
                B6.Text = _From
                B6.Location = New Point(_SizeX, _SizeY)
                ' B6.Name = "txt" & Month(_From) & Microsoft.VisualBasic.Day(_From)
                B6.BackColor = Color.LightYellow
                B6.ForeColor = Color.LightYellow
                B6.BorderStyle = BorderStyle.FixedSingle

                'Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Black)
                'Dim formGraphics As System.Drawing.Graphics
                'Dim dashValues As Single() = {2, 2, 2, 2}
                'formGraphics = Panel2.CreateGraphics()
                'myPen.DashPattern = dashValues
                'formGraphics.DrawLine(myPen, 560, 60, 560, 2200)

                vcWhere = "tmpDate='" & _From & "' and tmpMC_No='" & Trim(M01.Tables(0).Rows(i)("M36MC_No")) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetYarnDye_PLN", New SqlParameter("@cQryType", "SDP"), New SqlParameter("@vcWhereClause1", vcWhere))
                Z = 0
                For Each DTRow4 As DataRow In M02.Tables(0).Rows
                    Dim _FromTime As Date
                    Dim _ToTime As Date
                    Dim _SlipWidth As Integer


                    _FromTime = M02.Tables(0).Rows(Z)("tmpDate") & " " & "7:00 AM"
                    _ToTime = M02.Tables(0).Rows(Z)("tmpSTTime")

                    '  If Microsoft.VisualBasic.Left(M02.Tables(0).Rows(Z)("tmpSTTime"), 10) = Microsoft.VisualBasic.Left(M02.Tables(0).Rows(Z)("tmpEnd_Time"), 10) Then
                    Dim B7 As New Label
                    _TimeSpan = _ToTime.Subtract(_FromTime)
                    _SizeX1 = _SizeX1 + _TimeSpan.Hours * 5
                    _ToTime = M02.Tables(0).Rows(Z)("tmpEnd_Time")
                    _FromTime = M02.Tables(0).Rows(Z)("tmpSTTime")

                    _TimeSpan = _ToTime.Subtract(_FromTime)
                    _SlipWidth = _TimeSpan.Hours * 5


                    B6.Controls.Add(B7)
                    'UltraGroupBox6.Width = Me.Width - 50
                    B7.Font = New Font(B7.Font, FontStyle.Bold)
                    B7.AutoSize = False
                    B7.Width = _SlipWidth
                    B7.Height = 18
                    B7.TextAlign = ContentAlignment.MiddleCenter
                    B7.Text = M02.Tables(0).Rows(Z)("tmp15Class")
                    B7.Location = New Point(_SizeX1, 5)
                    B7.Name = "lbl" & M02.Tables(0).Rows(Z)("tmpRefNo")
                    
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
                    ToolTip1.SetToolTip(B7, B7.Text & ControlChars.NewLine & "Shade :" & M02.Tables(0).Rows(Z)("tmpShade"))
                    '  ToolTip1.SetToolTip(B7, "Shade :" & M02.Tables(0).Rows(Z)("tmpShade"))
                    B7.Cursor = Cursors.Hand
                    _SizeX1 = _SizeX1 + _SlipWidth
                    Z = Z + 1
                Next
                _SizeX = _SizeX + 120
                '_SizeY = _SizeY + 30
                '_SizeX = 5
                _From = _From.AddDays(+1)

            Next

            '--------------------------------------------------------------------
            _SizeY = _SizeY + 30
            _SizeX = 5

            '   Exit For

            i = i + 1
        Next
    End Function
End Class