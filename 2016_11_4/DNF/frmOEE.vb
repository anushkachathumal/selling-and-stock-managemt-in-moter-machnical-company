Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS
Public Class frmOEE
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub frmOEE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'lblC.Appearance.BackColor = Color.Black
        Call Clear_CalanderLR()
        Call Clear_WorkingTime()
        Call Clear_LoadingTime()
        Call Claer_OparatingTime()
        Call Clear_Net()
        Call Clear_Value()
        Call Claer_OEE()
        Call Clear_Other()
    End Sub

    Function Clear_CalanderLR()
        lblCLR1.Text = ""
        lblCLR2.Text = ""
        lblCLR3.Text = ""
        lblCLR4.Text = ""
        lblCLR5.Text = ""
        lblCLR6.Text = ""
        lblCLR7.Text = ""
        lblCLR8.Text = ""
        lblCLR9.Text = ""

        lblCE1.Text = ""
        lblCE10.Text = ""
        lblCE11.Text = ""
        lblCE12.Text = ""
        lblCE13.Text = ""
        lblCE14.Text = ""
        lblCE15.Text = ""
        lblCE16.Text = ""
        lblCE17.Text = ""
        lblCE2.Text = ""
        lblCE4.Text = ""
        lblCE6.Text = ""
        lblCE7.Text = ""
        lblCE8.Text = ""
        lblCE9.Text = ""
        lblCT21.Text = ""
        lblCT22.Text = ""
        lblCT23.Text = ""
    End Function
    Function Clear_WorkingTime()
        lblWE10.Text = ""
        lblWE11.Text = ""
        lblWE12.Text = ""
        lblWE13.Text = ""
        lblWE14.Text = ""
        lblWE15.Text = ""
        lblWE16.Text = ""
        lblWE17.Text = ""
        lblWE2.Text = ""
        lblWE3.Text = ""
        lblWE4.Text = ""
        lblWE4.Text = ""
        lblWE6.Text = ""
        lblWE7.Text = ""
        lblWE8.Text = ""
        lblWE9.Text = ""
        lblWLR1.Text = ""
        lblWLR2.Text = ""
        lblWLR3.Text = ""
        lblWLR4.Text = ""
        lblWLR5.Text = ""
        lblWLR6.Text = ""
        lblWLR7.Text = ""
        lblWLR8.Text = ""
        lblWLR9.Text = ""
        lblWT21.Text = ""
        lblWT22.Text = ""
        lblWT23.Text = ""
    End Function
    Function Clear_LoadingTime()
        lblLE10.Text = ""
        lblLE11.Text = ""
        lblLE12.Text = ""
        lblLE13.Text = ""
        lblLE14.Text = ""
        lblLE15.Text = ""
        lblLE16.Text = ""
        lblLE17.Text = ""
        lbllE2.Text = ""
        lblLE3.Text = ""
        lblLE6.Text = ""
        lblLE4.Text = ""
        lblLE6.Text = ""
        lblLE7.Text = ""
        lblLE8.Text = ""
        lblLE9.Text = ""
        lblLLR1.Text = ""
        lblLLR2.Text = ""
        lblLLR3.Text = ""
        lblLLR4.Text = ""
        lblLLR5.Text = ""
        lblLLR6.Text = ""
        lblLLR7.Text = ""
        lblLLR8.Text = ""
        lblLLR9.Text = ""
        lblLT21.Text = ""
        lblLT22.Text = ""
        lblLT23.Text = ""
    End Function
    Function Claer_OparatingTime()
        lblOE10.Text = ""
        lblOE11.Text = ""
        lblOE12.Text = ""
        lblOE13.Text = ""
        lblOE14.Text = ""
        lblOE15.Text = ""
        lblOE16.Text = ""
        lblOE17.Text = ""
        lblOE2.Text = ""
        lblOE3.Text = ""
        lblOE4.Text = ""
        lblOE6.Text = ""
        lblOE7.Text = ""
        lblOE8.Text = ""
        lblOE9.Text = ""
        lblOLR1.Text = ""
        lblOLR2.Text = ""
        lblOLR3.Text = ""
        lblOLR4.Text = ""
        lblOLR5.Text = ""
        lblOLR6.Text = ""
        lblOLR7.Text = ""
        lblOLR8.Text = ""
        lblOT21.Text = ""
        lblOT22.Text = ""
        lblOT23.Text = ""
        lblOLR9.Text = ""
    End Function
    Function Clear_Net()
        lblNE10.Text = ""
        lblNE11.Text = ""
        lblNE12.Text = ""
        lblNE13.Text = ""
        lblNE14.Text = ""
        lblNE15.Text = ""
        lblNE16.Text = ""
        lblNE17.Text = ""
        lblNE2.Text = ""
        lblNE3.Text = ""
        lblNE4.Text = ""
        lblNE6.Text = ""
        lblNE7.Text = ""
        lblNE8.Text = ""
        lblNE9.Text = ""
        lblNLR1.Text = ""
        lblNLR2.Text = ""
        lblNLR3.Text = ""
        lblNLR4.Text = ""
        lblNLR5.Text = ""
        lblNLR6.Text = ""
        lblNLR7.Text = ""
        lblNLR8.Text = ""
        lblNLR9.Text = ""
        lblNT21.Text = ""
        lblNT22.Text = ""
        lblNT23.Text = ""
    End Function

    Function Clear_Value()
        lblVE10.Text = ""
        lblVE11.Text = ""
        lblVE12.Text = ""
        lblVLR5.Text = ""
        lblVE13.Text = ""
        lblVE14.Text = ""
        lblVE15.Text = ""
        lblVE16.Text = ""
        lblVE17.Text = ""
        lblVE2.Text = ""
        lblVE3.Text = ""
        lblVE4.Text = ""
        lblVE6.Text = ""
        lblVE7.Text = ""
        lblVE8.Text = ""
        lblVE9.Text = ""
        lblVLR1.Text = ""
        lblVLR2.Text = ""
        lblVLR3.Text = ""
        lblVLR4.Text = ""
        lblCLR5.Text = ""
        lblVLR6.Text = ""
        lblVLR7.Text = ""
        lblVLR8.Text = ""
        lblVLR9.Text = ""
        lblVT21.Text = ""
        lblVT22.Text = ""
        lblVT23.Text = ""
    End Function
    Function Claer_OEE()
        lblOEE10.Text = ""
        lblOEE11.Text = ""
        lblOEE12.Text = ""
        lblOEE13.Text = ""
        lblOEE14.Text = ""
        lblOEE15.Text = ""
        lblOEE16.Text = ""
        lblOEE17.Text = ""
        lblOEE2.Text = ""
        lblOEE3.Text = ""
        lblOEE4.Text = ""
        lblOEE6.Text = ""
        lblOEE7.Text = ""
        lblOEE8.Text = ""
        lblOEE9.Text = ""
        lblOEL1.Text = ""
        lblOEL2.Text = ""
        lblOEL3.Text = ""
        lblOEL4.Text = ""
        lblOEL5.Text = ""
        lblOEL6.Text = ""
        lblOEL7.Text = ""
        lblOEL8.Text = ""
        lblOEL9.Text = ""
        lblOET21.Text = ""
        lblOET22.Text = ""
        lblOET23.Text = ""
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        cmdAdd.Enabled = False
        txtFromDate.Text = Today
        txtTodate.Text = Today
        ' cmdSave.Enabled = True
        cmdEdit.Enabled = True

        'Dim pen As New Drawing.Pen(System.Drawing.Color.Red, 1)
        'Me.CreateGraphics.DrawEllipse(pen, 0, 0, 100, 100)
        'pen.Dispose()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim TS As TimeSpan
        Dim n_Stop As Date
        Dim n_Start As Date
        Dim _HR As Integer
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim I As Integer
        Dim M01 As DataSet


        n_Start = txtFromDate.Text
        n_Stop = txtTodate.Text
        n_Stop = CDate(n_Stop).AddDays(+1)
        TS = n_Stop - n_Start
        _HR = TS.Days * 24

        'CALANDER TIME
        lblCLR1.Text = _HR
        lblCLR2.Text = _HR
        lblCLR3.Text = _HR
        lblCLR4.Text = _HR
        lblCLR5.Text = _HR
        lblCLR6.Text = _HR
        lblCLR7.Text = _HR
        lblCLR8.Text = _HR
        lblCLR9.Text = _HR

        lblCE1.Text = _HR
        lblCE10.Text = _HR
        lblCE11.Text = _HR
        lblCE12.Text = _HR
        lblCE13.Text = _HR
        lblCE14.Text = _HR
        lblCE15.Text = _HR
        lblCE16.Text = _HR
        lblCE17.Text = _HR
        lblCE2.Text = _HR
        lblCE4.Text = _HR
        lblCE6.Text = _HR
        lblCE7.Text = _HR
        lblCE8.Text = _HR
        lblCE9.Text = _HR
        lblCT21.Text = _HR
        lblCT22.Text = _HR
        lblCT23.Text = _HR

        '-----------------------------------------------------------------
        'WORKING HOUR
        lblWE10.Text = _HR
        lblWE11.Text = _HR
        lblWE12.Text = _HR
        lblWE13.Text = _HR
        lblWE14.Text = _HR
        lblWE15.Text = _HR
        lblWE16.Text = _HR
        lblWE17.Text = _HR
        lblWE2.Text = _HR
        lblWE3.Text = _HR
        lblWE4.Text = _HR
        lblWE4.Text = _HR
        lblWE6.Text = _HR
        lblWE7.Text = _HR
        lblWE8.Text = _HR
        lblWE9.Text = _HR
        lblWLR1.Text = _HR
        lblWLR2.Text = _HR
        lblWLR3.Text = _HR
        lblWLR4.Text = _HR
        lblWLR5.Text = _HR
        lblWLR6.Text = _HR
        lblWLR7.Text = _HR
        lblWLR8.Text = _HR
        lblWLR9.Text = _HR
        lblWT21.Text = _HR
        lblWT22.Text = _HR
        lblWT23.Text = _HR
        '-----------------------------------------------------------------------
        Dim _Loading As Double
        _Loading = 0
        I = 0
        SQL = "select * from T03Machine where T03Type='02'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            SQL = "SELECT * FROM M012Setup_Time WHERE m012mccode='" & M01.Tables(0).Rows(I)("T03code") & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                _Loading = T01.Tables(0).Rows(0)("M012PM") + T01.Tables(0).Rows(0)("M012Other")
                _Loading = _Loading / 60  '// Hour
                _Loading = _Loading / 365
                _Loading = _Loading * TS.Days
            End If

            If Trim(M01.Tables(0).Rows(I)("T03name")) = "LR1" Then
                lblLLR1.Text = CInt(Val(lblWLR1.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR2" Then
                lblLLR2.Text = CInt(Val(lblWLR2.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR3" Then
                lblLLR3.Text = CInt(Val(lblWLR3.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR4" Then
                lblLLR4.Text = CInt(Val(lblWLR4.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR5" Then
                lblLLR5.Text = CInt(Val(lblWLR5.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR6" Then
                lblLLR6.Text = CInt(Val(lblWLR6.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR7" Then
                lblLLR7.Text = CInt(Val(lblWLR7.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR8" Then
                lblLLR8.Text = CInt(Val(lblWLR8.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR9" Then
                lblLLR9.Text = CInt(Val(lblWLR9.Text) - _Loading)
            End If
            I = I + 1
        Next
        '------------------------------------------------------------------------------------------
        I = 0
        SQL = "select * from T03Machine where T03Type='01'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            SQL = "SELECT * FROM M012Setup_Time WHERE m012mccode='" & M01.Tables(0).Rows(I)("T03code") & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                _Loading = T01.Tables(0).Rows(0)("M012PM") + T01.Tables(0).Rows(0)("M012Other")
                _Loading = _Loading / 60  '// Hour
                _Loading = _Loading / 365
                _Loading = _Loading * TS.Days
            End If

            If Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco2" Then
                lbllE2.Text = CInt(Val(lblWE2.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco3" Then
                lblLE3.Text = CInt(Val(lblWE3.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco4" Then
                lblLE4.Text = CInt(Val(lblWE4.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco6" Then
                lblLE6.Text = CInt(Val(lblWE6.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco7" Then
                lblLE7.Text = CInt(Val(lblWE7.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco8" Then
                lblLE8.Text = CInt(Val(lblWE8.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco9" Then
                lblLE9.Text = CInt(Val(lblWE9.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco10" Then
                lblLE10.Text = CInt(Val(lblWE10.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco11" Then
                lblLE11.Text = CInt(Val(lblWE11.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco12" Then
                lblLE12.Text = CInt(Val(lblWE12.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco13" Then
                lblLE13.Text = CInt(Val(lblWE13.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco14" Then
                lblLE14.Text = CInt(Val(lblWE14.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 15" Then
                lblLE15.Text = CInt(Val(lblWE15.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 16" Then
                lblLE16.Text = CInt(Val(lblWE16.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 17" Then
                lblLE17.Text = CInt(Val(lblWE17.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 21" Then
                lblLT21.Text = CInt(Val(lblWT21.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 22" Then
                lblLT22.Text = CInt(Val(lblWT22.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 23" Then
                lblLT23.Text = CInt(Val(lblWT23.Text) - _Loading)
            End If
            I = I + 1
        Next
        '----------------------------------------------------------------------------------------------
        'oparating time
        n_Start = txtFromDate.Text & " " & "7:30 AM"
        n_Stop = txtTodate.Text & " " & "7:30 AM"
        I = 0
        SQL = "select sum(M04Taken) as M04Taken,T03Name from M04Lot inner join T03Machine on M04Machine_No=T03Code where M04ProgrameType in ('B','I') and M04Etime between '7/1/2013 7:30:00 AM' and '7/2/2013 7:30:00 AM' group by T03Name "
        M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            _Loading = 0
            _Loading = M01.Tables(0).Rows(I)("M04Taken")
            _Loading = _Loading / 60
            If Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco2" Then
                lblOE2.Text = CInt(Val(lbllE2.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco3" Then
                lblOE3.Text = CInt(Val(lblLE3.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco4" Then
                lblOE4.Text = CInt(Val(lblLE4.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco6" Then
                lblOE6.Text = CInt(Val(lblLE6.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco7" Then
                lblOE7.Text = CInt(Val(lblLE7.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco8" Then
                lblOE8.Text = CInt(Val(lblLE8.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco9" Then
                lblOE9.Text = CInt(Val(lblLE9.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco10" Then
                lblOE10.Text = CInt(Val(lblLE10.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco11" Then
                lblOE11.Text = CInt(Val(lblLE11.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco12" Then
                lblOE12.Text = CInt(Val(lblLE12.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco13" Then
                lblOE13.Text = CInt(Val(lblLE13.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco14" Then
                lblOE14.Text = CInt(Val(lblLE14.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 15" Then
                lblOE15.Text = CInt(Val(lblLE15.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 16" Then
                lblOE16.Text = CInt(Val(lblLE16.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 17" Then
                lblOE17.Text = CInt(Val(lblLE17.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 21" Then
                lblOT21.Text = CInt(Val(lblLT21.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 22" Then
                lblOT22.Text = CInt(Val(lblLT22.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 23" Then
                lblOT23.Text = CInt(Val(lblLT23.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR1" Then
                lblOLR1.Text = CInt(Val(lblLLR1.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR2" Then
                lblOLR2.Text = CInt(Val(lblLLR2.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR3" Then
                lblOLR3.Text = CInt(Val(lblLLR3.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR4" Then
                lblOLR4.Text = CInt(Val(lblLLR4.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR5" Then
                lblOLR5.Text = CInt(Val(lblLLR5.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR6" Then
                lblOLR6.Text = CInt(Val(lblLLR6.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR7" Then
                lblOLR7.Text = CInt(Val(lblLLR7.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR8" Then
                lblOLR8.Text = CInt(Val(lblLLR8.Text) - _Loading)
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR9" Then
                lblOLR9.Text = CInt(Val(lblLLR9.Text) - _Loading)

            End If


            I = I + 1
        Next
        '---------------------------------------------------------------------------------------------
        I = 0
        Dim x As Integer
        Dim T02 As DataSet
        Dim T03 As DataSet
        Dim y As Integer

        Dim _Wash As Integer
        Dim _Stip As Integer
        Dim _Standed As Integer
        SQL = "select * from T03Machine where T03Type='02'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            SQL = "select sum(M04STD) as M04STD,M04ProgrameType,COUNT(M04ProgrameType) as BCount from M04Lot where M04ProgrameType in ('N','R','S','W','O','Y') and M04Machine_No='" & M01.Tables(0).Rows(I)("t03code") & "' and  M04Etime between '" & n_Start & "' and '" & n_Stop & "' group by M04ProgrameType"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            x = 0
            For Each DTRow2 As DataRow In T01.Tables(0).Rows
                _Wash = 0
                _Stip = 0
                ' _Standed = 0
                _Standed = _Standed + T01.Tables(0).Rows(x)("M04STD")
                SQL = "select * from M012Setup_Time where M012MCCode='" & M01.Tables(0).Rows(I)("t03code") & "'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T02) Then
                    If Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "W" Or Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "Y" Or Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "O" Then
                        _Wash = _Wash + (T02.Tables(0).Rows(0)("M012CExtra") * T01.Tables(0).Rows(x)("BCount"))
                    ElseIf Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "S" Then
                        _Stip = T02.Tables(0).Rows(0)("M012CDark") * T01.Tables(0).Rows(x)("BCount")
                    ElseIf Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "N" Or Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "R" Then
                        'POLYESTER
                        SQL = "select COUNT(M04Shade_Type)as nCount,M08SubName from M04Lot inner join M08Sub_Shade on M04Shade_Type=M08Code where M04Machine_No='" & M01.Tables(0).Rows(I)("t03code") & "' and " & _
                        "M04Etime between '" & n_Start & "' and '" & n_Stop & "'  " & _
                        "and M04ProgrameType in ('N','R') and M04Quality in ('54336','53582') group by M08SubName"
                        T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        y = 0
                        For Each DTRow1 As DataRow In T03.Tables(0).Rows
                            If Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Light" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FLight") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Medium" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FMedium") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FDark") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Extra Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FExtra") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "White" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012CLight") * T03.Tables(0).Rows(y)("nCount"))
                            End If
                            y = y + 1
                        Next
                        '-----------------------------------------------------------------------------------------------------
                        'COTTON and VISCOSE
                        SQL = "select COUNT(M04Shade_Type)as nCount,M08SubName from M04Lot inner join M08Sub_Shade on M04Shade_Type=M08Code where M04Machine_No='" & M01.Tables(0).Rows(I)("t03code") & "' and " & _
                     "M04Etime between '" & n_Start & "' and '" & n_Stop & "'  " & _
                     "and M04ProgrameType in ('N','R') and M04Quality NOT in ('54336','53582') group by M08SubName"
                        T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        y = 0
                        For Each DTRow1 As DataRow In T03.Tables(0).Rows
                            If Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Light" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DLight") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Medium" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DMedium") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DDark") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Extra Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DExtra") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "White" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012CLight") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "MARL & Yarn dye" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012CMedium") * T03.Tables(0).Rows(y)("nCount"))
                            End If
                            y = y + 1
                        Next
                    End If
                End If
                _Standed = _Standed - (_Wash + _Stip)
                x = x + 1
            Next
            _Standed = _Standed / 60

            _Standed = Microsoft.VisualBasic.Format(_Standed, "#.00")
            If Trim(M01.Tables(0).Rows(I)("T03name")) = "LR1" Then
                lblNLR1.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR2" Then
                lblNLR2.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR3" Then
                lblNLR3.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR4" Then
                lblNLR4.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR5" Then
                lblNLR5.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR6" Then
                lblNLR6.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR7" Then
                lblNLR7.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR8" Then
                lblNLR8.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "LR9" Then
                lblNLR9.Text = _Standed
            End If

            I = I + 1
        Next
        '-----------------------------------------------------------------------------
        'ECO
        I = 0
        SQL = "select * from T03Machine where T03Type='01'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow4 As DataRow In M01.Tables(0).Rows
            _Standed = 0
            SQL = "select sum(M04STD) as M04STD,M04ProgrameType,COUNT(M04ProgrameType) as BCount from M04Lot where M04ProgrameType in ('N','R','S','W','O','Y') and M04Machine_No='" & M01.Tables(0).Rows(I)("t03code") & "' and  M04Etime between '" & n_Start & "' and '" & n_Stop & "' group by M04ProgrameType"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            x = 0
            For Each DTRow2 As DataRow In T01.Tables(0).Rows
                _Wash = 0
                _Stip = 0
                ' _Standed = 0
                _Standed = _Standed + T01.Tables(0).Rows(x)("M04STD")
                SQL = "select * from M012Setup_Time where M012MCCode='" & M01.Tables(0).Rows(I)("t03code") & "'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T02) Then
                    If Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "W" Or Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "Y" Or Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "O" Then
                        _Wash = _Wash + (T02.Tables(0).Rows(0)("M012CExtra") * T01.Tables(0).Rows(x)("BCount"))
                    ElseIf Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "S" Then
                        _Stip = T02.Tables(0).Rows(0)("M012CDark") * T01.Tables(0).Rows(x)("BCount")
                    ElseIf Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "N" Or Trim(T01.Tables(0).Rows(x)("M04ProgrameType")) = "R" Then
                        'POLYESTER
                        SQL = "select COUNT(M04Shade_Type)as nCount,M08SubName from M04Lot inner join M08Sub_Shade on M04Shade_Type=M08Code where M04Machine_No='" & M01.Tables(0).Rows(I)("t03code") & "' and " & _
                        "M04Etime between '" & n_Start & "' and '" & n_Stop & "'  " & _
                        "and M04ProgrameType in ('N','R') and M04Quality in ('54336','53582') group by M08SubName"
                        T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        y = 0
                        For Each DTRow1 As DataRow In T03.Tables(0).Rows
                            If Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Light" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FLight") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Medium" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FMedium") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FDark") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Extra Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012FExtra") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "White" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012CLight") * T03.Tables(0).Rows(y)("nCount"))
                            End If
                            y = y + 1
                        Next
                        '-----------------------------------------------------------------------------------------------------
                        'COTTON and VISCOSE
                        SQL = "select COUNT(M04Shade_Type)as nCount,M08SubName from M04Lot inner join M08Sub_Shade on M04Shade_Type=M08Code where M04Machine_No='" & M01.Tables(0).Rows(I)("t03code") & "' and " & _
                     "M04Etime between '" & n_Start & "' and '" & n_Stop & "'  " & _
                     "and M04ProgrameType in ('N','R') and M04Quality NOT in ('54336','53582') group by M08SubName"
                        T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        y = 0
                        For Each DTRow1 As DataRow In T03.Tables(0).Rows
                            If Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Light" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DLight") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Medium" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DMedium") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DDark") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "Extra Dark" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012DExtra") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "White" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012CLight") * T03.Tables(0).Rows(y)("nCount"))
                            ElseIf Trim(T03.Tables(0).Rows(y)("M08SubName")) = "MARL & Yarn dye" Then
                                _Standed = _Standed - (T02.Tables(0).Rows(0)("M012CMedium") * T03.Tables(0).Rows(y)("nCount"))
                            End If
                            y = y + 1
                        Next
                    End If
                End If
                _Standed = _Standed - (_Wash + _Stip)
                x = x + 1
            Next
            _Standed = _Standed / 60
            _Standed = Microsoft.VisualBasic.Format(_Standed, "#.00")
            If Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco2" Then
                lblNE2.Text = _Standed
                lblVE2.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco3" Then
                lblNE3.Text = _Standed
                lblVE3.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco4" Then
                lblNE4.Text = _Standed
                lblVE4.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco6" Then
                lblNE6.Text = _Standed
                lblVE6.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco7" Then
                lblNE7.Text = _Standed
                lblVE7.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco8" Then
                lblNE8.Text = _Standed
                lblVE8.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco9" Then
                lblNE9.Text = _Standed
                lblVE9.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco10" Then
                lblNE10.Text = _Standed
                lblVE10.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco11" Then
                lblNE11.Text = _Standed
                lblVE11.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco12" Then
                lblNE12.Text = _Standed
                lblVE12.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco13" Then
                lblNE13.Text = _Standed
                lblVE13.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco14" Then
                lblNE14.Text = _Standed
                lblVE14.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 15" Then
                lblNE15.Text = _Standed
                lblVE15.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 16" Then
                lblNE16.Text = _Standed
                lblVE16.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "Eco 17" Then
                lblNE17.Text = _Standed
                lblVE17.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 21" Then
                lblNT21.Text = _Standed
                lblVT21.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 22" Then
                lblNT22.Text = _Standed
                lblVT22.Text = _Standed
            ElseIf Trim(M01.Tables(0).Rows(I)("T03name")) = "IT 23" Then
                lblNT23.Text = _Standed
                lblVT23.Text = _Standed
            End If
            I = I + 1
        Next
        '----------------------------------------------------------------------------------------
        'VALUE ADDED OPERATING TIME
        lblVLR1.Text = lblNLR1.Text
        lblVLR2.Text = lblNLR2.Text
        lblVLR3.Text = lblNLR3.Text
        lblVLR4.Text = lblNLR4.Text
        lblVLR5.Text = lblNLR5.Text
        lblVLR6.Text = lblNLR6.Text
        lblVLR7.Text = lblNLR7.Text
        lblVLR8.Text = lblNLR8.Text
        lblVLR9.Text = lblNLR9.Text



        Dim _VALUE As Double
        SQL = "select T03Name,sum(M04Batchwt) as M04Batchwt,sum(T03Reject) as T03Reject,sum(M04STD)as M04STD from M04Lot  " & _
              "inner join T03DNH on M04Ref=T03Ecode inner join T03Machine on T03Code=M04Machine_No where M04ProgrameType='N' and M04Etime between " & _
              " '" & n_Start & "' and '" & n_Stop & "' group by T03Name"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        I = 0
        For Each DTRow1 As DataRow In T01.Tables(0).Rows
            _VALUE = 0
            _VALUE = T01.Tables(0).Rows(I)("T03Reject") / T01.Tables(0).Rows(I)("M04Batchwt")
            _VALUE = _VALUE * T01.Tables(0).Rows(I)("M04STD")
            _VALUE = _VALUE / 60
            'LR MACHINE
            If Trim(T01.Tables(0).Rows(I)("T03name")) = "LR1" Then
                lblVLR1.Text = CInt(Val(lblNLR1.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR2" Then
                lblVLR2.Text = CInt(Val(lblNLR2.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR3" Then
                lblVLR3.Text = CInt(Val(lblNLR3.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR4" Then
                lblVLR4.Text = CInt(Val(lblNLR4.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR5" Then
                lblVLR5.Text = CInt(Val(lblNLR5.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR6" Then
                lblVLR6.Text = CInt(Val(lblNLR6.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR7" Then
                lblVLR7.Text = CInt(Val(lblNLR7.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR8" Then
                lblVLR8.Text = CInt(Val(lblNLR8.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "LR9" Then
                lblVLR9.Text = CInt(Val(lblNLR9.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco2" Then

                lblVE2.Text = CInt(Val(lblNE2.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco3" Then
                lblVE3.Text = CInt(Val(lblNE3.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco4" Then
                lblVE4.Text = CInt(Val(lblNE4.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco6" Then
                lblVE6.Text = CInt(Val(lblNE6.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco7" Then

                lblVE7.Text = CInt(Val(lblNE7.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco8" Then
                lblVE8.Text = CInt(Val(lblNE8.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco9" Then
                lblVE9.Text = CInt(Val(lblNE9.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco10" Then
                lblVE10.Text = CInt(Val(lblNE10.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco11" Then
                lblVE11.Text = CInt(Val(lblNE11.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco12" Then
                lblVE12.Text = CInt(Val(lblNE12.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco13" Then
                lblVE13.Text = CInt(Val(lblNE13.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco14" Then
                lblVE14.Text = CInt(Val(lblNE14.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco 15" Then
                lblVE15.Text = CInt(Val(lblNE15.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco 16" Then

                lblVE16.Text = CInt(Val(lblNE16.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "Eco 17" Then

                lblVE17.Text = CInt(Val(lblNE17.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "IT 21" Then
                lblVT21.Text = CInt(Val(lblNT21.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "IT 22" Then

                lblVT22.Text = CInt(Val(lblNT22.Text) - _VALUE)
            ElseIf Trim(T01.Tables(0).Rows(I)("T03name")) = "IT 23" Then

                lblVT23.Text = CInt(Val(lblNT23.Text) - _VALUE)
            End If

            I = I + 1
        Next
        '--------------------------------------------------------
        'OEE
        lblOEL1.Text = CInt((lblVLR1.Text / lblLLR1.Text) * 100) & "%"
        lblOEL2.Text = CInt((lblVLR2.Text / lblLLR2.Text) * 100) & "%"
        lblOEL3.Text = CInt((lblVLR3.Text / lblLLR3.Text) * 100) & "%"
        lblOEL4.Text = CInt((lblVLR4.Text / lblLLR4.Text) * 100) & "%"
        lblOEL5.Text = CInt((lblVLR5.Text / lblLLR5.Text) * 100) & "%"
        lblOEL6.Text = CInt((lblVLR6.Text / lblLLR6.Text) * 100) & "%"
        lblOEL7.Text = CInt((lblVLR7.Text / lblLLR7.Text) * 100) & "%"
        lblOEL8.Text = CInt((lblVLR8.Text / lblLLR8.Text) * 100) & "%"
        lblOEL9.Text = CInt((lblVLR9.Text / lblLLR9.Text) * 100) & "%"

        lblOEE2.Text = CInt((lblVE2.Text / lbllE2.Text) * 100) & "%"
        lblOEE3.Text = CInt((lblVE3.Text / lblLE3.Text) * 100) & "%"
        lblOEE4.Text = CInt((lblVE4.Text / lblLE4.Text) * 100) & "%"
        lblOEE6.Text = CInt((lblVE6.Text / lblLE6.Text) * 100) & "%"
        lblOEE7.Text = CInt((lblVE7.Text / lblLE7.Text) * 100) & "%"
        lblOEE8.Text = CInt((lblVE8.Text / lblLE8.Text) * 100) & "%"
        lblOEE9.Text = CInt((lblVE9.Text / lblLE9.Text) * 100) & "%"
        lblOEE10.Text = CInt((lblVE10.Text / lblLE10.Text) * 100) & "%"
        lblOEE11.Text = CInt((lblVE11.Text / lblLE11.Text) * 100) & "%"
        lblOEE12.Text = CInt((lblVE12.Text / lblLE12.Text) * 100) & "%"
        lblOEE13.Text = CInt((lblVE13.Text / lblLE13.Text) * 100) & "%"
        lblOEE14.Text = CInt((lblVE14.Text / lblLE14.Text) * 100) & "%"
        lblOEE15.Text = CInt((lblVE15.Text / lblLE15.Text) * 100) & "%"
        lblOEE16.Text = CInt((lblVE16.Text / lblLE16.Text) * 100) & "%"
        lblOEE17.Text = CInt((lblVE17.Text / lblLE17.Text) * 100) & "%"
        lblOET21.Text = CInt((lblVT21.Text / lblLT21.Text) * 100) & "%"
        lblOET22.Text = CInt((lblVT22.Text / lblLT22.Text) * 100) & "%"
        lblOET23.Text = CInt((lblVT23.Text / lblLT23.Text) * 100) & "%"
        '-------------------------------------------------------------------------
        'SUMMERY
        lblCA.Text = Val(lblCLR1.Text) + Val(lblCLR2.Text) + Val(lblCLR3.Text) + Val(lblCLR4.Text) + Val(lblCLR5.Text) + Val(lblCLR6.Text) + Val(lblCLR7.Text) + Val(lblCLR8.Text) + Val(lblCLR9.Text)
        lblCW.Text = Val(lblCE1.Text) * 18
        lblCT.Text = Val(lblCA.Text) + Val(lblCW.Text)

        lblWA.Text = Val(lblWLR1.Text) * 9
        lblWW.Text = Val(lblWE2.Text) * 18
        lblWT.Text = Val(lblWA.Text) + Val(lblWW.Text)

        lblLA.Text = Val(lblLLR1.Text) + Val(lblLLR2.Text) + Val(lblLLR3.Text) + Val(lblLLR4.Text) + Val(lblLLR5.Text) + Val(lblLLR6.Text) + Val(lblLLR7.Text) + Val(lblLLR8.Text) + Val(lblLLR9.Text)
        lblLW.Text = Val(lbllE2.Text) + Val(lblLE3.Text) + Val(lblLE4.Text) + Val(lblLE12.Text) + Val(lblLE6.Text) + Val(lblLE7.Text) + Val(lblLE8.Text) + Val(lblLE9.Text) + Val(lblLE10.Text) + Val(lblLE11.Text) + Val(lblLE13.Text) + Val(lblLE14.Text) + Val(lblLE14.Text) + Val(lblLE16.Text) + Val(lblLE17.Text) + Val(lblLT21.Text) + Val(lblLT22.Text) + Val(lblLT23.Text)
        lblLT.Text = Val(lblLA.Text) + Val(lblLW.Text)

        lblOA.Text = Val(lblOLR1.Text) + Val(lblOLR2.Text) + Val(lblOLR3.Text) + Val(lblOLR4.Text) + Val(lblOLR5.Text) + Val(lblOLR6.Text) + Val(lblOLR7.Text) + Val(lblOLR8.Text) + Val(lblOLR9.Text)
        lblOW.Text = Val(lblOE2.Text) + Val(lblOE3.Text) + Val(lblOE4.Text) + Val(lblOE12.Text) + Val(lblOE6.Text) + Val(lblOE7.Text) + Val(lblOE8.Text) + Val(lblOE9.Text) + Val(lblOE10.Text) + Val(lblOE11.Text) + Val(lblOE13.Text) + Val(lblOE14.Text) + Val(lblOE14.Text) + Val(lblOE16.Text) + Val(lblOE17.Text) + Val(lblOT21.Text) + Val(lblOT22.Text) + Val(lblOT23.Text)
        lblOT.Text = Val(lblOA.Text) + Val(lblOW.Text)

        lblNA.Text = Val(lblNLR1.Text) + Val(lblNLR2.Text) + Val(lblNLR3.Text) + Val(lblNLR4.Text) + Val(lblNLR5.Text) + Val(lblNLR6.Text) + Val(lblNLR7.Text) + Val(lblNLR8.Text) + Val(lblNLR9.Text)
        lblNW.Text = Val(lblNE2.Text) + Val(lblNE3.Text) + Val(lblNE4.Text) + Val(lblNE12.Text) + Val(lblNE6.Text) + Val(lblNE7.Text) + Val(lblNE8.Text) + Val(lblNE9.Text) + Val(lblNE10.Text) + Val(lblNE11.Text) + Val(lblNE13.Text) + Val(lblNE14.Text) + Val(lblNE14.Text) + Val(lblNE16.Text) + Val(lblNE17.Text) + Val(lblNT21.Text) + Val(lblNT22.Text) + Val(lblNT23.Text)
        lblNT.Text = Val(lblNA.Text) + Val(lblNW.Text)

        lblVA.Text = Val(lblVLR1.Text) + Val(lblVLR2.Text) + Val(lblVLR3.Text) + Val(lblVLR4.Text) + Val(lblVLR5.Text) + Val(lblVLR6.Text) + Val(lblVLR7.Text) + Val(lblVLR8.Text) + Val(lblVLR9.Text)
        lblVW.Text = Val(lblVE2.Text) + Val(lblVE3.Text) + Val(lblVE4.Text) + Val(lblVE12.Text) + Val(lblVE6.Text) + Val(lblVE7.Text) + Val(lblVE8.Text) + Val(lblVE9.Text) + Val(lblVE10.Text) + Val(lblVE11.Text) + Val(lblVE13.Text) + Val(lblVE14.Text) + Val(lblVE14.Text) + Val(lblVE16.Text) + Val(lblVE17.Text) + Val(lblVT21.Text) + Val(lblVT22.Text) + Val(lblVT23.Text)
        lblVT.Text = Val(lblVA.Text) + Val(lblVW.Text)

        lblOEA.Text = CInt(Val(lblVA.Text) / Val(lblLA.Text) * 100) & "%"
        lblOEW.Text = CInt(Val(lblVW.Text) / Val(lblLW.Text) * 100) & "%"
        lblOET.Text = CInt(Val(lblVT.Text) / Val(lblLT.Text) * 100) & "%"
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_CalanderLR()
        Call Clear_WorkingTime()
        Call Clear_LoadingTime()
        Call Claer_OparatingTime()
        Call Clear_Net()
        Call Clear_Value()
        Call Claer_OEE()
        Call Clear_Other()
    End Sub

    Function Clear_Other()
        lblCA.Text = ""
        lblCW.Text = ""
        lblCT.Text = ""
        lblWA.Text = ""
        lblWW.Text = ""
        lblWT.Text = ""
        lblOA.Text = ""
        lblOW.Text = ""
        lblOT.Text = ""
        lblNA.Text = ""
        lblNW.Text = ""
        lblNT.Text = ""
        lblVA.Text = ""
        lblVW.Text = ""
        lblVT.Text = ""
        lblOEA.Text = ""
        lblOEW.Text = ""
        lblOET.Text = ""
        lblLA.Text = ""
        lblLW.Text = ""
        lblLT.Text = ""
    End Function
    Private Sub Label63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblVW.Click

    End Sub

    Private Sub Label31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label31.Click

    End Sub

    Private Sub Label31_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label31.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblA.Visible = False
    End Sub

    Private Sub Label31_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label31.MouseMove
        Me.Cursor = Cursors.Hand
        lblA.Visible = True
    End Sub

    Private Sub Label32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label32.Click

    End Sub

    Private Sub Label32_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label32.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblB.Visible = False
    End Sub

    Private Sub Label32_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label32.MouseMove
        Me.Cursor = Cursors.Hand
        lblB.Visible = True
    End Sub

    Private Sub Label33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label33.Click

    End Sub

    Private Sub Label33_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label33.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblC.Visible = False
    End Sub

    Private Sub Label33_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label33.MouseMove
        Me.Cursor = Cursors.Hand
        lblC.Visible = True
    End Sub

    Private Sub Label34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label34.Click

    End Sub

    Private Sub Label34_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label34.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblD.Visible = False
    End Sub

    Private Sub Label34_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label34.MouseMove
        Me.Cursor = Cursors.Hand
        lblD.Visible = True
    End Sub

    Private Sub Label35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label35.Click

    End Sub

    Private Sub Label35_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label35.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblE.Visible = False
    End Sub

    Private Sub Label35_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label35.MouseMove
        Me.Cursor = Cursors.Hand
        lblE.Visible = True
    End Sub

    Private Sub Label36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label36.Click

    End Sub

    Private Sub Label36_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label36.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblF.Visible = False
    End Sub

    Private Sub Label36_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label36.MouseMove
        Me.Cursor = Cursors.Hand
        lblF.Visible = True
    End Sub

    Private Sub Label40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label40.Click

    End Sub

    Private Sub Label40_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label40.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblG.Visible = False
    End Sub

    Private Sub Label40_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label40.MouseMove
        Me.Cursor = Cursors.Hand
        lblG.Visible = True
    End Sub

    Private Sub Label39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label39.Click

    End Sub

    Private Sub Label39_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label39.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblA1.Visible = False

    End Sub

    Private Sub Label39_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label39.MouseMove
        Me.Cursor = Cursors.Hand
        lblA1.Visible = True
    End Sub

    Private Sub Label38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label38.Click

    End Sub

    Private Sub Label38_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label38.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblB1.Visible = False
    End Sub

    Private Sub Label38_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label38.MouseMove
        Me.Cursor = Cursors.Hand
        lblB1.Visible = True
    End Sub

    Private Sub Label37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label37.Click
       
    End Sub

    Private Sub Label37_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label37.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblC1.Visible = False
    End Sub

    Private Sub Label37_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label37.MouseMove
        Me.Cursor = Cursors.Hand
        lblC1.Visible = True
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblD1.Visible = False
    End Sub

    Private Sub Label3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label3.MouseMove
        Me.Cursor = Cursors.Hand
        lblD1.Visible = True
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
       
    End Sub

    Private Sub Label2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblE1.Visible = False
    End Sub

    Private Sub Label2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseMove
        Me.Cursor = Cursors.Hand
        lblE1.Visible = True
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblF1.Visible = False
    End Sub

    Private Sub Label1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseMove
        Me.Cursor = Cursors.Hand
        lblF1.Visible = True
    End Sub

    Private Sub Label41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label41.Click

    End Sub

    Private Sub Label41_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label41.MouseLeave
        Me.Cursor = Cursors.Arrow
        lblG1.Visible = False
    End Sub

    Private Sub Label41_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label41.MouseMove
        Me.Cursor = Cursors.Hand
        lblG1.Visible = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter
        Dim exc As New Application

        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
   


        workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet1.Name = "OEE Report"


        worksheet1.Cells(2, 1) = "OEE Report"
        worksheet1.Rows(2).Font.Bold = True
        worksheet1.Rows(2).Font.size = 12
        worksheet1.Rows(2).Font.Name = "Times New Roman"

        worksheet1.Rows(2).rowheight = 20.25


        worksheet1.Rows(2).Font.Bold = True
        worksheet1.Rows(2).Font.size = 11
        worksheet1.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
        'range1 = worksheet1.Cells("B2:G2")

    End Sub
End Class