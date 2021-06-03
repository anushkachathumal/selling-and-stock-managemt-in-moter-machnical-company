
Imports System.Management
Imports System.Data.SqlClient
Imports DBLotVbnet.DBEngin


Public Class common
    Public Shared netCard As String
    Public Shared LoggedUser As String
    Public Shared VDominLoggedUser As String
    Public Shared VIP As String
    Public Shared VMName As String
    Public Shared VDataSource As String
    Public Shared VDataBase As String
    Public Shared VserverTime As DateTime

    Public Shared Function Ins_Length(ByVal Field As String) As String

        If Field = "Tkt_No" Then : Return "1, 8"
        ElseIf Field = "Str1" Then : Return "1, 8,9, 8,17, 8,25, 8,33, 8,41, 8,49, 8,57, 10,67, 3"
        ElseIf Field = "Total" Then : Return "69"
        End If


    End Function

    Public Shared Function GetMachineTime(ByVal machine As String) As String

        If Not machine.StartsWith("\\") Then
            machine = "\\" & machine
        End If

        Dim searcher As New ManagementObjectSearcher(machine & "\root\CIMV2", _
            "SELECT * FROM Win32_LocalTime")

        For Each timeObject As ManagementObject In searcher.Get()

            Dim hour As Integer = Convert.ToInt32(timeObject("Hour"))
            Dim minute As Integer = Convert.ToInt32(timeObject("Minute"))
            Dim second As Integer = Convert.ToInt32(timeObject("Second"))

            Dim time As New TimeSpan(hour, minute, second)

            Return time.ToString()

        Next

    End Function 'GetMachineTime'



    Public Shared Function ClearAll(ByVal ParamArray Sforms As Infragistics.Win.Misc.UltraGroupBox())
        On Error Resume Next
        Dim Ctl As Object
        Dim Sform As Infragistics.Win.Misc.UltraGroupBox

        For Each Sform In Sforms
            For Each Ctl In Sform.Controls

                If TypeOf Ctl Is System.Windows.Forms.TextBox Then Ctl.Text = ""
                If TypeOf Ctl Is System.Windows.Forms.ComboBox Then Ctl.Selectedindex = -1
                If TypeOf Ctl Is System.Windows.Forms.DateTimePicker Then Ctl.value = Date.Now
                If TypeOf Ctl Is System.Windows.Forms.CheckBox Then Ctl.checked = False
                If TypeOf Ctl Is System.Windows.Forms.RadioButton Then Ctl.checked = False
                If TypeOf Ctl Is System.Windows.Forms.ListBox Then Ctl.SelectedIndex = -1
                If TypeOf Ctl Is System.Windows.Forms.RichTextBox Then Ctl.Text = ""
                If TypeOf Ctl Is Infragistics.Win.UltraWinEditors.UltraTextEditor Then Ctl.Text = ""
                If TypeOf Ctl Is Infragistics.Win.UltraWinEditors.UltraComboEditor Then Ctl.Text = ""
                If TypeOf Ctl Is Infragistics.Win.UltraWinGrid.UltraCombo Then Ctl.Text = ""

            Next

            If Mid(Sform.Name, 1, 3) = "OPR" Then
                Sform.Enabled = False
            End If

            '  MDIMain.UltraStatusBar1.Panels(1).Text = ""

        Next

        
    End Function


    Public Shared Function isValidDataset(ByRef dataset As DataSet) As Boolean

        If dataset Is Nothing Then Return False
        If dataset.Tables.Count <= 0 Then Return False
        If dataset.Tables(0).Rows.Count <= 0 Then Return False
        Return True

    End Function

    Public Function isValidDatatable(ByRef dataTable As DataTable) As Boolean

        If dataTable Is Nothing Then Return False
        If dataTable.Rows.Count <= 0 Then Return False
        Return True

    End Function



    Public Shared Function SpellNumber(ByVal MyNumber As String)
        Dim Dollars, Cents, Temp
        Dim DecimalPlace, Count

        Dim Place(9) As String
        Place(2) = " Thousand "
        Place(3) = " Million "
        Place(4) = " Billion "
        Place(5) = " Trillion "

        ' String representation of amount
        MyNumber = Convert.ToString(MyNumber)

        ' Position of decimal place 0 if none
        DecimalPlace = InStr(MyNumber, ".")
        'Convert cents and set MyNumber to dollar amount
        If DecimalPlace > 0 Then
            Cents = GetTens(Microsoft.VisualBasic.Strings.Left(Mid(MyNumber, DecimalPlace + 1) & _
                                 "00", 2))
            MyNumber = Trim(Microsoft.VisualBasic.Strings.Left(MyNumber, DecimalPlace - 1))
        End If

        Count = 1
        Do While MyNumber <> ""
            Temp = GetHundreds(Microsoft.VisualBasic.Strings.Right(MyNumber, 3))
            If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
            If Len(MyNumber) > 3 Then
                MyNumber = Microsoft.VisualBasic.Strings.Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop

        Select Case Dollars
            Case ""
                Dollars = "zero Rupees"
            Case "One"
                Dollars = "One Rupees"
            Case Else
                Dollars = Dollars & " Rupees"
        End Select

        Select Case Cents
            Case ""
                Cents = " and zero Cents"
            Case "One"
                Cents = " and One Cent"
            Case Else
                Cents = " and " & Cents & " Cents"
        End Select

        SpellNumber = Dollars & Cents
    End Function

    Private Shared Function GetTens(ByVal TensText As String)
        Dim Result As String

        Result = ""           'null out the temporary function value
        If Val(Microsoft.VisualBasic.Strings.Left(TensText, 1)) = 1 Then   ' If value between 10-19
            Select Case Val(TensText)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else                               ' If value between 20-99
            Select Case Val(Microsoft.VisualBasic.Strings.Left(TensText, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select
            Result = Result & GetDigit(Microsoft.VisualBasic.Strings.Right(TensText, 1))
            'Retrieve ones place
        End If
        GetTens = Result
    End Function

    Private Shared Function GetDigit(ByVal Digit As String)
        Select Case Val(Digit)
            Case 1 : GetDigit = "One"
            Case 2 : GetDigit = "Two"
            Case 3 : GetDigit = "Three"
            Case 4 : GetDigit = "Four"
            Case 5 : GetDigit = "Five"
            Case 6 : GetDigit = "Six"
            Case 7 : GetDigit = "Seven"
            Case 8 : GetDigit = "Eight"
            Case 9 : GetDigit = "Nine"
            Case Else : GetDigit = ""
        End Select
    End Function
    Private Shared Function GetHundreds(ByVal MyNumber As String)
        Dim Result As String

        If Val(MyNumber) = 0 Then Exit Function
        MyNumber = Microsoft.VisualBasic.Strings.Right("000" & MyNumber, 3)

        'Convert the hundreds place
        If Mid(MyNumber, 1, 1) <> "0" Then
            Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
        End If

        'Convert the tens and ones place
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & GetTens(Mid(MyNumber, 2))
        Else
            Result = Result & GetDigit(Mid(MyNumber, 3))
        End If

        GetHundreds = Result
    End Function

    Public Shared Function Get_highestVouno(ByVal VQryType As String, ByVal VVouType As String) As String
        Dim con = New SqlConnection()
        Dim dsUser As DataSet
        Dim cMax As String

        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetParameter", New SqlParameter("@cQryType", VQryType), New SqlParameter("@vcCode", "BG"))
            If common.isValidDataset(dsUser) Then
                For Each DTRow As DataRow In dsUser.Tables(0).Rows
                    cMax = dsUser.Tables(0).Rows(0)("P01LastNo")
                    Return cMax
                Next
            Else
                MessageBox.Show("Record Not Found", "LMS", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            '====================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
        '=========================================================================
    End Function

    Public Shared Function Getservertime() As Date
        Dim con1 = New SqlConnection()
        Dim dsUser As DataSet
        Dim sqla As String
        Try
            SqlConnection.ClearAllPools()
            con1 = DBEngin.GetConnection()

            dsUser = DBEngin.ExecuteDataset(con1, Nothing, "SELECT GETDATE()as [Server time]")
            If isValidDataset(dsUser) = True Then
                VserverTime = dsUser.Tables(0).Rows(0)("Server time")
            End If
            Return VserverTime

            con1.ClearAllPools()
            con1.CLOSE()

        Catch
            MsgBox("Unable to return Database Details", MsgBoxStyle.Critical, "LMS")
            Exit Function
        End Try
    End Function

    
End Class


