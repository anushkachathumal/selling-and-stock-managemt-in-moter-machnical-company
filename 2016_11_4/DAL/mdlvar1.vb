Module mdlvar1

    Public Function Num2String(ByRef Qty As String) As String
        On Error GoTo flgerr
        'Numero a string
        Qty = VB6.Format(Qty, "####0.00")
        Dim temp As String
        If InStr(1, Qty, ".") > 0 Then
            temp = Right(Qty, Len(Qty) - InStr(1, Qty, "."))
            ' Qty = CStr(Fix(CShort(Qty)))
            Qty = Fix(Qty)
        End If
        Select Case Len(Qty)
            Case 1
                Num2String = Unidades(Qty, True)
            Case 2
                Num2String = Decenas(Qty)
            Case 3
                Num2String = Centenas(Qty)
            Case 4
                Num2String = UniMiles(Qty)
            Case 5
                Num2String = DecMiles(Qty)
            Case 6
                Num2String = CenMiles(Qty)
            Case 7
                'UPGRADE_WARNING: Couldn't resolve default property of object UniMillon(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Num2String = UniMillon(Qty)
        End Select

        If Len(temp) > 0 Then
            If CInt(temp) <> 0 Then
                Select Case Len(temp)
                    Case 1
                        temp = Unidades(temp, True)
                    Case 2
                        temp = Decenas(temp)
                    Case 3
                        temp = Centenas(temp)
                    Case 4
                        temp = UniMiles(temp)
                    Case 5
                        temp = DecMiles(temp)
                    Case 6
                        temp = CenMiles(temp)
                    Case 7
                        'UPGRADE_WARNING: Couldn't resolve default property of object UniMillon(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        temp = UniMillon(temp)
                End Select

                Num2String = Num2String & " and  " & temp & "Cents Only"
            Else
                Num2String = Num2String & " Only"
            End If
        End If
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function
    Private Function Unidades(ByRef Qty As String, Optional ByRef Un As Boolean = False) As String
        On Error GoTo flgerr
        'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Unidades = Choose(CInt(Qty) + 1, "Zero", IIf(Un, "One", "Uno"), "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

    Private Function Decenas(ByRef Qty As String) As String
        On Error GoTo flgerr
        Select Case CShort(Qty)
            Case 0
                Decenas = vbNullString
            Case 10
                Decenas = "Ten"
            Case 11
                Decenas = "Eleven"
            Case 12
                Decenas = "Twelve"
            Case 13
                Decenas = "Thirteen"
            Case 14
                Decenas = "Fourteen"
            Case 15
                Decenas = "Fifteen"
            Case 16
                Decenas = "Sixteen"
            Case 17
                Decenas = "Seventeen"
            Case 18
                Decenas = "Eighteen"
            Case 19
                Decenas = "Nineteen"
            Case 20
                Decenas = "Twenty"
            Case 21
                Decenas = "Twenty One"
            Case 22
                Decenas = "Twenty Two"
            Case 23
                Decenas = "Twenty Three"
            Case 24
                Decenas = "Twenty Four"
            Case 25
                Decenas = "Twenty Five"
            Case 26
                Decenas = "Twenty Six"
            Case 27
                Decenas = "Twenty Seven"
            Case 28
                Decenas = "Twenty Eight"
            Case 29
                Decenas = "Twenty Nine"
            Case 30
                Decenas = "Thirty"
            Case 40
                Decenas = "Fourty"
            Case 50
                Decenas = "Fifty"
            Case 60
                Decenas = "Sixty"
            Case 70
                Decenas = "Seventy"
            Case 80
                Decenas = "Eighty"
            Case 90
                Decenas = "Ninety"
            Case Else
                'Decenas = Decenas(Left$(Qty, 1) * 10) & IIf(Left$(Qty, 1) = 0, "", Space(1) & "y" + Space(1)) & Unidades(Right$(Qty, 1), True)
                Decenas = Decenas(CDbl(Left(Qty, 1)) * 10) & IIf(CDbl(Left(Qty, 1)) = 0, "", Space(1)) & Unidades(Right(Qty, 1), True)
        End Select
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

    Private Function Centenas(ByRef Qty As String) As String
        On Error GoTo flgerr
        Select Case Left(Qty, 1)
            Case CStr(1)
                If Qty = "100" Then : Centenas = "One Hundred" : Else : Centenas = "Hundred and"
                End If
            Case CStr(2)
                Centenas = "Two Hundred"
            Case CStr(3)
                Centenas = "Three Hundred"
            Case CStr(4)
                Centenas = "Four Hundred"
            Case CStr(5)
                Centenas = "Five Hundred"
            Case CStr(6)
                Centenas = "Six Hundred"
            Case CStr(7)
                Centenas = "Seven Hundred"
            Case CStr(8)
                Centenas = "Eight Hundred"
            Case CStr(9)
                Centenas = "Nine Hundred"
        End Select
        Centenas = Trim(Centenas & Space(1) & Trim(Decenas(Right(Qty, 2))))
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

    Private Function UniMiles(ByRef Qty As String) As String
        On Error GoTo flgerr
        UniMiles = Trim(Unidades(Left(Qty, 1), True) & " Thousand" & Space(1) & Trim(Centenas(Right(Qty, 3))))
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

    Private Function DecMiles(ByRef Qty As String) As String
        On Error GoTo flgerr
        DecMiles = Trim(Decenas(Left(Qty, 2)) & " Thousand " & Centenas(Right(Qty, 3)))
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

    Private Function CenMiles(ByRef Qty As String) As String
        On Error GoTo flgerr
        CenMiles = Trim(IIf(Len(Centenas(Left(Qty, 3))) = 0, Centenas(Right(Qty, 3)), Centenas(Left(Qty, 3)) & " thousand " & Centenas(Right(Qty, 3))))
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

    Private Function UniMillon(ByRef Qty As String) As Object
        On Error GoTo flgerr
        'UPGRADE_WARNING: Couldn't resolve default property of object UniMillon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        UniMillon = CDbl(Unidades(Left(Qty, 1), IIf(CDbl(Left(Qty, 1)) = 1, True, False))) + IIf(CDbl(Left(Qty, 1)) = 1, " Million ", " Millones ") + CDbl(CenMiles(Right(Qty, 6)))
        Exit Function
flgerr:
        '    ErrTrap "miscelaneous", Err.Description, Err.Number
    End Function

End Module

