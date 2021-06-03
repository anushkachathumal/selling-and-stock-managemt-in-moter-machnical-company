Module myReplace
    Public Enum types As Byte
        one
        two
        three
    End Enum

    Public Function [do](ByVal text As String, ByVal type As types) As String
        Select Case type
            Case types.one
                Return text.Replace(Chr(39), Chr(39) + Chr(39)) ' "
            Case types.two
                Return text.Replace(Chr(39), Chr(39) + Chr(39)) ' '
            Case types.three
                'so on ...
        End Select

    End Function
End Module
