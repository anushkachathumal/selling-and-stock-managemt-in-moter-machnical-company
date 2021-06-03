Module Module1
    Public Declare Auto Function SetCursorPos Lib "User32.dll" (ByVal X As Integer, ByVal Y As Integer) As Integer
    Public Declare Auto Function GetCursorPos Lib "User32.dll" (ByRef lpPoint As Point) As Integer
    Public Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)

    'Public Const MOUSEEVENTF_LEFTDOWN = &H2
    'Public Const MOUSEEVENTF_LEFTUP = &H4
    'Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
    'Public Const MOUSEEVENTF_MIDDLEUP = &H40
    'Public Const MOUSEEVENTF_RIGHTDOWN = &H8
    'Public Const MOUSEEVENTF_RIGHTUP = &H10
    'Public Const MOUSEEVENTF_MOVE = &H1


    Public Const MOUSE_EVENT_F_LEFT_DOWN As Short = 2
    Public Const MOUSE_EVENT_F_LEFT_UP As Short = 4
    Public Const MOUSE_EVENT_F_RIGHT_DOWN As Short = 8
    Public Const MOUSE_EVENT_F_RIGHT_UP As Short = 16

End Module
