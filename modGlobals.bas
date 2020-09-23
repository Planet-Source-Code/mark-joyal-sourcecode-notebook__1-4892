Attribute VB_Name = "modGlobals"
Public gblNewCode As Boolean
Public gblConnectString As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function StuffQuotes(strInput As String) As String
    'used to make text strings, SQL friendly
    Dim x As Integer
    Dim strTemp As String
    
    x = InStr(1, strInput, "'")
    If x = 0 Then
        StuffQuotes = strInput
        Exit Function
    End If
    
    strTemp = strInput
    Do While x <> 0
        strTemp = Left(strTemp, x) & "'" & Mid(strTemp, x + 1)
        x = InStr(x + 2, strTemp, "'")
    Loop
    StuffQuotes = strTemp
    
End Function

Public Function AppPath() As String
    'got this from PSC, kudos to whoever posted it
    If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
End Function

