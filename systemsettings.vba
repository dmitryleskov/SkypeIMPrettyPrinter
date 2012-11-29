Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Const LOCALE_SSHORTDATE = &H1F

Private Sub Command1_Click()
    Dim strResult As String
    Dim strInfo As String * 10
    lngIdentifier = GetUserDefaultLCID()
    lngResult = GetLocaleInfo(lngIdentifier, LOCALE_SSHORTDATE, strInfo, 10)
    strResult = "Short Date String = " & strInfo & vbLf
    Debug.Print strResult
    'MsgBox Replace(strResult, Chr(0), "")
End Sub

