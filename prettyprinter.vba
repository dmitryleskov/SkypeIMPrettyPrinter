Dim ClipboardData As New MSForms.DataObject

Type Message
    date As Date
    time As Date
    author As String
    text As String
End Type


Private Function ParseLine(line As String) As Message
    Dim skypeRE As Object
    Set skypeRE = CreateObject("vbscript.regexp")
    skypeRE.Pattern = "^\[(\d+)/(\d+)/(\d+) (\d+):(\d+):(\d+) (AM|PM)\] (.+): (.+)"
    Set m = skypeRE.Execute(line)
    If m Is Nothing Or m.Count = 0 Then
        Debug.Print "Can't parse"
        Exit Function
    End If
    If m(0).submatches.Count < 9 Then
        Debug.Print "Can't parse"
        Exit Function
    End If
    Dim msg As Message
    With m(0)
        msg.date = DateValue _
                        (.submatches(0) + "/" _
                        + .submatches(1) + "/" _
                        + .submatches(2))
        msg.time = TimeValue _
                        (.submatches(3) + ":" _
                        + .submatches(4) + ":" _
                        + .submatches(5) + " " _
                        + .submatches(6))
        msg.author = .submatches(7)
        msg.text = .submatches(8)
    End With
    ParseLine = msg
End Function

Public Sub ProcessClipboard()
    ClipboardData.GetFromClipboard
    Dim s As String
    s = ClipboardData.GetText
    Dim ss() As String
    ss = Split(s, vbLf)
    Dim author As String
    
    Debug.Print DateAdd("h", 1, Now)
    
    For i = 0 To UBound(ss)
        Dim msg As Message
        msg = ParseLine(ss(i))
        Debug.Print (msg.text)
    Next i
End Sub

