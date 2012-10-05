Option Explicit

Dim ClipboardData As New MSForms.DataObject

Type Message
    ' Extracted from clipboard
    append As Boolean
    timestamp As Date
    author As String
    text As String
    ' Computed later
    authorIndex As Integer
    firstByAuthor As Boolean
End Type

Type AuthorData
    fullName As String
    shortName As String
    initials As String
    color As String
End Type

Private Const COLORTABLE = "#4573a7,#aa4644,#89a54e,#71588f,#4298af,#db843d"

Private Function ParseLine(line As String) As Message
    Dim msg As Message
    msg.append = True ' Continue previous message
    msg.text = line
    Dim skypeRE As Object
    Set skypeRE = CreateObject("vbscript.regexp")
    skypeRE.Pattern = "^\[(\d+)/(\d+)/(\d+) (\d+):(\d+):(\d+) (AM|PM)\] (.+): (.+)"
    Dim m As Object
    Set m = skypeRE.Execute(line)
    If m Is Nothing Or m.Count = 0 Then
        ParseLine = msg
        Exit Function
    End If
    If m(0).submatches.Count < 9 Then
        ParseLine = msg
        Exit Function
    End If
    msg.append = False ' Begin new message
    With m(0)
        msg.timestamp = DateSerial( _
                        .submatches(2), _
                        .submatches(0), _
                        .submatches(1))
        msg.timestamp = DateAdd("h", .submatches(3), msg.timestamp)
        msg.timestamp = DateAdd("n", .submatches(4), msg.timestamp)
        msg.timestamp = DateAdd("s", .submatches(5), msg.timestamp)
        '.submatches(6))
        msg.author = .submatches(7)
        msg.text = .submatches(8)
    End With
    ParseLine = msg
End Function

Public Sub ProcessClipboard()
    ClipboardData.GetFromClipboard
    Dim text As String
    text = ClipboardData.GetText
    Dim lines() As String
    lines = Split(text, vbLf)
    Dim messages() As Message
    ReDim messages(UBound(lines))
    Dim authors() As AuthorData
    ReDim authors(1)
    Dim authorCount As Integer
    authorCount = 0
    Dim lastAuthor As String
    lastAuthor = ""
    Dim colors() As String
    colors = Split(COLORTABLE, ",")
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        messages(i) = ParseLine(lines(i))
        With messages(i)
            If Not .append Then
                If .author = lastAuthor Then
                    .append = True
                Else
                    lastAuthor = .author
                    .firstByAuthor = True
                    Dim a As Integer
                    For a = 0 To authorCount
                        If authors(a).fullName = .author Then
                            .firstByAuthor = False
                            .authorIndex = a
                            Exit For
                        End If
                    Next a
                    
                    If .firstByAuthor Then
                        Dim ad As AuthorData
                        ad.fullName = .author
                        ad.shortName = Split(.author)(0)
                        Debug.Print .author; authorCount; UBound(authors)
                        If authorCount = UBound(authors) + 1 Then
                            ReDim Preserve authors((UBound(authors) + 1) * 2)
                        End If
                        .authorIndex = authorCount
                        authors(authorCount) = ad
                        authorCount = authorCount + 1
                    End If
                End If
            End If
        End With
    Next i
    Dim lastTimestamp As Date
    lastTimestamp = DateSerial(1970, 1, 1)
    Dim color As String
    color = "#000000"
    text = ""
    For i = 0 To UBound(messages)
        Dim gap As Long
        gap = DateDiff("n", lastTimestamp, messages(i).timestamp)
        If gap > 30 Then
            text = text + "<p style='" _
                    + "margin-top: 0;" _
                    + "margin-bottom: 1em;" _
                    + "'>" + Format(messages(i).timestamp) + "</p>"
        End If
        If messages(i).append Then
            text = text + "<p style='" _
                    + "margin-top: 0;" _
                    + "margin-bottom: 0.5em;" _
                    + "margin-left: 3em;" _
                    + "color:" + color _
                    + "'>"
        Else
            ad = authors(messages(i).authorIndex)
            color = colors(messages(i).authorIndex)
            Dim authorName As String
            If messages(i).firstByAuthor Then
                authorName = messages(i).author
            Else
                authorName = ad.shortName
            End If
            text = text + "<p style='" _
                    + "margin-top: 0.5em;" _
                    + "margin-bottom: 0.5em;" _
                    + "margin-left: 3em;" _
                    + "text-indent: -3em;" _
                    + "color: " + color _
                    + "'>"
            text = text + "<b>" + authorName + ":</b><br>"
        End If
        text = text + messages(i).text
        text = text + "</p>" + vbLf
        lastTimestamp = messages(i).timestamp
    Next i
    text = text + "<body><html>"
    PutHTMLClipboard (Encode_UTF8(text))
    'ClipboardData.SetText (text)
    'ClipboardData.PutInClipboard
End Sub

