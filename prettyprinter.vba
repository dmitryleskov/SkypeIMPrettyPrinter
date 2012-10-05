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
    skypeRE.Pattern = "^\[(.+)\] (.+): (.+)"
    Dim m As Object
    Set m = skypeRE.Execute(line)
    If m Is Nothing Or m.Count = 0 Then
        ParseLine = msg
        Exit Function
    End If
    If m(0).submatches.Count < 3 Then
        ParseLine = msg
        Exit Function
    End If
    msg.append = False ' Begin new message
    With m(0)
        msg.timestamp = CDate(Split(.submatches(0), "|")(0))
        msg.author = .submatches(1)
        msg.text = .submatches(2)
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
    
    Dim URLRE As Object
    Set URLRE = CreateObject("vbscript.regexp")
    URLRE.Pattern = "^(.*)(https?:\/\/?[\da-z\.-]+\.[a-z\.]{2,6}([\/\w\.-]*)*\/?)(.*)$"
    
    Dim lastTimestamp As Date
    lastTimestamp = DateSerial(1970, 1, 1)
    Dim color As String
    color = "#000000"
    text = "<p>Chat</p>"
    For i = 0 To UBound(messages)
        If messages(i).append Then
            text = text + "<p style='" _
                    + "margin-top: 0;" _
                    + "margin-bottom: 0.5em;" _
                    + "margin-left: 3em;" _
                    + "color:" + color _
                    + "'>"
        Else
            Dim gap As Long
            gap = DateDiff("n", lastTimestamp, messages(i).timestamp)
            Debug.Print gap
            If gap > 30 Then
                Dim displayTimestamp As String
                If gap > 60 * 24 Or _
                    DatePart("y", lastTimestamp) <> DatePart("y", messages(i).timestamp) Then
                    ' Full date and time
                    displayTimestamp = Format(messages(i).timestamp)
                Else
                    ' Show time only
                    displayTimestamp = Format(messages(i).timestamp, "Medium Time")
                End If
                text = text + "<p style='" _
                        + "margin-top: 1em;" _
                        + "margin-bottom: 1em;" _
                        + "text-align: center;" _
                        + "font-size: 0.875em;" _
                        + "background-color: #eee;" _
                        + "'>" + displayTimestamp + "</p>"
            End If
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
        
        Dim m As Object
        Set m = URLRE.Execute(messages(i).text)
        If m.Count = 0 Then
            Debug.Print "No Match"
        Else
            Debug.Print m(0).submatches(0)
            Debug.Print m(0).submatches(1)
            Debug.Print m(0).submatches(2)
            Stop
        End If
        
        text = text + messages(i).text
        text = text + "</p>" + vbLf
        lastTimestamp = messages(i).timestamp
    Next i
    'PutHTMLClipboard (Encode_UTF8(text))
    
    Dim mail As MailItem
    Set mail = CreateItem(olMailItem)
    mail.HTMLBody = text
    mail.Display
    mail.GetInspector.Activate
    
    'ClipboardData.SetText (text)
    'ClipboardData.PutInClipboard
End Sub

