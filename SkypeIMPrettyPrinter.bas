Attribute VB_Name = "SkypeIMPrettyPrinter"
'
' SkypeIMPrettyPrinter.bas - skype IM chat log formatter for Micorsoft Outlook
'
' Author: Dmitry Leskov, www.dmitryleskov.com
'
' Copyright (c) 2012 Excelsior LLC, www.excelsior-usa.com
'
Option Explicit

''' CONFIGURATION PARAMETERS ''''''''''''''''''''''''''''''''''''''''
'''  (edit to your liking)   ''''''''''''''''''''''''''''''''''''''''

' Time gap between messages in a continuous conversaion, in minutes
' Exceeding forces timestamp insertion
Private Const MAXGAP = 20

' Colors assigned to authors
' RGB values must be prefixed with '#', color names also work
Private Const COLORTABLE = "#4573a7,#aa4644,#89a54e,#71588f,#4298af,#db843d"


''' DO NOT MODIFY ANYTHING BELOW THIS POINT '''''''''''''''''''''''''
'''   (unless you know what you are doing)  '''''''''''''''''''''''''

' Data type to represent a single message
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

' Data type to represent information about authors
Type AuthorData
    fullName As String   ' Same as Message.author
    shortName As String
    initials As String
    color As String
End Type

' Regular expression to match lines against.
' Lines containing a timestamp and an author name will match,
' indicating the start of a message
' Non-matching lines are appended to the current message
Dim skypeRE As Object
Private Sub InitSkypeRE()
    Set skypeRE = CreateObject("vbscript.regexp")
    skypeRE.Pattern = "^\[([^\]]+)\] ([^:]+): (.+)"
End Sub

Private Function ParseLine(line As String) As Message
    Dim msg As Message
    msg.append = True ' Previous message continues
    msg.text = line
    Dim m As Object
    Set m = skypeRE.Execute(line)
    If m Is Nothing Or m.count = 0 Then
        ParseLine = msg
        Exit Function
    End If
    If m(0).submatches.count < 3 Then
        ParseLine = msg
        Exit Function
    End If
    msg.append = False ' New message begins
    With m(0)
        ' Throw away the "Edited" timestamp
        msg.timestamp = CDate(Split(.submatches(0), "|")(0))
        If CDbl(msg.timestamp) < 1# Then
            ' Only time is present, add today's date
            msg.timestamp = CDate(CDbl(date) + CDbl(msg.timestamp))
        End If
        msg.author = .submatches(1)
        msg.text = .submatches(2)
    End With
    ParseLine = msg
End Function

Public Sub ProcessClipboard()
    InitSkypeRE
    ' Fetch text from clipboard
    Dim ClipboardData As New MSForms.DataObject
    ClipboardData.GetFromClipboard
    Dim text As String
    text = ClipboardData.GetText
    Dim lines() As String
    lines = Split(text, vbLf)
    Dim messages() As Message
    ReDim messages(UBound(lines))
    Dim authors() As AuthorData
    ' In most cases, there will be just two people in the chat
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
                    For a = 0 To authorCount - 1
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
    
    If authorCount > 0 Then
        ' Skype IM content detected
        text = "<p>Participants: "
        Dim sep As String
        sep = ""
        For a = 0 To authorCount - 1
            text = text & sep & authors(a).fullName
            sep = ", "
        Next a
        text = text & "</p>"
    End If
        
    i = 0
    
    ' Handle the leftover lines at the top or non-Skype content
    Do While i <= UBound(messages)
        If Not messages(i).append Then
            Exit Do
        End If
        text = text & "<p style='" _
                    & "margin-top: 0;" _
                    & "margin-bottom: 0.5em;" _
                    & "margin-left: 3em;" _
                    & "'>" & messages(i).text & "</p>" & vbLf
        i = i + 1
    Loop
    
    Dim lastTimestamp As Date
    lastTimestamp = DateSerial(1899, 1, 1)
    Dim color As String
    color = "#000000"
    Do While i <= UBound(messages)
        If messages(i).append Then
            text = text + "<p style='" _
                    + "margin-top: 0;" _
                    + "margin-bottom: 0.5em;" _
                    + "margin-left: 3em;" _
                    + "color:" + color _
                    + "'>"
        Else
            ' Delay from previous message in minutes
            Dim gap As Long
            gap = DateDiff("n", lastTimestamp, messages(i).timestamp)
            ' Has the conversation rolled over midnight?
            Dim sameDay As Boolean
            sameDay = DatePart("y", lastTimestamp) <> DatePart("y", messages(i).timestamp)

            lastTimestamp = messages(i).timestamp
            Debug.Print "timestamp="; messages(i).timestamp; " gap="; gap
            If gap > MAXGAP Then
                Dim displayTimestamp As String
                If gap > 24 * 60 Or Not sameDay Then
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
                        + "font-size: 0.83333em;" _
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
        
        text = text + messages(i).text
        text = text + "</p>" + vbLf
        i = i + 1
    Loop
    'PutHTMLClipboard (Encode_UTF8(text))
    
    Dim mail As MailItem
    Set mail = CreateItem(olMailItem)
    mail.HTMLBody = text
    mail.Display
    mail.GetInspector.Activate
    
    'ClipboardData.SetText (text)
    'ClipboardData.PutInClipboard
End Sub

