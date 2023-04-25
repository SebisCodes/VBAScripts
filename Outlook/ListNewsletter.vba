'Required References: mscorlib.dll, Microsoft VBScript Regular Expressions 1.0, Microsoft VBScript Regular Expressions 5.5

' A VBA-script to list all newsletter unsubscribe urls out of the inbox
'
' Contact: sebiscodes@gmail.com
' My Github: https://github.com/SebisCodes

Sub ListNewsletter()
    
    On Error Resume Next
    
    Dim filepath As String
    Dim keywords As ArrayList: Set keywords = New ArrayList
    
    'Add keywords in your language to find the "unsubscribe" urls in the newsletter
    keywords.Add "UNSUBSCRIBE"
    keywords.Add "ABBESTELLEN"
    keywords.Add "ABMELDEN"
    
    'Filepath of the output-file (Default in Documents-path)
    filepath = Environ$("USERPROFILE") & "\Documents\_Subscribed_URLs.html" 'Save in documents folder
    
    'Init variables
    Dim regEx As New RegExp
    Dim matches As MatchCollection
    Dim match As match

    Dim objNS As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim oAccount As Account
    Dim i As Integer
    Dim item As Outlook.MailItem

    'Get Namespace
    Set objNS = GetNamespace("MAPI")

    'Setup Regular Expressions
    regEx.Pattern = "<a[^>]*>(.*?)</a>"
    regEx.Global = True

    'Open file
    Open filepath For Output As #1
    
    'Iterate through accounts
    For Each oAccount In objNS.Accounts
        'Print accountname as title
        Print #1, "<hr><h1>" & oAccount & "</h1>"
        'Get Folder
        Set objFolder = oAccount.DeliveryStore.GetDefaultFolder(olFolderInbox)
        'Iterate through folders
        For i = 0 To objFolder.Items.Count
            Set item = objFolder.Items(i)
            If TypeName(item) = "MailItem" Then
                'Search for a-Tags
                Set matches = regEx.Execute(item.HTMLBody)
    
                'Go to first match and write it to the file
                For Each match In matches
                    If keywordsInString(match.Value, keywords) Then
                        Print #1, "<p>" & match.Value & " from <b>" & item.SenderEmailAddress & "</b> in <u>" & item.ConversationTopic & "</u></p>"
                        Exit For
                    End If
                Next match
            End If
            DoEvents
        Next i
    Next oAccount
    'Close file
    Close #1
End Sub

'Check if a list of keywords is in a string
Function keywordsInString(ByVal checkString As String, ByRef keywords As ArrayList) As Boolean
    Dim v As Variant
    Dim ret As Boolean
    ret = False
    For Each v In keywords
        If InStr(UCase(checkString), UCase(v)) > 0 Then
            ret = True
            Exit For
        End If
    Next
    keywordsInString = ret
End Function
