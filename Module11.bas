    Public agedItemsCount As Long
    Public agedUnreadItemsCount As Long

Sub CountSelectedItems()
    Dim olApp As Application
    Dim SelItems As Outlook.Selection
    Dim IntRes As Integer
    Dim strMsg As String
    Dim olMail As Variant
    Dim fldr As Folder
    Dim processed As Integer
    daysAgo = 3


    Set olApp = Outlook.Application
    Set SelItems = olApp.ActiveExplorer.Selection
            Set fldr = GetFolderPath("exemplaryMailbox\Inbox")

    i = 0
    j = 0
    
    If Weekday(Now()) = vbMonday Then
    
            For Each olMail In fldr.Items.Restrict("[ReceivedTime]>'" & Format(Date - daysAgo, "DDDDD HH:NN") & "'")
            j = j + 1
            If olMail.UnRead = True Then
                i = i + 1
            End If
        Next olMail
    
    Else
        For Each olMail In fldr.Items.Restrict("@SQL=%yesterday(""urn:schemas:httpmail:datereceived"")%")
            j = j + 1
            If olMail.UnRead = True Then
                i = i + 1
            End If
        Next olMail
    End If


    processed = j - i
    Call CountItemsInFolder
    
    strMsg = "Total: " & j & vbNewLine & "Processed: " & processed & vbNewLine & "Not processed: " & i & vbNewLine & "Breached: " & agedUnreadItemsCount
    IntRes = MsgBox(strMsg, vbOKOnly + vbInformation, "Count Selected Outlook Items")
    
    Call CreateNewMail(j, processed, i, agedUnreadItemsCount)
End Sub
' Use the GetFolderPath function to find a folder in non-default mailboxes
Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer

    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function

GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function

Sub CreateNewMail(total, processed, unprocessed, breached)
    Dim obApp As Object
    Dim NewMail As MailItem
    Dim MyDate

    MyDate = Date

    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(olMailItem)
    
    If MyDate = vbMonday Then
        With NewMail
            .Subject = "Processed/Unprocessed mailbox " & (Date - 1)
            .To = "exemplary@email.com"
            .Body = "Hi" & vbCrLf & vbCrLf & "As of " & (MyDate - 3) & " the KM Mailbox messages status is: " & vbCrLf & "Total: " & total & vbCrLf & "Processed: " & processed & vbCrLf & "Unprocessed: " & unprocessed & vbCrLf & "Overall breached: " & breached & vbCrLf & vbCrLf & "Kind regards"
            .Display
        End With
    Else
    
        With NewMail
            .Subject = "Processed/Unprocessed mailbox " & (Date - 1)
            .To = "exemplary@email.com"
            .Body = "Hi" & vbCrLf & vbCrLf & "As of " & (MyDate - 1) & " the Mailbox messages status is: " & vbCrLf & "Total: " & total & vbCrLf & "Processed: " & processed & vbCrLf & "Unprocessed: " & unprocessed & vbCrLf & "Overall breached: " & breached & vbCrLf & vbCrLf & "Kind regards"
            .Display
        End With
        
    End If
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub

Sub CountItemsInFolder()

    Dim strMsg As String

    Dim allItems As Items
    Dim unreadItems As Items

    Dim agedItems As Items
    Dim agedUnreadItems As Items

    Dim fldr As Folder

    Dim processed As Long

    Dim allItemsCount As Long
    Dim unreadItemsCount As Long


    Dim strFilterUnread As String
    Dim strFilterAged As String

    Set fldr = GetFolderPath("exemplarySharedMailbox\Inbox")
    'Set fldr = Session.GetDefaultFolder(olFolderInbox)
    'Debug.Print vbCr & "** folder: " & fldr

    Set allItems = fldr.Items
    allItemsCount = allItems.Count
    'Debug.Print "items in folder: " & allItemsCount

    ' ** filter for unread items
    strFilterUnread = "[unread]=true"
    'Debug.Print strFilterUnread

    Set unreadItems = allItems.Restrict(strFilterUnread)
    unreadItemsCount = unreadItems.Count
    'Debug.Print "unread items in " & fldr & ": " & unreadItemsCount & vbCr

    ' ** filter for aged items
    strFilterAged = "[ReceivedTime]<'" & Format(Date - 2, "DDDDD HH:NN") & "'"
    'Debug.Print strFilterAged

    Set agedItems = allItems.Restrict(strFilterAged)
    agedItemsCount = agedItems.Count
    'Debug.Print "aged items in " & fldr & ": " & agedItemsCount

    Set agedUnreadItems = agedItems.Restrict(strFilterUnread)
    agedUnreadItemsCount = agedUnreadItems.Count
    'Debug.Print "aged unread items in " & fldr & ": " & agedUnreadItemsCount & vbCr

    processed = allItemsCount - unreadItemsCount

    strMsg = "Breached: " & agedUnreadItemsCount

    'Debug.Print strMsg & vbCr
    'MsgBox strMsg, vbOKOnly + vbInformation, "Count Selected Outlook Items"
    'Call CreateNewMail(allItemsCount, processed, unreadItemsCount, agedUnreadItemsCount)

End Sub

