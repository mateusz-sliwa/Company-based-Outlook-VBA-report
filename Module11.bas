Attribute VB_Name = "Module11"
Sub CountSelectedItems()
    Dim olApp As Application
    Dim SelItems As Outlook.Selection
    Dim IntRes As Integer
    Dim StrMsg As String
    Dim olMail As Variant
    Dim Fldr As Folder
    Dim processed As Integer
 
    Set olApp = Outlook.Application
    Set SelItems = olApp.ActiveExplorer.Selection
    Set Fldr = GetFolderPath("sharedMailboxName\Inbox")
    
    i = 0
    j = 0
    h = 0
    For Each olMail In Fldr.Items.Restrict("@SQL=%today(""urn:schemas:httpmail:datereceived"")%")
        j = j + 1
        If olMail.UnRead Then
            i = i + 1
        ElseIf DateDiff("d", olMail.CreationTime, Now) > 2 Then
            h = h + 1
        End If
    Next olMail
    
    
    processed = j - i
    StrMsg = "Total: " & j & vbNewLine & "Processed: " & processed & vbNewLine & "Not processed: " & i
    IntRes = MsgBox(StrMsg, vbOKOnly + vbInformation, "Count Selected Outlook Items")
    Call CreateNewMail(j, processed, i, h)
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
    
    With NewMail
         .Subject = "Processed/Unprocessed KM mailbox " & Date
         .To = "recipient@exemplaryCompany.com"
         .Body = "Hi" & vbCrLf & vbCrLf & "As of " & MyDate & " the current  Mailbox messages status is: " & vbCrLf & "Total: " & total & vbCrLf & "Processed: " & processed & vbCrLf & "Unprocessed: " & unprocessed & vbCrLf & "Breached: " & breached & vbCrLf & vbCrLf & "Kind regards"
         .Display
    End With
 
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub
