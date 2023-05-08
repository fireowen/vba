Attribute VB_Name = "SaveEmailToProjectFolder"
Sub SaveEmailToProjectFolder()
    Dim searchStr As String
    Dim projFolder As String
    Dim emailSubject As String
    Dim emailDate As String
    Dim emailFileName As String
    Dim saveFolder As String
    Dim folderPath As String
    Dim subFolderPath As String
    Dim attachment As attachment
    Dim sel As Selection
    Dim targetFolder As Folder
    Dim folderName As String
    
    
    
    Set sel = Application.ActiveExplorer.Selection
    
    ' Get search string from user
    searchStr = InputBox("Enter search string:")
    If searchStr = "" Then Exit Sub
    
    ' Find project folder based on the search string
    folderPath = "C:\Users\Owen Tilley\OneDrive - Safyre Consulting\Safyre\Projects\"
    projFolder = Dir(folderPath & Left(searchStr, 3) & "*", vbDirectory)
    If projFolder = "" Then
        MsgBox "Project folder not found."
        Exit Sub
    End If
    folderPath = folderPath & projFolder & "\"
    
    ' Find subfolder containing the first two characters after the full stop
    subFolderPath = ""
    For Each subFolder In CreateObject("Scripting.FileSystemObject").GetFolder(folderPath).subFolders
        If InStr(subFolder.Name, Mid(searchStr, 5, 2)) > 0 Then
            subFolderPath = subFolder.Path
            Exit For
        End If
    Next
    If subFolderPath = "" Then
        MsgBox "Subfolder not found."
        Exit Sub
    End If
    folderPath = subFolderPath & "\"
    
    ' Find project folder containing the full search string
    projFolder = Dir(folderPath & "*" & searchStr & "*", vbDirectory)
    If projFolder = "" Then
        MsgBox "Project folder not found."
        Exit Sub
    End If
    folderPath = folderPath & projFolder & "\"
    
    ' Select email from Inbox
    Set selectedEmail = Application.ActiveExplorer.Selection.Item(1)
    
    ' Save email as .msg in project folder
    emailSubject = Replace(Replace(Replace(Replace(selectedEmail.Subject, ":", ""), "/", ""), "\", ""), "|", "")
    emailDate = Format(selectedEmail.ReceivedTime, "yyyymmdd_hhnnss")
    emailFileName = folderPath & "Emails\" & emailDate & "_" & emailSubject & ".msg"
    saveFolder = folderPath & "Emails\"
    
    If Dir(saveFolder, vbDirectory) = "" Then MkDir saveFolder
    selectedEmail.SaveAs emailFileName, olMSG
    
    ' Save email attachments in project folder
    For Each attachment In selectedEmail.Attachments
        If Not LCase(Right(attachment.FileName, 3)) = "png" And Not LCase(Right(attachment.FileName, 3)) = "gif" Then
            attachment.SaveAsFile folderPath & "Emails\" & attachment.FileName
        End If
    Next attachment
    
    MsgBox "Email saved to project folder."
    
    ' Move email to outlook folder
    For Each subFolder In GetInbox().Folders
        If Left(subFolder.Name, Len(searchStr)) = searchStr Then
            Set targetFolder = subFolder
            Exit For
        End If
    Next
    
    If targetFolder Is Nothing Then
        MsgBox "Folder not found"
        Exit Sub
        End If
    
    For Each Item In sel
        Item.UnRead = False
        Item.Move targetFolder
    Next

    Set targetFolder = Nothing
    
End Sub

Function GetInbox() As Folder
    Set GetInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
End Function

