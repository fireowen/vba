Attribute VB_Name = "SaveSentMail"
Private Sub SaveSentMail(Item As Outlook.MailItem)
  Dim F As Outlook.MAPIFolder

  If Item.DeleteAfterSubmit = False Then
    Set F = Application.Session.PickFolder
    If Not F Is Nothing Then
      Set Item.SaveSentMessageFolder = F
    End If
  End If
End Sub
