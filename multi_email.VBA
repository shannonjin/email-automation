Public Sub DraftOutlookEmails()
    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim lCounter As Long
    'Set objects
    Set objOutlook = Outlook.Application
    'Read details from Excel sheet and draft emails
    For lCounter = 6 To 8    'You can change the counter as per requirement
        'Create a new email item
        Set objMail = objOutlook.CreateItem(olMailItem)
        'To
        objMail.To = Sheet1.Range("A" & lCounter).Value
        'Cc
        objMail.CC = Sheet1.Range("B" & lCounter).Value
        'Subject
        objMail.Subject = Sheet1.Range("C" & lCounter).Value
        'Email Body
        objMail.Body = Sheet1.Range("D" & lCounter).Value
        'Add Attachment
        objMail.Attachments.Add (Sheet1.Range("E" & lCounter).Value)
        'Draft email
        objMail.Close (olSave)
        'Close the object
        Set objMail = Nothing
    Next
    'Show confirmation message to user
    MsgBox "Done", vbInformation
End Sub
