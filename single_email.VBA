Sub CreateFromTemplate()
Dim myOlApp As Outlook.Application
Dim MyItem As Outlook.MailItem
Set myOlApp = CreateObject("Outlook.Application")
'Create template
Set MyItem = myOlApp.CreateItem(olMailItem)
MyItem.Subject = "Status Report"
MyItem.To = "Dan Wilson"
MyItem.Body = "Hi, Name"
'MyItem.Display
MyItem.SaveAs "D:\statusrep.oft", OlSaveAsType.olTemplate
Dim cell As Range
For Each cell In Range("A2:A3")
Set MyItem = myOlApp.CreateItemFromTemplate("D:\statusrep.oft")
With MyItem
.To = cell.Offset(0, 1).Value
.Body = Replace(.Body, "Name", cell.Value)
.Send 'or use .Display to show the mailitem
End With
Next
End Sub
