Option Explicit

Dim iItemsUpdated As Integer
Dim Itemscount As Integer

Sub RemoveCopy()
Dim myolApp As Outlook.Application
Dim calendar As MAPIFolder
Rem Dim aItem As Object

Set myolApp = CreateObject("Outlook.Application")
Set calendar = myolApp.ActiveExplorer.CurrentFolder

Rem Dim iItemsUpdated As Integer
Rem Dim strTemp As String
Rem On Error Resume Next
iItemsUpdated = 0
Itemscount = 0
ProcessFolder calendar

MsgBox iItemsUpdated & " of " & Itemscount & " Meetings Updated"

End Sub


Private Sub ProcessFolder(ByVal oParent As MAPIFolder)

Dim olFolder As Outlook.Folder

Dim strTemp As String
Dim Item As Object
If (oParent.Folders.Count > 0) Then
     For Each olFolder In oParent.Folders
         ProcessFolder olFolder
     Next
End If

Itemscount = oParent.items.Count + Itemscount

For Each Item In oParent.items
    
     If Mid(Item.Subject, 1, 6) = "Copy: " Then
       strTemp = Replace(Item.Subject, "Copy: ", "")
       Item.Subject = strTemp  
       iItemsUpdated = iItemsUpdated + 1
       Item.Save
     End If
    
     If Item.Subject = "" Then
       Item.Subject = " "
       iItemsUpdated = iItemsUpdated + 1
       Item.Save
     End If
 
Next Item

End Sub
