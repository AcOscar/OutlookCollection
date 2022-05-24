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


Rem For Each Item In calendar.Items
Rem    If Mid(Item.Subject, 1, 6) = "Copy: " Then
    Rem If Mid(Item.Subject, 1, 6) = "" Then
Rem    If Item.Subject = "" Then
      Rem strTemp = Replace(Item.Subject, "Copy: ", "")
      Rem Item.Subject = strTemp
Rem      Item.Subject = " "
Rem      iItemsUpdated = iItemsUpdated + 1
Rem    End If
Rem    Item.Save
Rem Next Item

MsgBox iItemsUpdated & " of " & Itemscount & " Meetings Updated"

End Sub


Private Sub ProcessFolder(ByVal oParent As MAPIFolder)
Rem For Each subf In folder.Folders
Rem    removeinfolder (subf)
Rem Next subf
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
     Rem If Mid(Item.Subject, 1, 6) = "Copy: " Then
     If Item.Subject = "" Then
       Rem strTemp = Replace(Item.Subject, "Copy: ", "")
       Rem Item.Subject = strTemp
       Item.Subject = " "
       iItemsUpdated = iItemsUpdated + 1
       Item.Save
        
     End If
     
 Next Item

End Sub
