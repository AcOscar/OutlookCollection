' Microsoft Scripting Runtime muss gebunden werden
' (Extras - Verweise - Microsoft Scripting Runtime aktivieren)

    Dim dict As Dictionary

Sub DupeKiller()

' Skript fordert Auswahl eines Ordners auf
' Sucht alle mails mit gleichem Betreff und Sendedatum

    ' Dim Folder and ask User to select the folder
    Debug.Print "--- Pick Folder to check für duplicates"
    Dim objfolder As MAPIFolder
    Set objfolder = Outlook.GetNamespace("MAPI").PickFolder
  
    ' Create a dictionary instance.
    Debug.Print "--- Initializing Dictionary"
    Set dict = New Dictionary
    dict.CompareMode = BinaryCompare
  
  
  ProcessFolderDupKill objfolder

        
End Sub

Private Sub ProcessFolderDupKill(ByVal objfolder As MAPIFolder)
Rem For Each subf In folder.Folders
Rem    removeinfolder (subf)
Rem Next subf
    Dim olFolder As Outlook.Folder

    Dim strTemp As String
    Rem Dim Item As Object
    If (objfolder.Folders.Count > 0) Then
        For Each olFolder In objfolder.Folders
            ProcessFolderDupKill olFolder
        Next
    End If

    Debug.Print "--- Loading Items"
    Dim items As items
    Set items = objfolder.items
    
    Dim Item As Object  ' Generic object
    Set Item = items.GetLast
    
    Dim key As String
    Do While Not Item Is Nothing
    
        If InStr(1, Item.MessageClass, ".SMIME", vbTextCompare) > 0 Then
            key = Item.Subject
        ElseIf InStr(1, Item.MessageClass, "IPM.Note", vbTextCompare) > 0 Then
            Debug.Print "  Handling Note"
            key = Item.Subject & vbTab & Item.SentOn
        ElseIf InStr(1, Item.MessageClass, "IPM.Appointment", vbTextCompare) > 0 Then
            Debug.Print "  Handling Appointment"
            key = Item.Subject & vbTab & Item.Start & vbTab & Item.End
        Else
            key = ""
        End If
        
        If key <> "" Then
            Debug.Print "Item:" & key
            If dict.Exists(key) Then
                Debug.Print "  Duplicate Found. DELETE"
                Item.Delete
            Else
                Debug.Print "  First occurence. Add to Dictionary"
                dict.Add key, True
            End If
        Else
            Debug.Print "  Skip Mesageclass:" & Item.MessageClass
        End If
        
        
        Set Item = items.GetPrevious
    Loop
    
    
'Itemscount = oParent.items.Count + Itemscount
'
' For Each Item In oParent.items
'     Rem If Mid(Item.Subject, 1, 6) = "Copy: " Then
'     If Item.Subject = "" Then
'       Rem strTemp = Replace(Item.Subject, "Copy: ", "")
'       Rem Item.Subject = strTemp
'       Item.Subject = " "
'       iItemsUpdated = iItemsUpdated + 1
'       Item.Save
'
'     End If
'
' Next Item

End Sub
