' Microsoft Scripting Runtime muss gebunden werden
' (Extras - Reference - Microsoft Scripting Runtime)
Dim dict As Dictionary

Sub DupeKiller()
    ' first pick a folder
    ' looking for mails with same subject and date
    Debug.Print "--- Pick Folder to check for duplicates"
    Dim objfolder As MAPIFolder
    Set objfolder = Outlook.GetNamespace("MAPI").PickFolder
  
    ' Create a dictionary instance.
    Debug.Print "--- Initializing Dictionary"
    Set dict = New Dictionary
    dict.CompareMode = BinaryCompare
    
    ProcessFolderDupKill objfolder

End Sub

Private Sub ProcessFolderDupKill(ByVal objfolder As MAPIFolder)
On Error Resume Next
    Dim olFolder As Outlook.Folder

    If (objfolder.Folders.Count > 0) Then
        For Each olFolder In objfolder.Folders
            ProcessFolderDupKill olFolder
        Next
    End If
    
    Dim items As items
    Set items = objfolder.items
    
    Dim Item As Object  ' Generic object
    Set Item = items.GetLast
    
    Dim key As String
    Do While Not Item Is Nothing
    
        If InStr(1, Item.MessageClass, ".SMIME", vbTextCompare) > 0 Then
            key = Item.Subject
        ElseIf InStr(1, Item.MessageClass, "IPM.Note", vbTextCompare) > 0 Then
            'Debug.Print "  Handling Note"
            key = Item.SentOn & vbTab & Item.SenderEmailAddress & vbTab & Item.Subject
        ElseIf InStr(1, Item.MessageClass, "IPM.Appointment", vbTextCompare) > 0 Then
            'Debug.Print "  Handling Appointment"
            key = Item.Subject & vbTab & Item.Start & vbTab & Item.End
        Else
            key = ""
        End If
        
        If key <> "" Then
            If dict.Exists(key) Then
                Debug.Print "--- Folder: " & objfolder.Name
                Debug.Print key
                Rem Debug.Print "  Duplicate Found. DELETE"
                Item.Delete
            Else
                'Debug.Print "  First occurence. Add to Dictionary"
                dict.Add key, True
            End If
        Else
            'Debug.Print "  Skip Mesageclass:" & Item.MessageClass
        End If
        
        Set Item = items.GetPrevious
    Loop
End Sub

