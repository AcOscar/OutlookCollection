Public Sub MoveMailsByTable()
'moves mails from a lected folder to a destination defin by a table


    'User to select the folder
    Debug.Print "--- Pick Folder to move from"
    Dim objSrcFolder As MAPIFolder
    Set objSrcFolder = Outlook.GetNamespace("MAPI").PickFolder

    If Not objSrcFolder Is Nothing Then
        
        ProcessMoveMails objSrcFolder
        
    End If

End Sub

Sub ProcessMoveMails(ByVal objfolder As MAPIFolder)
    Dim objMsg As Outlook.MailItem

    Dim olFolder As Outlook.Folder
    
    If (objfolder.Folders.Count > 0) Then
        For Each olFolder In objfolder.Folders
            ProcessMoveMails olFolder
        Next
    End If
    Dim objfolderitem As Object
    Debug.Print "folder: " & objfolder.FolderPath
    For Each objfolderitem In objfolder.Items
        
        If TypeOf objfolderitem Is Outlook.MailItem Then
            Set objMsg = objfolderitem
          Dim myrec As String
          myrec = GetSMTPDomainAddressForRecipients(objMsg)
            
            Debug.Print myrec
        Else
            Debug.Print TypeName(objfolderitem)
            Rem Exit For
            GoTo ResumeNext:
        End If
ResumeNext:

    Next

ExitSub:


End Sub

Sub GetSMTPAddressForRecipients(mymail As Outlook.MailItem)

    Dim recips As Outlook.Recipients
    
    Dim recip As Outlook.Recipient
    
    Dim pa As Outlook.PropertyAccessor
    
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    
    Set recips = mymail.Recipients
    
    For Each recip In recips
        
        Set pa = recip.PropertyAccessor
        
        Debug.Print recip.Name & " SMTP=" & pa.GetProperty(PR_SMTP_ADDRESS)
        
    Next
    
End Sub

Function GetSMTPDomainAddressForRecipients(mymail As Outlook.MailItem) As String

    Dim recips As Outlook.Recipients
    
    Dim recip As Outlook.Recipient
    
    Dim pa As Outlook.PropertyAccessor
    
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    
    Set recips = mymail.Recipients
    Dim domain As String
    Dim pos As Long
    For Each recip In recips
        
        Set pa = recip.PropertyAccessor
        
        domain = pa.GetProperty(PR_SMTP_ADDRESS)
        
        Debug.Print recip.Name & " SMTP=" & domain
        
        pos = InStr(1, domain, "@", VbCompareMethod.vbTextCompare)
        
        domain = Mid(domain, pos + 1)
        
        GetSMTPDomainAddressForRecipients = GetSMTPDomainAddressForRecipients & domain & " "
        
    Next
    
End Function
