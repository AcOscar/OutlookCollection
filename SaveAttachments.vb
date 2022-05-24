Public Sub SaveAttachments()

Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

' Get the path to your My Documents folder
strFolderpath = "c:\temp"
On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

Dim cid As String



' Set the Attachment folder.
strFolderpath = strFolderpath & "\Attachments\"
Dim objattachment As attachment

Dim htmlbody As String
    Dim pa As PropertyAccessor

' Check each selected item for attachments. If attachments exist,
' save them to the strFolderPath folder and strip them from the item.
For Each objMsg In objSelection

    ' This code only strips attachments from mail items.
    ' If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
    strDeletedFiles = ""

    If lngCount > 0 Then

        ' We need to use a count down loop for removing items
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.

        For i = lngCount To 1 Step -1

            ' Save attachment before deleting from item.
            ' Get the file name.
            
            Set objattachment = objAttachments.Item(i)

            strFile = objattachment.FileName

            ' Combine with the path to the Temp folder.
            strFile = strFolderpath & strFile

            
            Set pa = objAttachments.Item(i).PropertyAccessor
            cid = pa.GetProperty(PR_ATTACH_CONTENT_ID)
            
            If Len(cid) > 0 Then
                If InStr(objMsg.htmlbody, cid) Then
                Else
                    'In case that PR_ATTACHMENT_HIDDEN does not exists,
                    'an error will occur. We simply ignore this error and
                    'treat it as false.
                    On Error Resume Next
                    If Not pa.GetProperty(PR_ATTACHMENT_HIDDEN) Then
                        'c = c + 1
                         objAttachments.Item(i).SaveAsFile strFile

                    End If
                    On Error GoTo 0
                End If
            Else
                'c = c + 1
            objAttachments.Item(i).SaveAsFile strFile
            End If
            
            
            ' Save the attachment as a file.
            

            ' Delete the attachment.
            'objAttachments.Item(i).Delete

            'write the save as path to a string to add to the message
            'check for html and use html tags in link
            'If objMsg.BodyFormat <> olFormatHTML Then
            '    strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
            'Else
            '    strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
             '   strFile & "'>" & strFile & "</a>"
            'End If

            'Use the MsgBox command to troubleshoot. Remove it from the final code.
            'MsgBox strDeletedFiles

        Next i

        ' Adds the filename string to the message body and save it
        ' Check for HTML body
        'If objMsg.BodyFormat <> olFormatHTML Then
        '    objMsg.Body = vbCrLf & "The file(s) were saved to " & strDeletedFiles & vbCrLf & objMsg.Body
        'Else
        '    objMsg.HTMLBody = "<p>" & "The file(s) were saved to " & strDeletedFiles & "</p>" & objMsg.HTMLBody
        'End If
        'objMsg.Save
    End If
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub
