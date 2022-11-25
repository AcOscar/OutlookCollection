Dim strFolderpath As String
Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
Dim fso As New Scripting.FileSystemObject

Public Sub SaveAttachments()

Dim objOL As Outlook.Application

Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String

' Get the path to your My Documents folder
strFolderpath = "O:\1212_USZ\12_CORRESPONDENCE\A\"
Rem strFolderpath = strFolderpath & "\Attachments\"

'On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

Dim cid As String
' Skript fordert Auswahl eines Ordners auf

' Dim Folder and ask User to select the folder
Debug.Print "--- Pick Folder to save attachements"
Dim objfolder As MAPIFolder
Set objfolder = Outlook.GetNamespace("MAPI").PickFolder

If Not objfolder Is Nothing Then
    
    ' Create a dictionary instance.
    Debug.Print "---"

    ProcessSaveAttachments objfolder
    
End If

End Sub

Sub ProcessSaveAttachments(ByVal objfolder As MAPIFolder)
Dim objMsg As Outlook.MailItem 'Object
'Set the Attachment folder.

Dim olFolder As Outlook.Folder

'Dim strTemp As String
Rem Dim Item As Object
If (objfolder.Folders.Count > 0) Then
    For Each olFolder In objfolder.Folders
        ProcessSaveAttachments olFolder
    Next
End If

Dim objattachment As Attachment
Debug.Print "folder: " & objfolder.FolderPath
Dim htmlbody As String
Dim pa As PropertyAccessor

' Check each selected item for attachments. If attachments exist,
' save them to the strFolderPath folder and strip them from the item.
For Each objFolderItem In objfolder.Items

If TypeOf objFolderItem Is Outlook.MailItem Then
    Set objMsg = objFolderItem
Else
    Debug.Print TypeName(objFolderItem)
    Rem Exit For
    GoTo ResumeNext:
End If

Dim Fldr As Scripting.Folder

    ' This code only strips attachments from mail items.
    ' Get the Attachments collection of the item.
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count

    If lngCount > 0 Then
    
        Debug.Print "message: " & objMsg.Subject

        Dim prjPath As String
        
        prjPath = objfolder.FolderPath
        
        prjPath = Mid(prjPath, 12)
        
        prjPath = strFolderpath & prjPath
    
        Dim attachmentfolder As String
        
        attachmentfolder = Format(objMsg.SentOn, "YYYY") & "\"

        attachmentfolder = attachmentfolder & Format(objMsg.SentOn, "YYMMDD")
    
        attachmentfolder = attachmentfolder & "_" & ClearSubject(objMsg.Subject)
    
        attachmentfolder = prjPath & "\" & attachmentfolder '& "\"
        
        For i = lngCount To 1 Step -1
            
            ' Save attachment before deleting from item.
            ' Get the file name.
            
            Set objattachment = objAttachments.Item(i)
            Debug.Print "attachment: " & objattachment.FileName

            strFile = objattachment.FileName

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
                    
                        If CheckCreateFolder(attachmentfolder) Then
                        
                            Set Fldr = fso.GetFolder(attachmentfolder)
                            strFile = attachmentfolder & "\" & strFile
                            objAttachments.Item(i).SaveAsFile strFile
                            Debug.Print "save: " & strFile
                        End If
 
                    End If
                    On Error GoTo 0
                End If
            Else
                If CheckCreateFolder(attachmentfolder) Then
                
                    Set Fldr = fso.GetFolder(attachmentfolder)
                    attachmentfolder = Fldr.ShortPath
                    objAttachments.Item(i).SaveAsFile attachmentfolder & "\" & strFile
                    
                    Debug.Print "save: " & attachmentfolder & "\" & strFile
                    
                End If
                 
            End If
            
        Next i

    End If
    
ResumeNext:

Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

End Sub

Function ClearSubject(ByVal Subject As String) As String

Const sBadChar As String = "&\/:*?<>|[]"""

ClearSubject = Subject
ClearSubject = Replace(ClearSubject, "RE:", "", vbTextCompare)
ClearSubject = Replace(ClearSubject, "AW:", "", vbTextCompare)
ClearSubject = Replace(ClearSubject, "FW:", "", vbTextCompare)
ClearSubject = Replace(ClearSubject, "WG:", "", vbTextCompare)

Dim CleanString As String
'we have now a minimum length of 1 :-)
CleanString = "A"

Dim CheckChar As String

Dim i As Long
Dim LenCS As Long

LenCS = Len(ClearSubject)

If LenCS > 70 Then
    LenCS = 70
End If

For i = 0 To LenCS - 1
    CheckChar = Mid(ClearSubject, i + 1, 1)
    'keep out a lot of not allowd chars
    If Asc(CheckChar) > 31 Then
        'no double whitespace
        If CheckChar = " " And Right(CleanString, 1) = " " Then
            'GoTo endnext
        Else
            'CleanString = CleanString & CheckChar
            'GoTo endnext
            If InStr(1, sBadChar, CheckChar, vbTextCompare) = 0 Then
                        CleanString = CleanString & CheckChar
                'GoTo endnext
            End If
        End If
        'looking for the rest of uglly chars
    End If
'endnext:
    
Next

'removing the leading A
CleanString = Mid(CleanString, 2, 100)

'removing leading whitespace
CleanString = LTrim(CleanString)


'check for points and whitespace at the end
For i = Len(CleanString) To 1 Step -1

    If Right(ClearSubject, 1) = " " Then
        CleanString = Left(CleanString, i - 1)
        i = i - 1
    
    End If
   
    If Right(CleanString, 1) = "." Then
        CleanString = Left(CleanString, i - 1)
        i = i - 1
    Else
        Exit For
    
    End If
    
Next

'CON , PRN, AUX, NUL
'COM1 , COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9
'LPT1 , LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8, LPT9

LenCS = Len(CleanString)
Dim searchstring As String

If LenCS = 4 Then

    searchstring = "COM1 COM2 COM3 COM4 COM5 COM6 COM7 COM8 COM9 LPT1 LPT2 LPT3 LPT4 LPT5 LPT6 LPT7 LPT8 LPT9"
    
    If InStr(searchstring, CleanString) Then
        CleanString = "-" & CleanString & "-"
    End If
    
ElseIf LenCS = 3 Then

    searchstring = "CON PRN AUX NUL"
    
    If InStr(searchstring, CleanString) Then
        CleanString = "-" & CleanString & "-"
    End If


End If

ClearSubject = CleanString
                            
End Function

Function CheckCreateFolder(FolderToCheckOrCreate As String) As Boolean

    Dim PathPArts As Variant
    
    PathPArts = Split(FolderToCheckOrCreate, "\")
    Dim testPath As String
    
    For j = 0 To UBound(PathPArts)
        testPath = testPath & Trim(PathPArts(j)) & "\"
    
        If Not CheckFolderExists(testPath) Then
            'MkDir (testPath)
            fso.CreateFolder (testPath)
            
            If Not CheckFolderExists(testPath) Then
                CheckCreateFolder = False
                Exit Function
            Else
                Debug.Print "folder Created: " & testPath
            End If
            
        End If
        
    Next j
    
    CheckCreateFolder = True

End Function

Function CheckFolderExists(strFolderName As String) As Boolean
 
    Dim strFolderExists As String
    
    If strFolderName = "\\" Then
    
        CheckFolderExists = True
        
        Exit Function
        
    End If
    
    If strFolderName = "\\.\" Then
    
        CheckFolderExists = True
        
        Exit Function
        
    End If
    
    If fso.FolderExists(strFolderName) Then
    
        CheckFolderExists = True
            
    Else
    
        CheckFolderExists = False
    
    End If
 
End Function
