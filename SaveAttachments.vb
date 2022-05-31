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

Rem Dim strDeletedFiles As String

' Get the path to your My Documents folder
strFolderpath = "C:\temp\Attachments\"
Rem strFolderpath = strFolderpath & "\Attachments\"

'On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Rem Set objSelection = objOL.ActiveExplorer.Selection


Dim cid As String
' Skript fordert Auswahl eines Ordners auf
' Sucht alle mails mit gleichem Betreff und Sendedatum

' Dim Folder and ask User to select the folder
Debug.Print "--- Pick Folder to check für duplicates"
Dim objfolder As MAPIFolder
Set objfolder = Outlook.GetNamespace("MAPI").PickFolder


If Not objfolder Is Nothing Then
    
    ' Create a dictionary instance.
    Debug.Print "---"
    Rem Set dict = New Dictionary
    Rem dict.CompareMode = BinaryCompare


ProcessSaveAttachments objfolder
End If

End Sub

Sub ProcessSaveAttachments(ByVal objfolder As MAPIFolder)
Dim objMsg As Outlook.MailItem 'Object
' Set the Attachment folder.

Dim olFolder As Outlook.Folder

'Dim strTemp As String
Rem Dim Item As Object
If (objfolder.Folders.Count > 0) Then
    For Each olFolder In objfolder.Folders
        ProcessSaveAttachments olFolder
    Next
End If

Dim objattachment As attachment
Debug.Print "folder: " & objfolder.FolderPath
Dim htmlbody As String
Dim pa As PropertyAccessor

' Check each selected item for attachments. If attachments exist,
' save them to the strFolderPath folder and strip them from the item.
For Each objFolderItem In objfolder.items

If TypeOf objFolderItem Is Outlook.MailItem Then
    Set objMsg = objFolderItem
Else
    Exit For
End If
Dim Fldr As Scripting.Folder


    ' This code only strips attachments from mail items.
    ' If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
    Rem strDeletedFiles = ""


    Debug.Print "message: " & objMsg.Subject

    If lngCount > 0 Then

        Dim prjPath As String
        
        prjPath = objfolder.FolderPath
        
        prjPath = Mid(prjPath, 3)
        
        prjPath = strFolderpath & prjPath
    
        Dim attachmentfolder As String
        
        attachmentfolder = Format(objMsg.SentOn, "YYYY") & "\"

        attachmentfolder = attachmentfolder & Format(objMsg.SentOn, "YYMMDD")
    
        attachmentfolder = attachmentfolder & "_" & ClearSubject(objMsg.Subject)
    
        attachmentfolder = prjPath & "\" & attachmentfolder '& "\"
        
        ' We need to use a count down loop for removing items
        
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.

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
                    Debug.Print "save: " & strFile
                End If
                 
            End If
            
        Next i

    End If
    
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub

Function ClearSubject(ByVal Subject As String) As String
ClearSubject = Subject
ClearSubject = Replace(ClearSubject, "RE:", "", vbTextCompare)
ClearSubject = Replace(ClearSubject, "AW:", "", vbTextCompare)
ClearSubject = Replace(ClearSubject, "FW:", "", vbTextCompare)
ClearSubject = Replace(ClearSubject, "WG:", "", vbTextCompare)

Const sBadChar As String = "\/:*?<>|[]"""
Dim i As Long

'Assume valid unless it isn't
  ValidFileName = True

'Loop through each "Bad Character" and test for an instance
  For i = 1 To Len(sBadChar)
    If InStr(ClearSubject, Mid$(sBadChar, i, 1)) > 0 Then
      ClearSubject = Replace(ClearSubject, Mid$(sBadChar, i, 1), " ", 1, -1, vbTextCompare)
    
    End If
  Next

ClearSubject = Replace(ClearSubject, "  ", " ", 1, -1, vbTextCompare)

ClearSubject = Left(ClearSubject, 70)

ClearSubject = Trim(ClearSubject)

For i = Len(ClearSubject) To 1 Step -1

    'If Right(ClearSubject, 2) = ".." Then Stop
    'If Right(ClearSubject, 1) = "." Then Stop
    
    If Right(ClearSubject, 1) = "." Then
        ClearSubject = Left(ClearSubject, i - 1)
        i = i - 1
    Else
        Exit For
    
    End If
    
    
    
    
    
Next
    If Right(ClearSubject, 1) = vbTab Then
        ClearSubject = Left(ClearSubject, i - 1)
       ' i = i - 1
    'Else
        'Exit For
    
    End If




ClearSubject = Trim(ClearSubject)

If ClearSubject = "." Then
    ClearSubject = "-"
End If






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
        Rem MsgBox "The selected folder exists"
        CheckFolderExists = False
    
    
    End If
    
    
    
    'strFolderExists = Dir(strFolderName, vbDirectory)
 
    'If strFolderExists = "" Then
        Rem MsgBox "The selected folder doesn't exist"
        'CheckFolderExists = False
    'Else
        Rem MsgBox "The selected folder exists"
        'CheckFolderExists = True
   
    'End If
 
End Function

