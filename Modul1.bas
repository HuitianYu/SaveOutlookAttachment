Attribute VB_Name = "Modul1"
' ---------- << New Features >> ----------
' Download and link files, delete the orignal attachments, add the html to the attachment are decoupled:
'   if one of the steps failes, just run it again.
' Failed replacement will only show a msgbox to avoid interruption of program.
' Signed emails will not be processed and a msgbox will show this info as the program terminates.
'   If you want to proceed signed emails, please select only one single signed email and run it.
' Multiple types of objects with attachments supported: mail, meeting, etc...
'   see the variable 'childrenNames'
' Better structure for different types of objects with attachments:
'   files of different types are stored in different subfolders under the parent folder.
' The path convention is only suitable for windows


Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' ---------- << Hints >> ----------
' PLEASE INITIALIZE YOUR PARENT PATH HERE BY YOUR FAVOURITE PATH.
Public Const parentPath_ As String = "H:\attachments\"


Function sp() As String
    sp = "\"
End Function

Function mySave(ByRef objItem_ As Object) As Boolean
    Dim sleepTime As Integer
    Dim saveSuccessful As Boolean
    Dim maxAttemp As Integer
    Dim currentAttemp As Integer
    
    sleepTime = 10 ' in milliseconds
    saveSuccessful = False
    maxAttemp = 500
    currentAttemp = 0
    Do While currentAttemp < maxAttemp
        On Error Resume Next
        ' CORE starting
        objItem_.Save
        ' CORE ended
        If Err.Number = 0 Then
            saveSuccessful = True
            Exit Do
        Else
            currentAttemp = currentAttemp + 1
            Sleep sleepTime
            If currentAttemp >= maxAttemp Then
                Exit Do
            End If
        End If
    Loop
    mySave = saveSuccessful
End Function

Function createFolder(ByRef objFSO_ As Object, ByVal path As String)
    If Not objFSO_.FolderExists(path) Then
        ' If the folder doesn't exist, create it
        objFSO_.createFolder path
    End If
End Function

Function childIndex(valToBeFound As Variant, arr As Variant)
    Dim element As Variant
    Dim iter As Integer
    iter = 0
    childIndex = -1  ' Assume the value is not in the array
    For Each element In arr
        If element = valToBeFound Then
            childIndex = iter
            Exit Function
        Else
            iter = iter + 1
        End If
    Next element
End Function

Function HashEntryID(entryID As String) As String
    Dim byteData() As Byte
    Dim objCrypto As Object
    Dim objStream As Object
    Dim hashedData() As Byte
    Dim i As Integer
    Dim hashValue As String

    ' Convert string to byte array
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8"
    objStream.WriteText entryID
    objStream.Position = 0
    objStream.Type = 1 ' adTypeBinary
    byteData = objStream.Read
    objStream.Close

    ' Hash the byte array
    Set objCrypto = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    hashedData = objCrypto.ComputeHash_2(byteData)

    ' Convert hashed byte array to hex string
    For i = 0 To UBound(hashedData)
        hashValue = hashValue & Right("0" & Hex(hashedData(i)), 2)
    Next i

    HashEntryID = hashValue
End Function

Function downLoadAndLink(ByRef objItem_ As Object, ByVal originalSubPath_ As String, ByRef objHTMLFile_ As Object)
    Dim fileCount As Integer
    Dim originalFileName As String
    Dim originalFilePath As String
    Dim fileLink As String
    
    fileCount = 0
    
    objHTMLFile_.WriteLine "<html><body>"
    While fileCount < objItem_.Attachments.count
        fileCount = fileCount + 1
        ' Save attachment with a unique name
        originalFileName = objItem_.Attachments(fileCount).FileName
        originalFilePath = originalSubPath_ & originalFileName
        objItem_.Attachments(fileCount).SaveAsFile originalFilePath
        ' Create a link in the temporary HTML file
        fileLink = "<a href='file:///" & originalFilePath & "'>" & originalFileName & "</a><br>"
        objHTMLFile_.WriteLine fileLink
    Wend

    ' Close the temporary HTML file
     objHTMLFile_.WriteLine "</body></html>"
     objHTMLFile_.Close
End Function


Function IsMailDigitallySigned(mail As MailItem) As Boolean
    Const PR_MESSAGE_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
    Dim securityFlags As Long

    On Error Resume Next
    securityFlags = mail.PropertyAccessor.GetProperty(PR_MESSAGE_SECURITY_FLAGS)
    On Error GoTo 0

    ' Check if the mail is digitally signed
    If (securityFlags And &H2) = &H2 Then
        IsMailDigitallySigned = True
    Else
        IsMailDigitallySigned = False
    End If
End Function


Function ReplaceAttachment(ByRef objFSO_ As Object, ByRef objNamespace_ As NameSpace, ByRef objItem_ As Object, ByVal parentPath_ As String, ByVal childrenNames_ As Variant) As Boolean
    Dim originalPath As String
    Dim originalSubPath As String
    Dim htmlPath As String
    Dim htmlFileName As String
    Dim objHTMLFile As Object
    Dim replaceSuccessful As Boolean
    
    replaceSuccessful = True
    
    ' Create folders to save files
    ID = objNamespace_.CurrentUser.Name & "_" & HashEntryID(objItem_.entryID)
    originalPath = parentPath_ & childrenNames_(childI) & sp() & "original" & sp()
    originalSubPath = originalPath & ID & sp()
    createFolder objFSO_, originalSubPath
    htmlPath = parentPath_ & childrenNames_(childI) & sp() & "html" & sp()
    htmlFileName = ID & "_link.html"

    ' Step 1 to 3 are decoupled: if the sucessing steps fail, they can be repeated based on the preceeding steps.
    ' Step 1: Loop through attachments in the object, wenn the files are not created
    If Not objFSO_.FileExists(htmlPath & htmlFileName) Then
        Set objHTMLFile = objFSO_.CreateTextFile(htmlPath & htmlFileName)
        downLoadAndLink objItem_, originalSubPath, objHTMLFile
    End If
    
    ' Step 2: Delete all the files in the attachment
    While objItem_.Attachments.count > 0
        objItem_.Attachments(1).Delete
    Wend
    If Not mySave(objItem_) Then
        replaceSuccessful = False
    End If
    
    ' Step 3: Add html file in the empty attachment
    If objItem_.Attachments.count = 0 Then
        objItem_.Attachments.Add htmlPath & htmlFileName
    End If
    If Not mySave(objItem_) Then
        saveSuccessful = False
    End If
    ReplaceAttachment = replaceSuccessful
End Function

Function CreatePaths(ByRef objFSO_ As Object, ByVal parentPath_ As String, ByRef childrenTypes_ As Variant, ByRef childrenNames_ As Variant)
    Dim childI As Integer
    ' Create folders if not exist
    createFolder objFSO_, parentPath_
    For childI = 0 To UBound(childrenNames_)
        createFolder objFSO_, parentPath_ & childrenNames_(childI) & sp()
        createFolder objFSO_, parentPath_ & childrenNames_(childI) & sp() & "original" & sp()
        createFolder objFSO_, parentPath_ & childrenNames_(childI) & sp() & "html" & sp()
    Next childI
    childI = -1
End Function



Sub ProcessAttachmentsAndCreateLinks()
    Dim objNamespace As NameSpace
    Dim objSelection As Selection
    Dim objFSO As Object
    Dim objItem As Object
    Dim childrenTypes() As Variant
    Dim childrenNames() As Variant
    Dim childI As Integer
    ' Dim parentPath As String
    Dim hasSignedMail As Boolean
    Dim replaceAllSuccessful As Boolean
    
    replaceAllSuccessful = True
    Set objNamespace = Application.GetNamespace("MAPI")
    
    ' TODO: select items also in extended search
    ' Get the selected items in Outlook
    Set objSelection = Application.ActiveExplorer.Selection
    ' Create File System Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    
    ' Check the correctness of the parent path
    If Right(parentPath_, 1) = sp() Then
        parentPath = parentPath_
    Else
        parentPath = parentPath_ & sp()
    End If
    
    ' Define types with attachments
    childrenTypes = Array(olMail, olAppointment, olContact, olTask, olJournal, olNote, olPost, olMeeting, olMeetingCanceled, olMeetingReceived, olMeetingReceivedAndCanceled, olNonMeeting, olReport, olDistributionList)
    childrenNames = Array("mail", "appointment", "contact", "task", "journal", "note", "post", "meeting", "meeting", "meeting", "meeting", "meeting", "report", "distributionlist")
    ' Create paths
    CreatePaths objFSO, parentPath, childrenTypes, childrenNames
    
    ' Loop through selected items
    For Each objItem In objSelection
        childI = childIndex(objItem.Class, childrenTypes)
        If childI > -1 Then
            ' If there are multiple objects selected and the current one is a mail object: If it is signed --> skip
            If objSelection.count > 1 And childI = 0 And IsMailDigitallySigned(objItem) Then
                ' Skip
                hasSignedMail = True
            Else
                If Not ReplaceAttachment(objFSO, objNamespace, objItem, parentPath, childrenNames) Then
                    replaceAllSuccessful = False
                End If
            End If
        End If
    Next objItem
    
    Dim iBeep As Integer
    For iBeep = 1 To 3 ' Loop 3 times.
        Beep ' Sound a tone.
    Next iBeep
    
    If hasSignedMail Then
        MsgBox "There are emails with signature. Please select only one signed mail of them at once and run the program again."
    End If
    If Not replaceAllSuccessful Then
        MsgBox "Done! There are some object (in additional to signed emails) not successfully replaced! You only have to select them and run the program again!"
    Else
        MsgBox "Done!"
    End If
End Sub
