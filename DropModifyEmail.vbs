' Replaces any file dropped on the app with the text "defang_" removed in the file content, and then emails it to mobile@res.cisco.com as attachment.
' Reference https://blogs.technet.microsoft.com/heyscriptingguy/2005/02/08/how-can-i-find-and-replace-text-in-a-text-file/

Set objArgs = Wscript.Arguments
Set objFso = createobject("scripting.filesystemobject")

'iterate through all the arguments passed
For i = 0 to objArgs.count
  on error resume next

  'try and treat the argument like a folder
  Set folder = objFso.GetFolder(objArgs(i))

  'if we get an error, we know it is a file
  If err.number <> 0 then
    'this is not a folder, treat as file
    ProcessFile(objArgs(i))
    AddAttachment(objArgs(i))
  Else

  'No error? This is a folder, process accordingly
    For Each file In folder.Files
        ProcessFile(file.path)
        AddAttachment(file.path)
    Next
  End if
  On Error Goto 0
Next


Function ProcessFile(sFilePath)
 'msgbox "Now processing file: " & sFilePath
 'Do something with the file here...
    Const ForReading = 1
    Const ForWriting = 2

    strFileName = sFilePath
    strOldText = "defang_"
    strNewText = ""

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFileName, ForReading)

    strText = objFile.ReadAll
    objFile.Close

    strNewText = Replace(strText, strOldText, strNewText)

    Set objFile = objFSO.OpenTextFile(strFileName, ForWriting)

    objFile.WriteLine strNewText
    objFile.Close
End Function

Function AddAttachment(sFilePath)
    Set outlook = createobject("Outlook.Application")
    Set message = outlook.createitem(olMailItem)
    message.Recipients.Add ("mobile@res.cisco.com")
    message.Subject = "Requesting Secure link"
    message.HTMLBody = "<html><body>Please send me the link which pertains to the attachment</body></html>"    'This line is added to force BodyFormat=olFormatHTML, otherwise attachment is embedded in RTF body.
    message.Attachments.Add sFilePath
    message.Display
End Function