Attribute VB_Name = "Module1"
Option Explicit

Public Sub saveAttachment(letter As Outlook.MailItem)

    Dim attachment As Outlook.attachment
    Dim folder As String
    folder = "D:\New folder"
    Dim appendix As String
    appendix = ""
    Dim i As Integer

    If Dir(folder, vbDirectory) = "" Then
        MsgBox "No such folder. [Alt+F11] to fix."
        Exit Sub
    End If

    For Each attachment In letter.Attachments
        For i = 1 To 100
            If Not Dir(folder & "\" & FileWithoutExtension(attachment.fileName) & appendix & FileExtension(attachment.fileName)) = "" Then
                appendix = "_" & i
            Else
                Exit For
           End If
       Next i

        attachment.SaveAsFile folder & "\" & FileWithoutExtension(attachment.fileName) & appendix & FileExtension(attachment.fileName)
        writeLog folder, attachment.fileName
        Set attachment = Nothing
    Next


    letter.Delete

End Sub

Function FileWithoutExtension(fullFileName As String) As String
 
    FileWithoutExtension = Mid(fullFileName, 1, Len(fullFileName) - 4)
 
End Function

Function FileExtension(fullFileName As String) As String
 
    FileExtension = Right(fullFileName, 4)
 
End Function


Sub writeLog(folderPath As String, textToWrite As String)

    Dim fileName As String
    fileName = "saved.log"
    Dim fso As New FileSystemObject
    Dim stream As TextStream
    
    Set stream = fso.OpenTextFile(folderPath & "\" & fileName, ForAppending, True, TristateUseDefault)
    stream.WriteLine textToWrite
    stream.Close

End Sub


