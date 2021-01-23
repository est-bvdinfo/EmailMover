Attribute VB_Name = "Generic"
Option Explicit
Public Sub fsoWriteFile(ByVal Content, ByVal filename, ByVal sExtention, ByVal sfolder)
Dim UseOfFolder
Dim oStreamRecoder
Dim fso
UseOfFolder = sfolder
Debug.Print "[fsoWriteFile]" & filename & "." & sExtention & "  in folder '" & UseOfFolder & "' about to be created"
Set fso = CreateObject("Scripting.FileSystemObject")
Set oStreamRecoder = fso.CreateTextFile(UseOfFolder & "\" & filename & "." & sExtention, True, False)
Debug.Print "[fsoWriteFile] type of content object" & TypeName(Content)
    If TypeName(Content) <> "Null" Then oStreamRecoder.Write (Content)
Set oStreamRecoder = Nothing
Debug.Print filename & "." & sExtention & " created in folder '" & UseOfFolder & "'"
Set fso = Nothing
 

End Sub
Public Function fsoFindValueIntoString(ByVal text, ByVal startText, ByVal stopText)
Dim startPos: startPos = InStr(text, startText)
Dim endPos
If startPos > 0 Then
startPos = startPos + Len(startText)
        If (stopText <> "") Then
            endPos = InStr(startPos, text, stopText)
            'Debug.Print "startPos: " & startPos & " endPos :" & endPos
            If endPos > 0 Then
                fsoFindValueIntoString = Trim(Mid(text, startPos, endPos - startPos))
            End If
        Else
            fsoFindValueIntoString = Right(text, Len(text) - startPos)
        End If
        
End If

End Function
Public Sub CopyAllFiles(originalFolderPath As String, targetFolderPath As String)
Dim fso As FileSystemObject
Dim originalFolder As Variant
Dim ofile As Variant

Set fso = New FileSystemObject
Set originalFolder = fso.GetFolder(originalFolderPath)

If Not fso.FileExists(targetFolderPath) Then
    originalFolder.Copy (targetFolderPath)
End If

For Each ofile In originalFolder.Files
    DeleteFile ofile.path
Next


End Sub

Sub DeleteFile(ByVal FileToDelete As String)
Dim oFso As FileSystemObject
Set oFso = New FileSystemObject
   If oFso.FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub


