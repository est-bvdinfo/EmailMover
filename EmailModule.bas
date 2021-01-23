Attribute VB_Name = "EmailModule"

Option Explicit
Public DefectCollection As Collection
Public Const TitleReplacement = "PROGR.PD_Tools_Phase_2 - Defect #"
Public EnvironementPath As String
Public StoragePath As String

Public Sub ReplaceCharsForFileName(sName As String, sChr As String)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, ",", sChr)
  sName = Replace(sName, "'", sChr)
  sName = Replace(sName, "&", "and")
  sName = Replace(sName, "%", "percentage")
  sName = Replace(sName, Chr(9), sChr)
  sName = Replace(sName, Chr(150), "-")
 
End Sub

Public Function GenerateMailFileName(oMail As Outlook.MailItem) As String
Dim sName As String
Dim dateFormat As String
 
   ' get the date
    dateFormat = Format(oMail.ReceivedTime, "(yyyy-mm-dd hh-nn-ss)")
    
    ' generate name of file
    sName = oMail.Subject
    If (Len(sName) > 150) Then
            sName = Left(sName, 150)
    End If
    
    sName = Replace(sName, TitleReplacement, "")
    ReplaceCharsForFileName sName, " "
    sName = Trim(sName) & dateFormat & ".msg"
    
GenerateMailFileName = sName

End Function

Public Sub testChar(strText As String)
    Dim lLoop As Long, lCount As Long
    Dim strChar As String
    lCount = Len(strText)
    ReDim strArray(lCount - 1)
     
    For lLoop = 0 To lCount - 1
        strChar = Mid(strText, lLoop + 1, 1)
        Debug.Print strChar & " " & Asc(strChar)
        If lLoop > 20 Then Exit For
    Next lLoop
End Sub

