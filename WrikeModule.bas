Attribute VB_Name = "WrikeModule"
Option Explicit
Public WrikeAuth As WrikeAuth2
Public Const CLIENT_ID = "FQ6PSocP"
Public Const CLIENT_SECRET = "QrXjG4AMM4Ivu2QTXPBYcoyETeQF7Wesjc2CKhaZQiBKzBiqTduBT5gIs5mWi6BU"
Public Initialised As Boolean
Public FolderStructure As FolderList
Public UserList As Dictionary
Public Enum requestType
    Wrike = 1
    DefectShooter = 2
End Enum



Private Sub InitialiseHandshake()

If Initialised = False Then
    'Set WrikeAuth = New WrikeAuth2
    WrikeAuth.InitEndPoints
    WrikeAuth.InitClientCredentials CLIENT_ID, CLIENT_SECRET
    Initialised = True
    ' get all the users
    PopulateUsers
End If




End Sub
Public Sub RetrieveFoldersAndLoadTree(treeView As clsTreeView)
Dim sJson As String
Dim dicJson As Object

sJson = HttpGET("https://www.wrike.com/api/v3/folders")
'parse the received json
Set dicJson = JSON.ParseJson(sJson)

''load a tree structure of it into dictionaries
Set FolderStructure = New FolderList
FolderStructure.CreateFolderStructure dicJson

'populate the treeview
FolderStructure.PopulateTreeView treeView


'fsoWriteFile sJson, "json_test", "txt", "C:\Temp"


End Sub
Public Function SendTaskToWrike(task As WrikeTask) As Boolean
Dim urlTask As String
Dim resultDic As Dictionary
Dim httpResult As String
Dim postData As String
Dim filename As String
'generate jsonDictionary to be serialised
If task Is Nothing Then
    Exit Function
End If


    ' first check if the tasks exists in Wrike
If (FindTaskInWrike(task) = True) Then
    urlTask = "https://www.wrike.com/api/v3/tasks/" & task.Id
    postData = task.ToPostUpdateRequest
    httpResult = HttpPut(urlTask, postData)
Else
    ' generate parameters to be sent to wrike
    urlTask = "https://www.wrike.com/api/v3/folders/" & FolderStructure.FolderTag & "/tasks"
    postData = task.ToPostCreationRequest
    httpResult = HttpPOST(urlTask, postData)
End If


Set resultDic = JSON.ParseJson(httpResult)

If Not resultDic.Exists("data") Then
    SendTaskToWrike = False
     filename = "WrikeLogs" & Format(Now, "(yyyy-mm-dd hh-nn-ss)")
    fsoWriteFile httpResult & vbCrLf & task.title & vbCrLf & postData, filename, "log", "C:\Temp"
    MsgBox "this task has not been sent to Wrike" & vbCrLf & filename
Else
    SendTaskToWrike = True
    
End If
        
End Function
Public Function FindTaskInWrike(ByRef task As WrikeTask) As Boolean
Dim httpResult As String
Dim DefectId As String
Dim resultDic As Dictionary
Dim urlSearch As String
FindTaskInWrike = False

'set handshake in case of not properly instantiated
InitialiseHandshake

'get the ID from the task title
DefectId = GetIdFromTitle(task.title)

'compose the get request
urlSearch = "https://www.wrike.com/api/v3/tasks?title=" & DefectId
httpResult = HttpGET(urlSearch)

'process the results
Set resultDic = JSON.ParseJson(httpResult)

Dim result As Variant
If (resultDic.Exists("data")) Then
    For Each result In resultDic.Item("data")
       If (GetIdFromTitle(result.Item("title")) = DefectId) Then
            task.UpdateFromFind result
            FindTaskInWrike = True
            Exit Function
       End If
    Next
Else
    Exit Function
End If



'parse the received json
 'fsoWriteFile sJson, "test", "log", "C:\Temp"


End Function
Function HttpGET(url As String)

    HttpGET = HttpRequest(url, "GET", Wrike)
    
End Function
Function HttpPOST(url As String, ByVal arguments)
    HttpPOST = HttpRequest(url, "POST", Wrike, arguments)
End Function
Function HttpPut(url As String, ByVal arguments)
    HttpPut = HttpRequest(url, "PUT", Wrike, arguments)
End Function
Function HttpPOSTXto(url As String, ByVal arguments)
    HttpPOSTXto = HttpRequest(url, "POST", DefectShooter, arguments)
End Function

Function HttpRequest(url As String, sType As String, requestType As requestType, Optional ByVal arguments As String = "")

 Dim Http As MSXML2.ServerXMLHTTP60
 Set Http = New MSXML2.ServerXMLHTTP60
 'filter empty namespace
 arguments = Replace(arguments, " xmlns=""""", "")
 
 On Error Resume Next
 
  Http.Open sType, url, False
  If (requestType = Wrike) Then
    Http.setRequestHeader "Authorization", WrikeAuth.AuthHeader
    Http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  Else
     Http.setRequestHeader "Content-Type", "text/xml"
  End If
  Http.Send arguments
  
  If Err.Number <> 0 Then
    Debug.Print "[HttpGET] " & " unable to reach " & url
    Debug.Print "[HttpGET] (" & Err.Number & ") :" & Err.Description
  Err.Clear
  End If
  HttpRequest = Http.responseText
 Set Http = Nothing

 End Function

Public Sub Main()

Load frmUserPickup
frmUserPickup.Show

End Sub
Public Sub MainExcel()

Load frmUserPickupExcel
frmUserPickupExcel.Show

End Sub
Public Function GetIdFromTitle(title As String)
Dim txtPos As Integer

txtPos = InStr(title, "#")

    If (txtPos > 0) Then
        GetIdFromTitle = Trim(Left(title, txtPos))
    End If
    
End Function

Public Sub AddUser(userID As String, userName As String)
    If Not UserList.Exists(userID) Then
        UserList.Add userID, userName
    End If
End Sub

Public Function GetUserName(userID As String)
     If UserList.Exists(userID) Then
        GetUserName = UserList(userID)
     Else
        GetUserName = "no user found"
     End If
End Function

Public Sub PopulateUsers()
Dim sJson As String
Dim dicJson As Object
Dim m_jsonData As Dictionary

'set the userList in case of later loadin
Set UserList = New Dictionary
'https://www.wrike.com/api/v3/comments?limit=3&plainText=true'

sJson = HttpGET("https://www.wrike.com/api/v3/contacts")
'parse the received json

Set dicJson = JSON.ParseJson(sJson)
Dim users As Object

If TypeOf dicJson Is Dictionary Then
    Set m_jsonData = dicJson
    'first convert all the item in a class object then store them into a dictionary
    For Each users In m_jsonData.Item("data")
        'add users
         AddUser users.Item("id"), users.Item("firstName") & "_" & users.Item("lastName")
    Next
End If
Set users = Nothing
Set dicJson = Nothing

End Sub
