VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WrikeTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sParents As Collection
Private m_sTitle As String
Private m_sStatus As String
Private m_objDates As WrikeDates
Private m_sDescription As String
Private m_sId As String
Private m_sAccountId As String
Private Sub Class_Terminate()
    Set m_objDates = Nothing
    Set m_Comments = Nothing

End Sub

Public Property Get parents() As Collection
    Set parents = m_sParents

End Property
Public Property Get title() As String
    title = m_sTitle

End Property
Public Property Get status() As String
    status = m_sStatus

End Property
Public Property Get Dates() As WrikeDates
    Set Dates = m_objDates

End Property
Public Property Get Description() As String
    Description = m_sDescription

End Property
Public Sub Initialize(title As String, Description As String, parents As Collection)
m_sTitle = title
m_sStatus = "Active"
ReplaceCharsForFileName Description, " "
m_sDescription = Replace(Description, vbCrLf, "<br>")


Set m_objDates = New WrikeDates
m_objDates.Initialize 60, Now, Now

Set m_sParents = parents

End Sub


'superTasks=["IEAAAL2YKQADFMIL"]&metadata=[{"key":"testMetaKey","value":"testMetaValue"}]
'&priorityBefore=IEAAAL2YKQADFMIL&importance=High&customFields=[{"id":"IEAAAL2YJUAAABKQ","value":"testValue"}]
'&description=Test task description&dates={"start":"2015-03-31","due":"2015-04-03"}
'&title=Test task&follow=true&followers=["KUAAJP7I"]&responsibles=["KUAAJP7I"]&shareds=["KUAAJP7I"]
'&parents=["IEAAAL2YI4ADFMII"]&status=Active" https://www.wrike.com/api/v3/folders/IEAAAL2YI4ADFMII/tasks

Public Function ToPostCreationRequest() As String
Dim postData As String

    BuildPostParemeters postData, "title", m_sTitle
    BuildPostParemeters postData, "description", m_sDescription
    BuildPostParemeters postData, "dates", m_objDates.EncodeToJson
    BuildPostParemeters postData, "parents", JSON.ConvertToJson(m_sParents)

ToPostCreationRequest = postData
End Function

'"removeSuperTasks=["IEAAAL2YKQADFMIL"]&priorityAfter=IEAAAL2YKQADFMIL&importance=Low&
'customFields=[{"id":"IEAAAL2YJUAAABKQ"},{"id":"IEAAAL2YJUAAAATE","value":"testValue"}]
'&description=New description&addResponsibles=["KUAAJP7I"]&dates={"start":"2015-03-31","due":"2015-04-03"}
'&addFollowers=["KUAAJP7I"]&addParents=["IEAAAL2YI4ADFMIJ"]&title=New title&status=Deferred"

Public Function ToPostUpdateRequest() As String
Dim postData As String

    BuildPostParemeters postData, "title", m_sTitle
    BuildPostParemeters postData, "description", m_sDescription
    BuildPostParemeters postData, "addParents", JSON.ConvertToJson(m_sParents)

ToPostUpdateRequest = postData
End Function
Private Sub BuildPostParemeters(ByRef postData As String, parameters As String, Value As String)

'first node
If Len(postData) < 2 Then
    postData = parameters & "=" & Value
Else
     postData = postData & "&" & parameters & "=" & Value
End If

End Sub

Public Property Get Id() As String

    Id = m_sId

End Property

Public Property Let Id(ByVal sId As String)

    m_sId = sId

End Property

Public Property Get AccountId() As String

    AccountId = m_sAccountId

End Property

Public Property Let AccountId(ByVal sAccountId As String)

    m_sAccountId = sAccountId

End Property
Public Sub UpdateFromFind(jsonDic As Variant)
If (TypeName(jsonDic) = "Dictionary") Then
    m_sId = jsonDic.Item("id")
    m_sAccountId = jsonDic.Item("accountId")
End If
End Sub

Public Function GetComments() As Collection
Dim httpResult As String
Dim resultDic As Dictionary
Dim urlSearch As String
'https://www.wrike.com/api/v3/tasks/IEAAALNZKQAC7MZH/comments?plainText=true'
Dim oComment As WrikeComment
Dim colComments As Collection
Set colComments = New Collection

'compose the get request
urlSearch = "https://www.wrike.com/api/v3/tasks/" & m_sId & "/comments?plainText=true"
httpResult = HttpGET(urlSearch)

'process the results
Set resultDic = JSON.ParseJson(httpResult)

Dim comment As Variant
If (resultDic.Exists("data")) Then
    For Each comment In resultDic.Item("data")
        Set oComment = New WrikeComment
        oComment.Initialised comment.Item("id"), comment.Item("authorId"), comment.Item("text"), comment.Item("updatedDate")

        'add it to the collection
        colComments.Add oComment
        
    Next
Else
    Exit Function
End If

Set GetComments = colComments

End Function

