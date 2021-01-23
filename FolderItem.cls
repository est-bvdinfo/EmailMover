VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sId As String
Private m_sTitle As String
Private m_colChildIds As Collection
Private m_sScope As String
Private m_colChildFolders As Collection

Public Property Get ChildFolders() As Collection

    Set ChildFolders = m_colChildFolders

End Property

Public Property Set ChildFolders(colChildFolders As Collection)

    Set m_colChildFolders = colChildFolders

End Property


Private Sub Class_Terminate()
    Set m_colChildIds = Nothing
    Set m_colChildFolders = Nothing
End Sub
Public Property Get Id() As String

    Id = m_sId

End Property
Public Property Get title() As String

    title = m_sTitle

End Property

Public Property Get ChildIds() As Collection

    Set ChildIds = m_colChildIds

End Property


Public Property Get Scope() As String

    Scope = m_sScope

End Property

Public Property Let Scope(ByVal sScope As String)

    m_sScope = sScope

End Property

Public Sub CreateFolder(jItem As Dictionary)

m_sId = jItem.Item("id")
m_sTitle = jItem.Item("title")
m_sScope = jItem.Item("scope")

Set m_colChildIds = jItem.Item("childIds")
Set m_colChildFolders = New Collection


End Sub
