VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WrikeComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sId As String
Private m_sAuthorID As String
Private m_sText As String
Private m_sUpdatedDate As String
Public Sub Initialised(Id As String, AuthorID As String, text As String, sDate As String)
    m_sId = Id
    m_sAuthorID = AuthorID
    m_sText = text
    m_sUpdatedDate = sDate

End Sub
Public Property Get Id() As String
    Id = m_sId
End Property
Public Property Get AuthorID() As String
    AuthorID = m_sAuthorID
End Property
Public Property Get text() As String
    text = m_sText
End Property
Public Property Get UpdatedDate() As String
    UpdatedDate = m_sUpdatedDate
End Property
Public Property Get AuthorName() As String
    AuthorName = GetUserName(m_sAuthorID)
End Property


