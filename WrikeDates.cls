VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WrikeDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sPlanificationType As String
Private m_iDurationMinutes As Integer
Private m_sStartDate As Date
Private m_sDueDate As Date
Public Property Get PlanificationType() As String
    PlanificationType = m_sPlanificationType

End Property
Public Property Get DurationMinutes() As Integer

    DurationMinutes = m_iDurationMinutes

End Property

Public Property Get StartDate() As String

    StartDate = m_sStartDate

End Property
Public Property Get DueDate() As String

    DueDate = m_sDueDate

End Property


Public Sub Initialize(durationInMinutes As Integer, StartDate As Date, DueDate As Date)
m_sStartDate = StartDate
m_sDueDate = DueDate
m_iDurationMinutes = durationInMinutes
m_sPlanificationType = "Planned"

End Sub
Public Function EncodeToJson() As String
Set m_Json = New Dictionary
    m_Json.Add "type", m_sPlanificationType
    m_Json.Add "duration", CStr(m_iDurationMinutes)
    m_Json.Add "start", Format(m_sStartDate, "yyyy-MM-dd")
    m_Json.Add "due", Format(m_sDueDate, "yyyy-MM-dd")
    
 EncodeToJson = JSON.ConvertToJson(m_Json)
Set m_Json = Nothing
End Function
