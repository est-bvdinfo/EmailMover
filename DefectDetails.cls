VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefectDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sDefectID As String
Private m_sSeverity As String
Private m_sDetector As String
Private m_sExternalID As String
Private m_sTitle As String
Private m_sDescriptions As String
Private m_colAttachmentFileNames As Collection
Private m_dateOfReception As String
Private m_sNameOfTheEmail As String
Public Property Get title() As String

    title = m_sTitle

End Property

Public Property Get Descriptions() As String

    Description = m_sDescriptions

End Property
Public Property Get DefectId() As String

    DefectId = m_sDefectID

End Property
Public Property Get DateOfReception() As String

    DateOfReception = m_dateOfReception

End Property


Public Property Get Severity() As String

    Severity = m_sSeverity

End Property

Public Property Get Detector() As String

    Detector = m_sDetector

End Property

Public Property Get ExternalID() As String

    ExternalID = m_sExternalID

End Property

Public Sub Initialise(oMail As Outlook.MailItem)
Dim emailContent As String
Dim emailTitle As String
Dim objAtt As Outlook.Attachment
Dim sPath As String
Dim fso As FileSystemObject
Dim attachmentFileName As String


emailContent = oMail.Body
emailTitle = oMail.Subject
sPath = EnvironementPath & "\Documents\temp\"

'Add attachements storage paths
Set fso = New FileSystemObject
m_sNameOfTheEmail = GenerateMailFileName(oMail)
Set m_colAttachmentFileNames = New Collection

'save it to temp folder if not exists yet
m_colAttachmentFileNames.Add m_sNameOfTheEmail
If Not fso.FileExists(sPath & m_sNameOfTheEmail) Then
    oMail.SaveAs sPath & m_sNameOfTheEmail, olMSG
End If

'save each underlying attachements
For Each objAtt In oMail.Attachments
attachmentFileName = objAtt.DisplayName
ReplaceCharsForFileName attachmentFileName, " "

    m_colAttachmentFileNames.Add attachmentFileName
    If Not fso.FileExists(sPath & attachmentFileName) Then
        objAtt.SaveAsFile sPath & attachmentFileName
    End If
Next

  m_dateOfReception = Format(oMail.ReceivedTime, "yyyy/mm/dd hh:nn:ss")

Dim details As String
Dim aryLines() As String
Dim propertyValue As String
Dim Line As Variant
'Remove tabs and other stuff
emailContent = Replace(emailContent, Chr(9), "")
emailContent = Replace(emailContent, "&", " and ")

m_sDescriptions = fsoFindValueIntoString(emailContent, "Description", "Defect Details")
details = fsoFindValueIntoString(emailContent, "Defect Details", "")

m_sTitle = Replace(emailTitle, TitleReplacement, "")
ReplaceCharsForFileName m_sTitle, " "


'details 1
aryLines = Split(details, vbCrLf)

For Each Line In aryLines
    
    'find defect id
    propertyValue = FindProperties(Line, "defect id")
    If (Len(propertyValue) > 0) Then m_sDefectID = propertyValue


    'find severity
    propertyValue = FindProperties(Line, "Severity")
    If (Len(propertyValue) > 0) Then m_sSeverity = propertyValue
    
    'find m_sExternalID
    propertyValue = FindProperties(Line, "External Defect ID")
    If (Len(propertyValue) > 0) Then m_sExternalID = propertyValue

Next Line


    'debug line
    'fsoWriteFile emailContent, "test_body", "txt", "C:\Temp"

End Sub
Public Sub InitialiseFromExcel(ByRef currentRow)
Dim details As String


m_sTitle = currentRow.Columns("L")
ReplaceCharsForFileName m_sTitle, " "

m_sDefectID = currentRow.Columns("A")
m_sExternalID = currentRow.Columns("Q")
m_sSeverity = currentRow.Columns("P")

m_sDescriptions = m_sTitle & vbCrLf & vbCrLf & currentRow.Columns("X")

ReplaceCharsForFileName m_sDescriptions, " "


    'debug line
    'fsoWriteFile emailContent, "test_body", "txt", "C:\Temp"

End Sub
Private Function FindProperties(Line As Variant, propertyToFind) As String
Dim convertedText As String
Dim textPos As Integer
    If (TypeName(Line) = "String") Then
        convertedText = LCase(Trim(Line))
        'Debug.Print convertedText
    
        textPos = InStr(convertedText, LCase(propertyToFind))
        If (textPos = 1) Then
            FindProperties = Trim(Right(convertedText, Len(convertedText) - Len(propertyToFind)))
        End If
        
    End If
    
End Function

Public Function GenerateXmlDescription() As String
Dim xmlDoc As DOMDocument60
Dim attachementFileName
Dim cdataSection As IXMLDOMCDATASection
    
Set xmlDoc = New DOMDocument60

' Create the root element.
Dim XmlDescription As IXMLDOMElement
Set XmlDescription = xmlDoc.createElement("XmlDescription")

'set iDeskInterface
Dim iDeskService As IXMLDOMElement
Set iDeskService = XmlDescription.OwnerDocument.createElement("iDeskService")
iDeskService.setAttribute "Method", "InsertDefect"

'add all the first levels of properties
AddXmlNode iDeskService, "DefectId", m_sDefectID
AddXmlNode iDeskService, "Title", m_sTitle
AddXmlNode iDeskService, "Status", "Open"
AddXmlNode iDeskService, "Priority", m_sExternalID
AddXmlNode iDeskService, "Description", Replace(m_sDescriptions, vbCrLf, "<br>")
AddXmlNode iDeskService, "Mail", m_sNameOfTheEmail
AddXmlNode iDeskService, "MailDate", m_dateOfReception

'add attachements
Dim attachementsNode As IXMLDOMElement
Set attachementsNode = iDeskService.OwnerDocument.createElement("attachments")

Dim attachementNode As IXMLDOMElement

For Each attachementFileName In m_colAttachmentFileNames
    Set attachementNode = AddXmlNode(attachementsNode, "attachment", StoragePath & "\" & attachementFileName)
    attachementNode.setAttribute "name", attachementFileName
    attachementsNode.appendChild attachementNode
Next

'Append the list of attachement
iDeskService.appendChild attachementsNode

'Append the IdeskService in the form of a cdata
Set cdataSection = xmlDoc.createCDATASection(iDeskService.XML)
XmlDescription.appendChild cdataSection

'return the value
GenerateXmlDescription = XmlDescription.XML
End Function
