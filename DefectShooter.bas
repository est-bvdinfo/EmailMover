Attribute VB_Name = "DefectShooter"

Public Const ClassName = "iDeskService"

Public Function SendToDefectShooter(defect As DefectDetails) As Boolean
Dim xmlDoc As DOMDocument60
Dim xmlSample As String
Dim oClientName As IXMLDOMNode
Dim oWebServiceUser As IXMLDOMNode
Dim oWebServicePassword As IXMLDOMNode
Dim oClassName As IXMLDOMNode
Dim oRequest As IXMLDOMNode
Dim oContext As IXMLDOMNode
Dim oEnveloppe As IXMLDOMElement
Dim errorParsing As String

'CAWS structure
xmlSample = "<?xml version=""1.0"" encoding=""utf-8""?> " _
 & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""><soap:Body>" _
 & "<ExecuteCustomAction xmlns=""http://fact.be.bvd/webservices/"">" _
 & "<context><ClientName>ifc</ClientName><WebServiceUser>Demo</WebServiceUser><WebServicePassword>test</WebServicePassword></context>" _
 & "<request><ClassName>iDeskService</ClassName>" & defect.GenerateXmlDescription & "</request>" _
 & "</ExecuteCustomAction></soap:Body></soap:Envelope>"

errorParsing = HttpPOSTXto("http://custo-bi-local/fact_cba_defects/Fact/WebServices/Functions/ActionService.asmx?", xmlSample)

If (InStr(errorParsing, "<Status>Ok</Status>") = 0) Then
    fsoWriteFile errorParsing, "Log_" & defect.DefectId, "log", StoragePath
    SendToDefectShooter = False
Else
    SendToDefectShooter = True
End If


 
End Function

Public Function UpdateXmlNode(ByRef oExecuteCustomAction As IXMLDOMNode, nodeName As String, valueForUpdate As String) As Boolean
Dim oNode As IXMLDOMNode

Set oNode = oExecuteCustomAction.OwnerDocument.createElement(nodeName)
oNode.text = valueForUpdate

oExecuteCustomAction.appendChild oNode

End Function

Public Function AddXmlNode(ByRef parentNode As IXMLDOMElement, nodeName As String, valueForUpdate As String) As IXMLDOMElement
Dim xmlDoc As DOMDocument60

Dim subNode As IXMLDOMElement
Set subNode = parentNode.OwnerDocument.createElement(nodeName)
subNode.text = valueForUpdate

parentNode.appendChild subNode
Set AddXmlNode = subNode

End Function
Public Sub AddSingleItem(objItem As Object, ByRef listOfmails As ListBox, ByRef DefectCollection As Collection, ByRef i As Integer)
Dim defect As DefectDetails
Dim oMail As Outlook.MailItem
Dim key As String
    
  If (TypeName(objItem) = "MailItem") Then

    'get the mail object
    Set oMail = objItem

    ' if not comming from bug report then skip
    If (oMail.To <> "CBA Bugs Reports") Then Exit Sub
    
    ' instanciate the defect based on the email
    Set defect = New DefectDetails
    defect.Initialise oMail
    
    
    key = defect.DefectId & "_" & i
    DefectCollection.Add defect, key
    
    If Not (listOfmails Is Nothing) Then
        'add to listbox
        listOfmails.AddItem
        
        listOfmails.List(i, 0) = key
        listOfmails.List(i, 1) = defect.DefectId
        listOfmails.List(i, 4) = defect.Severity
        listOfmails.List(i, 2) = defect.ExternalID
        If (Len(defect.title) > 230) Then
            listOfmails.List(i, 3) = Left(defect.title, 230)
        Else
            listOfmails.List(i, 3) = defect.title
        End If
        
        'increment next
    End If
    
    i = i + 1
  End If
  
  Set defect = Nothing
End Sub
